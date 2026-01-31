"""
XML Chord Analyzer - Advanced Music Theory Analysis Tool

A comprehensive application for analyzing chord progressions and harmonic structures
from MusicXML files. Features block chord detection, arpeggio analysis, anacrusis
handling, and advanced merging algorithms with visual timeline display.
For info: see Desire in Chromatic Harmony by Kenneth Smith (Oxford, 2020).
"""

import os
import platform
import sys
import threading
from typing import Callable, Dict, List, Optional, Tuple, Any, Set

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Text, BooleanVar, Frame, Label
import tkinter.font as tkfont

from PIL import Image, ImageDraw, ImageTk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.colors import black, HexColor
from music21 import converter, note, chord as m21chord, meter, stream

# Import MIDI library at top level for PyInstaller compatibility
try:
    print("Attempting to import mido...")
    import mido
    print("mido imported successfully")
    
    print("Attempting to import rtmidi...")
    import rtmidi  # Explicitly import the backend for PyInstaller
    print("rtmidi imported successfully")
    
    # Import all mido backends explicitly for PyInstaller
    try:
        print("Attempting to import mido.backends.rtmidi...")
        import mido.backends.rtmidi
        print("mido.backends.rtmidi imported successfully")
        
        print("Attempting to import mido.backends._rtmidi...")
        import mido.backends._rtmidi
        print("mido.backends._rtmidi imported successfully")
    except ImportError as e:
        print(f"Backend import failed: {e}")
        pass  # Some backends may not be available
    
    MIDO_AVAILABLE = True
    print("MIDO_AVAILABLE set to True")
except ImportError as e:
    print(f"MIDO import failed: {e}")
    import traceback
    print(f"Full traceback: {traceback.format_exc()}")
    mido = None
    MIDO_AVAILABLE = False

# Import pygame for MIDI output at top level for PyInstaller compatibility
try:
    import pygame
    import pygame.midi
    PYGAME_AVAILABLE = True
except ImportError:
    pygame = None
    PYGAME_AVAILABLE = False


def resource_path(relative_path: str) -> str:
    """Return absolute path to resource for both development and PyInstaller builds."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)

def debug_log(message: str) -> None:
    """Write debug messages to both console and log file for packaged executable debugging."""
    print(message)  # Console output
    try:
        # Also write to a log file in the same directory as the executable
        log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "midi_debug.log")
        with open(log_file, "a", encoding="utf-8") as f:
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        print(f"Failed to write to log file: {e}")

# Display symbols for chord analysis visualization
CLEAN_STACK_SYMBOL = "✅"  # Indicates clean stacked chord voicing
ROOT2_SYMBOL = "²"         # Second inversion marker
ROOT3_SYMBOL = "³"         # Third inversion marker

def beautify_chord(chord: str) -> str:
    """Convert flat (b) and sharp (#) symbols to proper musical notation."""
    return chord.replace("b", "♭").replace("#", "♯")

# Music theory constants and chord definitions
NOTE_TO_SEMITONE = {
    'C': 0, 'C#': 1, 'Db': 1, 'D': 2, 'D#': 3, 'Eb': 3, 'E': 4,
    'F': 5, 'F#': 6, 'Gb': 6, 'G': 7, 'G#': 8, 'Ab': 8, 'A': 9,
    'A#': 10, 'Bb': 10, 'B': 11
}

# Chord patterns defined as semitone intervals from root
CHORDS = {
    "C7": [0, 4, 7, 10], "C7b5": [0, 4, 6, 10], "C7#5": [0, 4, 8, 10],
    "Cm7": [0, 3, 7, 10], "Cø7": [0, 3, 6, 10], "C7m9noroot": [1, 4, 7, 10],
    "C7no3": [0, 7, 10], "C7no5": [0, 4, 10], "C7noroot": [4, 7, 10],
    "Caug": [0, 4, 8], "C": [0, 4, 7], "Cm": [0, 3, 7],
    "Cmaj7": [0, 4, 7, 11], "CmMaj7": [0, 3, 7, 11]
}

TRIADS = {"C", "Cm", "Caug"}  # Basic three-note chords
CIRCLE_OF_FIFTHS_ROOTS = ['F#', 'B', 'E', 'A', 'D', 'G', 'C', 'F', 'Bb', 'Eb', 'Ab', 'Db', 'Gb']

# Enharmonic equivalents for note normalization
ENHARMONIC_EQUIVALENTS = {
    # Common enharmonic pairs
    'A#': 'Bb', 'Bb': 'Bb',
    'C#': 'Db', 'Db': 'Db',
    'D#': 'Eb', 'Eb': 'Eb',
    'F#': 'F#', 'Gb': 'F#',
    'G#': 'Ab', 'Ab': 'Ab',
    'E#': 'F',  'Fb': 'E',
    'B#': 'C',  'Cb': 'B',
    # Double accidentals
    'A##': 'B', 'B##': 'C#', 'C##': 'D', 'D##': 'E', 'E##': 'F#', 'F##': 'G', 'G##': 'A',
    'Abb': 'G', 'Bbb': 'A', 'Cbb': 'Bb', 'Dbb': 'C', 'Ebb': 'D', 'Fbb': 'Eb', 'Gbb': 'F',
    # Triple accidentals (rare edge cases)
    'A###': 'B#', 'B###': 'C##', 'C###': 'D#', 'D###': 'E#', 'E###': 'F##', 'F###': 'G#', 'G###': 'A#',
    'Abbb': 'Gb', 'Bbbb': 'G', 'Cbbb': 'A', 'Dbbb': 'Bb', 'Ebbb': 'C', 'Fbbb': 'Db', 'Gbbb': 'Eb',
}

# Event merging algorithm parameters (default values for position 3 of 7-position slider)
MERGE_JACCARD_THRESHOLD = 0.60  # Chord similarity threshold (0.0-1.0, higher = stricter)
MERGE_BASS_OVERLAP = 0.50       # Required bass note overlap for merging (0.0-1.0)
MERGE_BAR_DISTANCE = 1          # Maximum bars apart for events to merge (0 = same bar only)
MERGE_DIFF_MAX = 1              # Maximum root differences allowed for simple merge path

class LoadOptionsDialog(tk.Toplevel):
    """Dialog for selecting MusicXML files and analysis options."""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.result = None
        self.include_triads_var = BooleanVar(value=True)
        self.sensitivity_var = tk.StringVar(value="Medium")
        self.selected_file = None
        self.build_ui()
    def build_ui(self):
        frame = tk.Frame(self, bg="black")
        frame.pack(padx=10, pady=10, fill="x")

        ttk.Checkbutton(
            frame, text="Include triads", variable=self.include_triads_var,
            style="White.TCheckbutton"
        ).pack(anchor="w", pady=5)

        ttk.Label(
            frame, text="Sensitivity level:", background="black", foreground="white"
        ).pack(anchor="w", pady=(10, 0))
        for level in ["High", "Medium", "Low"]:
            ttk.Radiobutton(
                frame, text=level, variable=self.sensitivity_var, value=level,
                style="White.TRadiobutton"
            ).pack(anchor="w")

        ttk.Button(frame, text="Select MusicXML File", command=self.select_file).pack(pady=10)

    def select_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("MusicXML files", "*.xml *.musicxml *.mxl"),
            ],
            title="Select a MusicXML file"
        )
        if path:
            self.selected_file = path
            self.result = {
                "file": path,
                "include_triads": self.include_triads_var.get(),
                "sensitivity": self.sensitivity_var.get()
            }
            self.destroy()

class MidiChordAnalyzer(tk.Tk):
    """
    Main application class for MIDI chord analysis.
    
    Analyzes MusicXML files to detect chord progressions using multiple algorithms:
    - Block chord detection from simultaneous notes
    - Arpeggio pattern recognition from melodic sequences  
    - Anacrusis handling for melodic resolution notes
    - Advanced event merging with configurable sensitivity
    """
    
    def debug_print_notes(self):
        """Print all notes and chords with bar, beat, and duration for debugging."""
        if not self.score:
            print("No score loaded.")
            return
        flat_notes = list(self.score.flatten().getElementsByClass([note.Note, m21chord.Chord]))
        # Use local offset_to_bar_beat from analyze_musicxml
        def offset_to_bar_beat(offset):
            # Find the last time signature whose offset is <= given offset
            time_signatures = []
            for ts in self.score.flatten().getElementsByClass(meter.TimeSignature):
                offset_ts = float(ts.offset)
                time_signatures.append((offset_ts, int(ts.numerator), int(ts.denominator)))
            time_signatures.sort(key=lambda x: x[0])
            if not time_signatures:
                num, denom = 4, 4
                return 1, int(offset) + 1, f"{num}/{denom}"
            elif time_signatures[0][0] > 0.0:
                first_num, first_den = time_signatures[0][1], time_signatures[0][2]
                time_signatures.insert(0, (0.0, first_num, first_den))
            bars_before = 0
            for i, (t_off, num, denom) in enumerate(time_signatures):
                next_off = time_signatures[i + 1][0] if i + 1 < len(time_signatures) else None
                beat_len = 4.0 / denom
                if next_off is None or offset < next_off:
                    beats_since_t = (offset - t_off) / beat_len
                    if beats_since_t < 0:
                        beats_since_t = (offset) / beat_len
                        bars = int(beats_since_t // num)
                        beat = int(beats_since_t % num) + 1
                        return bars + 1, beat, f"{num}/{denom}"
                    bar_in_segment = int(beats_since_t // num)
                    beat = int(beats_since_t % num) + 1
                    return bars_before + bar_in_segment + 1, beat, f"{num}/{denom}"
                else:
                    segment_beats = (next_off - t_off) / beat_len
                    bars_in_segment = int(segment_beats // num)
                    bars_before += bars_in_segment
            num, denom = time_signatures[-1][1], time_signatures[-1][2]
            beat_len = 4.0 / denom
            beats = offset / beat_len
            return int(beats // num) + 1, int(beats % num) + 1, f"{num}/{denom}"


    def __init__(self):
        super().__init__()
        self.title("🎵 MIDI Drive Analyzer")
        
        # Set window size and position intelligently
        window_width = 650
        window_height = 850
        settings_width = 350  # Anticipated settings dialog width
        
        # Get screen dimensions
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # Calculate position to ensure both main window and settings fit on screen
        total_width = window_width + settings_width
        
        # Position main window higher and ensure room for settings
        if total_width < screen_width:
            # Center the combined width, but position main window higher
            x = (screen_width - total_width) // 2
        else:
            # If too wide, position at left edge
            x = 20
            
        # Position window in upper third of screen, with taskbar consideration
        y = max(50, (screen_height - window_height) // 4)
        
        # Apply geometry with position
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.configure(bg="black")

        # Configure dark theme styles
        from tkinter import ttk
        style = ttk.Style()
        style.configure("White.TCheckbutton", background="black", foreground="white", focuscolor="black")
        style.configure("White.TRadiobutton", background="black", foreground="white", focuscolor="black")
        style.configure("White.TLabel", background="black", foreground="white")
        style.configure("TFrame", background="black")

        # Analysis algorithm settings
        self.include_triads = True
        self.sensitivity = "Low"
        self.min_duration = 0.0
        self.remove_repeats = True
        self.include_anacrusis = True
        self.include_non_drive_events = True
        self.arpeggio_searching = True
        self.neighbour_notes_searching = True
        self.arpeggio_block_similarity_threshold = 0.5
        self.pedal_mode = "Off"  # "Off", "Every Beat", "Strong Beats", "Every Bar", "Half Bar"
        
        # Time-segment analysis settings
        self.analysis_mode = "event"  # "event" or "time_segment"
        self.segment_size = "beats"   # "beats", "bars", "half_beats"

        # Event merging configuration
        self.collapse_similar_events = True
        self.collapse_sensitivity_pos = getattr(self, 'collapse_sensitivity_pos', 3)  # 7-position slider, default=3

        # Merge algorithm parameters (updated by slider)
        self.merge_jaccard_threshold = MERGE_JACCARD_THRESHOLD
        self.merge_bass_overlap = MERGE_BASS_OVERLAP
        self.merge_bar_distance = MERGE_BAR_DISTANCE
        self.merge_diff_max = MERGE_DIFF_MAX

        # Application state
        self.loaded_file_path = None
        self.score = None
        self.analyzed_events = None
        self.processed_events = None
        
        # Drive strength parameters (configurable via dialog)
        self.custom_strength_map = None
        self.custom_rule_params = None

        self.build_ui()
        self.show_splash()

    def get_effective_priority_list(self):
        """Generate priority list based on strength values (custom or default)"""
        # Default strength values (same as in DriveStrengthParametersDialog)
        default_strengths = {
            "7": 100, "7b5": 90, "7#5": 80, "m7": 70, "ø7": 65,
            "7m9noroot": 65, "7no3": 55, "7no5": 55, "7noroot": 50,
            "aug": 40, "": 42, "m": 35, "maj7": 30, "mMaj7": 25
        }
        
        # Use custom settings if available, otherwise use defaults
        strength_map = self.custom_strength_map or default_strengths
        
        # Filter out chord types with strength 0 (completely exclude from analysis)
        filtered_strengths = {chord_type: strength for chord_type, strength in strength_map.items() if strength > 0}
        
        # Sort remaining chord types by strength (highest first)
        sorted_chords = sorted(
            filtered_strengths.items(), 
            key=lambda x: x[1], 
            reverse=True
        )
        
        return [chord_type for chord_type, strength in sorted_chords]



    def build_ui(self):
        """Create the main user interface with cross-platform styling."""
        is_mac = platform.system() == "Darwin"
        top_pad = 30 if is_mac else 10  # Extra padding for macOS title bar
        frame = Frame(self, bg="black")
        frame.pack(pady=(top_pad, 10))

        # Platform-specific button styling
        if is_mac:
            btn_kwargs = {}
            disabled_fg = "#cccccc"
        else:
            btn_kwargs = {"bg": "#ffffff", "fg": "#000000", "activebackground": "#f0f0f0", 
                         "activeforeground": "#000000", "relief": "raised", "bd": 2, 
                         "font": ("Segoe UI", 10)}
            disabled_fg = "#808080"

        tk.Button(frame, text="Load XML", command=self.load_music_file, **btn_kwargs).pack(side="left", padx=5)
        self.settings_btn = tk.Button(
            frame,
            text="Settings",
            command=self.open_settings,
            disabledforeground=disabled_fg,
            **btn_kwargs
        )
        self.settings_btn.pack(side="left", padx=5)
        self.show_grid_btn = tk.Button(
            frame,
            text="Show Grid",
            command=self.show_grid_window,
            state="disabled",
            disabledforeground=disabled_fg,
            **btn_kwargs
        )
        self.show_grid_btn.pack(side="left", padx=5)

        def open_keyboard():
            # Create an embedded keyboard window (Toplevel) and run the embedded keyboard class.
            try:
                top = tk.Toplevel(self)
                from_types = (tk.Toplevel,)
                # Instantiate the embedded keyboard UI, passing the main app reference
                EmbeddedMidiKeyboard(top, main_app=self)
            except Exception as e:
                messagebox.showerror("Launch error", f"Failed to open embedded keyboard:\n{e}")

        kb_label = "Keyboard" if is_mac else "Open Keyboard"
        self.keyboard_btn = tk.Button(frame, text=kb_label, command=open_keyboard, **btn_kwargs)
        self.keyboard_btn.pack(side="left", padx=5)

        # Disabled until an analysis exists
        self.save_analysis_btn = tk.Button(
            frame,
            text="Save Analysis",
            command=self.save_analysis_txt,
            state="disabled",
            disabledforeground=disabled_fg,
            **btn_kwargs
        )
        self.save_analysis_btn.pack(side="left", padx=5)

        # Load Analysis should be available immediately
        self.load_analysis_btn = tk.Button(frame, text="Load Analysis", command=self.load_analysis_txt, **btn_kwargs)
        self.load_analysis_btn.pack(side="left", padx=5)

        # Main analysis results display with dark theme and proper text selection
        self.result_text = Text(
            self, bg="black", fg="white", font=("Segoe UI", 11),
            wrap="word", borderwidth=0, highlightthickness=0,
            selectbackground="magenta", selectforeground="white",  # Visible selection colors
            insertbackground="white", relief="flat", padx=0, pady=0,
            insertborderwidth=0, insertwidth=0, 
            highlightbackground="black", highlightcolor="black"
        )
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)

    def create_piano_image(self, octaves=2, key_width=40, key_height=150):
        """Generate a piano keyboard image with highlighted drive tones (G, B, D, F)."""
        white_keys = ['C', 'D', 'E', 'F', 'G', 'A', 'B']
        black_keys = ['C#', 'D#', '', 'F#', 'G#', 'A#', '']
        total_white_keys = 7 * octaves
        img_width = total_white_keys * key_width
        img_height = key_height
        img = Image.new('RGB', (img_width, img_height), color='white')
        draw = ImageDraw.Draw(img)

        for i in range(total_white_keys):
            x = i * key_width
            octave_idx = i // 7
            note = white_keys[i % 7]
            if (octave_idx == 0 and note in ['G', 'B']) or (octave_idx == 1 and note in ['D', 'F']):
                fill_color = '#ff00ff'
            else:
                fill_color = 'white'
            draw.rectangle([x, 0, x + key_width, key_height], fill=fill_color, outline='black')

        for octave in range(octaves):
            for i, key in enumerate(black_keys):
                if key != '':
                    x = (octave * 7 + i) * key_width + int(key_width*0.7)
                    draw.rectangle([x, 0, x + int(key_width*0.6), int(key_height*0.6)], fill='black')

        return img

    def show_splash(self):
        self.result_text.delete("1.0", "end")
        # Configure text spacing to eliminate gray stripes
        self.result_text.configure(spacing1=0, spacing2=0, spacing3=0)
        
        # Configure Segoe UI font tag for splash text
        self.result_text.tag_configure("splash_font", font=("Segoe UI", 11), foreground="white")
        # Insert the title.png image centered
        try:
            from PIL import Image, ImageTk
            img_path = resource_path(os.path.join("assets", "title.png"))
            title_img = Image.open(img_path)
            title_photo = ImageTk.PhotoImage(title_img)
            title_label = tk.Label(self.result_text, image=title_photo, bd=0, bg="black", highlightthickness=0)
            title_label.image = title_photo  # Keep a reference!
            self.result_text.window_create("1.0", window=title_label)
            self.result_text.insert("end", "\n")
        except Exception as e:
            # Insert fallback title with Segoe UI font
            start_pos = self.result_text.index("end")
            self.result_text.insert("end", "Harmonic Drive Analyzer\n")
            end_pos = self.result_text.index("end")
            self.result_text.tag_add("splash_font", start_pos, end_pos)
            print("Splash image load error:", e)
        description = (
"\n"
"• Analyze MusicXML scores\n"
"• Model patterns of harmonic tension\n"
"• Produce PDF graph of tension-release patterns\n"
"• Model harmonic entropy\n\n"
"For information on drive analysis, see: http://www.chromatic-harmony.com/drive_analysis.html\n"
"\n"
            "Kenneth Smith, Desire in Chromatic Harmony (New York: Oxford University Press, 2020).\n"
            "Kenneth Smith, “The Enigma of Entropy in Extended Tonality.” Music Theory Spectrum 43, no. 1 (2021): 1–18."
        )
        
        # Insert description text with Segoe UI font
        start_pos = self.result_text.index("end")
        self.result_text.insert("end", description)
        end_pos = self.result_text.index("end")
        self.result_text.tag_add("splash_font", start_pos, end_pos)
        
        # Add copyright notice
        copyright_text = "\n\n© Kenneth Smith, 2026"
        copyright_start = self.result_text.index("end")
        self.result_text.insert("end", copyright_text)
        copyright_end = self.result_text.index("end")
        self.result_text.tag_add("splash_font", copyright_start, copyright_end)
        
        self.result_text.configure(state="disabled")
       
    def preview_entropy(self, mode: str = "chord", base: int = 2):
        if not self.analyzed_events:
            print("No analyzed events available for entropy preview.")
            return
        ea = EntropyAnalyzer(
            self.analyzed_events, 
            symbol_mode=mode, 
            base=base, 
            logger=print,
            strength_map=self.custom_strength_map,
            rule_params=self.custom_rule_params
        )
        ea.register_step("Stage 1: Strength listing", lambda EA: EA.step_stage1_strengths(print_legend=False))
        ea.register_step("Stage 2: Strength entropy", lambda EA: EA.step_stage2_strength_entropy())
        ea.preview()

    def load_music_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("MusicXML files", "*.xml *.musicxml *.mxl")],
            title="Select a MusicXML file"
        )
        if not path:
            return
        else:
            self.loaded_file_path = path
            self.score = converter.parse(path)
            self.debug_print_notes()
            self.run_analysis()

    def run_analysis(self):
        """Execute full chord analysis pipeline and update UI."""
        min_duration = getattr(self, 'min_duration', 0.0)
        self.analyzed_events = None
        self.processed_events = None
        try:
            if self.analysis_mode == "time_segment":
                lines, events = self.analyze_musicxml_time_segments(self.score)
            else:
                lines, events = self.analyze_musicxml(self.score, min_duration=min_duration)
            self.analyzed_events = events

            self.display_results()
            
            # Generate entropy analysis for advanced statistics
            from io import StringIO
            entropy_buf = StringIO()
            analyzer = EntropyAnalyzer(
                self.analyzed_events, 
                logger=lambda x: print(x, file=entropy_buf),
                strength_map=self.custom_strength_map,
                rule_params=self.custom_rule_params
            )
            analyzer.step_stage1_strengths(print_legend=True)
            self.entropy_review_text = entropy_buf.getvalue()
            
            # Enable UI features after successful analysis
            self.show_grid_btn.config(state="normal")
            try:
                self.save_analysis_btn.config(state="normal")
            except Exception:
                pass
        except Exception as e:
            self.result_text.config(state="normal")
            self.result_text.delete("1.0", "end")
            self.result_text.insert("end", f"Error loading file:\n{e}")
            self.result_text.config(state="disabled")
            self.show_grid_btn.config(state="disabled")
            try:
                self.save_analysis_btn.config(state="disabled")
            except Exception:
                pass
            self.analyzed_events = None
            self.processed_events = None

    def open_settings(self):
        """Open analysis settings dialog with algorithm options and sensitivity controls."""
        dialog = tk.Toplevel(self)
        dialog.title("Analysis Settings")
        dialog.configure(bg="#f5f5f5")  # Set light grey background
        
        # Position dialog flush to the right of main window with matching height
        dialog.update_idletasks()
        
        # Get main window position and dimensions
        main_x = self.winfo_x()
        main_y = self.winfo_y() 
        main_width = self.winfo_width()
        main_height = self.winfo_height()
        
        # Settings dialog dimensions
        settings_width = 350
        settings_height = main_height  # Match main window height
        
        # Position exactly to the right of main window
        settings_x = main_x + main_width
        settings_y = main_y
        
        dialog.geometry(f"{settings_width}x{settings_height}+{settings_x}+{settings_y}")

        # Load and display settings title image at top center
        try:
            from PIL import Image, ImageTk
            title_img_path = resource_path(os.path.join("assets", "images", "settings_title.png"))
            title_img = Image.open(title_img_path)
            title_photo = ImageTk.PhotoImage(title_img)
            title_label = tk.Label(dialog, image=title_photo, bd=0, bg="#f5f5f5", highlightthickness=0)
            title_label.image = title_photo  # Keep a reference
            title_label.pack(pady=(10, 15))
        except Exception as e:
            print(f"Warning: Could not load settings title image: {e}")
            # Fallback text title if image fails
            fallback_title = tk.Label(dialog, text="Analysis Settings", font=("Segoe UI", 16, "bold"), 
                                    bg="#f5f5f5", fg="black")
            fallback_title.pack(pady=(10, 15))

        # Create scrollable frame structure
        main_frame = tk.Frame(dialog, bg="#f5f5f5")
        main_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Create canvas and scrollbar
        canvas = tk.Canvas(main_frame, bg="#f5f5f5", highlightthickness=0)
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f5f5f5")
        
        # Configure scrolling
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mouse wheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # Load info button image
        try:
            from PIL import Image, ImageTk
            info_img_path = resource_path(os.path.join("assets", "images", "info_button.png"))
            info_img = Image.open(info_img_path)
            info_photo = ImageTk.PhotoImage(info_img)
        except Exception as e:
            print(f"Warning: Could not load info button image: {e}")
            info_photo = None
        
        # Configure styles for settings dialog
        from tkinter import ttk
        style = ttk.Style()
        style.configure("Settings.TCheckbutton", background="#f5f5f5", foreground="black")
        style.configure("Settings.TLabel", background="#f5f5f5", foreground="black")
        style.configure("Settings.Heading.TLabel", background="#f5f5f5", foreground="black", font=("Segoe UI", 11))
        style.configure("Settings.Info.TLabel", background="#1f4788", foreground="white", font=("Segoe UI", 8, "bold"), 
                       relief="flat", borderwidth=0, anchor="center", width=2, padding=(8, 6))
        style.configure("Settings.TFrame", background="#f5f5f5")
        style.configure("Settings.TButton", background="white", foreground="black")
        style.configure("Settings.TSeparator", background="#cccccc")
        style.configure("Settings.TRadiobutton", background="#f5f5f5", foreground="black")
        style.configure("Settings.TCombobox", background="white", foreground="black")

        # Create function for adding setting sections with tooltips
        def create_setting_section(parent, title, tooltip, control_widget):
            # Heading frame with title and info symbol
            heading_frame = ttk.Frame(parent, style="Settings.TFrame")
            heading_frame.pack(fill="x", padx=12, pady=(8, 2))
            
            # Title label
            title_label = ttk.Label(heading_frame, text=title, style="Settings.Heading.TLabel")
            title_label.pack(side="left")
            
            # Info symbol with tooltip - PNG image
            if info_photo:
                info_label = tk.Label(heading_frame, image=info_photo, bd=0, bg="#f5f5f5", highlightthickness=0, cursor="hand2")
                info_label.image = info_photo  # Keep a reference
            else:
                # Fallback to text if image fails to load
                info_label = ttk.Label(heading_frame, text="i", style="Settings.Info.TLabel")
            info_label.pack(side="left", padx=(8, 0))
            
            # Add tooltip with proper hide on mouse leave
            tooltip_window_ref = [None]  # Use list to allow modification in nested functions
            
            def show_tooltip(event):
                # Hide any existing tooltip first
                if tooltip_window_ref[0]:
                    try:
                        tooltip_window_ref[0].destroy()
                    except:
                        pass
                
                # Simple tooltip implementation
                tooltip_window = tk.Toplevel(parent)
                tooltip_window_ref[0] = tooltip_window
                tooltip_window.wm_overrideredirect(True)
                tooltip_window.configure(bg="#333333")
                x = event.widget.winfo_rootx() + 20
                y = event.widget.winfo_rooty() + 20
                tooltip_window.geometry(f"+{x}+{y}")
                
                label = tk.Label(tooltip_window, text=tooltip, bg="#333333", fg="white", 
                               font=("Segoe UI", 9), wraplength=250, justify="left", padx=8, pady=4)
                label.pack()
            
            def hide_tooltip(event=None):
                if tooltip_window_ref[0]:
                    try:
                        tooltip_window_ref[0].destroy()
                        tooltip_window_ref[0] = None
                    except:
                        pass
            
            info_label.bind("<Enter>", show_tooltip)
            info_label.bind("<Leave>", hide_tooltip)
            
            # Control container
            control_frame = ttk.Frame(parent, style="Settings.TFrame")
            control_frame.pack(fill="x", padx=24, pady=(0, 4))
            
            return control_frame
        
        pad_opts = dict(anchor="w", padx=12, pady=6)
        
        # Analysis Mode
        analysis_mode_var = tk.StringVar(value=getattr(self, 'analysis_mode', 'event'))
        segment_size_var = tk.StringVar(value=getattr(self, 'segment_size', 'beats'))
        
        # Analysis algorithm toggles - define these early so they can be referenced
        include_triads_var = tk.BooleanVar(value=self.include_triads)
        remove_repeats_var = tk.BooleanVar(value=self.remove_repeats)
        include_anacrusis_var = tk.BooleanVar(value=self.include_anacrusis)
        arpeggio_searching_var = tk.BooleanVar(value=self.arpeggio_searching)
        neighbour_notes_var = tk.BooleanVar(value=getattr(self, 'neighbour_notes_searching', True))
        include_non_drive_var = tk.BooleanVar(value=self.include_non_drive_events)

        # Function to update checkbox states based on analysis mode - define early so radio buttons can reference it
        def update_checkbox_states():
            mode = analysis_mode_var.get()
            if mode == "time_segment":
                # Disable incompatible options for time-segment mode
                anacrusis_cb.config(state="disabled")
                arpeggio_cb.config(state="disabled")
                neighbour_cb.config(state="disabled")
                include_anacrusis_var.set(False)
                arpeggio_searching_var.set(False)
                neighbour_notes_var.set(False)
            else:
                # Re-enable all options for event-based mode
                anacrusis_cb.config(state="normal")
                arpeggio_cb.config(state="normal")
                neighbour_cb.config(state="normal")
        
        analysis_mode_frame = create_setting_section(scrollable_frame, "Analysis Mode", 
            "Choose between event-based analysis (recommended, detects actual musical events) or time-segment analysis (divides music into regular time intervals)", None)

        # Event-based mode radio button
        event_radio = ttk.Radiobutton(analysis_mode_frame, text="Event-based (current method)", variable=analysis_mode_var, 
                       value="event", style="Settings.TRadiobutton", command=update_checkbox_states)
        event_radio.pack(anchor="w", pady=2)
        
        # Time-segment mode radio button and options
        time_segment_frame = ttk.Frame(analysis_mode_frame, style="Settings.TFrame")
        time_segment_frame.pack(fill="x", pady=2)
        
        time_segment_radio = ttk.Radiobutton(time_segment_frame, text="Time-segment based", variable=analysis_mode_var,
                       value="time_segment", style="Settings.TRadiobutton", command=update_checkbox_states)
        time_segment_radio.pack(side="left")
        
        # Segment size dropdown
        ttk.Label(time_segment_frame, text="Size:", style="Settings.TLabel").pack(side="left", padx=(10, 5))
        segment_dropdown = ttk.Combobox(time_segment_frame, textvariable=segment_size_var,
                                      values=["half_beats", "beats", "bars"], state="readonly", width=12,
                                      style="Settings.TCombobox")
        segment_dropdown.pack(side="left")
        
        # Include Triads
        triads_frame = create_setting_section(scrollable_frame, "Include Triads", 
            "Detect and analyze three-note chords (triads) in addition to seventh chords", None)
        ttk.Checkbutton(triads_frame, text="Enabled", variable=include_triads_var, style="Settings.TCheckbutton").pack(anchor="w")
        
        # Include Anacrusis
        anacrusis_frame = create_setting_section(scrollable_frame, "Include Anacrusis", 
            "Analyze pickup notes (anacruses) that occur before the main beat", None)
        anacrusis_cb = ttk.Checkbutton(anacrusis_frame, text="Enabled", variable=include_anacrusis_var, style="Settings.TCheckbutton")
        anacrusis_cb.pack(anchor="w")
        
        # Arpeggio Searching
        arpeggio_frame = create_setting_section(scrollable_frame, "Arpeggio Searching", 
            "Detect arpeggiated chords (broken chords played in sequence) and group them together", None)
        arpeggio_cb = ttk.Checkbutton(arpeggio_frame, text="Enabled", variable=arpeggio_searching_var, style="Settings.TCheckbutton")
        arpeggio_cb.pack(anchor="w")
        
        # Neighbour Notes
        neighbour_frame = create_setting_section(scrollable_frame, "Neighbour Notes", 
            "Include neighbour note analysis (non-chord tones that move by step and return)", None)
        neighbour_cb = ttk.Checkbutton(neighbour_frame, text="Enabled", variable=neighbour_notes_var, style="Settings.TCheckbutton")
        neighbour_cb.pack(anchor="w")
        
        # Remove Repeated Patterns
        repeats_frame = create_setting_section(scrollable_frame, "Remove Repeated Patterns", 
            "Eliminate duplicate harmonic events that occur consecutively", None)
        ttk.Checkbutton(repeats_frame, text="Enabled", variable=remove_repeats_var, style="Settings.TCheckbutton").pack(anchor="w")
        
        # Include Non-Drive Events
        nondrive_frame = create_setting_section(scrollable_frame, "Include Non-Drive Events", 
            "Include events of 3 pitches or more, which don't contribute to harmonic drive (e.g. trichords that don't form diatonic chords). The bass note only will be recorded on the grid in a column of its own.", None)
        ttk.Checkbutton(nondrive_frame, text="Enabled", variable=include_non_drive_var, style="Settings.TCheckbutton").pack(anchor="w")
        
        # Sustain Pedal
        pedal_frame = create_setting_section(scrollable_frame, "Sustain Pedal", 
            "Simulate sustain pedal to hold notes across time boundaries, creating richer harmonic analysis", None)
        
        pedal_enabled_var = tk.BooleanVar(value=getattr(self, 'pedal_mode', 'Off') != 'Off')
        
        # Put checkbox and mode dropdown on the same line
        pedal_combo_frame = ttk.Frame(pedal_frame, style="Settings.TFrame")
        pedal_combo_frame.pack(fill="x")
        
        pedal_cb = ttk.Checkbutton(pedal_combo_frame, text="Enabled", variable=pedal_enabled_var, style="Settings.TCheckbutton")
        pedal_cb.pack(side="left")
        
        pedal_mode_var = tk.StringVar(value=getattr(self, 'pedal_mode', 'Auto') if getattr(self, 'pedal_mode', 'Off') != 'Off' else 'Auto')
        pedal_options = ["Every Beat", "Strong Beats", "Half Bar", "Every Bar", "Auto"]
        ttk.Label(pedal_combo_frame, text="Mode:", style="Settings.TLabel").pack(side="left", padx=(20, 5))
        pedal_combo = ttk.Combobox(pedal_combo_frame, textvariable=pedal_mode_var, values=pedal_options, state="readonly", width=12)
        pedal_combo.pack(side="left")
        
        def update_pedal_combo_state():
            if pedal_enabled_var.get():
                pedal_combo.config(state="readonly")
            else:
                pedal_combo.config(state="disabled")
        
        pedal_cb.config(command=update_pedal_combo_state)
        update_pedal_combo_state()  # Set initial state
        
        # Filter Short Durations
        duration_frame = create_setting_section(scrollable_frame, "Duration Filter", 
            "Exclude very short notes from analysis to focus on structurally significant events", None)
        
        duration_filter_enabled_var = tk.BooleanVar(value=getattr(self, 'min_duration', 0.0) > 0.0)
        
        # Put checkbox and dropdown on the same line
        duration_combo_frame = ttk.Frame(duration_frame, style="Settings.TFrame")
        duration_combo_frame.pack(fill="x")
        
        duration_filter_cb = ttk.Checkbutton(duration_combo_frame, text="Enabled", variable=duration_filter_enabled_var, style="Settings.TCheckbutton")
        duration_filter_cb.pack(side="left")
        
        # Map current min_duration to duration threshold
        current_min_duration = getattr(self, 'min_duration', 0.0)
        duration_to_name = {0.125: "32nd notes", 0.25: "16th notes", 0.5: "8th notes", 1.0: "Quarter notes"}
        current_duration = duration_to_name.get(current_min_duration, "8th notes")
        duration_threshold_var = tk.StringVar(value=current_duration)
        duration_options = ["32nd notes", "16th notes", "8th notes", "Quarter notes"]
        ttk.Label(duration_combo_frame, text="Shorter than:", style="Settings.TLabel").pack(side="left", padx=(20, 5))
        duration_combo = ttk.Combobox(duration_combo_frame, textvariable=duration_threshold_var, values=duration_options, state="readonly", width=12)
        duration_combo.pack(side="left")
        
        def update_duration_combo_state():
            if duration_filter_enabled_var.get():
                duration_combo.config(state="readonly")
            else:
                duration_combo.config(state="disabled")
        
        duration_filter_cb.config(command=update_duration_combo_state)
        update_duration_combo_state()  # Set initial state
        
        # Event Change Sensitivity
        sensitivity_frame = create_setting_section(scrollable_frame, "Event Change Sensitivity", 
            "Control how sensitive the analysis is to chord fluctuations - fine detail captures every change, broad grouping focuses on significant shifts", None)
        
        # Map current position to sensitivity level
        current_pos = getattr(self, 'collapse_sensitivity_pos', 3)
        pos_to_sensitivity = {1: "Record all variations", 2: "Capture small changes", 3: "Default", 5: "Focus on clear changes", 7: "Show only major shifts"}
        sensitivity_options = [
            "Record all variations",
            "Capture small changes", 
            "Default",
            "Focus on clear changes",
            "Show only major shifts"
        ]
        
        # Find current setting
        current_sensitivity_short = pos_to_sensitivity.get(current_pos, "Default")
        current_sensitivity_full = next((opt for opt in sensitivity_options if opt.startswith(current_sensitivity_short)), sensitivity_options[2])
        
        event_sensitivity_var = tk.StringVar(value=current_sensitivity_full)
        event_sensitivity_combo = ttk.Combobox(sensitivity_frame, textvariable=event_sensitivity_var, values=sensitivity_options, state="readonly")
        event_sensitivity_combo.pack(fill="x")

        # Initialize checkbox states based on current mode
        update_checkbox_states()
        
        # Drive Strength Configuration
        def open_strength_dialog():
            strength_dialog = DriveStrengthParametersDialog(
                dialog, 
                self.custom_strength_map, 
                self.custom_rule_params
            )
            dialog.wait_window(strength_dialog.window)
            if strength_dialog.result:
                self.custom_strength_map = strength_dialog.new_strength_map
                self.custom_rule_params = strength_dialog.new_rule_params
                # Re-run analysis if data is loaded
                if getattr(self, 'analyzed_events', None):
                    try:
                        self.display_results()
                    except Exception:
                        pass
        
        strength_frame = create_setting_section(scrollable_frame, "Drive Strength Configuration", 
            "Customize the harmonic drive strength values and analysis rules for different chord types", None)
        
        ttk.Button(
            strength_frame, 
            text="Configure", 
            command=open_strength_dialog,
            style="Settings.TButton"
        ).pack(anchor="w")

        def apply_settings():
            # Apply analysis mode settings
            self.analysis_mode = analysis_mode_var.get()
            self.segment_size = segment_size_var.get()
            
            # Apply basic boolean/text settings
            self.include_triads = include_triads_var.get()
            self.remove_repeats = remove_repeats_var.get()
            
            # For time-segment mode, force certain settings to False
            if self.analysis_mode == "time_segment":
                self.include_anacrusis = False
                self.arpeggio_searching = False
                self.neighbour_notes_searching = False
            else:
                self.include_anacrusis = include_anacrusis_var.get()
                self.arpeggio_searching = arpeggio_searching_var.get()
                self.neighbour_notes_searching = neighbour_notes_var.get()
                
            self.include_non_drive_events = include_non_drive_var.get()

            # Read duration filtering settings
            if duration_filter_enabled_var.get():
                duration_map = {
                    "32nd notes": 0.125,
                    "16th notes": 0.25, 
                    "8th notes": 0.5,
                    "Quarter notes": 1.0
                }
                self.min_duration = duration_map.get(duration_threshold_var.get(), 0.5)
                self.sensitivity = "Medium"  # Keep for backwards compatibility
            else:
                self.min_duration = 0.0
                self.sensitivity = "Low"
            
            # Read pedal mode setting directly from dropdown
            self.pedal_mode = pedal_mode_var.get()

            # Read event sensitivity setting and map to slider position
            selected_sensitivity = event_sensitivity_var.get()
            sensitivity_to_pos = {
                "Record all variations": 1,
                "Capture small changes": 2,
                "Default": 3,
                "Focus on clear changes": 5,
                "Show only major shifts": 7
            }
            pos = sensitivity_to_pos.get(selected_sensitivity, 3)
            self.collapse_sensitivity_pos = pos
            self.collapse_similar_events = True

            # Configure merge algorithm parameters based on slider position (1-7)
            presets = {
                1: {"jaccard": 0.95, "bass": 0.85, "bar": 0, "diff": 0},  # Minimal merging
                2: {"jaccard": 0.90, "bass": 0.80, "bar": 0, "diff": 0},  # Very low merging
                3: {"jaccard": 0.85, "bass": 0.70, "bar": 0, "diff": 0},  # Low merging (DEFAULT)
                4: {"jaccard": 0.70, "bass": 0.60, "bar": 1, "diff": 1},  # Medium-low merging
                5: {"jaccard": 0.60, "bass": 0.50, "bar": 1, "diff": 1},  # Medium merging
                6: {"jaccard": 0.55, "bass": 0.40, "bar": 1, "diff": 2},  # Medium-high merging
                7: {"jaccard": 0.45, "bass": 0.25, "bar": 2, "diff": 2},  # High merging
            }
            chosen = presets.get(pos, presets[3])
            self.merge_jaccard_threshold = chosen["jaccard"]
            self.merge_bass_overlap = chosen["bass"]
            self.merge_bar_distance = chosen["bar"]
            self.merge_diff_max = chosen["diff"]

            # Don't destroy dialog - let user close with X button
            # dialog.destroy()  # Removed - settings stay open after Apply
            
            # Re-run analysis with new settings if file is loaded
            if self.score:
                self.run_analysis()
            elif getattr(self, 'analyzed_events', None):
                try:
                    self.display_results()
                except Exception:
                    pass

            # Refresh grid window if open
            try:
                gw = getattr(self, '_grid_window', None)
                if gw and isinstance(gw, tk.Toplevel) and gw.winfo_exists():
                    gw.destroy()
                    self._grid_window = None
                    self._grid_window = None
                    # Reopen the grid if we still have analyzed events
                    if getattr(self, 'analyzed_events', None):
                        try:
                            self.show_grid_window()
                        except Exception:
                            pass
            except Exception:
                pass

        # Add buttons at bottom
        button_frame = tk.Frame(scrollable_frame, bg="#f5f5f5")
        button_frame.pack(fill=tk.X, pady=(15, 10), padx=0)
        
        # Restore defaults function
        def restore_defaults():
            # Reset all variables to their actual default values
            analysis_mode_var.set("event")
            segment_size_var.set("beats")
            include_triads_var.set(True)
            remove_repeats_var.set(True)
            include_anacrusis_var.set(True)
            arpeggio_searching_var.set(True)
            neighbour_notes_var.set(True)
            include_non_drive_var.set(True)
            pedal_enabled_var.set(False)
            pedal_mode_var.set("Off")
            duration_filter_enabled_var.set(False)
            duration_threshold_var.set("8th notes")
            event_sensitivity_var.set("Default")
        
        # Restore defaults button (left side) - now pink
        defaults_btn = tk.Button(button_frame, text="Restore defaults", command=restore_defaults,
                               bg="#ffffff", fg="#000000", font=("Segoe UI", 10),
                               relief="raised", bd=2, width=14)
        
        # Apply button (right side)
        apply_btn = tk.Button(button_frame, text="Apply", command=apply_settings,
                             bg="#ffffff", fg="#000000", font=("Segoe UI", 10),
                             relief="raised", bd=2, width=12)
        
        # Pack both buttons to the right side
        apply_btn.pack(side=tk.RIGHT, padx=(0, 0))
        defaults_btn.pack(side=tk.RIGHT, padx=(5, 5))

        # Set up normal dialog close (no auto-apply)
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)

    def display_results(self, lines: list[str] = None):
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        if self.analyzed_events:
            events = list(sorted(self.analyzed_events.items())) if self.analyzed_events else []

            # Remove immediately repeated patterns if option is enabled
            if self.analyzed_events and getattr(self, 'remove_repeats', False):
                def event_signature(event):
                    data = event[1]
                    chords = tuple(sorted(data["chords"]))
                    basses = tuple(sorted(data["basses"]))
                    return (chords, basses)

                filtered = []
                i = 0
                n = len(events)
                while i < n:
                    max_pat = (n - i) // 2
                    found_repeat = False
                    for pat_len in range(1, max_pat + 1):
                        pat = [event_signature(events[i + j]) for j in range(pat_len)]
                        repeat = True
                        for j in range(pat_len):
                            if i + pat_len + j >= n or event_signature(events[i + j]) != event_signature(events[i + pat_len + j]):
                                repeat = False
                                break
                        if repeat:
                            # keep the first occurrence, then skip any number of consecutive repeats
                            jpos = i + pat_len
                            while jpos + pat_len <= n:
                                all_match = True
                                for k in range(pat_len):
                                    if event_signature(events[i + k]) != event_signature(events[jpos + k]):
                                        all_match = False
                                        break
                                if all_match:
                                    jpos += pat_len
                                else:
                                    break
                            filtered.extend(events[i:i+pat_len])
                            i = jpos
                            found_repeat = True
                            break
                    if not found_repeat:
                        filtered.append(events[i])
                        i += 1
                events = filtered

            # Determine whether any drive (recognized chord) exists in the event list
            has_any_drives = any(data.get('chords') for (_, data) in events)

            if lines is not None:
                # Store the events as-is when displaying pre-formatted lines
                self.processed_events = events.copy()
                for line in lines:
                    self.result_text.insert("end", line)
            elif events:
                # Collect the ACTUAL events that get displayed after all filtering
                displayed_events = []
                output_lines = []
                prev_no_drive = False
                prev_bass = None
                for (bar, beat, ts), data in events:
                    chords = sorted(data["chords"])
                    chord_info = data.get("chord_info", {})
                    chord_strs = []
                    for chord in chords:
                        marker = ""
                        if chord_info.get(chord, {}).get("clean_stack"):
                            marker += CLEAN_STACK_SYMBOL
                        root_count = chord_info.get(chord, {}).get("root_count", 1)
                        if root_count == 2:
                            marker += ROOT2_SYMBOL
                        elif root_count >= 3:
                            marker += ROOT3_SYMBOL
                        chord_strs.append(f"{chord}{marker}")
                    chords_display = ", ".join(chord_strs) if chord_strs else ""
                    bass = "+".join(data["basses"])
                    is_no_drive = len(chord_strs) == 0
                    
                    # Format output based on whether there are known chords
                    if is_no_drive:
                        line_content = f"(bass = {bass})"
                    else:
                        line_content = f"{chords_display} (bass = {bass})"
                    # Deduplicate at output: skip if previous was also no known drive with same bass
                    if is_no_drive and prev_no_drive and bass == prev_bass:
                        continue
                    # Respect user preference for including non-drive events
                    if is_no_drive and not self.include_non_drive_events:
                        continue
                    
                    # This event will be displayed, so add it to our list
                    displayed_events.append(((bar, beat, ts), data))
                    output_lines.append(
                        f"Bar {bar}, Beat {beat} ({ts}): {line_content}\n"
                    )
                    prev_no_drive = is_no_drive
                    prev_bass = bass
                
                # Store only the events that actually got displayed
                self.processed_events = displayed_events
                output_lines.append(
                    f"\nLegend:\n{CLEAN_STACK_SYMBOL} = Clean stack   {ROOT2_SYMBOL} = Root doubled   {ROOT3_SYMBOL} = Root tripled or more\n"
                )
                # Replace musical symbols before displaying - only after note names
                import re
                final_output = "".join(output_lines)
                # Replace flats: note names followed by 'b' OR 'b' followed by numbers (chord extensions)
                final_output = re.sub(r'([ABCDEFG])b', r'\1♭', final_output)  # Direct note flats like Db
                final_output = re.sub(r'b([0-9]+)', r'♭\1', final_output)  # Chord extensions like b5, b9
                # Replace sharps: note names (A-G) followed by '#'
                final_output = re.sub(r'([ABCDEFG])#', r'\1♯', final_output)
                self.result_text.insert("end", final_output)
        self.result_text.config(state="disabled")


    def analyze_musicxml(self, score, min_duration=0.5):
        """
        Main chord analysis algorithm.
        
        Processes MusicXML score through multiple detection phases:
        1. Block chord detection from simultaneous notes
        2. Arpeggio pattern recognition from melodic sequences
        3. Anacrusis handling for melodic resolution notes
        4. Neighbor/passing note detection
        5. Event merging and post-processing
        """
        flat_notes = list(score.flatten().getElementsByClass([note.Note, m21chord.Chord]))

        # Extract time signatures for bar/beat calculation
        time_signatures = []
        for ts in score.flatten().getElementsByClass(meter.TimeSignature):
            offset = float(ts.offset)
            time_signatures.append((offset, int(ts.numerator), int(ts.denominator)))

        time_signatures.sort(key=lambda x: x[0])

        # Ensure we have at least one time signature
        if not time_signatures:
            time_signatures = [(0.0, 4, 4)]
        elif time_signatures[0][0] > 0.0:
            first_num, first_den = time_signatures[0][1], time_signatures[0][2]
            time_signatures.insert(0, (0.0, first_num, first_den))

        def get_time_signature(offset):
            # Find the last time signature whose offset is <= given offset
            ts = (4, 4)
            for t_off, n, d in time_signatures:
                if offset >= t_off:
                    ts = (n, d)
                else:
                    break
            return ts

        def offset_to_bar_beat(offset):
            # Map an absolute offset (in quarter lengths) to bar and beat
            if not time_signatures:
                num, denom = 4, 4
                return 1, int(offset) + 1, f"{num}/{denom}"

            # Walk through timesig segments and accumulate full bars
            bars_before = 0
            for i, (t_off, num, denom) in enumerate(time_signatures):
                next_off = time_signatures[i + 1][0] if i + 1 < len(time_signatures) else None
                beat_len = 4.0 / denom  # quarter lengths per beat

                if next_off is None or offset < next_off:
                    # offset lies in this segment
                    beats_since_t = (offset - t_off) / beat_len
                    if beats_since_t < 0:
                        # offset before first timesig marker
                        beats_since_t = (offset) / beat_len
                        bars = int(beats_since_t // num)
                        beat = int(beats_since_t % num) + 1
                        return bars + 1, beat, f"{num}/{denom}"
                    bar_in_segment = int(beats_since_t // num)
                    beat = int(beats_since_t % num) + 1
                    return bars_before + bar_in_segment + 1, beat, f"{num}/{denom}"
                else:
                    # full segment contributes whole bars
                    segment_beats = (next_off - t_off) / beat_len
                    bars_in_segment = int(segment_beats // num)
                    bars_before += bars_in_segment

            # Fallback (shouldn't normally reach here)
            num, denom = time_signatures[-1][1], time_signatures[-1][2]
            beat_len = 4.0 / denom
            beats = offset / beat_len
            return int(beats // num) + 1, int(beats % num) + 1, f"{num}/{denom}"

        def is_pedal_lift_point(offset, pedal_mode):
            """Determine if pedal lifts at this time point based on mode."""
            if pedal_mode == "Off":
                return False
            
            # Auto mode uses intelligent detection (handled in main loop)
            if pedal_mode == "Auto":
                return False  # Actual logic handled separately
            
            bar, beat, ts = offset_to_bar_beat(offset)
            num, denom = get_time_signature(offset)
            beat_len = 4.0 / denom  # quarter lengths per beat
            
            # Calculate the exact start time of this bar
            bar_start_offset = None
            bars_before = 0
            for i, (t_off, n, d) in enumerate(time_signatures):
                next_off = time_signatures[i + 1][0] if i + 1 < len(time_signatures) else None
                if next_off is None or offset < next_off:
                    # This time signature applies to our offset
                    segment_start = t_off
                    beats_since_start = (offset - t_off) / beat_len
                    bars_in_this_segment = int(beats_since_start // n)
                    bar_start_offset = t_off + (bars_in_this_segment * n * beat_len)
                    break
                else:
                    # Add complete bars from this segment
                    segment_beats = (next_off - t_off) / beat_len
                    bars_before += int(segment_beats // n)
            
            if bar_start_offset is None:
                return False
            
            # Only lift pedal at exact bar starts
            tolerance = 0.001  # Small tolerance for floating point precision
            is_exact_bar_start = abs(offset - bar_start_offset) < tolerance
            
            if pedal_mode == "Every Beat":
                # Pedal lifts at every beat boundary (exact beat starts)
                beat_start_offset = bar_start_offset + ((beat - 1) * beat_len)
                return abs(offset - beat_start_offset) < tolerance
            elif pedal_mode == "Strong Beats":
                # Pedal lifts at beats 1 and halfway point (exact beat starts)
                if beat == 1 or beat == (num // 2) + 1:
                    beat_start_offset = bar_start_offset + ((beat - 1) * beat_len)
                    return abs(offset - beat_start_offset) < tolerance
            elif pedal_mode == "Every Bar":
                # Pedal lifts only at exact bar starts
                return is_exact_bar_start
            elif pedal_mode == "Half Bar":
                # Pedal lifts at bar start and exact halfway through bar
                if beat == 1:
                    return is_exact_bar_start
                elif beat == (num // 2) + 1:
                    beat_start_offset = bar_start_offset + ((beat - 1) * beat_len)
                    return abs(offset - beat_start_offset) < tolerance
            return False

        note_events = []
        single_notes = []  # (start, end, pitch)
        for elem in flat_notes:
            if isinstance(elem, (note.Note, m21chord.Chord)):
                original_duration = elem.quarterLength
                
                # Filter out notes shorter than min_duration
                if original_duration < min_duration:
                    continue  # Skip this note entirely
                
                duration = original_duration
                start = elem.offset
                end = start + duration
                
                if isinstance(elem, m21chord.Chord):
                    pitches = [p.midi for p in elem.pitches]
                else:
                    pitches = [elem.pitch.midi]
                note_events.append((start, end, pitches))
                # Collect single melodic notes (not part of a chord, not doubled at start)
                if isinstance(elem, note.Note):
                    others_at_start = [e for e in flat_notes if e is not elem and hasattr(e, 'offset') and e.offset == start and e.quarterLength >= min_duration]
                    if not others_at_start:
                        single_notes.append((start, end, pitches[0]))

        time_points = sorted(set([t for start, end, _ in note_events for t in [start, end]]))



        events = {}
        active_notes = set()
        active_pitches = set()
        
        # Pedal-sustained notes (notes that started while pedal was down)
        pedal_sustained_notes = set()
        pedal_sustained_pitches = set()
        last_pedal_lift_time = None
        
        # Auto pedal tracking
        auto_pedal_collection = set()  # All pitch classes since last auto pedal lift
        auto_last_lift_time = 0.0

        # === PHASE 1: Block Chord Detection ===
        for i, time in enumerate(time_points):

            # Check if pedal lifts at this time point
            auto_pedal_lift = False
            if self.pedal_mode == "Auto":
                # Auto pedal logic - three conditions for lifting
                current_bar, current_beat, current_ts = offset_to_bar_beat(time)
                bar_length = 4.0 * get_time_signature(time)[0] / get_time_signature(time)[1]  # Quarter lengths per bar
                
                # Condition 1: Bar boundary (minimum frequency)
                if time >= auto_last_lift_time + bar_length:
                    auto_pedal_lift = True
                
                # Condition 2: Mass note ending (3+ simultaneous pitches end)
                if not auto_pedal_lift:
                    ending_count = 0
                    for start, end, pitches in note_events:
                        if end == time:
                            ending_count += len(pitches)
                    if ending_count >= 3:
                        auto_pedal_lift = True
                
                # Condition 3: Harmonic shift (pitch collection similarity < 0.5)
                if not auto_pedal_lift and auto_pedal_collection and active_notes:
                    intersection = auto_pedal_collection & active_notes
                    union = auto_pedal_collection | active_notes
                    similarity = len(intersection) / len(union) if union else 1.0
                    if similarity < 0.5:
                        auto_pedal_lift = True
                
                if auto_pedal_lift:
                    pedal_sustained_notes.clear()
                    pedal_sustained_pitches.clear()
                    auto_pedal_collection.clear()
                    auto_last_lift_time = time
            
            elif is_pedal_lift_point(time, self.pedal_mode):
                pedal_sustained_notes.clear()
                pedal_sustained_pitches.clear()
                last_pedal_lift_time = time
                
            # Track which notes end at this time
            ending_notes = []
            for start, end, pitches in note_events:
                if end == time:
                    ending_notes.extend(pitches)
                    active_notes.difference_update({p % 12 for p in pitches})
                    active_pitches.difference_update(pitches)

            # Track which notes start at this time        
            starting_notes = []
            for start, end, pitches in note_events:
                if start == time:
                    starting_notes.extend(pitches)
                    active_notes.update({p % 12 for p in pitches})
                    active_pitches.update(pitches)
                    
                    # Add to pedal-sustained notes if pedal is active (sustain ALL notes that sound)
                    if self.pedal_mode != "Off":
                        pedal_sustained_notes.update({p % 12 for p in pitches})
                        pedal_sustained_pitches.update(pitches)
                        
                        # For auto mode, also update the pitch collection
                        if self.pedal_mode == "Auto":
                            auto_pedal_collection.update({p % 12 for p in pitches})
            
            # Also add any currently active notes to pedal sustain (notes that were already sounding)
            if self.pedal_mode != "Off" and active_notes:
                pedal_sustained_notes.update(active_notes)
                pedal_sustained_pitches.update(active_pitches)

            # Combine active notes with pedal-sustained notes for chord detection
            test_notes = set(active_notes) | pedal_sustained_notes
            test_pitches = set(active_pitches) | pedal_sustained_pitches
            
            if self.include_anacrusis:
                anacrusis_added = []
                for s_start, s_end, s_pitch in single_notes:
                    # Only include anacrusis notes that were struck as single pitches in isolation
                    # and end exactly when the current chord analysis point occurs
                    # (single_notes already ensures the note was struck alone, not as part of a chord)
                    if s_end == time and (s_pitch % 12) not in test_notes:
                        test_notes.add(s_pitch % 12)
                        test_pitches.add(s_pitch)
                        anacrusis_added.append(s_pitch)

            if len(test_notes) >= 3:
                # Check for chord formation with sufficient note count
                bar, beat, ts = offset_to_bar_beat(time)
                
                # Analyze note collection for chord detection
                chords = self.detect_chords(test_notes, debug=False)
                bar, beat, ts = offset_to_bar_beat(time)
                key = (bar, beat, ts)
                # Event created; previously had diagnostic printing here which has been removed
                if chords:
                    bass_note = self.semitone_to_note(min(test_pitches) % 12)
                    if key not in events:
                        events[key] = {"chords": set(), "basses": set(), "event_notes": set(test_notes)}

                    
                    events[key]["chords"].update(chords)
                    events[key]["basses"].add(bass_note)
                    
                    events[key]["event_notes"] = set(test_notes)
                    

                    events[key]["event_pitches"] = set(test_pitches)
                else:
                    # No recognized chord, but 3+ notes: still set bass to lowest pitch
                    bass_note = self.semitone_to_note(min(test_pitches) % 12)
                    if key not in events:
                        events[key] = {"chords": set(), "basses": set(), "event_notes": set(test_notes), "event_pitches": set(test_pitches)}
                    events[key]["basses"].add(bass_note)

        # === PHASE 2: Arpeggio Pattern Detection ===
        if self.arpeggio_searching:
            # Build a list of all single notes (not chords) sorted by onset
            melodic_notes = [elem for elem in flat_notes if isinstance(elem, note.Note)]
            melodic_notes = sorted(melodic_notes, key=lambda n: n.offset)
            
            window_sizes = [3, 4]
            for w in window_sizes:
                for i in range(len(melodic_notes) - w + 1):
                    window = melodic_notes[i:i+w]
                    window_offsets = [n.offset for n in window]
                    # Only consider windows with strictly increasing onsets
                    if any(window_offsets[j] >= window_offsets[j+1] for j in range(w-1)):
                        continue
                    window_pitches = [n.pitch.midi for n in window]
                    window_pcs = {p % 12 for p in window_pitches}
                    if len(window_pcs) < 3:
                        continue
                    chords = self.detect_chords(window_pcs, debug=True)
                    if chords:
                        # Display arpeggio analysis for specified range
                        bar, beat, ts = offset_to_bar_beat(window[0].offset)
                        
                        # HARMONIC STABILITY CHECK: Only accept arpeggio if underlying harmony is stable
                        # Check all time points in the arpeggio window for existing block chords
                        underlying_chords = []
                        for note_elem in window:
                            note_bar, note_beat, note_ts = offset_to_bar_beat(note_elem.offset)
                            note_key = (note_bar, note_beat, note_ts)
                            existing_chords = events.get(note_key, {}).get('chords', set())
                            if existing_chords:
                                underlying_chords.append((note_key, existing_chords))
                        
                        # Decision logic:
                        # 1. No block chords throughout span → Accept arpeggio
                        # 2. Same block chord throughout span → Accept arpeggio  
                        # 3. Different block chords in span → Reject arpeggio
                        harmonic_stability = True
                        if underlying_chords:
                            # Check if all underlying chords are the same
                            first_chord_set = underlying_chords[0][1]
                            for key, chord_set in underlying_chords[1:]:
                                if chord_set != first_chord_set:
                                    harmonic_stability = False
                                    break
                        
                        if not harmonic_stability:
                            continue  # Skip this arpeggio
                        
                        # Before accepting arpeggio detection, compare to any simultaneous block event
                        key = (bar, beat, ts)
                        block_pcs = events.get(key, {}).get('event_notes', set())
                        if block_pcs:
                            union = window_pcs | block_pcs
                            inter = window_pcs & block_pcs
                            jaccard = (len(inter) / len(union)) if union else 0.0
                            # Evaluate arpeggio acceptance criteria
                            # Accept arpeggio if Jaccard passes OR if the detected arpeggio chord's root is present in the simultaneous block_pcs
                            accept_arpeggio = False
                            if jaccard >= getattr(self, 'arpeggio_block_similarity_threshold', 0.5):
                                accept_arpeggio = True
                            else:
                                # check whether any detected chord root is present in block_pcs
                                for chord_name in chords:
                                    root = next((n for n in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)) if chord_name.startswith(n)), None)
                                    if root is not None and (NOTE_TO_SEMITONE.get(root) % 12) in block_pcs:
                                        accept_arpeggio = True
                                        break
                            if not accept_arpeggio:
                                continue
                        # Accept arpeggio event
                        if key not in events:
                            events[key] = {"chords": set(), "basses": set(), "event_notes": set(window_pcs)}
                        events[key]["chords"].update(chords)
                        # Arpeggios contribute to chord identification but not bass detection
                        # Bass detection relies on actual bass line or block chord voicings
                        events[key]["event_notes"] = set(window_pcs)
                        events[key]["event_pitches"] = set(window_pitches)
                        events[key]["chords"].update(chords)
                        events[key]["basses"].add(self.semitone_to_note(min(window_pitches) % 12))
                        events[key]["event_notes"] = set(window_pcs)
                        events[key]["event_pitches"] = set(window_pitches)




        # === PHASE 3: Neighbor/Passing Note Detection ===
        # NEW ALGORITHM: Detect harmonic stability with melodic change
        # Look for cases where exactly one note changes while 2+ others are retained,
        # and BIND related events together at the foundational timing
        if getattr(self, 'neighbour_notes_searching', False):
            # Group all note events by bar for boundary respect
            notes_by_bar = {}
            for st, en, prs in note_events:
                bar, _, _ = offset_to_bar_beat(st)
                if bar not in notes_by_bar:
                    notes_by_bar[bar] = []
                notes_by_bar[bar].append((st, en, prs))
            
            # Track events to merge - store as {early_key: [later_keys_to_merge]}
            events_to_bind = {}
            
            # Process each bar separately
            for bar_num, bar_notes in notes_by_bar.items():
                # Sort all note events by start time
                bar_notes.sort(key=lambda x: x[0])
                
                # Track state changes: collect all unique time points where notes start or end
                time_points = set()
                for st, en, prs in bar_notes:
                    time_points.add(st)
                    time_points.add(en)
                time_points = sorted(time_points)
                
                # Analyze state at each time point
                for i in range(len(time_points) - 1):
                    current_time = time_points[i]
                    next_time = time_points[i + 1]
                    
                    # Find all notes sounding at current_time
                    current_state = set()
                    for st, en, prs in bar_notes:
                        if st <= current_time < en:
                            current_state.update({p % 12 for p in prs})
                    
                    # Find all notes sounding at next_time
                    next_state = set()
                    for st, en, prs in bar_notes:
                        if st <= next_time < en:
                            next_state.update({p % 12 for p in prs})
                    
                    # Check if we have exactly one note changing
                    if len(current_state) >= 3 and len(next_state) >= 3:
                        added = next_state - current_state
                        removed = current_state - next_state
                        retained = current_state & next_state
                        
                        # Exactly one note change: one added, one removed, 2+ retained
                        if len(added) == 1 and len(removed) == 1 and len(retained) >= 2:
                            old_note = list(removed)[0]
                            new_note = list(added)[0]
                            
                            # Evaluate chord formation with note substitution
                            test_pcs = {old_note, new_note} | retained
                            
                            # Look for passing notes within the duration of retained notes that might complete better chords
                            # Find the time span during which the retained notes are sounding
                            retained_start = current_time
                            retained_end = next_time
                            for st, en, prs in bar_notes:
                                if st <= current_time < en:
                                    pitch_classes = {p % 12 for p in prs}
                                    if pitch_classes & retained:  # If this contributes to retained notes
                                        retained_end = max(retained_end, en)
                            
                            # Look for any notes that sound during the retained note period
                            passing_pcs = set()
                            for st, en, prs in bar_notes:
                                # Include notes that start and end within the retained note duration
                                if retained_start <= st < retained_end and retained_start < en <= retained_end:
                                    passing_pcs.update({p % 12 for p in prs})
                            
                            # Include passing notes for enhanced chord analysis
                            enhanced_test_pcs = test_pcs | passing_pcs
                            
                            if len(enhanced_test_pcs) >= 4:  # Need at least 4 notes for seventh chord
                                # Try both versions and prefer the enhanced one if it produces better chords
                                basic_chords = self.detect_chords(test_pcs, debug=False)
                                enhanced_chords = self.detect_chords(enhanced_test_pcs, debug=False)
                                
                                # Use enhanced version if it found chords, otherwise fall back to basic
                                final_chords = enhanced_chords if enhanced_chords else basic_chords
                                final_test_pcs = enhanced_test_pcs if enhanced_chords else test_pcs
                                
                                if final_chords:
                                    
                                    # Find the foundational event (earliest time with retained notes)
                                    foundation_time = current_time
                                    foundation_bar, foundation_beat, foundation_ts = offset_to_bar_beat(foundation_time)
                                    foundation_key = (foundation_bar, foundation_beat, foundation_ts)
                                    
                                    # Find the completion event (when new note appears)
                                    completion_time = next_time
                                    completion_bar, completion_beat, completion_ts = offset_to_bar_beat(completion_time)
                                    completion_key = (completion_bar, completion_beat, completion_ts)
                                    
                                    # Plan to bind completion event into foundation event
                                    if foundation_key not in events_to_bind:
                                        events_to_bind[foundation_key] = []
                                    if completion_key != foundation_key:
                                        events_to_bind[foundation_key].append(completion_key)
                                    
                                    # Enhance the foundation event with the discovered chords
                                    if foundation_key not in events:
                                        events[foundation_key] = {"chords": set(), "basses": set(), "event_notes": set()}
                                    events[foundation_key]["chords"].update(final_chords)
                                    events[foundation_key]["event_notes"].update(final_test_pcs)
            
            # Execute the binding: merge later events into foundation events
            for foundation_key, completion_keys in events_to_bind.items():
                if foundation_key in events:
                    for completion_key in completion_keys:
                        if completion_key in events:
                            # Merge the completion event into the foundation event
                            completion_event = events[completion_key]
                            if foundation_key[0] == 3 and foundation_key[1] == 1:
                                print(f"DEBUG NEIGHBOR MERGE: Merging completion {completion_key} -> foundation {foundation_key}")
                                print(f"  Before: foundation chords = {events[foundation_key].get('chords', set())}")
                                print(f"  Adding: completion chords = {completion_event.get('chords', set())}")
                            events[foundation_key]["chords"].update(completion_event.get("chords", set()))
                            events[foundation_key]["basses"].update(completion_event.get("basses", set()))
                            events[foundation_key]["event_notes"].update(completion_event.get("event_notes", set()))
                            events[foundation_key]["event_pitches"] = events[foundation_key].get("event_pitches", set()) | completion_event.get("event_pitches", set())
                            if foundation_key[0] == 3 and foundation_key[1] == 1:
                                print(f"  After: foundation chords = {events[foundation_key].get('chords', set())}")
                            
                            # Remove the completion event since it's now merged
                            del events[completion_key]
        if not events:
            return ["No matching chords found."], {}

        return self._process_detected_events(events)

    def analyze_musicxml_time_segments(self, score):
        """
        Time-segment based analysis: divide music into regular time segments
        and analyze all pitches active during each segment.
        """
        flat_notes = list(score.flatten().getElementsByClass([note.Note, m21chord.Chord]))

        # Extract time signatures for bar/beat calculation
        time_signatures = []
        for ts in score.flatten().getElementsByClass(meter.TimeSignature):
            offset = float(ts.offset)
            time_signatures.append((offset, int(ts.numerator), int(ts.denominator)))

        time_signatures.sort(key=lambda x: x[0])

        # Ensure we have at least one time signature
        if not time_signatures:
            time_signatures = [(0.0, 4, 4)]
        elif time_signatures[0][0] > 0.0:
            first_num, first_den = time_signatures[0][1], time_signatures[0][2]
            time_signatures.insert(0, (0.0, first_num, first_den))

        def get_time_signature(offset):
            ts = (4, 4)
            for t_off, n, d in time_signatures:
                if offset >= t_off:
                    ts = (n, d)
                else:
                    break
            return ts

        def offset_to_bar_beat(offset):
            if not time_signatures:
                num, denom = 4, 4
                return 1, int(offset) + 1, f"{num}/{denom}"

            bars_before = 0
            for i, (t_off, num, denom) in enumerate(time_signatures):
                next_off = time_signatures[i + 1][0] if i + 1 < len(time_signatures) else None
                beat_len = 4.0 / denom

                if next_off is None or offset < next_off:
                    beats_since_t = (offset - t_off) / beat_len
                    if beats_since_t < 0:
                        beats_since_t = (offset) / beat_len
                        bars = int(beats_since_t // num)
                        beat = int(beats_since_t % num) + 1
                        return bars + 1, beat, f"{num}/{denom}"
                    bar_in_segment = int(beats_since_t // num)
                    beat = int(beats_since_t % num) + 1
                    return bars_before + bar_in_segment + 1, beat, f"{num}/{denom}"
                else:
                    segment_beats = (next_off - t_off) / beat_len
                    bars_in_segment = int(segment_beats // num)
                    bars_before += bars_in_segment

            num, denom = time_signatures[-1][1], time_signatures[-1][2]
            beat_len = 4.0 / denom
            beats = offset / beat_len
            return int(beats // num) + 1, int(beats % num) + 1, f"{num}/{denom}"

        # Build note events list
        note_events = []
        for elem in flat_notes:
            if isinstance(elem, (note.Note, m21chord.Chord)):
                start = elem.offset
                end = start + elem.quarterLength
                if isinstance(elem, m21chord.Chord):
                    pitches = [p.midi for p in elem.pitches]
                else:
                    pitches = [elem.pitch.midi]
                note_events.append((start, end, pitches))

        # Calculate segment boundaries
        segments = self._calculate_segment_boundaries(score, time_signatures, offset_to_bar_beat)
        
        events = {}
        
        # Process each segment
        for start_offset, end_offset, bar, beat, ts in segments:
            # Collect all pitches active during this segment
            active_pitches = set()
            for note_start, note_end, pitches in note_events:
                # Check if note overlaps with segment
                if note_start < end_offset and note_end > start_offset:
                    active_pitches.update(pitches)
            
            if len(active_pitches) >= 3:  # Need at least 3 notes for chord detection
                active_pcs = {p % 12 for p in active_pitches}
                chords = self.detect_chords(active_pcs, debug=False)
                
                key = (bar, beat, ts)
                bass_note = self.semitone_to_note(min(active_pitches) % 12) if active_pitches else None
                
                if chords or bass_note:
                    events[key] = {
                        "chords": set(chords) if chords else set(),
                        "basses": {bass_note} if bass_note else set(),
                        "event_notes": active_pcs,
                        "event_pitches": active_pitches
                    }

        if not events:
            return ["No matching chords found in time segments."], {}

        return self._process_detected_events(events)

    def _calculate_segment_boundaries(self, score, time_signatures, offset_to_bar_beat):
        """Calculate time segment boundaries based on selected segment size."""
        segments = []
        
        # Find the total duration of the piece
        flat_notes = list(score.flatten().getElementsByClass([note.Note, m21chord.Chord]))
        if not flat_notes:
            return segments
            
        total_duration = max(elem.offset + elem.quarterLength for elem in flat_notes)
        
        current_offset = 0.0
        
        while current_offset < total_duration:
            # Calculate segment duration based on current time signature and segment size
            num, denom = self._get_time_signature_at_offset(current_offset, time_signatures)
            beat_length = 4.0 / denom  # quarter note lengths per beat
            
            if self.segment_size == "half_beats":
                segment_duration = beat_length / 2
            elif self.segment_size == "beats":
                segment_duration = beat_length
            elif self.segment_size == "bars":
                segment_duration = beat_length * num
            else:
                segment_duration = beat_length  # Default to beats
            
            end_offset = min(current_offset + segment_duration, total_duration)
            bar, beat, ts = offset_to_bar_beat(current_offset)
            
            segments.append((current_offset, end_offset, bar, beat, ts))
            current_offset = end_offset
            
        return segments

    def _get_time_signature_at_offset(self, offset, time_signatures):
        """Helper function to get time signature at a given offset."""
        ts = (4, 4)
        for t_off, n, d in time_signatures:
            if offset >= t_off:
                ts = (n, d)
            else:
                break
        return ts

    def _is_clean_stack(self, chord_name: str, event_notes: set[int]) -> bool:
        """
        Returns True if all required chord notes are present and any extra notes are only outside the stack (not between lowest and highest chord tones, exclusive).
        chord_name: e.g. "C7", "Gm", etc.
        event_notes: set of MIDI pitch classes (0=C, 1=C#, ..., 11=B) present at this event.
        """
        root = next((note for note in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)) if chord_name.startswith(note)), None)
        if not root:
            return False
        base_chord = chord_name.replace(root, 'C')
        if base_chord not in CHORDS:
            return False

        root_pc = NOTE_TO_SEMITONE[root]
        expected_pcs = set((root_pc + i) % 12 for i in CHORDS[base_chord])
        event_notes = set(event_notes)

        # Must contain all required chord notes
        if not expected_pcs.issubset(event_notes):
            return False

        # If no extra notes, it's clean
        if event_notes == expected_pcs:
            return True

        # Find lowest and highest chord tones in the event
        chord_tones_sorted = sorted(expected_pcs)
        min_tone = chord_tones_sorted[0]
        max_tone = chord_tones_sorted[-1]

        # Check for extra notes that fall strictly between min and max chord tones
        for n in event_notes - expected_pcs:
            # Handle wrap-around (e.g., C-E-G, extra note B)
            if min_tone < max_tone:
                if min_tone < n < max_tone:
                    return False
            else:
                # e.g., chord tones G (7), C (0), E (4): min=0, max=7, so between is 1-6
                if (n > min_tone or n < max_tone):
                    return False

        return True
    
    def _count_root_in_pitches(self, chord_name: str, event_pitches: set[int]) -> int:
        """
        Returns how many times the root of chord_name appears in event_pitches (MIDI note numbers).
        """
        root = next((note for note in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)) if chord_name.startswith(note)), None)
        if not root:
            return 0
        root_pc = NOTE_TO_SEMITONE[root]
        return sum(1 for p in event_pitches if p % 12 == root_pc)    

    def _process_detected_events(self, events):
        """Process raw detected events into filtered events ready for display.

        Post-processing pipeline:
        1. Deduplicate chords by priority (higher priority wins per root)
        2. Optionally merge similar adjacent events using Jaccard similarity
        3. Apply repeat removal and filtering based on user settings
        """
        


        def chord_priority(chord_name: str) -> int:
            base = chord_name
            for n in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)):
                if chord_name.startswith(n):
                    base = chord_name.replace(n, 'C')
                    break
            # Remove the 'C' prefix to get the chord quality
            chord_quality = base[1:] if base.startswith('C') else base
            priority_list = self.get_effective_priority_list()
            return priority_list.index(chord_quality) if chord_quality in priority_list else 999

        def dedupe_chords_by_priority(chords_dict: Dict[str, Any]) -> Dict[str, str]:
            result = {}
            for root, chord in chords_dict.items():
                if isinstance(chord, list):
                    best = min(chord, key=chord_priority)
                else:
                    best = chord
                result[root] = best
            return result

        event_items = sorted(events.items())
        # Track tuples as: (key, chords_by_root, basses, event_notes, event_pitches)
        processed_events: List[Tuple[Tuple[int,int,str], Dict[str, Any], Any, Set[int], Set[int]]] = []

        for (bar, beat, ts), data in event_items:
            chords = data.get("chords", set())
            basses = data.get("basses", set())
            event_notes_set = set(data.get("event_notes", set()))
            event_pitches_set = set(data.get("event_pitches", set()))
            chords_by_root: Dict[str, Any] = {}
            for chord in chords:
                root = next((n for n in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)) if chord.startswith(n)), None)
                if not root:
                    continue
                base_chord = chord.replace(root, 'C')
                # Extract chord quality (remove 'C' prefix)
                chord_quality = base_chord[1:] if base_chord.startswith('C') else base_chord
                priority_list = self.get_effective_priority_list()
                current_priority = priority_list.index(chord_quality) if chord_quality in priority_list else 999
                prev_chord = chords_by_root.get(root)
                if prev_chord:
                    prev_base = prev_chord.replace(root, 'C')
                    prev_quality = prev_base[1:] if prev_base.startswith('C') else prev_base
                    prev_priority = priority_list.index(prev_quality) if prev_quality in priority_list else 999
                    if current_priority < prev_priority:
                        chords_by_root[root] = chord
                else:
                    chords_by_root[root] = chord
            processed_events.append(((bar, beat, ts), chords_by_root, basses, event_notes_set, event_pitches_set))

        # Remove trivial duplicates of same single-root chord across adjacent events
        # Skip this deduplication for time-segment analysis to maintain segment independence
        if getattr(self, 'analysis_mode', 'event') == 'event':
            # If an event loses ALL its chords due to deduplication, remove the entire event
            i = 0
            while i < len(processed_events) - 1:
                (event1, chords1, basses1, notes1, pitches1) = processed_events[i]
                (event2, chords2, basses2, notes2, pitches2) = processed_events[i + 1]
                
                # Track original chord count before deduplication
                original_chord_count = len(chords2)
                
                common_roots = set(chords1.keys()) & set(chords2.keys())
                for root in list(common_roots):
                    if len(chords1) == 1 and len(chords2) == 1:
                        # keep earlier occurrence only
                        del chords2[root]
                
                # Only remove the entire event if:
                # 1. It originally HAD chords (not a legitimate non-drive event)
                # 2. AND all chords were removed by deduplication
                if original_chord_count > 0 and not chords2:
                    processed_events.pop(i + 1)
                    # Don't increment i since we removed an element
                else:
                    i += 1

        # Now collapse strictly identical consecutive chord-sets by unioning basses
        # Skip this collapsing for time-segment analysis to maintain segment independence
        if getattr(self, 'analysis_mode', 'event') == 'event':
            final_filtered_events: List[Tuple[Tuple[int,int,str], Dict[str, Any], Any]] = []
            prev_chords_set = None
            prev_bass_set = set()
            prev_event = None

            for event in processed_events:
                chords_set = set(event[1].values())
                bass_set = set(event[2])
                notes_set = set(event[3])
                pitches_set = set(event[4])
                if chords_set and prev_chords_set and chords_set == prev_chords_set:
                    combined_bass = prev_bass_set | bass_set
                    combined_notes = prev_notes_set | notes_set
                    combined_pitches = prev_pitches_set | pitches_set
                    # keep the original key but update chords and basses and note/pitch unions
                    prev_event = (prev_event[0], event[1], combined_bass, combined_notes, combined_pitches)
                    final_filtered_events[-1] = prev_event
                    prev_bass_set = combined_bass
                    prev_notes_set = combined_notes
                    prev_pitches_set = combined_pitches
                else:
                    prev_event = event
                    final_filtered_events.append(event)
                    prev_chords_set = chords_set
                    prev_bass_set = bass_set
                    prev_notes_set = notes_set
                    prev_pitches_set = pitches_set
        else:
            # For time-segment mode: keep all events as-is without any collapsing
            final_filtered_events: List[Tuple[Tuple[int,int,str], Dict[str, Any], Any]] = []
            for event in processed_events:
                final_filtered_events.append(event)

        # === PHASE 4: Event Merging and Post-Processing ===
        # Skip merging entirely for time-segment analysis mode
        if getattr(self, 'collapse_similar_events', False) and final_filtered_events and getattr(self, 'analysis_mode', 'event') == 'event':
            merged: List[Tuple[Tuple[int,int,str], Dict[str, Any], Any]] = []
            for ev in final_filtered_events:
                if not merged:
                    merged.append(ev)
                    continue
                prev = merged[-1]
                prev_roots = set(prev[1].keys())
                cur_roots = set(ev[1].keys())
                if not prev_roots and not cur_roots:
                    merged.append(ev)
                    continue
                union = prev_roots | cur_roots
                inter = prev_roots & cur_roots
                jaccard = (len(inter) / len(union)) if union else 0.0
                diff = len(union) - len(inter)

                # Bass overlap requirement: at least some shared bass or at least 30% overlap
                prev_basses = set(prev[2])
                cur_basses = set(ev[2])
                bass_union = prev_basses | cur_basses
                bass_inter = prev_basses & cur_basses
                bass_overlap = (len(bass_inter) / len(bass_union)) if bass_union else 0.0

                # Only consider merging if the events are close in time (same bar or adjacent)
                prev_bar = prev[0][0]
                cur_bar = ev[0][0]
                bar_dist = abs(cur_bar - prev_bar)

                # Stricter thresholds to avoid over-collapsing
                should_merge = False
                if bar_dist <= getattr(self, 'merge_bar_distance', MERGE_BAR_DISTANCE):
                    if jaccard >= getattr(self, 'merge_jaccard_threshold', MERGE_JACCARD_THRESHOLD):
                        should_merge = True
                    elif diff <= getattr(self, 'merge_diff_max', MERGE_DIFF_MAX) and bass_overlap >= getattr(self, 'merge_bass_overlap', MERGE_BASS_OVERLAP):
                        should_merge = True
                    elif prev_roots.issubset(cur_roots) and bass_overlap >= (getattr(self, 'merge_bass_overlap', MERGE_BASS_OVERLAP) * 0.6):
                        should_merge = True
                if should_merge:
                    merged_chords: Dict[str, str] = {}
                    for root in union:
                        candidates: List[str] = []
                        if prev[1].get(root):
                            if isinstance(prev[1][root], list):
                                candidates.extend(prev[1][root])
                            else:
                                candidates.append(prev[1][root])
                        if ev[1].get(root):
                            if isinstance(ev[1][root], list):
                                candidates.extend(ev[1][root])
                            else:
                                candidates.append(ev[1][root])
                        if candidates:
                            merged_chords[root] = min(candidates, key=chord_priority)
                    merged_basses = set(prev[2]) | set(ev[2])
                    # union event notes and pitches to avoid losing pitch data during merge
                    prev_notes = set(prev[3]) if len(prev) > 3 else set()
                    prev_pitches = set(prev[4]) if len(prev) > 4 else set()
                    cur_notes = set(ev[3]) if len(ev) > 3 else set()
                    cur_pitches = set(ev[4]) if len(ev) > 4 else set()
                    merged_notes = prev_notes | cur_notes
                    merged_pitches = prev_pitches | cur_pitches
                    merged[-1] = (prev[0], merged_chords, merged_basses, merged_notes, merged_pitches)
                else:
                    merged.append(ev)
            final_filtered_events = merged

        output_lines: List[str] = []
        filtered_events: Dict[Tuple[int,int,str], Dict[str, Any]] = {}
        for (bar, beat, ts), chords_by_root, basses, event_notes, event_pitches in final_filtered_events:
            deduped_chords_by_root = dedupe_chords_by_priority(chords_by_root)
            chords_sorted = sorted(deduped_chords_by_root.values())
            bass_sorted = sorted(basses, key=lambda b: NOTE_TO_SEMITONE.get(b, 99))
            bass_string = " + ".join(beautify_chord(b) for b in bass_sorted)
            # Use the unioned event_notes and event_pitches carried through merges
            event_notes = set(event_notes or [])
            event_pitches = set(event_pitches or [])
            chord_info: Dict[str, Dict[str, Any]] = {}
            for chord in chords_sorted:
                chord_info[chord] = {
                    "clean_stack": self._is_clean_stack(chord, event_notes),
                    "root_count": self._count_root_in_pitches(chord, event_pitches)
                }
            filtered_events[(bar, beat, ts)] = {
                "chords": set(chords_sorted),
                "basses": bass_sorted,
                "chord_info": chord_info
            }
        return output_lines, filtered_events

    def detect_chords(self, semitones, debug: bool = False):
        """
        Detect chord names from pitch classes using pattern matching.
        
        Tests all possible roots and matches against known chord patterns.
        Includes special handling for "no3" chords to verify third presence.
        """
        if len(semitones) < 3:
            return []

        chords_found = []
        semitone_list = sorted(set(semitones))

        # First pass: try candidate roots that are present in the set
        for root in sorted(set(semitones)):
            normalized = {(n - root) % 12 for n in semitones}
            # Also collect basses and event pitches if available
            # Try to get the full set of event pitches and basses from the calling context
            # If not available, fallback to semitones only
            event_pitches = set()
            event_basses = set()
            # Try to get from the caller if possible
            import inspect
            frame = inspect.currentframe()
            try:
                outer_locals = frame.f_back.f_locals
                event_pitches = set(outer_locals.get('test_pitches', []))
                event_basses = set(outer_locals.get('basses', []))
            except Exception:
                pass
            finally:
                del frame

            for name in self.get_effective_priority_list():
                # Convert chord quality back to full chord name for CHORDS lookup
                full_name = 'C' + name
                if full_name in TRIADS and not self.include_triads:
                    continue
                if full_name not in CHORDS:
                    continue
                chord_pattern = set(CHORDS[full_name])
                # Special handling for 'no3' chords: only match if third is truly absent
                if "no3" in name:
                    third_major = (root + 4) % 12
                    third_minor = (root + 3) % 12
                    third_present = False
                    # Check in semitones
                    if third_major in semitones or third_minor in semitones:
                        third_present = True
                    # Check in event_pitches
                    if any((p % 12 == third_major or p % 12 == third_minor) for p in event_pitches):
                        third_present = True
                    # Check in event_basses
                    if any((self.semitone_to_note(b) == self.semitone_to_note(third_major) or self.semitone_to_note(b) == self.semitone_to_note(third_minor)) for b in event_basses):
                        third_present = True
                    if third_present:
                        continue  # Third is present, skip 'no3' chord
                    if chord_pattern.issubset(normalized):
                        matched = full_name.replace('C', self.semitone_to_note(root))
                        chords_found.append(matched)
                        break
                else:
                    if chord_pattern.issubset(normalized):
                        matched = full_name.replace('C', self.semitone_to_note(root))
                        chords_found.append(matched)
                        break

        # Second pass: try "noroot" style chords where the root pitch-class is absent
        for root in sorted(set(range(12)) - set(semitones)):
            normalized = {(n - root) % 12 for n in semitones}
            for name in self.get_effective_priority_list():
                if "noroot" not in name:
                    continue
                # Convert chord quality back to full chord name for CHORDS lookup
                full_name = 'C' + name
                if full_name in TRIADS and not self.include_triads:
                    continue
                if full_name not in CHORDS:
                    continue
                chord_pattern = set(CHORDS[full_name])
                if chord_pattern == normalized:
                    matched = full_name.replace('C', self.semitone_to_note(root))
                    chords_found.append(matched)
                    break

        return chords_found

    def semitone_to_note(self, semitone):
        """Convert semitone number to note name, preferring natural notes."""
        # First try natural notes (single character)
        for note in NOTE_TO_SEMITONE:
            if NOTE_TO_SEMITONE[note] == semitone and len(note) == 1:
                return note
        # Fallback to any matching note
        return next((note for note, val in NOTE_TO_SEMITONE.items() if val == semitone), "C")
        
    def save_analysis_txt(self):
        if not self.analyzed_events:
            tk.messagebox.showwarning("No Data", "No analysis to save.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                for (bar, beat, ts), data in sorted(self.analyzed_events.items()):
                    chord_info = data.get("chord_info", {})
                    chords = []
                    for chord in sorted(data["chords"]):
                        marker = ""
                        if chord_info.get(chord, {}).get("clean_stack"):
                            marker += CLEAN_STACK_SYMBOL
                        root_count = chord_info.get(chord, {}).get("root_count", 1)
                        if root_count == 2:
                            marker += ROOT2_SYMBOL
                        elif root_count >= 3:
                            marker += ROOT3_SYMBOL
                        chords.append(f"{chord}{marker}")
                    chords_str = ",".join(chords)
                    bass = "+".join(data["basses"])
                    f.write(f"{bar}|{beat}|{ts}|{chords_str}|{bass}\n")
                # Add legend at the end of the file
                f.write(f"\nLegend: {CLEAN_STACK_SYMBOL}=Clean stack, {ROOT2_SYMBOL}=Root doubled, {ROOT3_SYMBOL}=Root tripled or more\n")
            tk.messagebox.showinfo("Saved", f"Analysis saved to {file_path}")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to save analysis:\n{e}")

    def load_analysis_txt(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            analyzed_events = {}
            with open(file_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or line.startswith("Legend:"):
                        continue
                    parts = line.split("|")
                    if len(parts) != 5:
                        continue
                    bar, beat, ts, chords_str, bass_str = parts
                    chords = set()
                    chord_info = {}
                    for chord_entry in chords_str.split(","):
                        clean_stack = CLEAN_STACK_SYMBOL in chord_entry
                        root2 = ROOT2_SYMBOL in chord_entry
                        root3 = ROOT3_SYMBOL in chord_entry
                        chord = chord_entry.replace(CLEAN_STACK_SYMBOL, "").replace(ROOT2_SYMBOL, "").replace(ROOT3_SYMBOL, "").strip()
                        # Skip empty or invalid chord names
                        if chord and chord != "":
                            chords.add(chord)
                        root_count = 1
                        if root3:
                            root_count = 3
                        elif root2:
                            root_count = 2
                        chord_info[chord] = {"clean_stack": clean_stack, "root_count": root_count}
                    basses = bass_str.split("+")
                    analyzed_events[(int(bar), int(beat), ts)] = {
                        "chords": chords,
                        "basses": basses,
                        "chord_info": chord_info
                    }
            

            
            self.analyzed_events = analyzed_events

            self.display_results()
            
            # Generate entropy analysis for loaded data (same as in run_analysis)
            from io import StringIO
            entropy_buf = StringIO()
            analyzer = EntropyAnalyzer(
                self.analyzed_events, 
                logger=lambda x: print(x, file=entropy_buf),
                strength_map=self.custom_strength_map,
                rule_params=self.custom_rule_params
            )
            analyzer.step_stage1_strengths(print_legend=True)
            self.entropy_review_text = entropy_buf.getvalue()
            
            self.show_grid_btn.config(state="normal")
            try:
                self.save_analysis_btn.config(state="normal")
            except Exception:
                pass
            tk.messagebox.showinfo("Loaded", f"Analysis loaded from {file_path}")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to load analysis:\n{e}")
            

    def get_deduplicated_events(self, events):
        """Apply the same deduplication logic used in display_results to any event dictionary."""
        from typing import List, Tuple, Dict, Any, Set
        
        # Use the dynamic priority list from GUI settings for chord deduplication

        def chord_priority(chord_name: str) -> int:
            base = chord_name
            for root in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)):
                if chord_name.startswith(root):
                    base = chord_name[len(root):]
                    break
            priority_list = self.get_effective_priority_list()
            return priority_list.index(base) if base in priority_list else 999

        event_items = sorted(events.items())
        processed_events: List[Tuple[Tuple[int,int,str], Dict[str, Any], Any, Set[int], Set[int]]] = []

        # Process events and deduplicate by priority
        for (bar, beat, ts), data in event_items:
            chords = data.get("chords", set())
            basses = data.get("basses", set())
            event_notes_set = set(data.get("event_notes", set()))
            event_pitches_set = set(data.get("event_pitches", set()))
            chords_by_root: Dict[str, Any] = {}
            for chord in chords:
                root = next((n for n in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)) if chord.startswith(n)), None)
                if not root:
                    continue
                base_chord = chord.replace(root, 'C')
                # Extract chord quality (remove 'C' prefix)
                chord_quality = base_chord[1:] if base_chord.startswith('C') else base_chord
                priority_list = self.get_effective_priority_list()
                current_priority = priority_list.index(chord_quality) if chord_quality in priority_list else 999
                prev_chord = chords_by_root.get(root)
                if prev_chord:
                    prev_base = prev_chord.replace(root, 'C')
                    prev_quality = prev_base[1:] if prev_base.startswith('C') else prev_base
                    prev_priority = priority_list.index(prev_quality) if prev_quality in priority_list else 999
                    if current_priority < prev_priority:
                        chords_by_root[root] = chord
                else:
                    chords_by_root[root] = chord
            processed_events.append(((bar, beat, ts), chords_by_root, basses, event_notes_set, event_pitches_set))

        # Remove trivial duplicates of same single-root chord across adjacent events
        i = 0
        while i < len(processed_events) - 1:
            (event1, chords1, basses1, notes1, pitches1) = processed_events[i]
            (event2, chords2, basses2, notes2, pitches2) = processed_events[i + 1]
            common_roots = set(chords1.keys()) & set(chords2.keys())
            for root in list(common_roots):
                if len(chords1) == 1 and len(chords2) == 1:
                    # keep earlier occurrence only
                    del chords2[root]
            i += 1

        # Now collapse strictly identical consecutive chord-sets by unioning basses
        final_filtered_events: List[Tuple[Tuple[int,int,str], Dict[str, Any], Any, Set[int], Set[int]]] = []
        prev_chords_set = None
        prev_bass_set = set()
        prev_notes_set = set()
        prev_pitches_set = set()
        prev_event = None

        for event in processed_events:
            chords_set = set(event[1].values())
            bass_set = set(event[2])
            notes_set = set(event[3])
            pitches_set = set(event[4])
            if chords_set and prev_chords_set and chords_set == prev_chords_set:
                combined_bass = prev_bass_set | bass_set
                combined_notes = prev_notes_set | notes_set
                combined_pitches = prev_pitches_set | pitches_set
                # keep the original key but update chords and basses and note/pitch unions
                prev_event = (prev_event[0], event[1], combined_bass, combined_notes, combined_pitches)
                final_filtered_events[-1] = prev_event
                prev_bass_set = combined_bass
                prev_notes_set = combined_notes
                prev_pitches_set = combined_pitches
            else:
                prev_event = event
                final_filtered_events.append(event)
                prev_chords_set = chords_set
                prev_bass_set = bass_set
                prev_notes_set = notes_set
                prev_pitches_set = pitches_set

        # Convert back to the original events format
        deduplicated_events = {}
        for (bar, beat, ts), chords_by_root, basses, event_notes_set, event_pitches_set in final_filtered_events:
            deduplicated_events[(bar, beat, ts)] = {
                "chords": set(chords_by_root.values()) if chords_by_root else set(),
                "basses": basses,
                "event_notes": event_notes_set,
                "event_pitches": event_pitches_set
            }
            
        return deduplicated_events



    def show_grid_window(self):
        if not self.processed_events:
            return
        # If already open, raise it
        try:
            if getattr(self, '_grid_window', None) and self._grid_window.winfo_exists():
                try:
                    self._grid_window.lift()
                except Exception:
                    pass
                return
        except Exception:
            pass

        # Use the EXACT same processed events that the list display shows
        grid_events = dict(self.processed_events)
        gw = GridWindow(self, grid_events, main_app=self)
        # Keep a reference so it can be refreshed when settings change
        self._grid_window = gw

        # Ensure we clear the reference when the window is closed
        try:
            def _on_grid_close():
                try:
                    gw.destroy()
                finally:
                    try:
                        self._grid_window = None
                    except Exception:
                        pass
            gw.protocol("WM_DELETE_WINDOW", _on_grid_close)
        except Exception:
            pass

class EmbeddedMidiKeyboard:
    """Embedded version of midiv3.py keyboard UI adapted to be created inside a Toplevel or Frame.

    Usage: EmbeddedMidiKeyboard(parent_toplevel, main_app=main_application_instance)
    """
    def __init__(self, parent, main_app=None):
        self.parent = parent
        self.main_app = main_app  # Reference to main application
        # Create UI inside the provided parent (Toplevel)
        self.parent.title("🎹 Embedded Keyboard")
        # Window-wide dark theme
        try:
            self.parent.configure(bg="black")
        except Exception:
            pass

        # Try to initialize pygame for audio output (optional)
        if PYGAME_AVAILABLE:
            try:
                # Initialize both mixer and midi
                pygame.mixer.pre_init(frequency=22050, size=-16, channels=2, buffer=1024)
                pygame.mixer.init()
                pygame.midi.init()
                self.pygame = pygame
                self.pygame_midi = pygame.midi
                
                # Initialize audio synthesis for MIDI playback
                self.active_notes = {}  # Track currently playing notes
                
                try:
                    default_out = pygame.midi.get_default_output_id()
                except Exception:
                    default_out = None
                if default_out is not None and default_out >= 0:
                    try:
                        self.midi_out = pygame.midi.Output(default_out)
                    except Exception:
                        self.midi_out = None
                else:
                    self.midi_out = None
            except Exception:
                self.pygame = None
                self.pygame_midi = None
                self.midi_out = None
                self.active_notes = {}
        else:
            self.pygame = None
            self.pygame_midi = None
            self.midi_out = None
            self.active_notes = {}

        # Minimal constants
        WHITE_KEYS = ['C', 'D', 'E', 'F', 'G', 'A', 'B']
        BLACK_KEYS = ['C#\nDb', 'D#\nEb', '', 'F#\nGb', 'G#\nAb', 'A#\nBb', '']

        self.selected_notes = set()
        self.include_triads_var = tk.BooleanVar(value=True)

        # Build a simple layout on the parent (dark background)
        self.canvas = tk.Canvas(self.parent, width=700, height=200, bg="black", highlightthickness=0)
        self.canvas.pack(pady=10)

        self.white_key_width = 60
        self.white_key_height = 200
        self.black_key_width = 40
        self.black_key_height = 120

        total_white_width = len(WHITE_KEYS) * self.white_key_width
        self.offset_x = (700 - total_white_width) // 2

        self.white_keys_rects = []
        self.black_keys_rects = []

        for i, note in enumerate(WHITE_KEYS):
            x = self.offset_x + i * self.white_key_width
            rect = self.canvas.create_rectangle(
                x, 0, x + self.white_key_width, self.white_key_height,
                fill='white', outline='#555555', width=2, tags=("white_key", note)
            )
            self.white_keys_rects.append(rect)

        for i, note in enumerate(BLACK_KEYS):
            if note != '':
                x = self.offset_x + (i + 1) * self.white_key_width - self.black_key_width / 2
                rect = self.canvas.create_rectangle(
                    x, 0, x + self.black_key_width, self.black_key_height,
                    fill='black', outline='#222', width=1
                )
                self.black_keys_rects.append(rect)
                enharmonics = note.split('\n')
                self.canvas.addtag_withtag("black_key", rect)
                for enh_note in enharmonics:
                    self.canvas.addtag_withtag(enh_note, rect)

        # Bindings
        self.canvas.tag_bind("white_key", "<Button-1>", self._on_key_click)
        self.canvas.tag_bind("black_key", "<Button-1>", self._on_key_click)

        # Controls
        controls_frame = tk.Frame(self.parent, bg="black")
        controls_frame.pack(pady=10)

        # Clean toggle button for triads - same style as Clear button
        def toggle_triads():
            self.include_triads_var.set(not self.include_triads_var.get())
            # Update button text to reflect new state
            new_text = "Include triads: ON" if self.include_triads_var.get() else "Include triads: OFF"
            self.triads_btn.config(text=new_text)
            self.analyze_chord()  # refresh analysis when toggled
        
        # Platform-friendly triads button - same approach as Clear button
        import platform
        if platform.system() == "Darwin":  # Mac
            triads_btn_kwargs = {"font": ("Segoe UI", 10), "padx": 12, "pady": 5}
        else:  # PC/Linux
            triads_btn_kwargs = {"font": ("Segoe UI", 10, "bold"), "bg": "#444444", "fg": "white", 
                               "activebackground": "#666666", "activeforeground": "white", 
                               "bd": 0, "padx": 12, "pady": 5}
            
        self.triads_btn = tk.Button(
            controls_frame, text="Include triads: ON" if self.include_triads_var.get() else "Include triads: OFF",
            command=toggle_triads, **triads_btn_kwargs
        )
        self.triads_btn.pack(side="left", padx=10, pady=2)

        # Platform-friendly clear button
        if platform.system() == "Darwin":  # Mac
            clear_btn_kwargs = {"font": ("Segoe UI", 11), "padx": 12, "pady": 5}
        else:  # PC/Linux
            clear_btn_kwargs = {"font": ("Segoe UI", 11, "bold"), "bg": "#444444", "fg": "white", 
                               "activebackground": "#666666", "activeforeground": "white", 
                               "bd": 0, "padx": 12, "pady": 5}
        
        self.clear_button = tk.Button(
            controls_frame, text="Clear", command=self._clear_selection, **clear_btn_kwargs
        )
        self.clear_button.pack(side="left", padx=10)

        # MIDI Dropdown (if mido available)
        midi_frame = tk.Frame(self.parent, bg="black")
        midi_frame.pack(pady=5)
        tk.Label(midi_frame, text="MIDI Input:", font=("Segoe UI", 10), fg="white", bg="black").pack(side="left", padx=(0,5))
        if MIDO_AVAILABLE:
            # Enhanced MIDI port detection for packaged executables
            # Small delay to allow MIDI system to initialize
            self.parent.after(100, self._delayed_midi_init)
            self.midi_ports = []  # Will be populated by delayed init
        else:
            self.midi_ports = []
        self.midi_port_var = tk.StringVar()
        self.midi_dropdown = ttk.Combobox(midi_frame, textvariable=self.midi_port_var, values=self.midi_ports, state="readonly", width=40)
        if self.midi_ports:
            self.midi_port_var.set(self.midi_ports[0])
        self.midi_dropdown.pack(side="left")
        
        # Add refresh button for MIDI ports
        refresh_btn = tk.Button(midi_frame, text="↻", command=self._refresh_midi_ports, 
                               font=("Segoe UI", 8), width=3, height=1)
        refresh_btn.pack(side="left", padx=(5,0))
        
        # Bind MIDI port change
        try:
            self.midi_dropdown.bind("<<ComboboxSelected>>", self._on_midi_port_change)
        except Exception:
            pass

        # Start listener if ports exist
        if self.midi_ports:
            try:
                self.start_midi_listener(port_name=self.midi_ports[0])
            except Exception:
                pass

        # Result label on dark background with white text so 'drives' are visible
        self.result_label = tk.Label(self.parent, text="", font=("Segoe UI", 12), fg="white", bg="black", justify='left', wraplength=600)
        self.result_label.pack(pady=(10,5))

        # Close handler
        def _on_close():
            try:
                if hasattr(self, 'midi_out') and self.midi_out:
                    try:
                        self.midi_out.close()
                    except Exception:
                        pass
                if getattr(self, 'pygame_midi', None):
                    try:
                        self.pygame_midi.quit()
                    except Exception:
                        pass
            finally:
                try:
                    self.parent.destroy()
                except Exception:
                    pass

        try:
            self.parent.protocol("WM_DELETE_WINDOW", _on_close)
        except Exception:
            pass

    # Minimal event handlers (adapted from original)
    def _on_key_click(self, event):
        clicked = self.canvas.find_withtag('current')
        if not clicked:
            return
        clicked_id = clicked[0]
        tags = self.canvas.gettags(clicked_id)
        note = next((tag for tag in tags if tag in NOTE_TO_SEMITONE), None)
        if note is None:
            return
        semitone = NOTE_TO_SEMITONE[note]
        if semitone in self.selected_notes:
            self.selected_notes.remove(semitone)
            self._set_key_color(note, False)
            self._stop_note(semitone)
            # Update analysis after mouse-driven removal
            try:
                self.analyze_chord()
            except Exception:
                pass
        else:
            if len(self.selected_notes) >= 10:
                messagebox.showinfo("Limit Reached", "Maximum 10 notes can be selected.")
                return
            self.selected_notes.add(semitone)
            self._set_key_color(note, True)
            self._play_note(semitone)
            # Update analysis after mouse-driven addition
            try:
                self.analyze_chord()
            except Exception:
                pass

    def _set_key_color(self, note, selected):
        fluorescent_pink = '#ff00ff'
        for rect in self.white_keys_rects + self.black_keys_rects:
            if note in self.canvas.gettags(rect):
                if selected:
                    # played notes fluorescent pink
                    color = fluorescent_pink
                else:
                    color = 'white' if note in ['C','D','E','F','G','A','B'] else 'black'
                self.canvas.itemconfig(rect, fill=color)
                break

    def _clear_selection(self):
        for semitone in list(self.selected_notes):
            self._stop_note(semitone)
        self.selected_notes.clear()
        for rect in self.white_keys_rects + self.black_keys_rects:
            note_tags = self.canvas.gettags(rect)
            if any(t in ['C','D','E','F','G','A','B'] for t in note_tags):
                self.canvas.itemconfig(rect, fill='white')
            else:
                self.canvas.itemconfig(rect, fill='black')
        self.result_label.config(text="")

    def semitone_to_note(self, semitone):
        for note, val in NOTE_TO_SEMITONE.items():
            if val == semitone and len(note) == 1:
                return note
        for note, val in NOTE_TO_SEMITONE.items():
            if val == semitone:
                return note
        return "C"

    def _generate_sine_wave(self, frequency, duration=0.5, volume=0.3):
        """Generate a sine wave for audio synthesis."""
        if not PYGAME_AVAILABLE:
            return None
        
        try:
            import numpy as np
            sample_rate = 22050
            frames = int(duration * sample_rate)
            arr = np.zeros((frames, 2))
            
            # Generate sine wave
            for i in range(frames):
                wave = volume * np.sin(2 * np.pi * frequency * i / sample_rate)
                arr[i][0] = wave  # Left channel
                arr[i][1] = wave  # Right channel
            
            # Convert to pygame Sound
            arr = (arr * 32767).astype(np.int16)
            return pygame.sndarray.make_sound(arr)
        except ImportError:
            # Fallback: numpy not available, no audio synthesis
            return None
        except Exception:
            return None

    def _midi_to_frequency(self, midi_note):
        """Convert MIDI note number to frequency in Hz."""
        return 440.0 * (2.0 ** ((midi_note - 69) / 12.0))

    def _play_note(self, semitone, velocity=127):
        note_num = 60 + semitone
        
        # Send MIDI data only (clean approach like midiv3)
        if getattr(self, 'midi_out', None):
            try:
                self.midi_out.note_on(note_num, velocity)
            except Exception:
                pass

    def _stop_note(self, semitone):
        note_num = 60 + semitone
        
        # Send MIDI stop only
        if getattr(self, 'midi_out', None):
            try:
                self.midi_out.note_off(note_num, 0)
            except Exception:
                pass

    # ----- MIDI input handling and chord analysis -----
    def _on_midi_port_change(self, event=None):
        selected_port = self.midi_port_var.get()
        if hasattr(self, 'midi_in') and getattr(self, 'midi_in'):
            try:
                self.midi_in.close()
            except Exception:
                pass
        self.start_midi_listener(port_name=selected_port)

    def _get_midi_ports_safe(self):
        """
        Simple MIDI port detection - matches working midiv3.py approach.
        """
        try:
            if MIDO_AVAILABLE:
                ports = mido.get_input_names()
                debug_log(f"MIDI Debug: Found {len(ports)} MIDI ports: {ports}")
                return ports
            else:
                debug_log("MIDI Debug: mido not available")
                return []
        except Exception as e:
            debug_log(f"MIDI Debug: Error getting MIDI ports: {e}")
            return []

    def _delayed_midi_init(self):
        """Simple MIDI initialization - matches working midiv3.py approach."""
        self.midi_ports = self._get_midi_ports_safe()
        self.midi_dropdown['values'] = self.midi_ports
        if self.midi_ports:
            self.midi_port_var.set(self.midi_ports[0])
            # Auto-start listener with first port
            self.start_midi_listener(port_name=self.midi_ports[0])

    def _refresh_midi_ports(self):
        """Refresh the MIDI port list - simple approach like midiv3.py."""
        if MIDO_AVAILABLE:
            self.midi_ports = self._get_midi_ports_safe()
            self.midi_dropdown['values'] = self.midi_ports
            if self.midi_ports:
                self.midi_port_var.set(self.midi_ports[0])
                self.start_midi_listener(port_name=self.midi_ports[0])
            else:
                self.midi_port_var.set("")

    def start_midi_listener(self, port_name=None):
        if not MIDO_AVAILABLE:
            print("mido not available: MIDI input disabled")
            return

        if not port_name:
            ports = self._get_midi_ports_safe()
            if not ports:
                print("No MIDI input ports found.")
                return
            port_name = ports[0]

        try:
            self.midi_in = mido.open_input(port_name)
        except Exception as e:
            print(f"Failed to open MIDI input port: {e}")
            return

        def midi_loop():
            try:
                for msg in self.midi_in:
                    try:
                        if msg.type in ('note_on', 'note_off'):
                            pitch_class = msg.note % 12
                            if msg.type == 'note_on' and getattr(msg, 'velocity', 0) > 0:
                                self.parent.after(0, lambda pc=pitch_class: self.add_midi_note(pc))
                            else:
                                self.parent.after(0, lambda pc=pitch_class: self.remove_midi_note(pc))
                    except Exception as e:
                        print(f"MIDI message processing error: {e}")
                        continue
            except Exception as e:
                print(f"MIDI loop error: {e}")

        threading.Thread(target=midi_loop, daemon=True).start()

    def add_midi_note(self, semitone):
        if semitone not in self.selected_notes:
            if len(self.selected_notes) >= 10:
                return
            self.selected_notes.add(semitone)
            note_name = self.semitone_to_note(semitone)
            self._set_key_color(note_name, True)
            self._play_note(semitone)
            self.analyze_chord()

    def remove_midi_note(self, semitone):
        if semitone in self.selected_notes:
            self.selected_notes.remove(semitone)
            note_name = self.semitone_to_note(semitone)
            self._set_key_color(note_name, False)
            self._stop_note(semitone)
            self.analyze_chord()

    def analyze_chord(self):
        """Analyze selected notes using the main application's drive detection."""
        try:
            # Check if we have enough notes
            if len(self.selected_notes) < 3:
                self.result_label.config(text="🎵 Select at least 3 notes to analyze.")
                return

            # Check if main app is available
            if not self.main_app:
                self.result_label.config(text="Main application not available for drive analysis.")
                return

            # Use the main app's detect_chords method for full drive analysis
            detected_chords = self.main_app.detect_chords(self.selected_notes)
            
            if detected_chords:
                lines = []
                for chord_str in detected_chords:
                    # Clean up chord name formatting
                    chord_str = chord_str.replace("noroot", "no root")
                    
                    # Handle diminished chord alternatives for noroot chords
                    if "no root" in chord_str and len(self.selected_notes) >= 4:
                        # Try to identify the root from the chord name
                        root_name = chord_str.split(" ")[0].rstrip("7m9no")
                        if root_name in NOTE_TO_SEMITONE:
                            semitone = NOTE_TO_SEMITONE[root_name]
                            dim_semitone = (semitone + 1) % 12
                            dim_root = self.semitone_to_note(dim_semitone)
                            dim_chord_label = f"{dim_root}o7"
                            chord_str += f" [{dim_chord_label}]"

                    # Handle enharmonic equivalents for certain roots
                    root_name = chord_str.split(" ")[0]
                    root_note = root_name.rstrip("#b")
                    if root_note in NOTE_TO_SEMITONE:
                        semitone = NOTE_TO_SEMITONE[root_note]
                        if semitone in (8, 1, 3, 6, 10):  # G#, C#, D#, F#, A#
                            enh = ENHARMONIC_EQUIVALENTS.get(semitone, root_note)
                            if isinstance(enh, str) and '/' in enh:
                                enh_roots = enh.split('/')
                                if len(enh_roots) == 2 and root_note == enh_roots[0]:
                                    chord_str += f" ({enh_roots[1]} root)"

                    lines.append(chord_str)

                # Display detected drives/chords
                display_text = "\n".join(lines)
                if "no root" in display_text or len(lines) > 1:
                    display_text = "Drives detected:\n" + display_text
                else:
                    display_text = "Drive: " + display_text
                    
                self.result_label.config(text=display_text)
            else:
                # No drives detected
                self.result_label.config(text="No known drives detected.")
                
        except Exception as e:
            print(f"Error in keyboard chord analysis: {e}")
            import traceback
            print(f"Full traceback: {traceback.format_exc()}")
            self.result_label.config(text="Error analyzing chord. Check console for details.")


class GridWindow(tk.Toplevel):
    """
    Visual timeline display of chord analysis results.
    
    Shows chord progressions in a grid format with color-coded chord strengths,
    entropy analysis, and interactive features for detailed analysis review.
    """
    
    
    SUPERSCRIPT_MAP = {
        'no1': "ⁿᵒ¹",
        'no3': "ⁿᵒ³",
        'no5': "ⁿᵒ⁵",
        'noroot': "ⁿᵒ¹",  # optional alias for clarity
    }

    CELL_SIZE = 50
    PADDING = 40

    # For on-screen (Tkinter) - System B: More subtle gradations
    STRENGTH_COLORS_TK = {
        "60+": "#000000",      # Black - strongest chords
        "50-59": "#2A2A2A",    # Very dark grey
        "40-49": "#444444",    # Dark grey
        "30-39": "#666666",    # Medium-dark grey
        "25-29": "#888888",    # Medium grey
        "20-24": "#AAAAAA",    # Medium-light grey
        "15-19": "#CCCCCC",    # Light grey
        "0-14": "#EEEEEE",     # Very light grey (not pure white)
    }

    # For PDF (ReportLab) - System B: More subtle gradations
    STRENGTH_COLORS_PDF = {
        k: HexColor(v) for k, v in STRENGTH_COLORS_TK.items()
    }

    def _dedupe_for_grid(self, raw_events: Dict[Tuple[int, int, str], Dict[str, Any]]) -> Dict[Tuple[int, int, str], Dict[str, Any]]:
        """Return events dict with immediate repeated patterns removed to match main display logic.
        Implements the same sliding-window dedupe algorithm used in MidiChordAnalyzer.display_results.
        """
        events_list = list(sorted(raw_events.items()))
        if not getattr(self.parent, 'remove_repeats', False):
            return dict(events_list)

        # Copy of the dedupe algorithm from display_results
        filtered = []
        i = 0
        n = len(events_list)
        while i < n:
            max_pat = (n - i) // 2
            found_repeat = False
            for pat_len in range(1, max_pat + 1):
                pat = [
                    (tuple(sorted(events_list[i + j][1].get('chords', []))), tuple(sorted(events_list[i + j][1].get('basses', []))))
                    for j in range(pat_len)
                ]
                repeat = True
                for j in range(pat_len):
                    if i + pat_len + j >= n:
                        repeat = False
                        break
                    sig1 = (
                        tuple(sorted(events_list[i + j][1].get('chords', []))),
                        tuple(sorted(events_list[i + j][1].get('basses', [])))
                    )
                    sig2 = (
                        tuple(sorted(events_list[i + pat_len + j][1].get('chords', []))),
                        tuple(sorted(events_list[i + pat_len + j][1].get('basses', [])))
                    )
                    if sig1 != sig2:
                        repeat = False
                        break
                if repeat:
                    # keep the first occurrence, then skip any number of consecutive repeats
                    jpos = i + pat_len
                    while jpos + pat_len <= n:
                        all_match = True
                        for k in range(pat_len):
                            sig_a = (
                                tuple(sorted(events_list[i + k][1].get('chords', []))),
                                tuple(sorted(events_list[i + k][1].get('basses', [])))
                            )
                            sig_b = (
                                tuple(sorted(events_list[jpos + k][1].get('chords', []))),
                                tuple(sorted(events_list[jpos + k][1].get('basses', [])))
                            )
                            if sig_a != sig_b:
                                all_match = False
                                break
                        if all_match:
                            jpos += pat_len
                        else:
                            break
                    filtered.extend(events_list[i:i+pat_len])
                    i = jpos
                    found_repeat = True
                    break
            if not found_repeat:
                filtered.append(events_list[i])
                i += 1

        return {k: v for k, v in filtered}

    def __init__(self, parent, events, main_app=None):
        super().__init__(parent)
        self.title("Chord Grid Visualization")
        self.configure(bg="white")
        self.main_app = main_app  # Store reference to main application
        
        # Configure white theme for GridWindow ttk widgets
        import platform
        from tkinter import ttk
        style = ttk.Style()
        style.configure("GridWindow.TFrame", background="white")
        
        # Platform-specific text color for checkboxes (white on Mac, black on PC)
        checkbox_fg = "white" if platform.system() == "Darwin" else "black"
        style.configure("GridWindow.TCheckbutton", background="white", foreground=checkbox_fg)
        style.configure("GridWindow.TButton", background="white", foreground="black")

        self.parent = parent
        
        # Store references to custom parameters from parent
        self.custom_strength_map = getattr(parent, 'custom_strength_map', None)
        self.custom_rule_params = getattr(parent, 'custom_rule_params', None)
        
        # Apply same filtering as main window (respect include_non_drive_events)
        raw_events = {k: v for k, v in events.items()} if events else {}
        if hasattr(parent, 'include_non_drive_events') and not parent.include_non_drive_events:
            raw_events = {k: v for k, v in raw_events.items() if v.get('chords') and len(v['chords']) > 0}



        # Events are already fully processed by the parent - use them directly
        self.events = raw_events
        self.sorted_events = sorted(self.events.keys())

        # Remove Gb row from the circle of fifths for this grid
        self.root_list = [r for r in CIRCLE_OF_FIFTHS_ROOTS if r != 'Gb']
        self.root_to_row = {root: i for i, root in enumerate(self.root_list)}

        canvas_width = self.PADDING * 2 + len(self.sorted_events) * self.CELL_SIZE
        canvas_height = self.PADDING * 2 + len(self.root_list) * self.CELL_SIZE

        # --- Controls frame ---
        controls_frame = ttk.Frame(self, style="GridWindow.TFrame")
        controls_frame.pack(side="top", fill="x", pady=5)

        # Create all controls and pack them in a row
        self.show_resolutions_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            controls_frame,
            text="Show Resolution Patterns",
            variable=self.show_resolutions_var,
            command=self.redraw,
        ).pack(side="left", padx=5)

        self.color_pdf_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            controls_frame,
            text="Color-code Chords",
            variable=self.color_pdf_var,
            command=self.redraw
        ).pack(side="left", padx=5)

        self.show_entropy_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            controls_frame,
            text="Show Entropy",
            variable=self.show_entropy_var,
            command=self.redraw_entropy
        ).pack(side="left", padx=5)
        
        ttk.Button(
            controls_frame,
            text="Export as PDF",
            command=self.export_pdf
        ).pack(side="left", padx=5)

        # Center the whole frame
        controls_frame.update_idletasks()
        frame_width = controls_frame.winfo_width()
        controls_frame.pack_configure(anchor="center")

        # --- Canvas container below controls ---
        container = ttk.Frame(self, style="GridWindow.TFrame")
        container.pack(fill="both", expand=True)

        # Left (frozen) column canvas for root labels (wider to fit enharmonic alternatives)
        left_col_width = max(100, self.PADDING * 3)
        self.left_canvas = tk.Canvas(container, width=left_col_width, height=canvas_height, bg="white", highlightthickness=0)
        self.left_canvas.pack(side="left", fill="y")

        # Right scrollable area for the grid
        right_frame = ttk.Frame(container, style="GridWindow.TFrame")
        right_frame.pack(side="left", fill="both", expand=True)

        # Canvas with dynamic width (wide for many columns), fixed height
        self.canvas = tk.Canvas(right_frame, width=min(canvas_width, 800), height=canvas_height, bg="white", highlightthickness=0)
        self.canvas.pack(side="top", fill="both", expand=True)

        # Horizontal scrollbar
        h_scroll = ttk.Scrollbar(right_frame, orient="horizontal", command=self.canvas.xview)
        h_scroll.pack(side="bottom", fill="x")

        self.canvas.configure(xscrollcommand=h_scroll.set)

        self.tooltip = tk.Label(self.canvas, bg="#008080", fg="white", font=("Segoe UI", 10), bd=1, relief="solid")
        self.tooltip.place_forget()
        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<Leave>", lambda e: self.tooltip.place_forget())

        self.chord_positions = []
        self.draw_grid()

        # Set scroll region after drawing
        self.canvas.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        # Populate frozen left column with enharmonic alternatives and labels
        enh_map = {
            'F#': 'F#/Gb',
            'Db': 'Db/C#',
            'Ab': 'Ab/G#',
            'Eb': 'Eb/D#'
        }
        # Load and display image labels for chord roots
        for root, row in self.root_to_row.items():
            y = self.PADDING + row * self.CELL_SIZE + self.CELL_SIZE // 2
            # Reverse image numbering to match flipped grid order
            # Row 0 (F#, now at top) gets image 12, Row 11 (Db, now at bottom) gets image 1
            image_number = len(self.root_list) - row
            
            try:
                # Load the corresponding numbered image (use os.path.join for cross-platform paths)
                image_path = resource_path(os.path.join("assets", "images", f"{image_number}.png"))
                img = Image.open(image_path)
                photo = ImageTk.PhotoImage(img)
                
                # Create image on canvas, centered in the left column
                x_center = left_col_width // 2
                self.left_canvas.create_image(x_center, y, image=photo, anchor='center')
                
                # Keep a reference to prevent garbage collection
                if not hasattr(self, '_label_images'):
                    self._label_images = []
                self._label_images.append(photo)
                
            except Exception as e:
                # Fallback to text if image loading fails
                print(f"[WARNING] Failed to load image {image_number}.png: {e}")

                try:
                    # Simple fallback text
                    fallback_text = root.replace('b', '♭').replace('#', '♯')
                    self.left_canvas.create_text(left_col_width - 8, y, text=fallback_text, anchor='e', font=("Segoe UI", 12), fill="black")
                except Exception as ex:
                    print(f"[ERROR] Failed to create fallback label for root {root}: {ex}")
        
        #Inside your GridWindow __init__ method or GUI setup:
 
    def toggle_entropy(self):
        if self.show_entropy_var.get():
            print("Entropy graph should appear here!")
        else:
            print("Entropy graph should be hidden!")    

    # Inside GridWindow

    def show_entropy_info_window(self, entropy_text):
        import platform
        info_win = tk.Toplevel(self)
        info_win.title("Entropy Review")
        info_win.configure(bg="white")
        
        # Use platform-appropriate monospace fonts for better column alignment
        if platform.system() == "Darwin":  # Mac
            mono_font = ("Monaco", 10)
        elif platform.system() == "Windows":
            mono_font = ("Consolas", 10)
        else:  # Linux
            mono_font = ("DejaVu Sans Mono", 10)
            
        # Create scrollable text frame
        text_frame = tk.Frame(info_win, bg="white")
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, wrap="none", bg="white", fg="black", font=mono_font)
        
        # Add scrollbars
        v_scroll = tk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        h_scroll = tk.Scrollbar(text_frame, orient="horizontal", command=text_widget.xview)
        text_widget.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # Pack scrollbars and text widget
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")
        text_widget.pack(side="left", fill="both", expand=True)
        
        text_widget.insert("1.0", entropy_text)
        text_widget.config(state="disabled")
        
        # Calculate optimal window size based on text content
        lines = entropy_text.split('\n')
        max_line_length = max(len(line) for line in lines) if lines else 50
        num_lines = len(lines)
        
        # Estimate character width and height based on font
        char_width = 7 if platform.system() == "Darwin" else 8  # Monaco vs Consolas
        char_height = 15
        
        # Calculate window dimensions with padding for scrollbars and margins
        content_width = max_line_length * char_width + 60  # +60 for scrollbar and margins
        content_height = min(num_lines * char_height + 100, 600)  # Cap at 600px height
        
        # Set minimum and maximum sizes
        window_width = max(400, min(content_width, 1400))  # Between 400-1400px
        window_height = max(300, content_height)
        
        info_win.geometry(f"{window_width}x{window_height}")
        
        # Center the window on screen
        info_win.update_idletasks()
        x = (info_win.winfo_screenwidth() - window_width) // 2
        y = (info_win.winfo_screenheight() - window_height) // 2
        info_win.geometry(f"{window_width}x{window_height}+{x}+{y}")
        def save_entropy_info():
            from tkinter import filedialog
            path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], title="Save Entropy Info")
            if path:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(entropy_text)
        # Platform-friendly button styling
        if platform.system() == "Darwin":  # Mac
            save_btn_kwargs = {}
        else:  # PC/Linux
            save_btn_kwargs = {"bg": "#ff00ff", "fg": "#fff", "font": ("Segoe UI", 10, "bold")}
        save_btn = tk.Button(info_win, text="Save", command=save_entropy_info, **save_btn_kwargs)
        save_btn.pack(pady=(0,10))


    def compute_entropy(self, event_key: Tuple[int, int, str]) -> float:
        """
        Compute weighted chord strength entropy for a given event.
        Uses the EntropyAnalyzer class for calculation.
        """
        payload = self.events.get(event_key, {})
        chords = payload.get("chords", [])

        if not chords:
            return 0.0

        analyzer = EntropyAnalyzer(
            {event_key: payload}, 
            base=2, 
            logger=lambda x: None,
            strength_map=self.custom_strength_map,
            rule_params=self.custom_rule_params
        )    
        scores = []
        for chord in chords:
            score = analyzer._compute_score(chord)
            if isinstance(score, tuple):
                for x in score:
                    if isinstance(x, (int, float)):
                        score = x
                        break
            scores.append(score)
        H = analyzer._weighted_entropy(scores, base=2)
        return H

    def redraw_entropy(self):
        # --- Entropy review info window logic ---
        if self.show_entropy_var.get():
            # Compose entropy review text (replace with your actual entropy review logic)
            entropy_review = self.main_app.entropy_review_text if (self.main_app and hasattr(self.main_app, 'entropy_review_text')) else "Entropy review information not available."
            self.show_entropy_info_window(entropy_review)
            # Enlarge window to accommodate entropy band below the grid
            try:
                if not hasattr(self, '_prev_geom'):
                    self._prev_geom = self.geometry()
                self.geometry('1200x900')
            except Exception:
                pass

        grid_rows = len(self.root_to_row)
        ENTROPY_OFFSET = 110   # space below grid
        ENTROPY_SCALE = 20    # pixels per entropy unit
        DOT_RADIUS = 3
        buffer = 10            # extra space so dots aren’t clipped

        # If entropy display is turned off, clear any stored points and remove drawing
        if not self.show_entropy_var.get():
            self.canvas.delete("entropy_graph")
            self.entropy_points = []
            # Restore previous geometry
            try:
                if hasattr(self, '_prev_geom') and self._prev_geom:
                    self.geometry(self._prev_geom)
            except Exception:
                pass
            return

        # Remove any previous entropy graph
        self.canvas.delete("entropy_graph")

        # Calculate axis positions
        axis_x = self.PADDING - 14  # a little to the left of the grid
        canvas_height = self.PADDING * 2 + grid_rows * self.CELL_SIZE + ENTROPY_OFFSET
        y_base = self.PADDING + grid_rows * self.CELL_SIZE + ENTROPY_OFFSET - buffer
        y_top = y_base - 4 * ENTROPY_SCALE

        # Calculate entropy points (store H so tooltips can show values)
        points = []
        for idx, event_key in enumerate(self.sorted_events):
            H = self.compute_entropy(event_key)
            x = self.PADDING + idx * self.CELL_SIZE + self.CELL_SIZE // 2
            y = y_base - H * ENTROPY_SCALE
            points.append((x, y, H))

        # Draw Y-axis
        self.canvas.create_line(axis_x, y_base, axis_x, y_top, fill="black", width=2, tags="entropy_graph")

        # --- Draw horizontal X-axis at entropy = 0 ---
        self.canvas.create_line(
            self.PADDING,
            y_base,
            self.PADDING + len(self.sorted_events) * self.CELL_SIZE,
            y_base,
            fill="black",
            width=1,
            dash=(2, 2),
            tags="entropy_graph"
        )

        # --- Draw Y-axis tick marks and labels (0..4) ---
        for H_val in range(5):
            y = y_base - H_val * ENTROPY_SCALE
            self.canvas.create_line(axis_x - 5, y, axis_x + 5, y, fill="black", tags="entropy_graph")
            self.canvas.create_text(axis_x - 10, y, text=f"{H_val}", anchor="e", font=("Segoe UI", 9), tags="entropy_graph")

        # --- Draw connecting lines for entropy points ---
        for i in range(len(points) - 1):
            x1, y1, _ = points[i]
            x2, y2, _ = points[i + 1]
            self.canvas.create_line(x1, y1, x2, y2, fill="red", width=2, tags="entropy_graph")

        # --- Draw dots ---
        for x, y, _ in points:
            self.canvas.create_oval(
                x - DOT_RADIUS, y - DOT_RADIUS, x + DOT_RADIUS, y + DOT_RADIUS,
                fill="red", outline="", tags="entropy_graph"
            )

        # Store points with entropy values for the tooltip handler
        self.entropy_points = points

        # --- Update scroll region ---
        scroll_width = self.PADDING + len(self.sorted_events) * self.CELL_SIZE
        self.canvas.config(scrollregion=(0, 0, scroll_width, canvas_height))


    def classify_chord_type(self, chord):
        """Classify chord type for shape determination (kept for triangle/circle shapes)."""
        chord = chord.replace("♭", "b").replace("♯", "#")
        if "no" in chord.lower():
            return "no"

        suffix = chord.lstrip("ABCDEFGabcdefg#b")

        if suffix == "":
            return "maj"  # C, D, E, F, G = major triad
        elif suffix == "m":
            return "min"  # Cm, Dm, etc.

        # Add more specific checks for known types
        elif "maj7" in suffix or "mMaj7" in suffix:
            return "maj7"
        elif "m7" in suffix:
            return "m7"
        elif "ø7" in suffix:
            return "ø7"
        elif "7" in suffix:
            return "7"
        elif "dim" in suffix or "o7" in suffix:
            return "dim"
        elif "aug" in suffix:
            return "aug"
        else:
            return "other"

    def get_chord_strength_category(self, chord, event_key):
        """Calculate chord strength percentage and return color category."""
        # Get the event data
        event_data = self.events.get(event_key, {})
        
        # Calculate chord strength using entropy analyzer
        analyzer = EntropyAnalyzer(
            {event_key: event_data}, 
            base=2, 
            logger=lambda x: None,
            strength_map=self.custom_strength_map,
            rule_params=self.custom_rule_params
        )
        
        # Get all chord strengths for this event to calculate probabilities
        chords = event_data.get("chords", [])
        basses = event_data.get("basses", [])
        
        if not chords:
            return "0-29"  # No chords means white
        
        # Calculate scores for all chords in this event
        chord_scores = []
        for c in chords:
            score, _ = analyzer._compute_score(c, basses, event_data)
            chord_scores.append((c, score))
        
        # Find the score for our specific chord
        target_score = None
        for c, score in chord_scores:
            if c == chord:
                target_score = score
                break
        
        if target_score is None:
            return "0-29"
        
        # Calculate total score for probability calculation
        total_score = sum(score for _, score in chord_scores)
        
        if total_score == 0:
            return "0-29"
        
        # Calculate probability percentage
        probability = (target_score / total_score) * 100
        
        # Return color category based on probability ranges (System B - 8 categories)
        if probability >= 60:
            return "60+"
        elif probability >= 50:
            return "50-59"
        elif probability >= 40:
            return "40-49"
        elif probability >= 30:
            return "30-39"
        elif probability >= 25:
            return "25-29"
        elif probability >= 20:
            return "20-24"
        elif probability >= 15:
            return "15-19"
        else:
            return "0-14"

    def export_pdf(self):
        use_color = self.color_pdf_var.get()
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.colors import black, HexColor
        from reportlab.pdfgen import canvas as pdf_canvas

        import os, math
        # Always use the bundled DejaVuSans.ttf from assets/fonts
        font_path = resource_path(os.path.join('assets', 'fonts', 'DejaVuSans.ttf'))
        try:
            if os.path.exists(font_path):
                pdfmetrics.registerFont(TTFont('DejaVuSans', font_path))
        except Exception as e:
            print(f"Warning: Could not register DejaVuSans font: {e}")
            # silent fallback to default fonts
            pass

        pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Visualization as PDF"
        )
        if not pdf_path:
            return

        try:
            c = pdf_canvas.Canvas(pdf_path, pagesize=landscape(A4))
            width, height = landscape(A4)

            # margins
            margin_left = 80
            margin_right = 20
            margin_y = 40
            EXTRA_COLS = 2

            grid_cols = len(self.sorted_events)
            grid_rows = len(self.root_list)
            if grid_rows == 0 or grid_cols == 0:
                messagebox.showinfo("Export", "Nothing to export.")
                return

            # entropy band vertical reservation
            show_entropy_pdf = self.show_entropy_var.get()
            ENTROPY_SCALE_PDF = 12.0
            ENTROPY_OFFSET_PDF = 16.0
            ENTROPY_TOP_PAD = 6.0
            ENTROPY_HEIGHT_PDF = 4 * ENTROPY_SCALE_PDF + ENTROPY_OFFSET_PDF + ENTROPY_TOP_PAD
            available_grid_height = (height - 2 * margin_y) - (ENTROPY_HEIGHT_PDF if show_entropy_pdf else 0)

            # Vertical-limited cell size
            vertical_cell = available_grid_height / grid_rows

            # Try to allow EXTRA_COLS by computing desired cols and horizontal-limited cell size
            # but we'll choose the final cell_size to be the smaller of the vertical and horizontal options
            # so content always fits horizontally.
            # Compute a tentative number of columns that would fit vertically-sized cells
            tentative_cols_fit = max(1, int((width - margin_left - margin_right) // vertical_cell))
            desired_cols = max(1, tentative_cols_fit + EXTRA_COLS)

            # Horizontal cell size if we try to fit desired_cols
            horizontal_cell_for_desired = (width - margin_left - margin_right) / desired_cols

            # Final cell size is the smaller of vertical_cell and horizontal_cell_for_desired
            cell_size = min(vertical_cell, horizontal_cell_for_desired)

            # Enforce a minimum cell size so shapes/labels remain readable
            MIN_CELL = 14.0
            if cell_size < MIN_CELL:
                cell_size = MIN_CELL

            # Now compute page columns based on final cell_size (guaranteed to fit horizontally)
            page_grid_cols = max(1, int((width - margin_left - margin_right) // cell_size))
            num_pages = (grid_cols + page_grid_cols - 1) // page_grid_cols

            radius = int(cell_size * 0.65 / 2)  # Reduced from 0.85 to make triangles smaller
            circle_radius = int(cell_size * 0.80 / 2)  # Larger radius for circles only in PDF

            entropies_all = {ek: self.compute_entropy(event_key=ek) for ek in self.sorted_events}

            for page in range(num_pages):
                start_col = page * page_grid_cols
                end_col = min(start_col + page_grid_cols, grid_cols)
                visible_events = self.sorted_events[start_col:end_col]
                visible_cols = len(visible_events)

                # Row labels + horizontal grid lines
                for root, row in self.root_to_row.items():
                    y_center = height - (margin_y + row * cell_size + cell_size / 2)
                    c.setFont("DejaVuSans", 12)
                    enh_map = {'F#': 'F#/Gb', 'Db': 'Db/C#', 'Ab': 'Ab/G#', 'Eb': 'Eb/D#'}
                    label_raw = enh_map.get(root, root)
                    note_label = label_raw.replace('b', '♭').replace('#', '♯')
                    c.drawRightString(margin_left - 8, y_center - 4, note_label)

                    y_line = height - (margin_y + row * cell_size)
                    c.setStrokeColor(HexColor("#dddddd"))
                    c.line(margin_left, y_line, margin_left + visible_cols * cell_size, y_line)

                # Column labels + vertical lines
                for col_idx, (bar, beat, ts) in enumerate(visible_events):
                    x = margin_left + col_idx * cell_size + cell_size / 2
                    label = f"{bar}.{beat}"
                    c.setFont("Helvetica", 10)
                    c.drawCentredString(x, height - (margin_y - 18), label)

                    x_line = margin_left + col_idx * cell_size
                    c.setStrokeColor(HexColor("#dddddd"))
                    c.line(x_line, height - margin_y, x_line, height - (margin_y + grid_rows * cell_size))

                # Optional resolution arrows (drawn after grid lines but before chord shapes)
                if self.show_resolutions_var.get():
                    pos_dict = {}
                    for col_idx, event_key in enumerate(visible_events):
                        event_data = self.events[event_key]
                        chords_by_root = {}
                        for chord in event_data.get("chords", []):
                            root = self.get_root(chord)
                            chords_by_root[root] = chord
                        for root, chord in chords_by_root.items():
                            if root not in self.root_to_row:
                                continue
                            row = self.root_to_row[root]
                            x = margin_left + col_idx * cell_size + cell_size / 2
                            y = height - (margin_y + row * cell_size + cell_size / 2)
                            pos_dict[(col_idx, row)] = (x, y, chord)

                    # Arrows start from grid center and appear behind chord shapes
                    end_offset = cell_size * 0.55  # Reduced from 0.75 to make arrows longer
                    for (col, row), (x1, y1, chord1) in pos_dict.items():
                        diag_pos = (col + 1, row + 1)
                        if diag_pos in pos_dict:
                            x2, y2, chord2 = pos_dict[diag_pos]
                            dx = x2 - x1
                            dy = y2 - y1
                            dist = (dx**2 + dy**2) ** 0.5
                            if dist == 0:
                                continue
                            dx_norm = dx / dist
                            dy_norm = dy / dist
                            # Start from grid center (not circle edge)
                            start_x = x1
                            start_y = y1
                            end_x = x2 - dx_norm * end_offset
                            end_y = y2 - dy_norm * end_offset
                            
                            # Draw arrow line
                            c.setStrokeColor(black)
                            c.setLineWidth(1.5)
                            c.setLineCap(1)
                            c.line(start_x, start_y, end_x, end_y)
                            
                            # Draw arrowhead
                            arrow_size = 6  # Increased to make PDF arrowheads more prominent
                            angle = math.atan2(dy_norm, dx_norm)
                            left_angle = angle + math.pi / 6
                            right_angle = angle - math.pi / 6
                            
                            tip_x = end_x
                            tip_y = end_y
                            left_x = tip_x - arrow_size * math.cos(left_angle)
                            left_y = tip_y - arrow_size * math.sin(left_angle)
                            right_x = tip_x - arrow_size * math.cos(right_angle)
                            right_y = tip_y - arrow_size * math.sin(right_angle)
                            
                            c.setFillColor(black)
                            c.setStrokeColor(black)
                            c.setLineWidth(0.5)
                            path = c.beginPath()
                            path.moveTo(tip_x, tip_y)
                            path.lineTo(left_x, left_y)
                            path.lineTo(right_x, right_y)
                            path.close()
                            c.drawPath(path, stroke=1, fill=1)

                # Chords
                for col_idx, event_key in enumerate(visible_events):
                    event_data = self.events[event_key]
                    chords_by_root = {}
                    for chord in event_data.get("chords", []):
                        root = self.get_root(chord)
                        chords_by_root[root] = chord

                    for root, chord in chords_by_root.items():
                        if root not in self.root_to_row:
                            continue
                        row = self.root_to_row[root]
                        x = margin_left + col_idx * cell_size + cell_size / 2
                        y = height - (margin_y + row * cell_size + cell_size / 2)

                        chord_type = self.classify_chord_type(chord)
                        strength_category = self.get_chord_strength_category(chord, event_key)
                        fill_color = self.STRENGTH_COLORS_PDF.get(strength_category, HexColor("#CCCCCC")) if use_color else HexColor("#FFFFFF")
                        c.setFillColor(fill_color)
                        c.setStrokeColor(black)

                        if chord_type == "maj":
                            path = c.beginPath()
                            path.moveTo(x, y + radius)
                            path.lineTo(x - radius, y - radius)
                            path.lineTo(x + radius, y - radius)
                            path.close()
                            c.drawPath(path, stroke=1, fill=1 if use_color else 0)
                        elif chord_type == "min":
                            path = c.beginPath()
                            path.moveTo(x, y - radius)
                            path.lineTo(x - radius, y + radius)
                            path.lineTo(x + radius, y + radius)
                            path.close()
                            c.drawPath(path, stroke=1, fill=1 if use_color else 0)
                        else:
                            c.circle(x, y, circle_radius, stroke=1, fill=1 if use_color else 0)

                        if chord_type not in ("maj", "min"):
                            function_label = chord[len(root):] or "–"
                            function_label_lower = function_label.lower()
                            replaced = False
                            for key, superscript in self.SUPERSCRIPT_MAP.items():
                                if key in function_label_lower:
                                    function_label = function_label_lower.replace(key, superscript)
                                    replaced = True
                                    break
                            if not replaced:
                                function_label = beautify_chord(function_label)
                            # Use white text on dark backgrounds, black text on light backgrounds
                            # For System B: white text on the darkest 4 categories, black text on the lighter 4
                            text_color = HexColor("#FFFFFF") if strength_category in ["60+", "50-59", "40-49", "30-39"] else HexColor("#000000")
                            c.setFillColor(text_color)
                            c.setFont("DejaVuSans", 8)
                            c.drawCentredString(x, y - 4, function_label)





                c.setStrokeColor(HexColor("#dddddd"))
                c.setLineWidth(1)
                c.rect(
                    margin_left,
                    height - margin_y - grid_rows * cell_size,
                    visible_cols * cell_size,
                    grid_rows * cell_size,
                    stroke=1,
                    fill=0
                )

                # Draw bass dots AFTER grid lines to ensure they appear on top
                for col_idx, event_key in enumerate(visible_events):
                    event_data = self.events[event_key]
                    for bass in event_data.get("basses", []):
                        bass_root = self.get_root(bass)
                        if bass_root not in self.root_to_row:
                            continue
                        brow = self.root_to_row[bass_root]
                        bx = margin_left + col_idx * cell_size + cell_size / 2
                        by = height - (margin_y + brow * cell_size + cell_size / 2)
                        dot_radius = 2.5
                        
                        # Determine which radius to use based on chord type at this position
                        # Check if there's a chord at this position to determine shape size
                        chords_at_position = [chord for chord in event_data.get("chords", []) 
                                            if self.get_root(chord) == bass_root]
                        if chords_at_position:
                            chord = chords_at_position[0]
                            chord_type = self.classify_chord_type(chord)
                            shape_radius = circle_radius if chord_type not in ("maj", "min") else radius
                        else:
                            shape_radius = radius  # Default to triangle radius if no chord
                        
                        # Position dot at bottom edge of shape, matching tkinter positioning
                        # PDF coordinates: Y increases upward, tkinter increases downward
                        # In tkinter: by + radius places dot at bottom of shape
                        # In PDF: by - radius places dot at bottom of shape
                        dot_y_position = by - shape_radius
                        c.setFillColor(black)
                        c.circle(bx, dot_y_position, dot_radius, fill=1, stroke=0)

                # PDF entropy band
                if show_entropy_pdf:
                    y_base = margin_y + ENTROPY_OFFSET_PDF
                    axis_x = margin_left - 14

                    c.setStrokeColor(black)
                    c.setLineWidth(1)
                    c.line(axis_x, y_base, axis_x, y_base + 4 * ENTROPY_SCALE_PDF)

                    c.setFont("Helvetica", 8)
                    for h in range(5):
                        y_tick = y_base + h * ENTROPY_SCALE_PDF
                        c.line(axis_x - 3, y_tick, axis_x + 3, y_tick)
                        c.drawRightString(axis_x - 6, y_tick - 2, f"{h}")

                    c.setDash(2, 2)
                    c.line(margin_left, y_base, margin_left + visible_cols * cell_size, y_base)
                    c.setDash()

                    pts = []
                    for col_idx, ek in enumerate(visible_events):
                        H = float(entropies_all.get(ek, 0.0))
                        x = margin_left + col_idx * cell_size + cell_size / 2
                        y = y_base + H * ENTROPY_SCALE_PDF
                        pts.append((x, y))

                    c.setStrokeColor(HexColor("#cc0000"))
                    c.setLineWidth(1.5)
                    for i in range(len(pts) - 1):
                        x1, y1 = pts[i]
                        x2, y2 = pts[i + 1]
                        c.line(x1, y1, x2, y2)

                    c.setFillColor(HexColor("#cc0000"))
                    dot_r = 1.8
                    for x, y in pts:
                        c.circle(x, y, dot_r, fill=1, stroke=0)

                c.setFont("Helvetica", 9)
                c.setFillColor(black)
                
                # Draw page number in center
                c.drawCentredString(width / 2, 20, f"Page {page + 1} of {num_pages}")
                
                # Draw filename on the left (if available)
                if self.main_app and hasattr(self.main_app, 'loaded_file_path') and self.main_app.loaded_file_path:
                    import os
                    filename = os.path.basename(self.main_app.loaded_file_path)
                    c.drawString(30, 20, filename)

                c.showPage()

            c.save()
            messagebox.showinfo("Export Successful", f"PDF saved to:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export PDF:\n{e}")
            
    def redraw(self):
        self.canvas.delete("all")
        self.draw_grid()
        self.canvas.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def draw_grid(self):
        radius = int(self.CELL_SIZE * 0.65 / 2)  # Reduced from 0.85 to make triangles smaller

        def beautify_note_name(note):
            return note.replace("b", "♭").replace("#", "♯")

        # Draw horizontal grid lines (row labels live in the frozen left column)
        for root, row in self.root_to_row.items():
            y = self.PADDING + row * self.CELL_SIZE + self.CELL_SIZE // 2
            y_line = self.PADDING + row * self.CELL_SIZE
            self.canvas.create_line(self.PADDING, y_line, self.PADDING + len(self.sorted_events) * self.CELL_SIZE, y_line, fill="#ddd")

        # Draw column labels and vertical grid lines
        for col, (bar, beat, ts) in enumerate(self.sorted_events):
            x = self.PADDING + col * self.CELL_SIZE + self.CELL_SIZE // 2
            label = f"{bar}.{beat}"
            self.canvas.create_text(x, self.PADDING - 20, text=label, anchor="center", font=("Segoe UI", 12))
            x_line = self.PADDING + col * self.CELL_SIZE
            self.canvas.create_line(x_line, self.PADDING, x_line, self.PADDING + len(self.root_list) * self.CELL_SIZE, fill="#ddd")

        # Draw resolution arrows AFTER grid lines but BEFORE chord shapes (so arrows appear behind shapes)
        if self.show_resolutions_var.get():
            # First, collect all chord positions
            chord_positions = []
            for col, event_key in enumerate(self.sorted_events):
                event_data = self.events[event_key]
                chords = event_data.get("chords", set())
                chords_by_root = {}
                for chord in chords:
                    root = self.get_root(chord)
                    chords_by_root[root] = chord
                
                for root, chord in chords_by_root.items():
                    if root not in self.root_to_row:
                        continue
                    row = self.root_to_row[root]
                    x = self.PADDING + col * self.CELL_SIZE + self.CELL_SIZE // 2
                    y = self.PADDING + row * self.CELL_SIZE + self.CELL_SIZE // 2
                    chord_positions.append((col, row, x, y, chord))
            
            # Draw arrows from grid centers (will be hidden behind chord shapes)
            pos_dict = {(col, row): (x, y, chord) for col, row, x, y, chord in chord_positions}
            end_offset = self.CELL_SIZE * 0.55  # Reduced from 0.75 to make arrows longer
            for (col, row), (x1, y1, chord1) in pos_dict.items():
                diag_pos = (col + 1, row + 1)
                if diag_pos in pos_dict:
                    x2, y2, chord2 = pos_dict[diag_pos]
                    dx = x2 - x1
                    dy = y2 - y1
                    dist = (dx**2 + dy**2) ** 0.5
                    if dist == 0:
                        continue
                    dx_norm = dx / dist
                    dy_norm = dy / dist
                    # Start from grid center (not circle edge)
                    start_x = x1
                    start_y = y1
                    end_x = x2 - dx_norm * end_offset
                    end_y = y2 - dy_norm * end_offset
                    self.canvas.create_line(start_x, start_y, end_x, end_y, arrow=tk.LAST, fill="black", width=3)

        self.chord_positions.clear()

        # Draw chords as circles/triangles
        for col, event_key in enumerate(self.sorted_events):
            event_data = self.events[event_key]
            chords = event_data.get("chords", set())
            bass_notes = event_data.get("basses", [])

            chords_by_root = {}
            for chord in chords:
                root = self.get_root(chord)
                chords_by_root[root] = chord

            # Draw chord shapes (if any chords)
            if chords_by_root:
                for root, chord in chords_by_root.items():
                    if root not in self.root_to_row:
                        continue
                    row = self.root_to_row[root]
                    x = self.PADDING + col * self.CELL_SIZE + self.CELL_SIZE // 2
                    y = self.PADDING + row * self.CELL_SIZE + self.CELL_SIZE // 2

                    # Get chord type for shape determination
                    chord_type = self.classify_chord_type(chord)
                    # Get strength category for color determination  
                    strength_category = self.get_chord_strength_category(chord, event_key)
                    fill_color = self.STRENGTH_COLORS_TK.get(strength_category, "#CCCCCC") if self.color_pdf_var.get() else "white"

                    if chord_type == "maj":
                        # Upward pointing triangle
                        points = [
                            x, y - radius,  # top vertex
                            x - radius, y + radius,  # bottom-left vertex
                            x + radius, y + radius,  # bottom-right vertex
                        ]
                        self.canvas.create_polygon(points, fill=fill_color, outline="black")
                    elif chord_type == "min":
                        # Downward pointing triangle
                        points = [
                            x, y + radius,  # bottom vertex
                            x - radius, y - radius,  # top-left vertex
                            x + radius, y - radius,  # top-right vertex
                        ]
                        self.canvas.create_polygon(points, fill=fill_color, outline="black")
                    else:
                        self.canvas.create_oval(x - radius, y - radius, x + radius, y + radius, fill=fill_color, outline="black")

                    self.chord_positions.append((col, row, x, y, chord))
            # If no chords, leave column blank but show bass dots below

        # ALWAYS draw bass dots for each column based on event_data["basses"]
        for col, event_key in enumerate(self.sorted_events):
            event_data = self.events[event_key]
            bass_notes = event_data.get("basses", [])
            for bass in bass_notes:
                bass_root = self.get_root(bass)
                if bass_root not in self.root_to_row:
                    continue
                row = self.root_to_row[bass_root]
                bx = self.PADDING + col * self.CELL_SIZE + self.CELL_SIZE // 2
                by = self.PADDING + row * self.CELL_SIZE + self.CELL_SIZE // 2
                dot_radius = 4
                self.canvas.create_oval(
                    bx - dot_radius, by + radius - dot_radius,
                    bx + dot_radius, by + radius + dot_radius,
                    fill="black", outline=""
                )



        # Thin outer boundary matching grid lines
        self.canvas.create_line(
            self.PADDING, self.PADDING,
            self.PADDING + len(self.sorted_events) * self.CELL_SIZE, self.PADDING,
            fill="#ddd", width=1
        )
        self.canvas.create_line(
            self.PADDING, self.PADDING + len(self.root_list) * self.CELL_SIZE,
            self.PADDING + len(self.sorted_events) * self.CELL_SIZE, self.PADDING + len(self.root_list) * self.CELL_SIZE,
            fill="#ddd", width=1
        )
        self.canvas.create_line(
            self.PADDING, self.PADDING,
            self.PADDING, self.PADDING + len(self.root_list) * self.CELL_SIZE,
            fill="#ddd", width=1
        )
        self.canvas.create_line(
            self.PADDING + len(self.sorted_events) * self.CELL_SIZE, self.PADDING,
            self.PADDING + len(self.sorted_events) * self.CELL_SIZE, self.PADDING + len(self.root_list) * self.CELL_SIZE,
            fill="#ddd", width=1
        )


    def get_root(self, chord_name):
        for note in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)):
            if chord_name.startswith(note):
                canonical = ENHARMONIC_EQUIVALENTS.get(note, note)
                return canonical
        return None

    def on_mouse_move(self, event):
        # Adjust for canvas scroll offset
        mx = self.canvas.canvasx(event.x)
        my = self.canvas.canvasy(event.y)
        hover_radius = max(30, self.CELL_SIZE // 2)  # Larger radius for easier hit detection
        closest = None
        tooltip_text = None

        # --- Check chord hover (existing logic) ---
        for col, row, x, y, chord in self.chord_positions:
            if abs(mx - x) < hover_radius and abs(my - y) < hover_radius:
                dist = ((mx - x) ** 2 + (my - y) ** 2) ** 0.5
                if dist < hover_radius:
                    closest = (x, y)
                    tooltip_text = beautify_chord(chord)
                    break

        # --- If no chord found, check entropy hover ---
        if tooltip_text is None and hasattr(self, "entropy_points"):
            hover_radius = 6  # tighter tolerance for entropy dots
            for x, y, H in self.entropy_points:
                if abs(mx - x) < hover_radius and abs(my - y) < hover_radius:
                    dist = ((mx - x) ** 2 + (my - y) ** 2) ** 0.5
                    if dist < hover_radius:
                        closest = (x, y)
                        tooltip_text = f"H = {H:.3f}"
                        break

        # --- Show tooltip if something is hovered ---
        if closest and tooltip_text:
            # Place tooltip at mouse position relative to the canvas widget (not scrolled area)
            self.tooltip.config(text=tooltip_text)
            self.tooltip.place(x=event.x + 10, y=event.y - 10)
        else:
            self.tooltip.place_forget()
            
from typing import List, Tuple, Dict, Any, Optional, Callable, Set
from collections import Counter
from math import log2

class DriveStrengthParametersDialog:
    """Dialog for configuring drive strength parameters and rule bonuses."""
    
    # Default values
    DEFAULT_STRENGTH_MAP = {
        "7": 100, "7b5": 90, "7#5": 80, "m7": 70, "ø7": 65,
        "7m9noroot": 65, "7no3": 55, "7no5": 55, "7noroot": 50,
        "aug": 40, "": 42, "m": 35, "maj7": 30, "mMaj7": 25
    }
    
    DEFAULT_RULE_PARAMS = {
        "rule1_bass_support": 20,
        "rule2_tonic_dominant": 50,
        "rule2_selected_tonic": "No Tonic",  # Default: disabled
        "rule3_root_repetition": 2,
        "rule4_resolution_max": 10,
        "rule5_clean_voicing": 10,
        "rule6_same_chord": 5,
        "rule6_dominant_prep": 10,
        "rule7_root_doubled": 5,
        "rule7_root_tripled": 10
    }
    
    def __init__(self, parent, current_strength_map=None, current_rule_params=None):
        self.parent = parent
        self.result = None
        self.new_strength_map = None
        self.new_rule_params = None
        
        # Initialize with current or default values
        self.strength_map = current_strength_map or self.DEFAULT_STRENGTH_MAP.copy()
        self.rule_params = current_rule_params or self.DEFAULT_RULE_PARAMS.copy()
        
        # Create dialog window
        self.window = tk.Toplevel(parent)
        self.window.title("Drive Strength Configuration")
        self.window.withdraw()  # Hide initially to prevent flash
        self.window.geometry("600x580")
        self.window.resizable(True, True)
        self.window.transient(parent)
        self.window.grab_set()
        self.window.configure(bg="#f5f5f5")  # Set light gray background
        
        # Configure styles to match main settings dialog
        import platform
        from tkinter import ttk
        style = ttk.Style()
        style.configure("Dialog.TFrame", background="#f5f5f5")
        style.configure("Dialog.TLabel", background="#f5f5f5", foreground="black", font=("Segoe UI", 9))
        style.configure("Dialog.TNotebook", background="#f5f5f5", borderwidth=0)
        style.configure("Dialog.TNotebook.Tab", 
                       background="#e0e0e0", 
                       foreground="black", 
                       padding=[12, 8],
                       font=("Segoe UI", 9, "bold"))
        style.map("Dialog.TNotebook.Tab",
                 background=[("selected", "#f5f5f5"), ("active", "#d0d0d0")])
        style.configure("Dialog.TButton", background="#e0e0e0", foreground="black", font=("Segoe UI", 9))
        style.configure("Dialog.TEntry", background="white", foreground="black", font=("Segoe UI", 9))
        style.configure("Dialog.TCombobox", 
                       fieldbackground="white", 
                       background="#e0e0e0", 
                       foreground="black",
                       arrowcolor="black",
                       font=("Segoe UI", 9))
        
        # Configure LabelFrame style (needed for rules tab)
        style.configure("Dialog.TLabelFrame", background="#f5f5f5", foreground="black", font=("Segoe UI", 9))
        style.configure("Dialog.TLabelFrame.Label", background="#f5f5f5", foreground="black", font=("Segoe UI", 9, "bold"))
        
        # Center the window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.window.winfo_screenheight() // 2) - (580 // 2)
        self.window.geometry(f"600x580+{x}+{y}")
        
        # Variables for UI
        self.strength_vars = {}
        self.rule_vars = {}
        
        self.setup_ui()
        
        # Show the window now that everything is set up
        self.window.deiconify()
        
        # Bind close event
        self.window.protocol("WM_DELETE_WINDOW", self.cancel)
    
    def setup_ui(self):
        """Create the complete UI for the dialog."""
        # Main frame with scrollable content
        main_frame = ttk.Frame(self.window, style="Dialog.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create custom tab buttons with clean white styling to match main GUI
        tab_button_kwargs = {"bg": "#ffffff", "fg": "#000000", "activebackground": "#f0f0f0", 
                           "activeforeground": "#000000", "relief": "raised", "bd": 2, 
                           "font": ("Segoe UI", 10), "width": 20, 
                           "highlightbackground": "#b0b0b0", "highlightcolor": "#b0b0b0"}
        
        # Tab button frame
        tab_frame = tk.Frame(main_frame, bg="#f5f5f5")
        tab_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Tab buttons
        self.chord_tab_btn = tk.Button(tab_frame, text="Chord Base Strengths", 
                                      command=lambda: self.switch_tab(0), **tab_button_kwargs)
        self.chord_tab_btn.pack(side=tk.LEFT, padx=(0, 2))
        
        self.rules_tab_btn = tk.Button(tab_frame, text="Factor Bonuses", 
                                      command=lambda: self.switch_tab(1), **tab_button_kwargs)
        self.rules_tab_btn.pack(side=tk.LEFT)
        
        # Content frame for tab panels
        content_frame = tk.Frame(main_frame, bg="#f5f5f5", relief="sunken", bd=1)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Chord Strengths Tab
        self.strength_frame = ttk.Frame(content_frame, style="Dialog.TFrame")
        self.strength_frame.pack(fill=tk.BOTH, expand=True)
        
        # Rule Parameters Tab  
        self.rules_frame = ttk.Frame(content_frame, style="Dialog.TFrame")
        
        # Track current tab
        self.current_tab = 0
        
        self.setup_strength_tab(self.strength_frame)
        self.setup_rules_tab(self.rules_frame)

        # Load current values after both tabs are set up
        self.load_current_values()
        
        # Initialize first tab
        self.switch_tab(0)
        
        # Button frame
        button_frame = ttk.Frame(main_frame, style="Dialog.TFrame")
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Reset Defaults", command=self.reset_defaults, style="Dialog.TButton").pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Save Preset...", command=self.save_preset, style="Dialog.TButton").pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(button_frame, text="Load Preset...", command=self.load_preset, style="Dialog.TButton").pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self.cancel, style="Dialog.TButton").pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Apply", command=self.apply, style="Dialog.TButton").pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="OK", command=self.ok, style="Dialog.TButton").pack(side=tk.RIGHT, padx=(5, 0))
    
    def switch_tab(self, tab_index):
        """Switch between tabs using custom button styling."""
        # Hide all frames
        self.strength_frame.pack_forget()
        self.rules_frame.pack_forget()
        
        # Update button appearances
        unselected_style = {"bg": "#d0d0d0", "fg": "black", "relief": "raised", "bd": 1, 
                           "highlightbackground": "#b0b0b0"}
        selected_style = {"bg": "#808080", "fg": "white", "relief": "sunken", "bd": 1, 
                         "highlightbackground": "#b0b0b0"}
        
        if tab_index == 0:
            self.strength_frame.pack(fill=tk.BOTH, expand=True)
            self.chord_tab_btn.config(**selected_style)
            self.rules_tab_btn.config(**unselected_style)
        else:
            self.rules_frame.pack(fill=tk.BOTH, expand=True)
            self.rules_tab_btn.config(**selected_style)
            self.chord_tab_btn.config(**unselected_style)
        
        self.current_tab = tab_index
    
    def setup_strength_tab(self, parent):
        """Setup the chord strength configuration tab."""
        # Scrollable frame
        canvas = tk.Canvas(parent, bg="#f5f5f5")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Dialog.TFrame")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Header
        header_frame = ttk.Frame(scrollable_frame, style="Dialog.TFrame")
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(header_frame, text="Chord Base Strengths", font=("Segoe UI", 12, "bold"), style="Dialog.TLabel").pack()
        ttk.Label(header_frame, text="Configure the base strength points for each chord type (Range: 0-100)", 
                 font=("Segoe UI", 9), style="Dialog.TLabel").pack(pady=(5, 0))
        ttk.Label(header_frame, text="A score of 0 will remove this drive from the search", 
                 font=("Segoe UI", 9, "italic"), foreground="#808080", style="Dialog.TLabel").pack(pady=(2, 0))
        
        # Chord strength entries with proper centering
        strength_entries_frame = ttk.Frame(scrollable_frame, style="Dialog.TFrame")
        strength_entries_frame.pack(expand=True, padx=80, pady=10)
        
        # Define chord descriptions
        chord_descriptions = {
            "7": "Dominant 7th",
            "7b5": "Dominant 7th ♭5",
            "7#5": "Dominant 7th ♯5",
            "m7": "Minor 7th",
            "ø7": "Half-diminished 7th",
            "7m9noroot": "Minor 9th (no root)",
            "7no3": "Dominant 7th (no 3rd)",
            "7no5": "Dominant 7th (no 5th)",
            "7noroot": "Dominant 7th (no root)",
            "aug": "Augmented triad",
            "": "Major triad",
            "m": "Minor triad",
            "maj7": "Major 7th",
            "mMaj7": "Minor-major 7th"
        }
        
        row = 0
        for chord_symbol, description in chord_descriptions.items():
            frame = ttk.Frame(strength_entries_frame, style="Dialog.TFrame")
            frame.pack(fill=tk.X, pady=2)
            
            # Center the content using a container frame
            center_frame = ttk.Frame(frame, style="Dialog.TFrame")
            center_frame.pack(expand=True, fill=tk.X)
            
            # Label (centered with proper width)
            label_text = f"{description} ({chord_symbol if chord_symbol else 'major'}):"
            ttk.Label(center_frame, text=label_text, width=35, anchor="center", style="Dialog.TLabel").pack(side=tk.LEFT, expand=True, fill=tk.X)
            
            # Entry with validation
            var = tk.StringVar()
            self.strength_vars[chord_symbol] = var
            
            entry = ttk.Entry(center_frame, textvariable=var, width=10, justify="center", style="Dialog.TEntry")
            entry.pack(side=tk.LEFT, padx=(10, 5))
            
            # Range label
            ttk.Label(center_frame, text="(0-100)", font=("Segoe UI", 8), foreground="gray", style="Dialog.TLabel").pack(side=tk.LEFT, padx=(5, 0))
            
            row += 1
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def setup_rules_tab(self, parent):
        """Setup the rule parameters configuration tab."""
        # Scrollable frame
        canvas = tk.Canvas(parent, bg="#f5f5f5")
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style="Dialog.TFrame")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Header
        header_frame = ttk.Frame(scrollable_frame, style="Dialog.TFrame")
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(header_frame, text="Factor Bonus Parameters", font=("Segoe UI", 12, "bold"), style="Dialog.TLabel").pack()
        
        # Rules entries with proper centering
        rules_entries_frame = ttk.Frame(scrollable_frame, style="Dialog.TFrame")
        rules_entries_frame.pack(expand=True, padx=80, pady=10)
        
        # Load info button image for tooltips
        try:
            from PIL import Image, ImageTk
            info_img_path = resource_path(os.path.join("assets", "images", "info_button.png"))
            info_img = Image.open(info_img_path)
            info_photo = ImageTk.PhotoImage(info_img)
        except Exception as e:
            print(f"Warning: Could not load info button image: {e}")
            info_photo = None
        
        # Define rule descriptions with enhanced tooltips
        rule_descriptions = [
            ("rule1_bass_support", "Factor 1: Bass Support", "Adds points when the bass note supports the drive's root. Bass foundation strengthens harmonic clarity.", "(0-100)"),
            ("rule2_tonic_dominant", "Factor 2: Tonic-Dominant Relationship", "Bonus points awarded to the dominant drive of the selected tonic key. Strengthens tonal center relationships. Only select tonic if this is clear.", "(0-100)"),
            ("rule3_root_repetition", "Factor 3: Drive Repetition (per occurrence)", "Cumulative bonus for each repeated drive of any type associated with a given root note. Reinforcement through repetition increases drive strength.", "(0-100)"),
            ("rule4_resolution_max", "Factor 4: Proportional Resolution (maximum)", "Bonus given to proportional representation of the drive in the piece.", "(0-100)"),
            ("rule5_clean_voicing", "Factor 5: Clean Voicing", "Bonus for clean drive presentation. Clear voicing enhances harmonic impact.", "(0-100)"),
            ("rule6_same_chord", "Factor 6a: Same Chord Continuation", "Bonus when the same drive continues from the previous event.", "(0-100)"),
            ("rule6_dominant_prep", "Factor 6b: Dominant Preparation", "Bonus when the previous drive was a dominant.", "(0-100)"),
            ("rule7_root_doubled", "Factor 7a: Root Doubled", "Bonus when the drive root appears exactly twice. Root doubling strengthens harmonic foundation.", "(0-100)"),
            ("rule7_root_tripled", "Factor 7b: Root Tripled+", "Bonus when the drive root appears three or more times. Multiple roots create powerful harmonic drive.", "(0-100)")
        ]
        
        for rule_key, rule_name, tooltip, range_text in rule_descriptions:
            # Rule frame with title and info button
            rule_frame = ttk.LabelFrame(rules_entries_frame, text="", padding=10)
            rule_frame.pack(fill=tk.X, pady=5)
            
            # Header frame for title and info button
            header_frame = ttk.Frame(rule_frame, style="Dialog.TFrame")
            header_frame.pack(fill=tk.X, pady=(0, 10))
            
            # Title label
            title_label = ttk.Label(header_frame, text=rule_name, font=("Segoe UI", 11, "bold"), style="Dialog.TLabel")
            title_label.pack(side=tk.LEFT)
            
            # Info button with tooltip
            if info_photo:
                info_label = tk.Label(header_frame, image=info_photo, bd=0, bg="#f5f5f5", highlightthickness=0, cursor="hand2")
                info_label.image = info_photo  # Keep a reference
            else:
                # Fallback to text if image fails to load
                info_label = tk.Label(header_frame, text="i", bg="#1f4788", fg="white", 
                                    font=("Segoe UI", 8, "bold"), width=2, height=1, relief="flat", cursor="hand2")
            info_label.pack(side=tk.LEFT, padx=(8, 0))
            
            # Add tooltip with proper hide on mouse leave
            tooltip_window_ref = [None]  # Use list to allow modification in nested functions
            
            def show_tooltip(event, tooltip_text=tooltip):
                # Hide any existing tooltip first
                if tooltip_window_ref[0]:
                    try:
                        tooltip_window_ref[0].destroy()
                    except:
                        pass
                
                # Simple tooltip implementation
                tooltip_window = tk.Toplevel(rule_frame)
                tooltip_window_ref[0] = tooltip_window
                tooltip_window.wm_overrideredirect(True)
                tooltip_window.configure(bg="#333333")
                x = event.widget.winfo_rootx() + 20
                y = event.widget.winfo_rooty() + 20
                tooltip_window.geometry(f"+{x}+{y}")
                
                label = tk.Label(tooltip_window, text=tooltip_text, bg="#333333", fg="white", 
                               font=("Segoe UI", 9), wraplength=300, justify="left", padx=8, pady=4)
                label.pack()
            
            def hide_tooltip(event=None):
                if tooltip_window_ref[0]:
                    try:
                        tooltip_window_ref[0].destroy()
                        tooltip_window_ref[0] = None
                    except:
                        pass
            
            info_label.bind("<Enter>", show_tooltip)
            info_label.bind("<Leave>", hide_tooltip)
            
            # Content frame for entry controls
            content_frame = ttk.Frame(rule_frame, style="Dialog.TFrame")
            content_frame.pack(fill=tk.X, pady=(5, 0))
            
            # Special handling for Factor 2 (Tonic-Dominant)
            if rule_key == "rule2_tonic_dominant":
                # Points label and entry
                ttk.Label(content_frame, text="Points:", style="Dialog.TLabel").pack(side=tk.LEFT, padx=(0, 5))
                
                var = tk.StringVar()
                self.rule_vars[rule_key] = var
                
                entry = ttk.Entry(content_frame, textvariable=var, width=10, justify="center", style="Dialog.TEntry")
                entry.pack(side=tk.LEFT, padx=(5, 10))
                
                # Tonic key selector
                ttk.Label(content_frame, text="Tonic:", style="Dialog.TLabel").pack(side=tk.LEFT, padx=(0, 5))
                
                tonic_var = tk.StringVar()
                self.rule_vars["rule2_selected_tonic"] = tonic_var
                
                tonic_keys = ['No Tonic', 'C', 'C#', 'Db', 'D', 'D#', 'Eb', 'E', 'F', 'F#', 'Gb', 'G', 'G#', 'Ab', 'A', 'A#', 'Bb', 'B']
                tonic_combo = ttk.Combobox(content_frame, textvariable=tonic_var, values=tonic_keys, 
                                         width=6, justify="center", style="Dialog.TCombobox")
                tonic_combo.pack(side=tk.LEFT, padx=(5, 10))
                
                # Range label
                ttk.Label(content_frame, text=range_text, font=("Segoe UI", 8), 
                         foreground="gray", style="Dialog.TLabel").pack(side=tk.LEFT, padx=(5, 0))
            else:
                # Standard handling for other factors
                ttk.Label(content_frame, text="Points:", style="Dialog.TLabel").pack(side=tk.LEFT, padx=(0, 5))
                
                # Entry with validation
                var = tk.StringVar()
                self.rule_vars[rule_key] = var
                
                entry = ttk.Entry(content_frame, textvariable=var, width=10, justify="center", style="Dialog.TEntry")
                entry.pack(side=tk.LEFT, padx=(5, 0))
                
                # Range label
                ttk.Label(content_frame, text=range_text, font=("Segoe UI", 8), 
                         foreground="gray", style="Dialog.TLabel").pack(side=tk.LEFT, padx=(5, 0))
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def load_current_values(self):
        """Load current values into the UI."""
        # Load strength values
        for chord_symbol, var in self.strength_vars.items():
            # Use current values if available, otherwise use defaults
            current_value = self.strength_map.get(chord_symbol, self.DEFAULT_STRENGTH_MAP.get(chord_symbol, 0))
            var.set(str(current_value))
        
        # Load rule values  
        for rule_key, var in self.rule_vars.items():
            # Use current values if available, otherwise use defaults
            current_value = self.rule_params.get(rule_key, self.DEFAULT_RULE_PARAMS.get(rule_key, 0))
            var.set(str(current_value))
    
    def validate_inputs(self):
        """Validate all input values."""
        errors = []
        
        # Validate strength values
        for chord_symbol, var in self.strength_vars.items():
            try:
                value = int(var.get())
                if not (0 <= value <= 100):
                    chord_desc = {
                        "7": "Dominant 7th", "7b5": "Dominant 7th ♭5", "7#5": "Dominant 7th ♯5",
                        "m7": "Minor 7th", "ø7": "Half-diminished 7th", "7m9noroot": "Minor 9th (no root)",
                        "7no3": "Dom 7th (no 3rd)", "7no5": "Dom 7th (no 5th)", "7noroot": "Dom 7th (no root)",
                        "aug": "Augmented", "": "Major triad", "m": "Minor triad",
                        "maj7": "Major 7th", "mMaj7": "Minor-major 7th"
                    }
                    errors.append(f"{chord_desc.get(chord_symbol, chord_symbol)}: must be 0-100 (got {value})")
            except ValueError:
                errors.append(f"Chord strength for {chord_symbol}: must be a number")
        
        # Validate rule values with universal 0-100 range
        rule_ranges = {
            "rule1_bass_support": (0, 100),
            "rule3_root_repetition": (0, 100),
            "rule4_resolution_max": (0, 100),
            "rule5_clean_voicing": (0, 100),
            "rule6_same_chord": (0, 100),
            "rule6_dominant_prep": (0, 100),
            "rule7_root_doubled": (0, 100),
            "rule7_root_tripled": (0, 100)
        }
        
        rule_names = {
            "rule1_bass_support": "Factor 1 (Bass Support)",
            "rule3_root_repetition": "Factor 3 (Root Repetition)",
            "rule4_resolution_max": "Factor 4 (Resolution Max)",
            "rule5_clean_voicing": "Factor 5 (Clean Voicing)",
            "rule6_same_chord": "Factor 6a (Same Chord)",
            "rule6_dominant_prep": "Factor 6b (Dominant Prep)",
            "rule7_root_doubled": "Factor 7a (Root Doubled)",
            "rule7_root_tripled": "Factor 7b (Root Tripled)"
        }
        
        for rule_key, var in self.rule_vars.items():
            # Skip validation for tonic selector (it's a string, not a number)
            if rule_key == "rule2_selected_tonic":
                continue
                
            try:
                value = int(var.get())
                min_val, max_val = rule_ranges.get(rule_key, (0, 100))
                if not (min_val <= value <= max_val):
                    errors.append(f"{rule_names.get(rule_key, rule_key)}: must be 0-100 (got {value})")
            except ValueError:
                errors.append(f"{rule_names.get(rule_key, rule_key)}: must be a number")
        
        return errors
    
    def apply(self):
        """Apply the current settings."""
        errors = self.validate_inputs()
        if errors:
            messagebox.showerror("Validation Error", "\n".join(errors))
            return False
        
        # Update strength map
        new_strength_map = {}
        for chord_symbol, var in self.strength_vars.items():
            new_strength_map[chord_symbol] = int(var.get())
        
        # Update rule params
        new_rule_params = {}
        for rule_key, var in self.rule_vars.items():
            if rule_key == "rule2_selected_tonic":
                # Keep tonic selector as string
                new_rule_params[rule_key] = var.get()
            else:
                # Convert numeric values to integers
                new_rule_params[rule_key] = int(var.get())
        
        self.new_strength_map = new_strength_map
        self.new_rule_params = new_rule_params
        self.result = "apply"
        return True
    
    def ok(self):
        """OK button handler."""
        if self.apply():
            self.window.destroy()
    
    def cancel(self):
        """Cancel button handler."""
        self.result = None
        self.window.destroy()
    
    def save_preset(self):
        """Save current parameters to a preset file."""
        import json
        import datetime
        from tkinter import filedialog, simpledialog
        
        # Get preset name from user
        preset_name = simpledialog.askstring(
            "Save Preset", 
            "Enter a name for this preset:",
            parent=self.window
        )
        
        if not preset_name:
            return
        
        # Get current values
        current_strength_map = {}
        for chord_symbol, var in self.strength_vars.items():
            try:
                current_strength_map[chord_symbol] = int(var.get())
            except ValueError:
                current_strength_map[chord_symbol] = 0
        
        current_rule_params = {}
        for rule_key, var in self.rule_vars.items():
            try:
                current_rule_params[rule_key] = int(var.get())
            except ValueError:
                current_rule_params[rule_key] = 0
        
        # Create preset data
        preset_data = {
            "name": preset_name,
            "strength_map": current_strength_map,
            "rule_params": current_rule_params,
            "created": str(datetime.datetime.now()),
            "version": "1.0"
        }
        
        # Ask for save location
        filename = filedialog.asksaveasfilename(
            title="Save Drive Strength Preset",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=f"{preset_name.replace(' ', '_')}_preset.json",
            parent=self.window
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(preset_data, f, indent=2)
                messagebox.showinfo("Success", f"Preset '{preset_name}' saved successfully!", parent=self.window)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save preset:\n{str(e)}", parent=self.window)
    
    def load_preset(self):
        """Load parameters from a preset file."""
        import json
        from tkinter import filedialog
        
        # Ask for file to load
        filename = filedialog.askopenfilename(
            title="Load Drive Strength Preset",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            parent=self.window
        )
        
        if not filename:
            return
        
        try:
            with open(filename, 'r') as f:
                preset_data = json.load(f)
            
            # Validate preset data
            if not isinstance(preset_data, dict):
                raise ValueError("Invalid preset file format")
            
            if "strength_map" not in preset_data or "rule_params" not in preset_data:
                raise ValueError("Preset file missing required data")
            
            # Load strength values
            strength_map = preset_data["strength_map"]
            for chord_symbol, var in self.strength_vars.items():
                if chord_symbol in strength_map:
                    var.set(str(strength_map[chord_symbol]))
            
            # Load rule values
            rule_params = preset_data["rule_params"]
            for rule_key, var in self.rule_vars.items():
                if rule_key in rule_params:
                    var.set(str(rule_params[rule_key]))
            
            preset_name = preset_data.get("name", "Unknown")
            messagebox.showinfo("Success", f"Preset '{preset_name}' loaded successfully!", parent=self.window)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load preset:\n{str(e)}", parent=self.window)
    
    def reset_defaults(self):
        """Reset all values to defaults."""
        # Reset strength values
        for chord_symbol, var in self.strength_vars.items():
            var.set(str(self.DEFAULT_STRENGTH_MAP.get(chord_symbol, 0)))
        
        # Reset rule values
        for rule_key, var in self.rule_vars.items():
            var.set(str(self.DEFAULT_RULE_PARAMS.get(rule_key, 0)))

class EntropyAnalyzer:
    """
    Advanced statistical analysis of chord progressions.
    
    Provides two-stage analysis:
    1. Chord strength assessment based on harmonic complexity
    2. Information entropy calculation for progression predictability
    """
    NOTE_TO_SEMITONE = {
        'C': 0, 'C#': 1, 'Db': 1, 'D': 2, 'D#': 3, 'Eb': 3, 'E': 4,
        'F': 5, 'F#': 6, 'Gb': 6, 'G': 7, 'G#': 8, 'Ab': 8, 'A': 9,
        'A#': 10, 'Bb': 10, 'B': 11
    }
    _STRENGTH_MAP = {
        "7": 100,
        "7b5": 90,
        "7#5": 80,
        "m7": 70,
        "ø7": 65,
        "7m9noroot": 65,
        "7no3": 55,
        "7no5": 55,
        "7noroot": 50,
        "aug": 40,
        "": 42,       # pure major triad
        "m": 35,
        "maj7": 30,
        "mMaj7": 25,
    }

    def __init__(
        self,
        events: Dict[Tuple[int, int, str], Dict[str, Any]],
        symbol_mode: str = "chord",
        base: int = 2,
        logger: Callable[[str], None] = print,
        strength_map: Optional[Dict[str, int]] = None,
        rule_params: Optional[Dict[str, int]] = None
    ):
        self.events = events
        self.symbol_mode = symbol_mode
        self.base = base
        self.logger = logger
        self.custom_steps: List[Tuple[str, Callable[["EntropyAnalyzer"], None]]] = []
        
        # Use provided parameters or defaults
        self.strength_map = strength_map if strength_map is not None else self._STRENGTH_MAP.copy()
        
        # Default rule parameters (synchronized with dialog defaults)
        default_rule_params = {
            "rule1_bass_support": 20,
            "rule2_tonic_dominant": 50,
            "rule2_selected_tonic": "No Tonic",  # Default: disabled
            "rule3_root_repetition": 20,
            "rule4_resolution_max": 50,
            "rule5_clean_voicing": 50,
            "rule6_same_chord": 33,
            "rule6_dominant_prep": 50,
            "rule7_root_doubled": 33,
            "rule7_root_tripled": 50
        }
        self.rule_params = rule_params if rule_params is not None else default_rule_params

    # --------------------------
    # Stage 1: Chord strengths
    # --------------------------
    def _fourth_up(self, root: str) -> str:
        """Return the note a perfect fourth above the given root."""
        # Handle empty or None input
        if not root or root.strip() == "":
            return ""
            
        root = root.strip()
        chromatic_sharps = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
        flats_to_sharps = {'Db': 'C#', 'Eb': 'D#', 'Fb': 'E', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#', 'Cb': 'B'}
        
        # Extract just the root note (remove chord quality, extensions, etc.)
        # Handle chord symbols like "C7", "Dm", "F#maj7", etc.
        root_note = ""
        for i, char in enumerate(root):
            if i == 0 and char in 'ABCDEFG':
                root_note = char
            elif i == 1 and char in '#b':
                root_note += char
                break
            else:
                break
        
        if not root_note:
            self.logger(f"[Warning] _fourth_up: cannot extract root note from '{root}'")
            return root
            
        note = flats_to_sharps.get(root_note, root_note)
        if note not in chromatic_sharps:
            self.logger(f"[Warning] _fourth_up: unknown note '{root_note}' from chord '{root}'")
            return root
        index = chromatic_sharps.index(note)
        fourth_index = (index + 5) % 12  # perfect fourth = +5 semitones
        return chromatic_sharps[fourth_index]

    def _fifth_up(self, root: str) -> str:
        """Return the note a perfect fifth above the given root."""
        # Handle empty or None input
        if not root or root.strip() == "":
            return ""
            
        root = root.strip()
        chromatic_sharps = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
        flats_to_sharps = {'Db': 'C#', 'Eb': 'D#', 'Fb': 'E', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#', 'Cb': 'B'}
        
        # Extract just the root note (remove chord quality, extensions, etc.)
        root_note = ""
        for i, char in enumerate(root):
            if i == 0 and char in 'ABCDEFG':
                root_note = char
            elif i == 1 and char in '#b':
                root_note += char
                break
            else:
                break
        
        if not root_note:
            return root
            
        note = flats_to_sharps.get(root_note, root_note)
        if note not in chromatic_sharps:
            return root
        index = chromatic_sharps.index(note)
        fifth_index = (index + 7) % 12  # perfect fifth = +7 semitones
        return chromatic_sharps[fifth_index]


    def step_stage1_strengths(self, print_legend: bool = True):
        import re
        root_counter: Dict[str, int] = {}
        prev_event_roots: Set[str] = set()
        pending_roots: Set[str] = set()
        resolution_count: Dict[str, int] = {}
        total_resolutions: int = 0
        prev_event_chords: Set[str] = set()

        # Prepare table data
        table_rows = []
        rule_names = [f"F{i+1}" for i in range(7)]

        for (bar, beat, ts), payload in self.events.items():
            chords = payload.get("chords", [])
            basses = payload.get("basses", [])
            chord_scores: List[Tuple[str, float, List[str]]] = []
            current_event_roots: Set[str] = set()
            event_label = f"Bar {bar}, Beat {beat} ({ts})"

            # For each chord, collect which rules applied
            for chord in chords:
                root, quality = self._split_chord(chord)
                # Pass root_counter for R3
                base_score, rule_msgs = self._compute_score(chord, basses, payload, root_counter=root_counter)
                applied_rules = rule_msgs[:]

                # Rule 6: Previous event contains same chord or dominant chord
                rule6_bonus = 0
                dominant_root = self._fifth_up(root)
                dominant_chord = dominant_root + quality
                if chord in prev_event_chords:
                    rule6_same = self.rule_params.get("rule6_same_chord", 5)
                    rule6_bonus += rule6_same
                    applied_rules.append(f"Rule 6: Previous event contained {chord} → +{rule6_same}")
                if dominant_chord in prev_event_chords:
                    rule6_dom = self.rule_params.get("rule6_dominant_prep", 10)
                    rule6_bonus += rule6_dom
                    applied_rules.append(f"Rule 6: Previous event contained dominant {dominant_chord} → +{rule6_dom}")
                base_score += rule6_bonus

                # Rule 4: proportional resolution
                if total_resolutions > 0:
                    ratio = resolution_count.get(root, 0) / total_resolutions
                    rule4_max = self.rule_params.get("rule4_resolution_max", 10)
                    r4_bonus = rule4_max * ratio
                    if r4_bonus > 0:
                        applied_rules.append(f"Rule 4: Resolution ratio {ratio:.2f} → +{int(round(r4_bonus))}")
                    base_score += r4_bonus

                chord_scores.append((chord, base_score, applied_rules))
                current_event_roots.add(root)

                # Update root_counter for R3
                prev_count = root_counter.get(root, 0)
                root_counter[root] = prev_count + 1

            # Update Rule4 counters
            for prev_root in pending_roots:
                for cur_root in current_event_roots:
                    if cur_root == self._fourth_up(prev_root):
                        resolution_count[prev_root] = resolution_count.get(prev_root, 0) + 1
                        total_resolutions += 1

            # Build table row for each chord
            for chord, score, rules in chord_scores:
                row = [event_label + f" {chord}"]
                # For each rule, extract just the bonus points (e.g., +10, +5)
                for i in range(1, 8):
                    found = next((r for r in rules if r.startswith(f"Rule {i}:") or r.startswith(f"Rule {i} ")), "")
                    # Extract only the bonus after a + or - sign (not the rule number)
                    match = re.search(r"([+-]\d+(?:\.\d+)?)", found)
                    cell = match.group(1) if match else ""
                    row.append(cell)
                table_rows.append(row)

            # Prepare for next event
            pending_roots = current_event_roots.copy()
            prev_event_roots = current_event_roots.copy()
            prev_event_chords = set(chords)

        # Print table - wider first column for longer chord names
        col_widths = [35, 6] + [4]*7 + [6]
        header = ["Event/Chord", "base"] + rule_names + ["Total"]
        header_line = " | ".join(h.ljust(w) for h, w in zip(header, col_widths))
        sep_line = "-+-".join("-"*w for w in col_widths)
        self.logger(header_line)
        self.logger(sep_line)

        for row in table_rows:
            label = row[0]
            chord_name = label.split()[-1]
            # Get base strength from strength_map
            _, quality = self._split_chord(chord_name)
            base_strength = self.strength_map.get(quality or "", 0)
            total = base_strength
            for cell in row[1:]:
                try:
                    total += float(cell) if cell else 0
                except Exception:
                    pass
            # Format total as int if possible
            total_str = str(int(total)) if total == int(total) else f"{total:.2f}"
            # Insert base_strength as the second column
            self.logger(" | ".join(str(cell).rjust(w) for cell, w in zip([label, base_strength] + row[1:] + [total_str], col_widths)))

        self.logger("")  # Add a blank line before the legend

        # --- Compute and print average and maximum entropy ---
        entropy_values = []
        for payload in self.events.values():
            chords = payload.get("chords", [])
            if not chords:
                continue
            # For each event, compute entropy of the chord strengths (base + bonuses)
            event_scores = []
            for chord in chords:
                score, _ = self._compute_score(chord, payload.get("basses", []), payload)
                event_scores.append(score)
            if event_scores:
                # Use Shannon entropy of the event's chord scores
                from math import log2
                from collections import Counter
                counts = Counter(event_scores)
                total = sum(counts.values())
                entropy = -sum((count/total) * log2(count/total) for count in counts.values()) if total > 0 else 0.0
                entropy_values.append(entropy)
        if entropy_values:
            avg_entropy = sum(entropy_values) / len(entropy_values)
            max_entropy = max(entropy_values)
            self.logger(f"Average entropy = {avg_entropy:.3f} bits")
            self.logger(f"Maximum entropy = {max_entropy:.3f} bits")
        else:
            self.logger("Average entropy = 0.000 bits")
            self.logger("Maximum entropy = 0.000 bits")

        legend = (
            "Legend for Entropy Grid:\n"
            "  Factor 1 - Is the drive supported by the bass?\n"
            "  Factor 2 - Tonic-dominant relationship (configurable)\n"
            "  Factor 3 - How often is the drive recurring?\n"
            "  Factor 4 - How often is the drive discharging?\n"
            f"  Factor 5 - Is the drive articulated as a clean stack in the texture? {CLEAN_STACK_SYMBOL}\n"
            "  Factor 6 - Was the drive itself, or its dominant in the previous event?\n"
            f"  Factor 7 - Is the root of the drive doubled at the octave? {ROOT2_SYMBOL}\n"
        )
        self.logger(legend)

    @staticmethod
    def _get_chord_scores_static(payload, self_ref):
        chords = payload.get("chords", [])
        basses = payload.get("basses", [])
        chord_scores: list = []
        for chord in chords:
            score, rules = self_ref._compute_score(chord, basses, payload)
            chord_scores.append((chord, score, rules))
        return chord_scores

    # --------------------------
    # Stage 2: Strength entropy
    # --------------------------
    def _weighted_entropy(self, scores: List[int], base: int = 2) -> float:
        if not scores:
            return 0.0
        total = sum(scores)
        if total == 0:
            return 0.0
        probs = [s / total for s in scores]
        return -sum(p * log2(p) / log2(base) for p in probs if p > 0)

    def step_stage2_strength_entropy(self):
        scores = self._make_score_sequence()
        if not scores:
            self.logger("[Phase7] No scores for entropy calculation.")
            return
        H = self._weighted_entropy(scores, base=self.base)
        self.logger(f"[Phase7] Weighted strength-sequence entropy H (base {self.base}): {H:.4f} bits")

    # --------------------------
    # New helper: compute score with modifiers
    # --------------------------
    def _get_dominant_of_tonic(self, tonic: str) -> str:
        """Return the dominant (5th) of the given tonic key."""
        # Circle of fifths: each key's dominant is a perfect 5th up
        dominant_map = {
            'C': 'G', 'G': 'D', 'D': 'A', 'A': 'E', 'E': 'B', 'B': 'F#', 'F#': 'C#',
            'C#': 'G#', 'G#': 'D#', 'D#': 'A#', 'A#': 'F', 'F': 'C',
            # Enharmonic equivalents
            'Db': 'Ab', 'Ab': 'Eb', 'Eb': 'Bb', 'Bb': 'F',
            'Gb': 'Db'
        }
        return dominant_map.get(tonic, 'G')  # Default to G if unknown tonic

    def _compute_score(self, chord: str, basses: Optional[List[str]] = None, event_payload: Optional[dict] = None, root_counter: Optional[Dict[str, int]] = None) -> Tuple[int, List[str]]:
        root, quality = self._split_chord(chord)
        score = self.strength_map.get(quality or "", 0)
        messages: List[str] = []

        # Rule 1: Bass support
        if basses and root in basses:
            rule1_bonus = self.rule_params.get("rule1_bass_support", 20)
            score += rule1_bonus
            messages.append(f"Rule 1: Bass supports {chord} → +{rule1_bonus} bonus")

        # Rule 2: Tonic-Dominant relationship
        selected_tonic = self.rule_params.get("rule2_selected_tonic", "No Tonic")
        if selected_tonic != "No Tonic":
            dominant_of_tonic = self._get_dominant_of_tonic(selected_tonic)
            if root == dominant_of_tonic:
                rule2_bonus = self.rule_params.get("rule2_tonic_dominant", 50)
                score += rule2_bonus
                messages.append(f"Rule 2: {chord} is dominant of {selected_tonic} → +{rule2_bonus} bonus")

        # Rule 3: root repetition (now always included)
        if root_counter is not None:
            prev_count = root_counter.get(root, 0)
            rule3_multiplier = self.rule_params.get("rule3_root_repetition", 2)
            r3_bonus = rule3_multiplier * prev_count if prev_count > 0 else 0
            if r3_bonus > 0:
                messages.append(f"Rule 3: Root {root} repeated → +{r3_bonus}")
            score += r3_bonus

        if event_payload is not None:
            chord_info = event_payload.get("chord_info", {})
            # Rule 5
            if chord_info.get(chord, {}).get("clean_stack"):
                rule5_bonus = self.rule_params.get("rule5_clean_voicing", 10)
                score += rule5_bonus
                messages.append(f"Rule 5: Clean chord {chord} → +{rule5_bonus} bonus")
            # Rule 7
            root_count = chord_info.get(chord, {}).get("root_count", 1)
            if root_count == 2:
                rule7_doubled = self.rule_params.get("rule7_root_doubled", 5)
                score += rule7_doubled
                messages.append(f"Rule 7: Root doubled in chord {chord} → +{rule7_doubled} bonus")
            elif root_count >= 3:
                rule7_tripled = self.rule_params.get("rule7_root_tripled", 10)
                score += rule7_tripled
                messages.append(f"Rule 7: Root tripled+ in chord {chord} → +{rule7_tripled} bonus")

        return score, messages


    def _make_score_stream(self) -> List[List[int]]:
        scored_events: List[List[int]] = []
        for payload in self.events.values():
            chords = payload.get("chords", [])
            basses = payload.get("basses", [])
            event_scores: List[int] = []
            for chord in chords:
                score, _ = self._compute_score(chord, basses, payload)  # Pass payload here!
                event_scores.append(score)
            scored_events.append(event_scores)
        return scored_events

    def _make_score_sequence(self) -> List[int]:
        seq: List[int] = []
        for scores in self._make_score_stream():
            seq.extend(scores)
        return seq

    def _split_chord(self, chord: str) -> Tuple[str, str]:
        if not chord:
            return ("", "")
        if len(chord) > 1 and chord[1] in ["#", "b", "♯", "♭"]:
            root = chord[:2]
            quality = chord[2:]
        else:
            root = chord[0]
            quality = chord[1:]
        return root, quality

    def _shannon_entropy(self, seq: List[Any], base: int = 2) -> float:
        if not seq:
            return 0.0
        counts = Counter(seq)
        total = len(seq)
        return -sum((count / total) * log2(count / total) / log2(base) for count in counts.values())

    # --------------------------
    # Public API
    # --------------------------
    def register_step(self, name: str, func: Callable[["EntropyAnalyzer"], None]):
        self.custom_steps.append((name, func))

    def preview(self):
        self.logger("[Phase7] --- Basic stats ---")
        seq = self._make_score_sequence()
        if seq:
            self.logger(f"[Phase7] Total scores: {len(seq)}, Unique: {len(set(seq))}")
        else:
            self.logger("[Phase7] No scores available.")
        self.logger("[Phase7] --- Custom steps ---")
        for name, func in self.custom_steps:
            self.logger(f"[Phase7] Step: {name}")
            func(self)
    # Legend print handled in step_stage1_strengths

if __name__ == "__main__":
    debug_log("=== Application Starting ===")
    debug_log(f"Python version: {sys.version}")
    debug_log(f"Platform: {platform.platform()}")
    debug_log(f"MIDO_AVAILABLE: {MIDO_AVAILABLE}")
    debug_log(f"PYGAME_AVAILABLE: {PYGAME_AVAILABLE}")
    app = MidiChordAnalyzer()
    app.mainloop()
