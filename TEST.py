"""
MIDI Chord Analyzer - Advanced Music Theory Analysis Tool

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
    import mido
    import rtmidi  # Explicitly import the backend for PyInstaller
    MIDO_AVAILABLE = True
except ImportError:
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

# Chord detection priority (more specific chords checked first)
PRIORITY = [
    "C7", "C7b5", "C7#5", "Cm7", "Cø7", "C7m9noroot",
    "C7no3", "C7no5", "C7noroot", "Caug", "C", "Cm",
    "Cmaj7", "CmMaj7"
]

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

        print("Note List:")
        for elem in flat_notes:
            if isinstance(elem, note.Note):
                name = elem.nameWithOctave
                midi = elem.pitch.midi
                dur = elem.quarterLength
                offset = elem.offset
                bar, beat, ts = offset_to_bar_beat(offset)
                print(f"Note: {name} (MIDI {midi}) | Bar {bar}, Beat {beat} ({ts}) | Duration: {dur}")
            elif isinstance(elem, m21chord.Chord):
                names = [p.nameWithOctave for p in elem.pitches]
                midis = [p.midi for p in elem.pitches]
                dur = elem.quarterLength
                offset = elem.offset
                bar, beat, ts = offset_to_bar_beat(offset)
                print(f"Chord: {names} (MIDIs {midis}) | Bar {bar}, Beat {beat} ({ts}) | Duration: {dur}")
    def __init__(self):
        super().__init__()
        self.title("🎵 MIDI Drive Analyzer")
        self.geometry("650x850")
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
        self.sensitivity = "Medium"
        self.remove_repeats = True
        self.include_anacrusis = True
        self.include_non_drive_events = True
        self.arpeggio_searching = True
        self.neighbour_notes_searching = True
        self.arpeggio_block_similarity_threshold = 0.5

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

        self.build_ui()
        self.show_splash()



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
            btn_kwargs = {"bg": "#ff00ff", "fg": "#fff", "activebackground": "#ff33ff", 
                         "activeforeground": "#fff", "relief": "raised", "bd": 2, 
                         "font": ("Segoe UI", 10, "bold")}
            disabled_fg = "black"

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
                # Instantiate the embedded keyboard UI
                EmbeddedMidiKeyboard(top)
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
            self, bg="black", fg="white", font=("Consolas", 11),
            wrap="word", borderwidth=0, highlightthickness=0,
            selectbackground="blue", selectforeground="white",  # Visible selection colors
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
            self.result_text.insert("end", "Harmonic Drive Analyzer\n")
            print("Splash image load error:", e)
        description = (
            "• Analyze MusicXML scores\n"
            "• Model patterns of harmonic tension\n"
            "• Produce PDF graph of tension-release patterns\n"
            "• Model harmonic entropy\n\n"
            "For information on drive analysis, see: http://www.chromatic-harmony.com/drive_analysis.html\n"
            "\n"
            "Kenneth Smith, Desire in Chromatic Harmony (New York: Oxford University Press, 2020).\n"
            "Kenneth Smith, “The Enigma of Entropy in Extended Tonality.” Music Theory Spectrum 43, no. 1 (2021): 1–18."
        )
        self.result_text.insert("end", description)
        self.result_text.configure(state="disabled")
       
    def preview_entropy(self, mode: str = "chord", base: int = 2):
        if not self.analyzed_events:
            print("No analyzed events available for entropy preview.")
            return
        ea = EntropyAnalyzer(self.analyzed_events, symbol_mode=mode, base=base, logger=print)
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
        min_duration = {"High": 0.1, "Medium": 0.5, "Low": 1.0}[self.sensitivity]
        self.analyzed_events = None
        self.processed_events = None
        try:
            lines, events = self.analyze_musicxml(self.score, min_duration=min_duration)
            self.analyzed_events = events

            self.display_results()
            
            # Generate entropy analysis for advanced statistics
            from io import StringIO
            entropy_buf = StringIO()
            analyzer = EntropyAnalyzer(self.analyzed_events, logger=lambda x: print(x, file=entropy_buf))
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
        dialog.geometry("420x480")

        # Analysis algorithm toggles
        include_triads_var = tk.BooleanVar(value=self.include_triads)
        remove_repeats_var = tk.BooleanVar(value=self.remove_repeats)
        include_anacrusis_var = tk.BooleanVar(value=self.include_anacrusis)
        arpeggio_searching_var = tk.BooleanVar(value=self.arpeggio_searching)
        neighbour_notes_var = tk.BooleanVar(value=getattr(self, 'neighbour_notes_searching', True))
        include_non_drive_var = tk.BooleanVar(value=self.include_non_drive_events)

        # Merging sensitivity slider (7 positions: 1=minimal merging, 7=aggressive merging)
        sensitivity_scale_var = tk.IntVar(value=getattr(self, 'collapse_sensitivity_pos', 3))

        pad_opts = dict(anchor="w", padx=12, pady=6)
        ttk.Checkbutton(dialog, text="Include triads", variable=include_triads_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Include anacrusis", variable=include_anacrusis_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Arpeggio searching", variable=arpeggio_searching_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Neighbour notes", variable=neighbour_notes_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Remove repeated patterns", variable=remove_repeats_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Include non-drive events", variable=include_non_drive_var).pack(**pad_opts)

        # Separator line
        ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=12, pady=8)

        ttk.Label(dialog, text="Merge similar events together (1=minimal merging, 7=high merging):").pack(anchor="w", padx=12, pady=(8, 0))
        sens_scale = tk.Scale(dialog, from_=1, to=7, orient="horizontal", variable=sensitivity_scale_var)
        sens_scale.pack(fill="x", padx=12)

        ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=12, pady=8)

        ttk.Label(dialog, text="Note detail (1=fewer details, 5=more details):").pack(anchor="w", padx=12, pady=(8, 0))
        current_sensitivity = getattr(self, 'sensitivity', 'Medium')
        sensitivity_to_pos = {"Low": 1, "Medium": 3, "High": 5}
        detail_scale_var = tk.IntVar(value=sensitivity_to_pos.get(current_sensitivity, 3))
        detail_scale = tk.Scale(dialog, from_=1, to=5, orient="horizontal", variable=detail_scale_var)
        detail_scale.pack(fill="x", padx=12)

        def apply_settings():
            # Apply basic boolean/text settings
            self.include_triads = include_triads_var.get()
            self.remove_repeats = remove_repeats_var.get()
            self.include_anacrusis = include_anacrusis_var.get()
            self.arpeggio_searching = arpeggio_searching_var.get()
            self.neighbour_notes_searching = neighbour_notes_var.get()
            self.include_non_drive_events = include_non_drive_var.get()

            # Read detail slider and convert to sensitivity
            detail_pos = int(detail_scale_var.get())
            pos_to_sensitivity = {1: "Low", 2: "Low", 3: "Medium", 4: "High", 5: "High"}
            self.sensitivity = pos_to_sensitivity.get(detail_pos, "Medium")

            # Read merge slider position and persist it
            pos = int(sensitivity_scale_var.get())
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

            dialog.destroy()
            
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

        # Separator line
        ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=12, pady=8)

        ttk.Button(dialog, text="Apply", command=apply_settings).pack(pady=(6, 12))

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
                    chords_display = ", ".join(chord_strs) if chord_strs else "no known drive"
                    bass = "+".join(data["basses"])
                    is_no_drive = len(chord_strs) == 0
                    # Deduplicate at output: skip if previous was also no known drive with same bass
                    if is_no_drive and prev_no_drive and bass == prev_bass:
                        continue
                    # Respect user preference for including non-drive events
                    if is_no_drive and not self.include_non_drive_events:
                        continue
                    
                    # This event will be displayed, so add it to our list
                    displayed_events.append(((bar, beat, ts), data))
                    output_lines.append(
                        f"Bar {bar}, Beat {beat} ({ts}): {chords_display} (bass = {bass})\n"
                    )
                    prev_no_drive = is_no_drive
                    prev_bass = bass
                
                # Store only the events that actually got displayed
                self.processed_events = displayed_events
                output_lines.append(
                    f"\nLegend:\n{CLEAN_STACK_SYMBOL} = Clean stack   {ROOT2_SYMBOL} = Root doubled   {ROOT3_SYMBOL} = Root tripled or more\n"
                )
                self.result_text.insert("end", "".join(output_lines))
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

        note_events = []
        single_notes = []  # (start, end, pitch)
        for elem in flat_notes:
            if isinstance(elem, (note.Note, m21chord.Chord)):
                duration = max(elem.quarterLength, min_duration)
                start = elem.offset
                end = start + duration
                if isinstance(elem, m21chord.Chord):
                    pitches = [p.midi for p in elem.pitches]
                else:
                    pitches = [elem.pitch.midi]
                note_events.append((start, end, pitches))
                # Collect single melodic notes (not part of a chord, not doubled at start)
                if isinstance(elem, note.Note):
                    others_at_start = [e for e in flat_notes if e is not elem and hasattr(e, 'offset') and e.offset == start]
                    if not others_at_start:
                        single_notes.append((start, end, pitches[0]))

        time_points = sorted(set([t for start, end, _ in note_events for t in [start, end]]))



        events = {}
        active_notes = set()
        active_pitches = set()

        # === PHASE 1: Block Chord Detection ===
        for i, time in enumerate(time_points):

                
            for start, end, pitches in note_events:
                if end == time:
                    active_notes.difference_update({p % 12 for p in pitches})
                    active_pitches.difference_update(pitches)

                        
            for start, end, pitches in note_events:
                if start == time:
                    active_notes.update({p % 12 for p in pitches})
                    active_pitches.update(pitches)


            test_notes = set(active_notes)
            test_pitches = set(active_pitches)
            if self.include_anacrusis:
                bar, beat, ts = offset_to_bar_beat(time)
                if bar == 3 and beat == 1:
                    print(f"DEBUG ANACRUSIS: Before anacrusis - test_notes={sorted(test_notes)}")
                for s_start, s_end, s_pitch in single_notes:
                    if s_end == time and (s_pitch % 12) not in test_notes:
                        if bar == 3 and beat == 1:
                            print(f"  Adding anacrusis note: {s_pitch} (pc={s_pitch % 12}) ended at time {time}")
                        test_notes.add(s_pitch % 12)
                        test_pitches.add(s_pitch)
                if bar == 3 and beat == 1:
                    print(f"DEBUG ANACRUSIS: After anacrusis - test_notes={sorted(test_notes)}")

            if len(test_notes) >= 3:
                # Check for chord formation with sufficient note count
                bar, beat, ts = offset_to_bar_beat(time)
                
                # Analyze note collection for chord detection
                if bar == 3 and beat == 1:
                    print(f"DEBUG Bar {bar}, Beat {beat}: test_notes={sorted(test_notes)}, test_pitches={sorted(test_pitches)}")
                    print(f"  Active notes before anacrusis: {sorted(active_notes)}")
                
                chords = self.detect_chords(test_notes, debug=(bar == 3 and beat == 1))
                bar, beat, ts = offset_to_bar_beat(time)
                key = (bar, beat, ts)
                # Event created; previously had diagnostic printing here which has been removed
                if bar == 3 and beat == 1:
                    print(f"DEBUG BLOCK CHORD: detect_chords returned: {chords}")
                if chords:
                    bass_note = self.semitone_to_note(min(test_pitches) % 12)
                    if key not in events:
                        events[key] = {"chords": set(), "basses": set(), "event_notes": set(test_notes)}

                    
                    events[key]["chords"].update(chords)
                    events[key]["basses"].add(bass_note)
                    if bar == 3 and beat == 1:
                        print(f"DEBUG EVENT CREATION: Added chords {chords} to event {key}")
                    

                    
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
            
            # Display single note analysis for specific range
            print("[ARPEGGIO DEBUG] Single notes in Bar 10-12 range:")
            for note_elem in melodic_notes:
                bar, beat, ts = offset_to_bar_beat(note_elem.offset)
                if 10 <= bar <= 12:
                    print(f"  Offset {note_elem.offset}: Bar {bar}.{beat} - Pitch {note_elem.pitch.midi} ({note_elem.pitch.name})")
            print()
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
                        if 10 <= bar <= 12:
                            print(f"[ARPEGGIO WINDOW] Window size {w}: notes at offsets {[n.offset for n in window]}")
                            print(f"[ARPEGGIO WINDOW] Pitches: {[n.pitch.midi for n in window]} -> PCs: {list(window_pcs)}")
                            print(f"[ARPEGGIO WINDOW] Detected chords: {chords}")
                            print(f"[ARPEGGIO WINDOW] Assigned to: Bar {bar}.{beat} (first note offset {window[0].offset})")
                        
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
                                    if 10 <= bar <= 12:
                                        print(f"[ARPEGGIO REJECT] Harmonic instability: {underlying_chords[0][0]} has {list(first_chord_set)}, {key} has {list(chord_set)}")
                                    break
                            if harmonic_stability and 10 <= bar <= 12:
                                print(f"[ARPEGGIO ACCEPT] Harmonic stability: consistent chord {list(first_chord_set)} throughout")
                        else:
                            if 10 <= bar <= 12:
                                print(f"[ARPEGGIO ACCEPT] No underlying block chords - pure melodic arpeggio")
                        
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
                        else:
                            print(f"[ARPEGGIO OVERWRITE] {key}: {list(events[key]['event_notes'])} -> {list(window_pcs)}")
                        events[key]["chords"].update(chords)
                        # Arpeggios contribute to chord identification but not bass detection
                        # Bass detection relies on actual bass line or block chord voicings
                        events[key]["event_notes"] = set(window_pcs)
                        events[key]["event_pitches"] = set(window_pitches)
                        print(f"[ARPEGGIO CREATE] {key}: chords={chords}, window_pcs={list(window_pcs)}")
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
            return PRIORITY.index(base) if base in PRIORITY else 999

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
                current_priority = PRIORITY.index(base_chord) if base_chord in PRIORITY else 999
                prev_chord = chords_by_root.get(root)
                if prev_chord:
                    prev_priority = PRIORITY.index(prev_chord.replace(root, 'C')) if prev_chord.replace(root, 'C') in PRIORITY else 999
                    if current_priority < prev_priority:
                        chords_by_root[root] = chord
                else:
                    chords_by_root[root] = chord
            processed_events.append(((bar, beat, ts), chords_by_root, basses, event_notes_set, event_pitches_set))

        # Remove trivial duplicates of same single-root chord across adjacent events
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

        # === PHASE 4: Event Merging and Post-Processing ===
        if getattr(self, 'collapse_similar_events', False) and final_filtered_events:
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
            if debug:
                print(f"  CHORD DEBUG: Trying root {root}, normalized = {sorted(normalized)}")
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

            for name in PRIORITY:
                if name in TRIADS and not self.include_triads:
                    continue
                chord_pattern = set(CHORDS[name])
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
                        matched = name.replace('C', self.semitone_to_note(root))
                        chords_found.append(matched)
                        break
                else:
                    if chord_pattern.issubset(normalized):
                        matched = name.replace('C', self.semitone_to_note(root))
                        if debug:
                            print(f"    FOUND: {name} -> {matched} (pattern {sorted(chord_pattern)} matches)")
                        chords_found.append(matched)
                        break

        # Second pass: try "noroot" style chords where the root pitch-class is absent
        for root in sorted(set(range(12)) - set(semitones)):
            normalized = {(n - root) % 12 for n in semitones}
            if debug:
                print(f"  NOROOT DEBUG: Trying absent root {root}, normalized = {sorted(normalized)}")
            for name in PRIORITY:
                if "noroot" not in name:
                    continue
                if name in TRIADS and not self.include_triads:
                    continue
                chord_pattern = set(CHORDS[name])
                if chord_pattern == normalized:
                    matched = name.replace('C', self.semitone_to_note(root))
                    if debug:
                        print(f"    NOROOT FOUND: {name} -> {matched} (pattern {sorted(chord_pattern)} matches)")
                    chords_found.append(matched)
                    break

        # BUT WAIT - we're still missing regular chord detection! Let me check if Bm is being found normally
        if debug:
            print(f"  REGULAR CHORD CHECK: Looking for Bm pattern [0,3,7] in any normalization...")
            for root in range(12):
                normalized = {(n - root) % 12 for n in semitones}
                if set([0, 3, 7]).issubset(normalized):
                    print(f"    Bm pattern found with root {root}: normalized = {sorted(normalized)}")

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
            analyzer = EntropyAnalyzer(self.analyzed_events, logger=lambda x: print(x, file=entropy_buf))
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
        
        # Use the global PRIORITY list for chord deduplication

        def chord_priority(chord_name: str) -> int:
            base = chord_name
            for root in sorted(NOTE_TO_SEMITONE.keys(), key=lambda x: -len(x)):
                if chord_name.startswith(root):
                    base = chord_name[len(root):]
                    break
            return PRIORITY.index(base) if base in PRIORITY else 999

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
                current_priority = PRIORITY.index(base_chord) if base_chord in PRIORITY else 999
                prev_chord = chords_by_root.get(root)
                if prev_chord:
                    prev_priority = PRIORITY.index(prev_chord.replace(root, 'C')) if prev_chord.replace(root, 'C') in PRIORITY else 999
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
        gw = GridWindow(self, grid_events)
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

    Usage: EmbeddedMidiKeyboard(parent_toplevel)
    """
    def __init__(self, parent):
        self.parent = parent
        # Create UI inside the provided parent (Toplevel)
        self.parent.title("🎹 Embedded Keyboard")
        # Window-wide dark theme
        try:
            self.parent.configure(bg="black")
        except Exception:
            pass

        # Try to initialize pygame.midi for sound output (optional)
        if PYGAME_AVAILABLE:
            try:
                pygame.midi.init()
                self.pygame = pygame
                self.pygame_midi = pygame.midi
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
        else:
            self.pygame = None
            self.pygame_midi = None
            self.midi_out = None

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
            self.midi_ports = mido.get_input_names()
        else:
            self.midi_ports = []
        self.midi_port_var = tk.StringVar()
        self.midi_dropdown = ttk.Combobox(midi_frame, textvariable=self.midi_port_var, values=self.midi_ports, state="readonly", width=40)
        if self.midi_ports:
            self.midi_port_var.set(self.midi_ports[0])
        self.midi_dropdown.pack(side="left")
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

    def _play_note(self, semitone, velocity=127):
        note_num = 60 + semitone
        if getattr(self, 'midi_out', None):
            try:
                self.midi_out.note_on(note_num, velocity)
            except Exception:
                pass

    def _stop_note(self, semitone):
        note_num = 60 + semitone
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

    def start_midi_listener(self, port_name=None):
        if not MIDO_AVAILABLE:
            print("mido not available: MIDI input disabled")
            return

        if not port_name:
            ports = mido.get_input_names()
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
            for msg in self.midi_in:
                try:
                    if msg.type in ('note_on', 'note_off'):
                        pitch_class = msg.note % 12
                        if msg.type == 'note_on' and getattr(msg, 'velocity', 0) > 0:
                            self.parent.after(0, lambda pc=pitch_class: self.add_midi_note(pc))
                        else:
                            self.parent.after(0, lambda pc=pitch_class: self.remove_midi_note(pc))
                except Exception:
                    continue

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
        # Use the same chord-detection logic as the main analyzer
        if len(self.selected_notes) < 3:
            self.result_label.config(text="🎵 Select at least 3 notes to analyze.")
            return

        selected = set(self.selected_notes)
        detected_chords = {}

        for root in selected:
            normalized = set((note - root) % 12 for note in selected)
            for chord_name in PRIORITY:
                if chord_name in TRIADS and not self.include_triads_var.get():
                    continue
                chord_intervals = CHORDS[chord_name]
                chord_set = set(chord_intervals)
                if chord_set.issubset(normalized):
                    root_name = self.semitone_to_note(root)
                    detected_chords[root_name] = chord_name.replace('C', root_name)
                    break

        candidate_roots = set(range(12))
        for assumed_root in candidate_roots:
            if assumed_root in selected:
                continue
            normalized = set((note - assumed_root) % 12 for note in selected)
            for chord_name in PRIORITY:
                if "noroot" not in chord_name:
                    continue
                if chord_name in TRIADS and not self.include_triads_var.get():
                    continue
                chord_intervals = CHORDS[chord_name]
                chord_set = set(chord_intervals)
                if chord_set == normalized:
                    root_name = self.semitone_to_note(assumed_root)
                    detected_chords[root_name] = chord_name.replace('C', root_name)
                    break

        if detected_chords:
            lines = []
            for root_name, chord_str in detected_chords.items():
                semitone = NOTE_TO_SEMITONE.get(root_name, 0)
                if "noroot" in chord_str and len(self.selected_notes) >= 4:
                    dim_semitone = (semitone + 1) % 12
                    dim_root = self.semitone_to_note(dim_semitone)
                    dim_chord_label = f"{dim_root}o7"
                    chord_str = chord_str.replace("noroot", "no root")
                    chord_str += f" [{dim_chord_label}]"
                else:
                    chord_str = chord_str.replace("noroot", "no root")

                if semitone in (8, 1, 3, 6, 10):
                    enh = ENHARMONIC_EQUIVALENTS.get(semitone, root_name)
                    if isinstance(enh, str) and '/' in enh:
                        enh_roots = enh.split('/')
                        if len(enh_roots) == 2:
                            chord_str += f" ({enh_roots[1]} root)"

                lines.append(chord_str)

            self.result_label.config(text="\n".join(lines))
        else:
            self.result_label.config(text="No matching chords found.")


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

    def __init__(self, parent, events):
        super().__init__(parent)
        self.title("Chord Grid Visualization")
        self.configure(bg="white")
        
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
                    self.left_canvas.create_text(left_col_width - 8, y, text=fallback_text, anchor='e', font=("Arial", 12), fill="black")
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

        analyzer = EntropyAnalyzer({event_key: payload}, base=2, logger=lambda x: None)    
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
            entropy_review = self.parent.entropy_review_text if hasattr(self.parent, 'entropy_review_text') else "Entropy review information not available."
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
        analyzer = EntropyAnalyzer({event_key: event_data}, base=2, logger=lambda x: None)
        
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
                if hasattr(self.parent, 'loaded_file_path') and self.parent.loaded_file_path:
                    import os
                    filename = os.path.basename(self.parent.loaded_file_path)
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
        "": 40,       # pure major triad
        "m": 35,
        "maj7": 30,
        "mMaj7": 25,
    }

    def __init__(
        self,
        events: Dict[Tuple[int, int, str], Dict[str, Any]],
        symbol_mode: str = "chord",
        base: int = 2,
        logger: Callable[[str], None] = print
    ):
        self.events = events
        self.symbol_mode = symbol_mode
        self.base = base
        self.logger = logger
        self.custom_steps: List[Tuple[str, Callable[["EntropyAnalyzer"], None]]] = []

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
        rule_names = [f"R{i+1}" for i in range(7)]

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
                    rule6_bonus += 5
                    applied_rules.append(f"Rule 6: Previous event contained {chord} → +5")
                if dominant_chord in prev_event_chords:
                    rule6_bonus += 10
                    applied_rules.append(f"Rule 6: Previous event contained dominant {dominant_chord} → +10")
                base_score += rule6_bonus

                # Rule 4: proportional resolution
                if total_resolutions > 0:
                    ratio = resolution_count.get(root, 0) / total_resolutions
                    r4_bonus = 10 * ratio
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
            # Get base strength from _STRENGTH_MAP
            _, quality = self._split_chord(chord_name)
            base_strength = self._STRENGTH_MAP.get(quality or "", 0)
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
            "  R1 - Is the drive supported by the bass?\n"
            "  R2 - Is a tonal centre established? (N/A)\n"
            "  R3 - How often is the drive recurring?\n"
            "  R4 - How often is the drive discharging?\n"
            f"  R5 - Is the drive articulated as a clean stack in the texture? {CLEAN_STACK_SYMBOL}\n"
            "  R6 - Was the drive itself, or its dominant in the previous event?\n"
            f"  R7 - Is the root of the drive doubled at the octave? {ROOT2_SYMBOL}\n"
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
    def _compute_score(self, chord: str, basses: Optional[List[str]] = None, event_payload: Optional[dict] = None, root_counter: Optional[Dict[str, int]] = None) -> Tuple[int, List[str]]:
        root, quality = self._split_chord(chord)
        score = self._STRENGTH_MAP.get(quality or "", 0)
        messages: List[str] = []

        # Rule 1: Bass support
        if basses and root in basses:
            score += 20
            messages.append(f"Rule 1: Bass supports {chord} → +20 bonus")

        # Rule 3: root repetition (now always included)
        if root_counter is not None:
            prev_count = root_counter.get(root, 0)
            r3_bonus = 2 * prev_count if prev_count > 0 else 0
            if r3_bonus > 0:
                messages.append(f"Rule 3: Root {root} repeated → +{r3_bonus}")
            score += r3_bonus

        if event_payload is not None:
            chord_info = event_payload.get("chord_info", {})
            # Rule 5
            if chord_info.get(chord, {}).get("clean_stack"):
                score += 10
                messages.append(f"Rule 5: Clean chord {chord} → +10 bonus")
            # Rule 7
            root_count = chord_info.get(chord, {}).get("root_count", 1)
            if root_count == 2:
                score += 5
                messages.append(f"Rule 7: Root doubled in chord {chord} → +5 bonus")
            elif root_count >= 3:
                score += 10
                messages.append(f"Rule 7: Root tripled+ in chord {chord} → +10 bonus")

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
    app = MidiChordAnalyzer()
    app.mainloop()
