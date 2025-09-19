import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Text, BooleanVar, Frame, Label
import platform
import tkinter.font as tkfont
from PIL import Image, ImageDraw, ImageTk
import mido
import sys
import os
from math import log2
import threading
from collections import Counter
from typing import Callable, Dict, List, Optional, Tuple, Any, Set
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.colors import black, HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from music21 import converter, note, chord as m21chord, meter, stream
import subprocess
import sys
import os


def resource_path(relative_path: str) -> str:
    """Return absolute path to resource, working for dev and PyInstaller bundles."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    return os.path.join(base_path, relative_path)

# Marker symbols for chord analysis
CLEAN_STACK_SYMBOL = "✅"
ROOT2_SYMBOL = "²"
ROOT3_SYMBOL = "³"

def beautify_chord(chord: str) -> str:
    return chord.replace("b", "♭").replace("#", "♯")

NOTE_TO_SEMITONE = {
    'C': 0, 'C#': 1, 'Db': 1, 'D': 2, 'D#': 3, 'Eb': 3, 'E': 4,
    'F': 5, 'F#': 6, 'Gb': 6, 'G': 7, 'G#': 8, 'Ab': 8, 'A': 9,
    'A#': 10, 'Bb': 10, 'B': 11
}

CHORDS = {
    "C7": [0, 4, 7, 10], "C7b5": [0, 4, 6, 10], "C7#5": [0, 4, 8, 10],
    "Cm7": [0, 3, 7, 10], "Cø7": [0, 3, 6, 10], "C7m9noroot": [1, 4, 7, 10],
    "C7no3": [0, 7, 10], "C7no5": [0, 4, 10], "C7noroot": [4, 7, 10],
    "Caug": [0, 4, 8], "C": [0, 4, 7], "Cm": [0, 3, 7],
    "Cmaj7": [0, 4, 7, 11], "CmMaj7": [0, 3, 7, 11]
}

PRIORITY = [
    "C7", "C7b5", "C7#5", "Cm7", "Cø7", "C7m9noroot",
    "C7no3", "C7no5", "C7noroot", "Caug", "C", "Cm",
    "Cmaj7", "CmMaj7"
]

TRIADS = {"C", "Cm", "Caug"}
CIRCLE_OF_FIFTHS_ROOTS = ['Gb', 'Db', 'Ab', 'Eb', 'Bb', 'F', 'C', 'G', 'D', 'A', 'E', 'B', 'F#']

ENHARMONIC_EQUIVALENTS = {
    # Single sharps/flats
    'A#': 'Bb', 'Bb': 'Bb',
    'C#': 'Db', 'Db': 'Db',
    'D#': 'Eb', 'Eb': 'Eb',
    'F#': 'F#', 'Gb': 'F#',
    'G#': 'Ab', 'Ab': 'Ab',
    'E#': 'F',  'Fb': 'E',
    'B#': 'C',  'Cb': 'B',
    # Double sharps
    'A##': 'B', 'B##': 'C#', 'C##': 'D', 'D##': 'E', 'E##': 'F#', 'F##': 'G', 'G##': 'A',
    # Double flats
    'Abb': 'G', 'Bbb': 'A', 'Cbb': 'Bb', 'Dbb': 'C', 'Ebb': 'D', 'Fbb': 'Eb', 'Gbb': 'F',
    # Triple sharps (rare, but for completeness)
    'A###': 'B#', 'B###': 'C##', 'C###': 'D#', 'D###': 'E#', 'E###': 'F##', 'F###': 'G#', 'G###': 'A#',
    # Triple flats (rare)
    'Abbb': 'Gb', 'Bbbb': 'G', 'Cbbb': 'A', 'Dbbb': 'Bb', 'Ebbb': 'C', 'Fbbb': 'Db', 'Gbbb': 'Eb',
}

# Merge sensitivity params (tweak these)
# Central (medium) position presets are tuned so the slider position 3 matches these values.
# Jaccard threshold: how similar root-sets must be to consider merging (higher = stricter). If collapsing too much, raise this number.
MERGE_JACCARD_THRESHOLD = 0.60
# Bass overlap proportion required on the secondary merge path (0.0..1.0). If collapsing too much, lower this number.
MERGE_BASS_OVERLAP = 0.50
# How many bars apart events can be to allow merging (0 = same bar only)
MERGE_BAR_DISTANCE = 1
# Maximum number of differing roots allowed as a simple-diff test
MERGE_DIFF_MAX = 1

class LoadOptionsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.result = None  # <-- Fix: always define self.result
        self.include_triads_var = BooleanVar(value=True)
        self.sensitivity_var = tk.StringVar(value="Medium")
        self.selected_file = None
        self.build_ui()
# ...existing code...
    def build_ui(self):
        frame = tk.Frame(self, bg="#2b2b2b")
        frame.pack(padx=10, pady=10, fill="x")

        ttk.Checkbutton(
            frame, text="Include triads", variable=self.include_triads_var,
            style="White.TCheckbutton"
        ).pack(anchor="w", pady=5)

        ttk.Label(
            frame, text="Sensitivity level:", background="#2b2b2b", foreground="white"
        ).pack(anchor="w", pady=(10, 0))
        for level in ["High", "Medium", "Low"]:
            ttk.Radiobutton(
                frame, text=level, variable=self.sensitivity_var, value=level,
                style="White.TRadiobutton"
            ).pack(anchor="w")

        # Only allow XML files
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
    def __init__(self):
        super().__init__()
        self.title("🎵 MIDI Drive Analyzer")
        self.geometry("650x600")
        self.configure(bg="#2b2b2b")

        # Analysis options (defaults)
        self.include_triads = True
        self.sensitivity = "Medium"
        self.remove_repeats = True
        self.include_anacrusis = True
        self.include_non_drive_events = True
        self.arpeggio_searching = True
        # New: neighbour/passing notes detection (OFF by default)
        self.neighbour_notes_searching = False
        # Arpeggio vs. block-chord similarity threshold (0.0..1.0). If an arpeggio's pitch-class
        # set shares less than this proportion with the simultaneous block-chord, discard it.
        self.arpeggio_block_similarity_threshold = 0.5

        # Collapse similar events is always enabled; sensitivity controls how strict
        self.collapse_similar_events = True
        # Persisted slider position (1..5). Default center = 3
        self.collapse_sensitivity_pos = getattr(self, 'collapse_sensitivity_pos', 3)

        # Instance-level merge parameters (initialized from top-level constants)
        self.merge_jaccard_threshold = MERGE_JACCARD_THRESHOLD
        self.merge_bass_overlap = MERGE_BASS_OVERLAP
        self.merge_bar_distance = MERGE_BAR_DISTANCE
        self.merge_diff_max = MERGE_DIFF_MAX

        # Runtime state
        self.loaded_file_path = None
        self.score = None
        self.analyzed_events = None

        # Build UI and show splash
        self.build_ui()
        self.show_splash()



    def build_ui(self):
        is_mac = platform.system() == "Darwin"
        # Add extra top padding for macOS
        top_pad = 30 if is_mac else 10
        frame = Frame(self, bg="#2b2b2b")
        frame.pack(pady=(top_pad, 10))

        # Button style logic
        if is_mac:
            btn_kwargs = {}
            disabled_fg = "#cccccc"
        else:
            btn_kwargs = {"bg": "#ff00ff", "fg": "#fff", "activebackground": "#ff33ff", "activeforeground": "#fff", "relief": "raised", "bd": 2, "font": ("Segoe UI", 10, "bold")}
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

        self.result_text = Text(
            self, bg="#1e1e1e", fg="white", font=("Consolas", 11),
            wrap="word", borderwidth=0
        )
        self.result_text.pack(fill="both", expand=True, padx=10, pady=10)

    def create_piano_image(self, octaves=2, key_width=40, key_height=150):
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
        # Insert the title.png image centered
        try:
            from PIL import Image, ImageTk
            img_path = resource_path("title.png")
            title_img = Image.open(img_path)
            title_photo = ImageTk.PhotoImage(title_img)
            title_label = tk.Label(self.result_text, image=title_photo, bd=0)
            title_label.image = title_photo  # Keep a reference!
            text_width_chars = int(self.result_text['width'])
            img_width_chars = max(int(title_photo.width() / 8), 1)
            padding = max((text_width_chars - img_width_chars) // 2, 0)
            self.result_text.insert("end", " " * padding)
            self.result_text.window_create("end", window=title_label)
            self.result_text.insert("end", "\n\n")
        except Exception as e:
            self.result_text.insert("end", "Harmonic Drive Analyzer\n\n")
            print("Splash image load error:", e)
        # Now insert the piano image as before
        piano_img = self.create_piano_image(octaves=2)
        self.tk_piano_img = ImageTk.PhotoImage(piano_img)
        label = tk.Label(self.result_text, image=self.tk_piano_img, bd=0)
        label.image = self.tk_piano_img
        text_width_chars = int(self.result_text['width'])
        img_width_chars = max(int(self.tk_piano_img.width() / 8), 1)
        padding = max((text_width_chars - img_width_chars) // 2, 0)
        self.result_text.insert("end", " " * padding)
        self.result_text.window_create("end", window=label)
        self.result_text.insert("end", "\n\n")
        description = (
            "• Analyze MIDI / MusicXML files\n"
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
            print("[Phase7] No analyzed events yet.")
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
        self.loaded_file_path = path
        self.score = converter.parse(path)
        self.run_analysis()

    def run_analysis(self):
        # Use self.include_triads and self.sensitivity
        min_duration = {"High": 0.1, "Medium": 0.5, "Low": 1.0}[self.sensitivity]
        self.analyzed_events = None
        try:
            lines, events = self.analyze_musicxml(self.score, min_duration=min_duration)
            self.analyzed_events = events
            print("[DEBUG] analyzed_events keys before display:", list(self.analyzed_events.keys()))
            self.display_results()
            # Show entropy grid/log with legend after analysis
            print("\nENTROPY ANALYSIS\n")
            # --- Generate entropy review text and store it ---
            from io import StringIO
            entropy_buf = StringIO()
            analyzer = EntropyAnalyzer(self.analyzed_events, logger=lambda x: print(x, file=entropy_buf))
            analyzer.step_stage1_strengths(print_legend=True)
            self.entropy_review_text = entropy_buf.getvalue()
            self.show_grid_btn.config(state="normal")
            # Enable saving once an analysis exists (same behavior as Show Grid)
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

    def open_settings(self):
        """Open simple settings dialog and apply choices."""
        dialog = tk.Toplevel(self)
        dialog.title("Analysis Settings")
        dialog.geometry("420x380")

        # Option variables
        include_triads_var = tk.BooleanVar(value=self.include_triads)
        sensitivity_var = tk.StringVar(value=self.sensitivity)
        remove_repeats_var = tk.BooleanVar(value=self.remove_repeats)
        include_anacrusis_var = tk.BooleanVar(value=self.include_anacrusis)
        arpeggio_searching_var = tk.BooleanVar(value=self.arpeggio_searching)
        neighbour_notes_var = tk.BooleanVar(value=getattr(self, 'neighbour_notes_searching', True))
        include_non_drive_var = tk.BooleanVar(value=self.include_non_drive_events)

        # Slider for collapse sensitivity (1..5). Center (3) maps to the medium presets.
        sensitivity_scale_var = tk.IntVar(value=getattr(self, 'collapse_sensitivity_pos', 3))

        pad_opts = dict(anchor="w", padx=12, pady=6)
        ttk.Checkbutton(dialog, text="Include triads", variable=include_triads_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Include anacrusis", variable=include_anacrusis_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Arpeggio searching", variable=arpeggio_searching_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Neighbour notes", variable=neighbour_notes_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Remove repeated patterns", variable=remove_repeats_var).pack(**pad_opts)
        ttk.Checkbutton(dialog, text="Include non-drive events", variable=include_non_drive_var).pack(**pad_opts)

        ttk.Label(dialog, text="Collapse sensitivity (1=Low, 5=High):").pack(anchor="w", padx=12, pady=(8, 0))
        # Normal endpoints so left shows Low (1) and right shows High (5)
        sens_scale = tk.Scale(dialog, from_=1, to=5, orient="horizontal", variable=sensitivity_scale_var)
        sens_scale.pack(fill="x", padx=12)

        sens_frame = ttk.Frame(dialog)
        sens_frame.pack(anchor="w", padx=12, pady=(10, 12))
        ttk.Label(sens_frame, text="Note duration sensitivity: ").pack(side="left")
        for level in ["Low", "Medium", "High"]:
            ttk.Radiobutton(sens_frame, text=level, variable=sensitivity_var, value=level).pack(side="left", padx=6)

        def apply_settings():
            # Apply basic boolean/text settings
            self.include_triads = include_triads_var.get()
            self.sensitivity = sensitivity_var.get()
            self.remove_repeats = remove_repeats_var.get()
            self.include_anacrusis = include_anacrusis_var.get()
            self.arpeggio_searching = arpeggio_searching_var.get()
            self.neighbour_notes_searching = neighbour_notes_var.get()
            self.include_non_drive_events = include_non_drive_var.get()

            # Read slider position and persist it
            pos = int(sensitivity_scale_var.get())
            self.collapse_sensitivity_pos = pos
            # Collapsing is always enabled
            self.collapse_similar_events = True

            # Map slider position to merge parameter presets
            presets = {
                1: {"jaccard": 0.45, "bass": 0.25, "bar": 2, "diff": 2},
                2: {"jaccard": 0.55, "bass": 0.40, "bar": 1, "diff": 2},
                3: {"jaccard": 0.60, "bass": 0.50, "bar": 1, "diff": 1},
                4: {"jaccard": 0.70, "bass": 0.60, "bar": 1, "diff": 1},
                5: {"jaccard": 0.85, "bass": 0.70, "bar": 0, "diff": 0},
            }
            chosen = presets.get(pos, presets[3])
            self.merge_jaccard_threshold = chosen["jaccard"]
            self.merge_bass_overlap = chosen["bass"]
            self.merge_bar_distance = chosen["bar"]
            self.merge_diff_max = chosen["diff"]

            dialog.destroy()
            # If a score is loaded, re-run analysis so events are rebuilt and reprocessed with new merge params.
            if self.score:
                self.run_analysis()
            else:
                # If no score but we have existing analyzed_events (e.g., loaded from file), refresh display.
                if getattr(self, 'analyzed_events', None):
                    try:
                        # Re-display using current settings
                        self.display_results()
                    except Exception:
                        pass

            # If a GridWindow is currently open, refresh it to reflect the new settings/analysis.
            try:
                gw = getattr(self, '_grid_window', None)
                if gw and isinstance(gw, tk.Toplevel) and gw.winfo_exists():
                    try:
                        gw.destroy()
                    except Exception:
                        pass
                    self._grid_window = None
                    # Reopen the grid if we still have analyzed events
                    if getattr(self, 'analyzed_events', None):
                        try:
                            self.show_grid_window()
                        except Exception:
                            pass
            except Exception:
                pass

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
                for line in lines:
                    self.result_text.insert("end", line)
            elif events:
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
                    # Only show no-known-drive events when there are literally no other drives.
                    # If any drive exists anywhere in the displayed events, suppress non-drive entries.
                    if is_no_drive:
                        if has_any_drives:
                            continue
                        # If there are no drives at all, respect the include_non_drive_events flag
                        if not self.include_non_drive_events:
                            continue
                    output_lines.append(
                        f"Bar {bar}, Beat {beat} ({ts}): {chords_display} (bass = {bass})\n"
                    )
                    prev_no_drive = is_no_drive
                    prev_bass = bass
                output_lines.append(
                    f"\nLegend:\n{CLEAN_STACK_SYMBOL} = Clean stack   {ROOT2_SYMBOL} = Root doubled   {ROOT3_SYMBOL} = Root tripled or more\n"
                )
                self.result_text.insert("end", "".join(output_lines))
        self.result_text.config(state="disabled")


    def analyze_musicxml(self, score, min_duration=0.5):
        flat_notes = list(score.flat.getElementsByClass([note.Note, m21chord.Chord]))
        print("[DEBUG] flat_notes content:")
        for elem in flat_notes:
            if isinstance(elem, m21chord.Chord):
                print(f"  Chord at offset {elem.offset}: {[p.midi for p in elem.pitches]}")
            elif isinstance(elem, note.Note):
                print(f"  Note at offset {elem.offset}: {elem.pitch.midi}")
            else:
                print(f"  Other element at offset {getattr(elem, 'offset', '?')}: {type(elem)}")

        # Collect time signatures using the flattened score so offsets are absolute
        time_signatures = []
        for ts in score.flat.getElementsByClass(meter.TimeSignature):
            # ts.offset on a flat stream is the absolute offset in quarter lengths
            offset = float(ts.offset)
            time_signatures.append((offset, int(ts.numerator), int(ts.denominator)))

        # Ensure time signatures are sorted by offset
        time_signatures.sort(key=lambda x: x[0])

        # If no explicit time signature found, default to 4/4 at start
        if not time_signatures:
            time_signatures = [(0.0, 4, 4)]
        # If the first time signature doesn't start at 0, assume that meter applies from start
        elif time_signatures[0][0] > 0.0:
            first_num, first_den = time_signatures[0][1], time_signatures[0][2]
            time_signatures.insert(0, (0.0, first_num, first_den))

        # Debug: print collected time signatures and segments
        print(f"[DEBUG] collected time_signatures: {time_signatures}")
        for i, (t_off, num, den) in enumerate(time_signatures):
            next_off = time_signatures[i + 1][0] if i + 1 < len(time_signatures) else None
            beat_len = 4.0 / den
            if next_off is None:
                print(f"[DEBUG] timesig segment {i}: offset={t_off}, ts={num}/{den}, beat_len={beat_len}, next_off=None")
            else:
                seg_beats = (next_off - t_off) / beat_len
                print(f"[DEBUG] timesig segment {i}: offset={t_off}, ts={num}/{den}, beat_len={beat_len}, next_off={next_off}, segment_beats={seg_beats}")

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

        print(f"[DEBUG] time_points: {time_points}")
        print(f"[DEBUG] first 5 note_events: {note_events[:5]}")

        events = {}
        active_notes = set()
        active_pitches = set()

        # --- Block chord event-building (as before) ---
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
                for s_start, s_end, s_pitch in single_notes:
                    if s_end == time and (s_pitch % 12) not in test_notes:
                        test_notes.add(s_pitch % 12)
                        test_pitches.add(s_pitch)

            if len(test_notes) >= 3:
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

        # --- Arpeggio event-building (sliding window) ---
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
                        # Before accepting arpeggio detection, compare to any simultaneous block event
                        bar, beat, ts = offset_to_bar_beat(window[0].offset)
                        key = (bar, beat, ts)
                        block_pcs = events.get(key, {}).get('event_notes', set())
                        if block_pcs:
                            union = window_pcs | block_pcs
                            inter = window_pcs & block_pcs
                            jaccard = (len(inter) / len(union)) if union else 0.0
                            # Debug when F (pc=5) involved to inspect why it may be ignored
                            # (debug print removed)
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
                        events[key]["basses"].add(self.semitone_to_note(min(window_pitches) % 12))
                        events[key]["event_notes"] = set(window_pcs)
                        events[key]["event_pitches"] = set(window_pitches)

        # --- Neighbour / passing-note detection ---
        # Detect cases where two consecutive melodic notes together with sustained/supporting notes
        # form a full chord. This treats brief passing tones flanked by sustained notes as part of the chord.
        if getattr(self, 'neighbour_notes_searching', False):
            # single_notes: (start, end, pitch) from earlier
            # Build a sorted list of single-note entries by onset
            melodic_single = sorted(single_notes, key=lambda x: x[0])
            for i in range(len(melodic_single) - 1):
                s1, e1, p1 = melodic_single[i]
                s2, e2, p2 = melodic_single[i + 1]
                # Only consider notes that are consecutive (no intervening single with same or earlier onset)
                if s2 <= s1:
                    continue
                # Gather supporting pitches that are sounding at the time span covering both notes
                span_start = s1
                span_end = e2
                supporting_pcs = set()
                supporting_pitches = set()
                for st, en, prs in note_events:
                    # Treat as supporting if the pitch is sounding at either end of the span:
                    #  - struck earlier and still sounding at the span start, or
                    #  - struck before the span end and still sounding at or after the span end.
                    # This captures pitches that were struck just before the two-note window
                    # but sustained into it, while still excluding short passing notes that
                    # both start and end inside the span.
                    if (st <= span_start and en > span_start) or (st < span_end and en >= span_end):
                        supporting_pcs.update({pp % 12 for pp in prs})
                        supporting_pitches.update(prs)
                combined_pcs = {p1 % 12, p2 % 12} | supporting_pcs
                if len(combined_pcs) < 3:
                    continue
                # Try to detect chords from the combined pcs
                chords = self.detect_chords(combined_pcs, debug=False)
                if chords:
                    bar, beat, ts = offset_to_bar_beat(span_start)
                    key = (bar, beat, ts)
                    if key not in events:
                        events[key] = {"chords": set(), "basses": set(), "event_notes": set(combined_pcs)}
                    events[key]["chords"].update(chords)
                    events[key]["basses"].add(self.semitone_to_note(min(supporting_pitches | {p1, p2}) % 12))
                    events[key]["event_notes"] = set(combined_pcs)
                    events[key]["event_pitches"] = set(supporting_pitches | {p1, p2})
                    # (debug print removed)

        print(f"[DEBUG] Number of events built: {len(events)}")
        print(f"[DEBUG] Event keys: {list(events.keys())}")
        if not events:
            return ["No matching chords found."], {}

        return self._process_detected_events(events)

# ...rest of your code unchanged...

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

        Behavior:
        - Choose highest-priority chord per root when multiple chords agree.
        - Optionally collapse adjacent similar events into a single column when
          `self.collapse_similar_events` is True. Collapsing uses a Jaccard-like
          similarity and merges chords choosing the strongest-priority chord per root.
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

        # Optionally collapse adjacent similar events (merge columns)
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
                    merged_pitches = prev_pitches | cur_c
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
        """Detect chord names from a set of pitch-classes (semitones).

        If debug=True, print intermediate normalization and matched patterns to help
        diagnose why particular chord names were chosen.
        """
        if len(semitones) < 3:
            return []

        chords_found = []
        semitone_list = sorted(set(semitones))

        # First pass: try candidate roots that are present in the set
        for root in sorted(set(semitones)):
            normalized = {(n - root) % 12 for n in semitones}
            for name in PRIORITY:
                if name in TRIADS and not self.include_triads:
                    continue
                chord_pattern = set(CHORDS[name])
                if chord_pattern.issubset(normalized):
                    matched = name.replace('C', self.semitone_to_note(root))
                    chords_found.append(matched)
                    break

        # Second pass: try "noroot" style chords where the root pitch-class is absent
        for root in sorted(set(range(12)) - set(semitones)):
            normalized = {(n - root) % 12 for n in semitones}
            for name in PRIORITY:
                if "noroot" not in name:
                    continue
                if name in TRIADS and not self.include_triads:
                    continue
                chord_pattern = set(CHORDS[name])
                if chord_pattern == normalized:
                    matched = name.replace('C', self.semitone_to_note(root))
                    chords_found.append(matched)
                    break

        return chords_found

    def semitone_to_note(self, semitone):
        for note in NOTE_TO_SEMITONE:
            if NOTE_TO_SEMITONE[note] == semitone and len(note) == 1:
                return note
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
                        chord = chord_entry.replace(CLEAN_STACK_SYMBOL, "").replace(ROOT2_SYMBOL, "").replace(ROOT3_SYMBOL, "")
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
            self.show_grid_btn.config(state="normal")
            try:
                self.save_analysis_btn.config(state="normal")
            except Exception:
                pass
            tk.messagebox.showinfo("Loaded", f"Analysis loaded from {file_path}")
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to load analysis:\n{e}")
            

    def show_grid_window(self):
        if not self.analyzed_events:
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

        gw = GridWindow(self, self.analyzed_events)
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
        try:
            import pygame
            import pygame.midi
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

        # Checkbox with a clear visual indicator (green = ON, gray = OFF)
        chk_frame = tk.Frame(controls_frame, bg="black")
        chk_frame.pack(side="left", padx=10)

        self.checkbox = tk.Checkbutton(
            chk_frame, text="Include triads", variable=self.include_triads_var,
            font=("Segoe UI", 11), fg="white", bg="black", selectcolor="#ffd6ff",
            activebackground="black", activeforeground="white",
            bd=0, highlightthickness=0, anchor="w",
            command=self.analyze_chord  # refresh analysis when toggled
        )
        self.checkbox.pack(side="left")

        # Small colored square indicator to show ON/OFF clearly (works across themes)
        self._triad_indicator = tk.Canvas(chk_frame, width=16, height=16, bg="black", highlightthickness=0)
        self._indicator_rect = self._triad_indicator.create_rectangle(
            2, 2, 14, 14,
            fill="#00cc66" if self.include_triads_var.get() else "#444444",
            outline=""
        )
        self._triad_indicator.pack(side="left", padx=(6,0), pady=2)

        def _update_triads_indicator(*args):
            state = self.include_triads_var.get()
            color = "#00cc66" if state else "#444444"
            try:
                self._triad_indicator.itemconfig(self._indicator_rect, fill=color)
            except Exception:
                # Fallback: recreate rectangle if needed
                try:
                    self._triad_indicator.delete("all")
                except Exception:
                    pass
                self._indicator_rect = self._triad_indicator.create_rectangle(2, 2, 14, 14, fill=color, outline="")

        # Attach variable trace (compatible with older tkinter versions)
        try:
            self.include_triads_var.trace_add("write", _update_triads_indicator)
        except Exception:
            self.include_triads_var.trace("w", lambda *a: _update_triads_indicator())

        self.clear_button = tk.Button(
            controls_frame, text="Clear", font=("Segoe UI", 11, "bold"),
            bg="#444444", fg="white", activebackground="#666666", activeforeground="white",
            bd=0, padx=12, pady=5, command=self._clear_selection
        )
        self.clear_button.pack(side="left", padx=10)

        # MIDI Dropdown (if mido available)
        midi_frame = tk.Frame(self.parent, bg="black")
        midi_frame.pack(pady=5)
        tk.Label(midi_frame, text="MIDI Input:", font=("Segoe UI", 10), fg="white", bg="black").pack(side="left", padx=(0,5))
        try:
            import mido
            self.midi_ports = mido.get_input_names()
        except Exception:
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
        try:
            import mido
        except Exception:
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


class GridWindow(tk.Toplevel):#
    
    
    SUPERSCRIPT_MAP = {
        'no1': "ⁿᵒ¹",
        'no3': "ⁿᵒ³",
        'no5': "ⁿᵒ⁵",
        'noroot': "ⁿᵒ¹",  # optional alias for clarity
    }

    CELL_SIZE = 50
    PADDING = 40

    # For on-screen (Tkinter)
    CHORD_TYPE_COLORS_TK = {
        "maj": "#FFFFFF",
        "min": "#FFFFFF",
        "7": "#666666",
        "dim": "#CCCCCC",
        "aug": "#CCCCCC",
        "maj7": "#FFFFFF",
        "m7": "#AAAAAA",
        "ø7": "#AAAAAA",
        "no": "#CCCCCC",          # ← new
        "other": "#FFFFFF",
    }

    # For PDF (ReportLab)
    CHORD_TYPE_COLORS_PDF = {
        k: HexColor(v) for k, v in CHORD_TYPE_COLORS_TK.items()
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

        self.parent = parent
        # Apply same filtering as main window (respect include_non_drive_events)
        raw_events = {k: v for k, v in events.items()} if events else {}
        if hasattr(parent, 'include_non_drive_events') and not parent.include_non_drive_events:
            raw_events = {k: v for k, v in raw_events.items() if v.get('chords') and len(v['chords']) > 0}

        # Deduplicate repeated patterns the same way the main display does
        self.events = self._dedupe_for_grid(raw_events)
        self.sorted_events = sorted(self.events.keys())

        # Remove Gb row from the circle of fifths for this grid
        self.root_list = [r for r in CIRCLE_OF_FIFTHS_ROOTS if r != 'Gb']
        self.root_to_row = {root: i for i, root in enumerate(self.root_list)}

        canvas_width = self.PADDING * 2 + len(self.sorted_events) * self.CELL_SIZE
        canvas_height = self.PADDING * 2 + len(self.root_list) * self.CELL_SIZE

        # --- Controls frame ---
        controls_frame = ttk.Frame(self)
        controls_frame.pack(side="top", fill="x", pady=5)

        # Create all controls and pack them in a row
        self.show_resolutions_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            controls_frame,
            text="Show Resolution Patterns",
            variable=self.show_resolutions_var,
            command=self.redraw
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
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        # Left (frozen) column canvas for root labels (wider to fit enharmonic alternatives)
        left_col_width = max(100, self.PADDING * 3)
        self.left_canvas = tk.Canvas(container, width=left_col_width, height=canvas_height, bg="white", highlightthickness=0)
        self.left_canvas.pack(side="left", fill="y")

        # Right scrollable area for the grid
        right_frame = ttk.Frame(container)
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
        # Probe available fonts once to choose best rendering for accidentals
        try:
            # tkfont.families() returns an iterable of family-name strings
            available_fonts = set(tkfont.families())
        except Exception:
            # If probing fails (rare on some embedded Tk builds), fall back to empty set
            available_fonts = set()
        prefer_dejavu = any(name.lower().startswith('dejavu') for name in available_fonts)
        segui_available = any('segoe' in name.lower() for name in available_fonts)

        for root, row in self.root_to_row.items():
            y = self.PADDING + row * self.CELL_SIZE + self.CELL_SIZE // 2
            label_text = enh_map.get(root, root)
            label_text = label_text.replace('b', '♭').replace('#', '♯')

            # Choose font: prefer DejaVuSans when accidentals are present (better glyph coverage),
            # otherwise prefer Segoe UI for system consistency. Fall back to default Tk font.
            if ('♭' in label_text or '♯' in label_text) and prefer_dejavu:
                family = next((n for n in available_fonts if n.lower().startswith('dejavu')), 'DejaVu Sans')
            elif segui_available:
                family = next((n for n in available_fonts if 'segoe' in n.lower()), 'Segoe UI')
            else:
                family = tkfont.nametofont('TkDefaultFont').cget('family')

            try:
                font_choice = (family, 14)
                self.left_canvas.create_text(left_col_width - 8, y, text=label_text, anchor='e', font=font_choice, fill="black")
            except Exception:
                # If font creation or drawing fails for any reason, log and fallback to default font
                try:
                    self.left_canvas.create_text(left_col_width - 8, y, text=label_text, anchor='e', fill="black")
                except Exception as ex:
                    import traceback
                    print("[ERROR] Failed to create left-column label for root:", root, "label:", label_text)
                    traceback.print_exc()
        
        #Inside your GridWindow __init__ method or GUI setup:
 
    def toggle_entropy(self):
        if self.show_entropy_var.get():
            print("Entropy graph should appear here!")
        else:
            print("Entropy graph should be hidden!")    

    # Inside GridWindow

    def show_entropy_info_window(self, entropy_text):
        info_win = tk.Toplevel(self)
        info_win.title("Entropy Review")
        info_win.configure(bg="white")
        info_win.geometry("1000x400")
        text_widget = tk.Text(info_win, wrap="word", bg="white", fg="black", font=("Consolas", 11))
        text_widget.insert("1.0", entropy_text)
        text_widget.config(state="disabled")
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        def save_entropy_info():
            from tkinter import filedialog
            path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], title="Save Entropy Info")
            if path:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(entropy_text)
        save_btn = tk.Button(info_win, text="Save", command=save_entropy_info, bg="#ff00ff", fg="#fff", font=("Segoe UI", 10, "bold"))
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

    def export_pdf(self):
        use_color = self.color_pdf_var.get()
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.colors import black, HexColor
        from reportlab.pdfgen import canvas as pdf_canvas

        import os, math
        # Prefer to use bundled DejaVu fonts placed into assets/fonts by the CI build
        try:
            assets_fonts_dir = resource_path(os.path.join('assets', 'fonts'))
            dejavu_candidates = []
            if os.path.isdir(assets_fonts_dir):
                for fn in os.listdir(assets_fonts_dir):
                    if fn.lower().startswith('dejavu') and fn.lower().endswith('.ttf'):
                        dejavu_candidates.append(os.path.join(assets_fonts_dir, fn))
            # Fallback to older expected folder name if CI used a different layout
            if not dejavu_candidates:
                legacy_path = resource_path(os.path.join('dejavu-fonts-ttf-2.37', 'ttf', 'DejaVuSans.ttf'))
                if os.path.exists(legacy_path):
                    dejavu_candidates.append(legacy_path)

            # Prefer a non-italic/oblique candidate when possible (avoid files like "-Oblique" or "-Italic").
            pdf_font_name = 'Helvetica'  # fallback
            registered_path = None
            if dejavu_candidates:
                chosen = None
                for p in dejavu_candidates:
                    lower = os.path.basename(p).lower()
                    if ('oblique' in lower) or ('italic' in lower):
                        continue
                    # prefer explicit DejaVuSans.ttf filename
                    if 'dejavusans.ttf' in lower or 'dejavu' in lower:
                        chosen = p
                        break
                if not chosen:
                    # If only italic/oblique variants present, fall back to first candidate
                    chosen = dejavu_candidates[0]

                try:
                    # Register under a deterministic internal name so we can force usage
                    pdf_font_name = 'DejaVuForce'
                    pdfmetrics.registerFont(TTFont(pdf_font_name, chosen))
                    registered_path = chosen
                except Exception:
                    # registration failed; fall back to Helvetica
                    pdf_font_name = 'Helvetica'

            # Debug: print which font (if any) was registered for PDF output
            try:
                if registered_path:
                    print(f"[DEBUG] Registered PDF font '{pdf_font_name}' from: {registered_path}")
                else:
                    print("[DEBUG] No bundled DejaVu font registered for PDF; using fallback fonts.")
            except Exception:
                pass
        except Exception:
            # If resource_path or os.listdir fails, ignore and continue with defaults
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

            radius = int(cell_size * 0.85 / 2)

            entropies_all = {ek: self.compute_entropy(event_key=ek) for ek in self.sorted_events}

            for page in range(num_pages):
                start_col = page * page_grid_cols
                end_col = min(start_col + page_grid_cols, grid_cols)
                visible_events = self.sorted_events[start_col:end_col]
                visible_cols = len(visible_events)

                # Row labels + horizontal grid lines
                for root, row in self.root_to_row.items():
                    y_center = height - (margin_y + row * cell_size + cell_size / 2)
                    c.setFont(pdf_font_name, 12)
                    enh_map = {'F#': 'F#/Gb', 'Db': 'Db/C#', 'Ab': 'Ab/G#', 'Eb': 'Eb/D#'}
                    label_raw = enh_map.get(root, root)
                    note_label = label_raw.replace('b', '♭').replace('#', '♯')
                    c.drawRightString(margin_left - 8, y_center + 4, note_label)

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
                        fill_color = self.CHORD_TYPE_COLORS_PDF.get(chord_type, HexColor("#CCCCCC")) if use_color else HexColor("#FFFFFF")
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
                            c.circle(x, y, radius, stroke=1, fill=1 if use_color else 0)

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
                            text_color = HexColor("#FFFFFF") if chord_type == "7" else HexColor("#000000")
                            c.setFillColor(text_color)
                            c.setFont(pdf_font_name, 8)
                            c.drawCentredString(x, y - 4, function_label)

                # Bass dots
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
                        c.setFillColor(black)
                        c.circle(bx, by - radius + dot_radius, dot_radius, fill=1, stroke=0)

                # Optional resolution arrows
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

                    arrow_length = radius
                    for (col, row), (x1, y1, chord1) in pos_dict.items():
                        diag_pos = (col + 1, row + 1)
                        if diag_pos in pos_dict:
                            x2, y2, chord2 = pos_dict[diag_pos]
                            dx, dy = x2 - x1, y2 - y1
                            dist = (dx**2 + dy**2) ** 0.5
                            if dist == 0:
                                continue
                            dx_norm, dy_norm = dx / dist, dy / dist
                            start_x = x1 + dx_norm * arrow_length
                            start_y = y1 + dy_norm * arrow_length
                            end_x = x2 - dx_norm * arrow_length
                            end_y = y2 - dy_norm * arrow_length
                            c.setStrokeColor(black)
                            c.setLineWidth(2)
                            c.line(start_x, start_y, end_x, end_y)
                            arrow_size = 6
                            angle = math.atan2(dy, dx)
                            left_angle = angle + math.pi / 6
                            right_angle = angle - math.pi / 6
                            left_x = end_x - arrow_size * math.cos(left_angle)
                            left_y = end_y - arrow_size * math.sin(left_angle)
                            right_x = end_x - arrow_size * math.cos(right_angle)
                            right_y = end_y - arrow_size * math.sin(right_angle)
                            c.line(end_x, end_y, left_x, left_y)
                            c.line(end_x, end_y, right_x, right_y)

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
                c.drawCentredString(width / 2, 20, f"Page {page + 1} of {num_pages}")

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
        radius = int(self.CELL_SIZE * 0.85 / 2)

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

                    chord_type = self.classify_chord_type(chord)
                    fill_color = self.CHORD_TYPE_COLORS_TK.get(chord_type, "#CCCCCC") if self.color_pdf_var.get() else "white"

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

        # Draw resolution pattern arrows if enabled (unchanged)
        if self.show_resolutions_var.get():
            pos_dict = {(col, row): (x, y, chord) for col, row, x, y, chord in self.chord_positions}
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
                    start_x = x1 + dx_norm * radius
                    start_y = y1 + dy_norm * radius
                    end_x = x2 - dx_norm * radius
                    end_y = y2 - dy_norm * radius
                    self.canvas.create_line(start_x, start_y, end_x, end_y, arrow=tk.LAST, fill="black", width=2)

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
    Phase 7 entropy analysis on analyzed_events from MidiChordAnalyzer.
    Produces both human-readable per-event chord strength listing (Stage 1)
    and numeric-sequence entropy calculation (Stage 2).
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
        chromatic_sharps = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
        flats_to_sharps = {'Db': 'C#', 'Eb': 'D#', 'Fb': 'E', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#', 'Cb': 'B'}
        note = flats_to_sharps.get(root, root)
        if note not in chromatic_sharps:
            self.logger(f"[Warning] _fourth_up: unknown note {root}")
            return root
        index = chromatic_sharps.index(note)
        fourth_index = (index + 5) % 12  # perfect fourth = +5 semitones
        return chromatic_sharps[fourth_index]

    def _fifth_up(self, root: str) -> str:
        """Return the note a perfect fifth above the given root."""
        chromatic_sharps = ['C', 'C#', 'D', 'D#', 'E', 'F', 'F#', 'G', 'G#', 'A', 'A#', 'B']
        flats_to_sharps = {'Db': 'C#', 'Eb': 'D#', 'Fb': 'E', 'Gb': 'F#', 'Ab': 'G#', 'Bb': 'A#', 'Cb': 'B'}
        note = flats_to_sharps.get(root, root)
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

        # Print table
        col_widths = [30, 6] + [4]*7 + [6]
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
