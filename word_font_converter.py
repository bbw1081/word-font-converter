"""
Word Font Converter
--------------------
An accessibility tool that toggles a Word document between a large,
easy-to-read font and a standard font. Works on old .doc files and
modern .docx files alike, because it drives Microsoft Word directly
(via COM) rather than reading the file format itself.

Copyright (C) 2026  Richard Wilkinson

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser, ttk

print("<program>  Copyright (C) <year>  <name of author> \n" \
"This program comes with ABSOLUTELY NO WARRANTY; for details type `show w'. \n" \
"This is free software, and you are welcome to redistribute it\n" \
"under certain conditions; type `show c' for details.")

try:
    import win32com.client as win32
except ImportError:
    win32 = None


def app_dir():
    """Folder the exe/script lives in, so settings persist next to it."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), "font_settings.json")

DEFAULT_CONFIG = {
    "large": {"font_name": "Arial", "font_size": 28, "color": [0, 128, 0], "margin_in": 1.5},
    "standard": {"font_name": "Calibri", "font_size": 11, "color": [0, 0, 0], "margin_in": 1.0},
    "last_file": "",
}


def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                data = json.load(f)
            for key, val in DEFAULT_CONFIG.items():
                data.setdefault(key, val)
            for profile_key in ("large", "standard"):
                for sub_key, sub_val in DEFAULT_CONFIG[profile_key].items():
                    data[profile_key].setdefault(sub_key, sub_val)
            return data
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_CONFIG))


def save_config(cfg):
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)


def rgb_to_word_color(rgb):
    r, g, b = rgb
    return r + (g << 8) + (b << 16)


def convert_document(filepath, profile):
    """Open filepath in Word, apply the font profile to the whole
    document -- including the Normal style, so any new typing after
    conversion matches too -- along with the profile's page margins,
    then save (overwriting the original) and close."""
    if win32 is None:
        raise RuntimeError("pywin32 is not installed. Run: pip install pywin32")

    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(filepath))
        try:
            color = rgb_to_word_color(profile["color"])

            normal_style = doc.Styles("Normal").Font
            normal_style.Name = profile["font_name"]
            normal_style.Size = profile["font_size"]
            normal_style.Color = color

            rng = doc.Content
            rng.Font.Name = profile["font_name"]
            rng.Font.Size = profile["font_size"]
            rng.Font.Color = color

            # Word's PageSetup margins are in points; 1 inch = 72 points.
            margin_pts = profile["margin_in"] * 72
            for section in doc.Sections:
                section.PageSetup.TopMargin = margin_pts
                section.PageSetup.BottomMargin = margin_pts
                section.PageSetup.LeftMargin = margin_pts
                section.PageSetup.RightMargin = margin_pts

            doc.Save()
        finally:
            doc.Close()
    finally:
        word.Quit()


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, cfg, on_save):
        super().__init__(parent)
        self.title("Font Settings")
        self.resizable(False, False)
        self.transient(parent)
        self.cfg = cfg
        self.on_save = on_save
        self.vars = {}

        row = 0
        for profile_key, label in (("large", "Large Print"), ("standard", "Standard")):
            ttk.Label(self, text=label, font=("Segoe UI", 10, "bold")).grid(
                row=row, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 2)
            )
            row += 1

            ttk.Label(self, text="Font name:").grid(row=row, column=0, sticky="w", padx=10)
            name_var = tk.StringVar(value=cfg[profile_key]["font_name"])
            ttk.Entry(self, textvariable=name_var, width=20).grid(row=row, column=1, columnspan=2, padx=5, pady=2)
            self.vars[f"{profile_key}_name"] = name_var
            row += 1

            ttk.Label(self, text="Font size:").grid(row=row, column=0, sticky="w", padx=10)
            size_var = tk.IntVar(value=cfg[profile_key]["font_size"])
            ttk.Spinbox(self, from_=6, to=96, textvariable=size_var, width=6).grid(row=row, column=1, sticky="w", padx=5, pady=2)
            self.vars[f"{profile_key}_size"] = size_var
            row += 1

            ttk.Label(self, text="Color:").grid(row=row, column=0, sticky="w", padx=10)
            swatch = tk.Label(self, text="        ", bg=self._rgb_hex(cfg[profile_key]["color"]), relief="sunken")
            swatch.grid(row=row, column=1, sticky="w", padx=5, pady=2)
            self.vars[f"{profile_key}_color"] = cfg[profile_key]["color"]
            self.vars[f"{profile_key}_swatch"] = swatch

            def make_picker(pk=profile_key):
                def pick():
                    initial = self._rgb_hex(self.vars[f"{pk}_color"])
                    result = colorchooser.askcolor(color=initial, title="Choose color", parent=self)
                    if result[0]:
                        rgb = [int(c) for c in result[0]]
                        self.vars[f"{pk}_color"] = rgb
                        self.vars[f"{pk}_swatch"].config(bg=self._rgb_hex(rgb))
                    # Old Tk on Windows 7 can leave the window unpainted
                    # after the native color dialog closes -- force a
                    # full repaint and restore focus so it's usable.
                    self.lift()
                    self.focus_force()
                    self.update_idletasks()
                    self.update()
                return pick

            ttk.Button(self, text="Choose...", command=make_picker()).grid(row=row, column=2, padx=5, pady=2)
            row += 1

            ttk.Label(self, text="Margin (inches):").grid(row=row, column=0, sticky="w", padx=10)
            margin_var = tk.DoubleVar(value=cfg[profile_key]["margin_in"])
            ttk.Spinbox(
                self, from_=0.25, to=3.0, increment=0.25, textvariable=margin_var, width=6
            ).grid(row=row, column=1, sticky="w", padx=5, pady=2)
            self.vars[f"{profile_key}_margin"] = margin_var
            row += 1

        ttk.Button(self, text="Save", command=self._save).grid(row=row, column=0, columnspan=3, pady=12)

    @staticmethod
    def _rgb_hex(rgb):
        return "#%02x%02x%02x" % tuple(rgb)

    def _save(self):
        for pk in ("large", "standard"):
            self.cfg[pk]["font_name"] = self.vars[f"{pk}_name"].get()
            self.cfg[pk]["font_size"] = self.vars[f"{pk}_size"].get()
            self.cfg[pk]["color"] = self.vars[f"{pk}_color"]
            self.cfg[pk]["margin_in"] = self.vars[f"{pk}_margin"].get()
        save_config(self.cfg)
        self.on_save()
        self.destroy()


LABEL_FONT = ("Segoe UI", 16)
ENTRY_FONT = ("Segoe UI", 16)
BUTTON_FONT = ("Segoe UI", 16, "bold")
STATUS_FONT = ("Segoe UI", 20, "bold")


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word Font Converter")
        self.geometry("1600x800")
        self.resizable(True, True)

        # Scale up every ttk widget's default font/padding so the whole
        # UI reads comfortably for someone with low vision.
        style = ttk.Style(self)
        style.configure("TButton", font=BUTTON_FONT, padding=14)
        style.configure("TLabel", font=LABEL_FONT)
        style.configure("TEntry", font=ENTRY_FONT, padding=8)

        self.converted_path = None

        self.cfg = load_config()
        self.filepath = tk.StringVar(value=self.cfg.get("last_file", ""))

        ttk.Label(self, text="Word document:").pack(anchor="w", padx=20, pady=(20, 6))
        row = ttk.Frame(self)
        row.pack(fill="x", padx=20)
        ttk.Entry(row, textvariable=self.filepath).pack(side="left", fill="x", expand=True, ipady=6)
        ttk.Button(row, text="Browse...", command=self.browse).pack(side="left", padx=(10, 0))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=35)
        ttk.Button(
            btn_frame, text="Convert to\nLARGE Print", command=lambda: self.convert("large")
        ).grid(row=0, column=0, padx=15, ipadx=20, ipady=20)
        ttk.Button(
            btn_frame, text="Convert to\nStandard Print", command=lambda: self.convert("standard")
        ).grid(row=0, column=1, padx=15, ipadx=20, ipady=20)

        open_row = ttk.Frame(self)
        open_row.pack(pady=(0, 20))
        self.open_btn = ttk.Button(
            open_row, text="Open Converted File", command=self.open_converted_file, state="disabled"
        )
        self.open_btn.grid(row=0, column=0, padx=8, ipadx=10, ipady=10)
        self.open_folder_btn = ttk.Button(
            open_row, text="Open File Location", command=self.open_converted_folder, state="disabled"
        )
        self.open_folder_btn.grid(row=0, column=1, padx=8, ipadx=10, ipady=10)

        ttk.Button(self, text="Edit Font Settings", command=self.open_settings).pack(pady=(0, 15))

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(
            self, textvariable=self.status, foreground="#1a1a1a", font=STATUS_FONT, anchor="center", wraplength=1500
        ).pack(pady=(0, 20), fill="x", padx=20)

    def set_status(self, text):
        """Clear the status text and force a redraw before showing the
        new message, so old text never lingers behind the new text."""
        self.status.set("")
        self.update_idletasks()
        self.status.set(text)
        self.update_idletasks()

    def browse(self):
        path = filedialog.askopenfilename(
            title="Select a Word document",
            filetypes=[("Word documents", "*.docx *.doc"), ("All files", "*.*")],
        )
        if path:
            self.filepath.set(path)
            self.cfg["last_file"] = path
            save_config(self.cfg)

    def open_settings(self):
        SettingsDialog(self, self.cfg, on_save=lambda: self.set_status("Settings saved."))

    def convert(self, profile):
        path = self.filepath.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("No file", "Please choose a valid Word document first.")
            return
        # Grey out the open buttons for the duration of the conversion so
        # they can't be clicked while the file is being rewritten.
        self.open_btn.config(state="disabled")
        self.open_folder_btn.config(state="disabled")
        self.update_idletasks()
        self.set_status(f"Converting to {profile}...")
        try:
            convert_document(path, self.cfg[profile])
            self.set_status(f"Done \u2014 converted to {profile} print.")
            self.converted_path = path
            self.open_btn.config(state="normal")
            self.open_folder_btn.config(state="normal")
        except Exception as e:
            self.set_status("Error.")
            messagebox.showerror("Conversion failed", str(e))
            # If a previous successful conversion left a valid file behind,
            # restore the open buttons rather than leaving them stuck off.
            if self.converted_path and os.path.exists(self.converted_path):
                self.open_btn.config(state="normal")
                self.open_folder_btn.config(state="normal")

    def open_converted_file(self):
        if not self.converted_path or not os.path.exists(self.converted_path):
            messagebox.showerror("File not found", "The converted file could not be located.")
            return
        try:
            os.startfile(self.converted_path)
        except Exception as e:
            messagebox.showerror("Couldn't open file", str(e))

    def open_converted_folder(self):
        if not self.converted_path or not os.path.exists(self.converted_path):
            messagebox.showerror("File not found", "The converted file could not be located.")
            return
        try:
            # Opens Explorer with the converted file selected/highlighted
            os.startfile(os.path.dirname(self.converted_path))
        except Exception as e:
            messagebox.showerror("Couldn't open folder", str(e))


if __name__ == "__main__":
    MainApp().mainloop()
