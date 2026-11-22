"""Reusable na piraso ng UI (stat card, file slot, section title, card)."""
import os
import tkinter as tk

from ui import theme as T


def page_header(parent, title, subtitle):
    """Malaking title + maliit na subtitle. Returns (frame, title_box) para madagdagan sa kanan."""
    head = tk.Frame(parent, bg=T.BG)
    head.pack(fill="x")
    left = tk.Frame(head, bg=T.BG)
    left.pack(side="left")
    tk.Label(left, text=title, bg=T.BG, fg="#111", font=(T.FONT, 26, "bold")).pack(anchor="w")
    tk.Label(left, text=subtitle, bg=T.BG, fg="#bbb", font=(T.FONT, 10)).pack(anchor="w")
    return head


def section_title(parent, icon, title):
    tk.Label(parent, text=f"{icon}  {title}", bg=T.BG, fg="#111",
             font=(T.FONT, 12, "bold")).pack(anchor="w", pady=(8, 4))


def card(parent):
    c = tk.Frame(parent, bg=T.BG, highlightbackground=T.BORDER, highlightthickness=1)
    c.pack(fill="x")
    return c


class StatCard(tk.Frame):
    """Bilog na icon + numero + label (Student Processed, etc.)."""

    def __init__(self, parent, text, color, glyph):
        super().__init__(parent, bg=T.BG)
        c = tk.Canvas(self, width=40, height=40, bg=T.BG, highlightthickness=0)
        c.create_oval(2, 2, 38, 38, fill=color, outline=color)
        c.create_text(20, 20, text=glyph, fill="white", font=(T.FONT, 13, "bold"))
        c.pack(side="left", padx=(0, 8))
        box = tk.Frame(self, bg=T.BG)
        box.pack(side="left")
        self.var = tk.StringVar(value="0")
        tk.Label(box, textvariable=self.var, bg=T.BG, font=(T.FONT, 18, "bold")).pack(anchor="w")
        tk.Label(box, text=text, bg=T.BG, fg=T.MUTED, font=(T.FONT, 8)).pack(anchor="w")

    def set(self, value):
        self.var.set(str(value))


class FileSlot(tk.Frame):
    """Title (nagiging pangalan ng file) + hint + Browse button."""

    def __init__(self, parent, default_title, var, btn_text, command, divider=True):
        super().__init__(parent, bg=T.BG)
        self.pack(fill="x")
        inner = tk.Frame(self, bg=T.BG)
        inner.pack(fill="x", padx=16, pady=8)
        self.title = tk.Label(inner, text="🗎  " + default_title, bg=T.BG,
                              font=(T.FONT, 13), anchor="w")
        self.title.pack(anchor="w")
        tk.Label(inner, text="Select File to This Area", bg=T.BG, fg=T.MUTED,
                 font=(T.FONT, 8)).pack(anchor="w")
        tk.Button(inner, text=btn_text, bg=T.GRAY, fg="#555", relief="flat",
                  font=(T.FONT, 8), cursor="hand2", command=command
                  ).pack(anchor="w", pady=4, ipadx=14)
        if divider:
            tk.Frame(self, bg=T.DIVIDER, height=1).pack(fill="x", padx=16)

        def refresh(*_):
            p = var.get()
            self.title.config(text="🗎  " + (os.path.basename(p) if p else default_title))
        var.trace_add("write", refresh)
        refresh()
