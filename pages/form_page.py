"""FORM PAGE - isang UI na gamit ng SF10 at Form 9. Ang lohika ay nasa controllers/base_controller.py."""
import os
import tkinter as tk
from tkinter import ttk, messagebox

from pages.base_page import BasePage
from ui import theme as T
from ui.widgets import FileSlot, StatCard, card, page_header, section_title


class FormPage(BasePage):

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.c = controller
        self._style()

        # header + Form / Settings pill
        head = page_header(self, self.c.page_title, "Excel to Unified")
        pill = tk.Frame(head, bg=T.BG, highlightbackground="#888", highlightthickness=1)
        pill.pack(side="right", anchor="n")
        self.tab_form = tk.Label(pill, text="Form", font=(T.FONT, 8), padx=18, pady=2, cursor="hand2")
        self.tab_set = tk.Label(pill, text="Settings", font=(T.FONT, 8), padx=12, pady=2, cursor="hand2")
        self.tab_form.pack(side="left")
        self.tab_set.pack(side="left")
        self.tab_form.bind("<Button-1>", lambda e: self.show_tab("form"))
        self.tab_set.bind("<Button-1>", lambda e: self.show_tab("settings"))

        # stats
        stats = tk.Frame(self, bg=T.BG)
        stats.pack(fill="x", pady=(8, 6))
        self.stat_cards = {}
        for i, (key, text, color, glyph) in enumerate([
            ("processed", "Student Processed", T.GREEN, "✓"),
            ("success", "Succesful Transfer", "#c9c9c9", "↗"),
            ("missing", "Missing Data", "#3b5bdb", "✔"),
            ("failed", "Failed Sheet", "#e0d90f", "◔"),
        ]):
            sc = StatCard(stats, text, color, glyph)
            sc.grid(row=0, column=i, sticky="w", padx=(0, 50))
            self.stat_cards[key] = sc

        # body
        body = tk.Frame(self, bg=T.BG)
        body.pack(fill="both", expand=True)
        self.form_frame = tk.Frame(body, bg=T.BG)
        self.settings_frame = tk.Frame(body, bg=T.BG)
        self._build_form(self.form_frame)
        self._build_settings(self.settings_frame)
        self.show_tab("form")

        # bottom buttons
        bar = tk.Frame(self, bg=T.BG)
        bar.pack(fill="x", pady=(8, 0))
        self.btn_start = tk.Button(bar, text="▷  Start Processing", bg=T.GREEN, fg="white", relief="flat",
                                   font=(T.FONT, 9), activebackground=T.GREEN_DARK,
                                   activeforeground="white", cursor="hand2", command=self.c.start)
        self.btn_start.pack(side="left", fill="x", expand=True, ipady=10)
        self.btn_save = tk.Button(bar, text="💾  Save As", bg=T.BG, relief="solid", bd=1, font=(T.FONT, 9),
                                  cursor="hand2", state="disabled", command=self.c.save)
        self.btn_save.pack(side="left", padx=8, ipady=8, ipadx=30)
        tk.Button(bar, text="📂  Open", bg=T.BG, relief="solid", bd=1, font=(T.FONT, 9),
                  cursor="hand2", command=self.c.open_out).pack(side="left", ipady=8, ipadx=30)

    # ---------- build ----------
    def _style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Maroon.Horizontal.TProgressbar", troughcolor=T.GRAY, background=T.MAROON,
                    bordercolor=T.GRAY, lightcolor=T.MAROON, darkcolor=T.MAROON, thickness=14)

    def _build_form(self, parent):
        section_title(parent, "▥", "File Configuration")
        files = card(parent)
        FileSlot(files, "GRADE 11 AUTOMATED CONSO AND ECARD", self.c.g11_path,
                 "Browse Folder", self.c.pick_g11)
        FileSlot(files, "GRADE 12 AUTOMATED CONSO AND ECARD", self.c.g12_path,
                 "Browse Folder", self.c.pick_g12, divider=False)

        section_title(parent, "◔", "Processing")
        prog = card(parent)
        top = tk.Frame(prog, bg=T.BG)
        top.pack(fill="x", padx=18, pady=(10, 4))
        tk.Label(top, text="Overall Progress", bg=T.BG, fg="#3a0d0d",
                 font=(T.FONT, 15, "bold")).pack(side="left")
        self.pct_var = tk.StringVar(value="0%")
        tk.Label(top, textvariable=self.pct_var, bg=T.BG, fg="#3a0d0d",
                 font=(T.FONT, 15, "bold")).pack(side="right")
        self.pbar = ttk.Progressbar(prog, style="Maroon.Horizontal.TProgressbar", maximum=100)
        self.pbar.pack(fill="x", padx=18, pady=4)
        self.status_var = tk.StringVar(value="Waiting File")
        tk.Label(prog, textvariable=self.status_var, bg=T.BG, fg=T.MUTED,
                 font=(T.FONT, 9)).pack(anchor="w", padx=18)
        self.badge = tk.Label(prog, text="Ready", bg=T.GREEN, fg="white", font=(T.FONT, 8), padx=18, pady=2)
        self.badge.pack(anchor="w", padx=18, pady=(4, 12))

    def _build_settings(self, parent):
        section_title(parent, "⚙", "Settings")
        box = card(parent)
        f = tk.Frame(box, bg=T.BG)
        f.pack(fill="x", padx=16, pady=10)
        tk.Label(f, text=f"{self.c.template_label} ({self.c.template_name})", bg=T.BG, font=(T.FONT, 11)).pack(anchor="w")
        row = tk.Frame(f, bg=T.BG)
        row.pack(fill="x", pady=4)
        self.template_choice = tk.StringVar()
        self.template_combo = ttk.Combobox(row, textvariable=self.template_choice, state="disabled")
        self.template_combo.pack(side="left", fill="x", expand=True, ipady=2)
        self.template_combo.bind("<<ComboboxSelected>>", lambda _: self.c.select_template(self.template_choice.get()))
        tk.Button(row, text="Choose Folder", bg=T.GRAY, relief="flat",
                  command=self.c.pick_template).pack(side="left", padx=6, ipadx=10)
        self.template_folder_var = tk.StringVar()
        tk.Label(f, textvariable=self.template_folder_var, bg=T.BG, fg=T.MUTED,
                 font=(T.FONT, 8), anchor="w").pack(anchor="w")
        tk.Label(box, text=f'Pagkatapos ng processing, kusang lalabas ang Save As: pumili ng destination at ilagay ang pangalan ng folder (default: "{self.c.output_folder_name}").',
                 bg=T.BG, fg=T.MUTED, font=(T.FONT, 8)).pack(anchor="w", padx=16, pady=(0, 10))

    def set_template_files(self, filenames, folder):
        self.template_folder_var.set(folder)
        self.template_combo.configure(values=filenames, state="readonly" if filenames else "disabled")
        current_file = os.path.basename(self.c.template_path.get()) if self.c.template_path.get() else ""
        self.template_choice.set(current_file if current_file in filenames else "")

    # ---------- tabs ----------
    def show_tab(self, which):
        self.form_frame.pack_forget()
        self.settings_frame.pack_forget()
        on, off = dict(bg=T.MAROON, fg="white"), dict(bg=T.BG, fg="#111")
        if which == "form":
            self.form_frame.pack(fill="both", expand=True)
            self.tab_form.config(**on)
            self.tab_set.config(**off)
        else:
            self.settings_frame.pack(fill="both", expand=True)
            self.tab_set.config(**on)
            self.tab_form.config(**off)

    # ---------- tinatawag ng controller ----------
    def reset(self):
        for sc in self.stat_cards.values():
            sc.set(0)
        self.pbar["value"] = 0
        self.pct_var.set("0%")

    def set_running(self, running):
        self.btn_start.config(state="disabled" if running else "normal")
        if running:
            self.btn_save.config(state="disabled")
            self.badge.config(text="Running", bg=T.ORANGE)

    def update_progress(self, pct, text, stats):
        self.pbar["value"] = pct
        self.pct_var.set(f"{pct}%")
        self.status_var.set(text)
        for k, v in stats.items():
            self.stat_cards[k].set(v)

    def show_processed(self, count):
        """Tapos na ang processing. Kusang lalabas ang Save As (galing sa controller)."""
        self.badge.config(text="Ready to Save", bg="#3b5bdb")
        if count > 0:
            self.btn_save.config(state="normal")
            self.status_var.set(f"Tapos na ang {count} {self.c.item_label}. Piliin kung saan ise-save ang folder...")
        else:
            self.status_var.set(f"Walang nagawang {self.c.item_label}. Tingnan ang Failed Sheet / Missing Data.")

    def show_cancelled(self):
        """Kinansela ang Save As. Hindi pa naka-save; pwede pang i-click ang Save As."""
        self.badge.config(text="Ready to Save", bg="#3b5bdb")
        self.status_var.set("Hindi pa naka-save. I-click ang 💾 Save As kapag handa ka na.")

    def show_saved(self, folder, count):
        self.badge.config(text="Saved", bg=T.GREEN)
        self.status_var.set(f"Na-save ang {count} file")
        messagebox.showinfo("Na-save", f"Na-save ang {count} {self.c.item_label} sa folder:\n\n{folder}")

    def show_error(self, msg):
        self.badge.config(text="Error", bg=T.RED)
        self.status_var.set("May error")
        messagebox.showerror("Error", msg)