"""
BASE CONTROLLER - pinagsasaluhan ng SF10 at Form 9 (parehong daloy, iba lang ang engine at template).
Hawak nito ang mga napiling file, ang thread na nagpoproseso, at nag-uupdate ng page.

Dalawang hakbang:
  1. Start Processing -> ginagawa ang mga file sa memory (wala pang naka-save)
  2. Save             -> pipili ang user ng location, gagawa ng folder at dun isusulat
"""
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import config
from services.excel_utils import find_template


class BaseController:
    # --- ito ang binabago ng bawat subclass ---
    engine_cls = None                 # class ng engine (SF10Engine / Form9Engine)
    page_title = "Form Automation"    # malaking title sa taas ng page
    template_name = ""                # default na pangalan ng template file
    template_keywords = ()            # panghanap kung bahagyang iba ang pangalan
    output_folder_name = "OUTPUT"     # pangalan ng folder na gagawin sa Save
    template_label = "Template"       # pangalan ng template sa mga mensahe
    item_label = "file"               # tawag sa nagagawa (SF10 / Form 9) sa mga mensahe

    def __init__(self, root):
        self.root = root
        self.view = None
        self.g11_path = tk.StringVar(master=root)
        self.g12_path = tk.StringVar(master=root)
        self.template_path = tk.StringVar(master=root, value=self._default_template())
        self.engine = None        # may laman pagkatapos ng processing (hindi pa naka-save)
        self.last_out = ""        # huling folder na na-save
        self.q = queue.Queue()
        self.running = False
        self.root.after(100, self._poll)

    def attach(self, view):
        self.view = view

    # ---------- template ----------
    def _default_template(self):
        return find_template((config.BASE_DIR, os.getcwd()), self.template_name, self.template_keywords)

    # ---------- pickers ----------
    def _pick_excel(self, var, title):
        p = filedialog.askopenfilename(title=title, filetypes=[("Excel", "*.xlsx *.xlsm"), ("All", "*.*")])
        if p:
            var.set(p)

    def pick_g11(self):
        self._pick_excel(self.g11_path, "Grade 11 Automated Conso")

    def pick_g12(self):
        self._pick_excel(self.g12_path, "Grade 12 Automated Conso")

    def pick_template(self):
        self._pick_excel(self.template_path, self.template_name)

    # ---------- run (hakbang 1) ----------
    def start(self):
        if self.running:
            return
        g11, g12, tpl = self.g11_path.get(), self.g12_path.get(), self.template_path.get()
        if not g11 or not g12:
            messagebox.showwarning("Kulang", "Pumili muna ng Grade 11 at Grade 12 file.")
            return
        if not tpl or not os.path.exists(tpl):
            messagebox.showwarning("Template", f"Hanapin ang {self.template_name} sa Settings.")
            self.view.show_tab("settings")
            return
        self.engine = None
        self.view.reset()
        self.running = True
        self.view.set_running(True)
        engine = self.engine_cls(g11, g12, tpl, lambda p, t, s: self.q.put(("progress", p, t, s)))

        def work():
            try:
                engine.run()
                self.q.put(("done", engine))
            except Exception as e:  # noqa
                self.q.put(("error", str(e)))

        threading.Thread(target=work, daemon=True).start()

    # ---------- save (hakbang 2) - user ang pipindot ----------
    def _unique_folder(self, parent):
        """Folder sa loob ng napiling location; hindi nag-o-overwrite."""
        base = os.path.join(parent, self.output_folder_name)
        path, n = base, 2
        while os.path.exists(path):
            path = f"{base} ({n})"
            n += 1
        return path

    def save(self):
        if self.running:
            return
        if self.engine is None or not self.engine.results:
            messagebox.showinfo("Save", "Wala pang mase-save. I-click muna ang Start Processing.")
            return
        opts = {"title": "Piliin kung saan ise-save ang folder"}
        if self.g11_path.get():
            opts["initialdir"] = os.path.dirname(self.g11_path.get())
        parent = filedialog.askdirectory(**opts)
        if not parent:          # kinansela ng user
            return
        folder = self._unique_folder(parent)
        try:
            count = self.engine.save_to(folder)
        except OSError as e:
            messagebox.showerror("Save", f"Hindi na-save ang mga file:\n\n{e}")
            return
        self.last_out = folder
        self.view.show_saved(folder, count)

    def open_out(self):
        p = self.last_out
        if not p or not os.path.isdir(p):
            messagebox.showinfo("Open", "Wala pang na-save na folder.")
            return
        if sys.platform.startswith("win"):
            os.startfile(p)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])

    # ---------- polling ng thread ----------
    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    _, p, t, s = msg
                    self.view.update_progress(p, t, s)
                elif msg[0] == "done":
                    self.running = False
                    self.engine = msg[1]
                    self.view.set_running(False)
                    self.view.show_processed(len(self.engine.results))
                elif msg[0] == "error":
                    self.running = False
                    self.view.set_running(False)
                    self.view.show_error(msg[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll)