"""
Trace College - Form Automation
MAIN: dito nagsisimula ang app. Run:  python app.py

Structure:
    app.py            -> main window, sidebar, paglipat ng page
    config.py         -> cell maps ng SF10
    config_form9.py   -> cell maps ng Form 9
    pages/            -> UI ng bawat page (SF10, Grad Form 9)
    controllers/      -> lohika sa pagitan ng page at engine
    services/         -> processing (Excel read/write), walang tkinter
    ui/               -> kulay, font, reusable widgets
"""
import tkinter as tk

import config
from controllers.form9_controller import Form9Controller
from controllers.sf10_controller import SF10Controller
from pages.grad_form9_page import GradForm9Page
from pages.sf10_page import SF10Page
from ui import theme as T


class Sidebar(tk.Frame):
    def __init__(self, parent, items, on_select):
        super().__init__(parent, bg=T.BG, width=170, highlightbackground=T.SIDEBAR_BORDER,
                         highlightthickness=1)
        self.pack_propagate(False)
        self.on_select = on_select
        self.buttons = {}

        tk.Label(self, text="🛡 Trace College", bg=T.BG, fg="#111",
                 font=(T.FONT, 11, "bold")).pack(anchor="w", padx=14, pady=(16, 24))
        for key, text in items:
            b = tk.Label(self, text=text, bg=T.BG, fg="#111", font=(T.FONT, 9),
                         anchor="w", padx=14, pady=6, cursor="hand2")
            b.pack(fill="x")
            b.bind("<Button-1>", lambda e, k=key: self.on_select(k))
            self.buttons[key] = b

        up = tk.Frame(self, bg="#f4f4f4", height=60)
        up.pack(side="bottom", fill="x", padx=14, pady=16)
        tk.Label(up, text="Updates plan", bg="#f4f4f4", font=(T.FONT, 9)).pack(pady=18)

    def highlight(self, active):
        for key, b in self.buttons.items():
            if key == active:
                b.config(bg=T.NAV_ACTIVE_BG, fg=T.MAROON, font=(T.FONT, 9, "bold"))
            else:
                b.config(bg=T.BG, fg="#111", font=(T.FONT, 9))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(config.APP_TITLE)
        self.geometry("1080x700")
        self.minsize(980, 660)
        self.configure(bg=T.BG)

        self.sf10_controller = SF10Controller(self)
        self.form9_controller = Form9Controller(self)

        self.sidebar = Sidebar(self, [("sf10", "  SF10 GENERATE"),
                                      ("form9", "     GRAD FORM 9")], self.show_page)
        self.sidebar.pack(side="left", fill="y")

        self.container = tk.Frame(self, bg=T.BG)
        self.container.pack(side="left", fill="both", expand=True, padx=18, pady=10)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        # i-register dito ang lahat ng page
        self.pages = {
            "sf10": SF10Page(self.container, self.sf10_controller),
            "form9": GradForm9Page(self.container, self.form9_controller),
        }
        self.sf10_controller.attach(self.pages["sf10"])
        self.form9_controller.attach(self.pages["form9"])
        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

        self.show_page("sf10")

    def show_page(self, key):
        page = self.pages[key]
        page.tkraise()
        page.on_show()
        self.sidebar.highlight(key)


if __name__ == "__main__":
    App().mainloop()