"""Pinagmulan ng lahat ng page. Ang bawat page ay isang Frame na ipinapakita sa kanan ng sidebar."""
import tkinter as tk

from ui import theme as T


class BasePage(tk.Frame):
    title = "Page"

    def __init__(self, parent):
        super().__init__(parent, bg=T.BG)

    def on_show(self):
        """Tinatawag tuwing lilipat sa page na ito (override kung kailangan)."""
