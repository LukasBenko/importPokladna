from __future__ import annotations

import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from converter import convert_xlsx_to_xml_file

# =========================
# Brand styles (ARGANIA)
# =========================
PRIMARY_PURPLE = "#3d1f59"
PRIMARY_ORANGE = "#f6915b"
BG_LIGHT = "#fef8f4"
GRAY_50 = "#808080"
WHITE = "#ffffff"


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Import pokladňa")
        self.geometry("720x420")
        self.resizable(False, False)
        self.configure(bg=BG_LIGHT)

        self.xlsx_path = tk.StringVar()
        self.xml_path = tk.StringVar()
        self.sheet = tk.StringVar(value="0")
        self.mandant = tk.StringVar(value="1")

        self._build()

    # =========================
    # UI BUILD
    # =========================
    def _build(self):
        self._build_header()
        self._build_form()
        self._build_actions()
        self._build_log()

        self._log("Pripravené.\n")

    def _build_header(self):
        header = tk.Frame(self, bg=PRIMARY_PURPLE, height=64)
        header.pack(fill="x")

        title = tk.Label(
            header,
            text="IMPORT POKLADŇA",
            bg=PRIMARY_PURPLE,
            fg=WHITE,
            font=("Segoe UI", 16, "bold"),
        )
        title.pack(side="left", padx=20, pady=16)

    def _build_form(self):
        form = tk.Frame(self, bg=BG_LIGHT)
        form.pack(fill="x", padx=20, pady=20)

        # Excel
        tk.Label(form, text="Excel (.xlsx):", bg=BG_LIGHT).grid(row=0, column=0, sticky="w", pady=6)
        tk.Entry(form, textvariable=self.xlsx_path, width=48).grid(row=0, column=1, padx=10)
        tk.Button(form, text="Vybrať…", command=self.pick_xlsx).grid(row=0, column=2)

        # XML
        tk.Label(form, text="Výstup (.xml):", bg=BG_LIGHT).grid(row=1, column=0, sticky="w", pady=6)
        tk.Entry(form, textvariable=self.xml_path, width=48).grid(row=1, column=1, padx=10)
        tk.Button(form, text="Uložiť ako…", command=self.pick_xml).grid(row=1, column=2)

        # Options
        tk.Label(form, text="Sheet (index / názov):", bg=BG_LIGHT).grid(row=2, column=0, sticky="w", pady=6)
        tk.Entry(form, textvariable=self.sheet, width=10).grid(row=2, column=1, sticky="w", padx=10)

        tk.Label(form, text="Mandant ID:", bg=BG_LIGHT).grid(row=2, column=1, sticky="e", padx=(0, 70))
        tk.Entry(form, textvariable=self.mandant, width=6).grid(row=2, column=2, sticky="w")

    def _build_actions(self):
        actions = tk.Frame(self, bg=BG_LIGHT)
        actions.pack(fill="x")

        convert_btn = tk.Button(
            actions,
            text="KONVERTOVAŤ",
            command=self.run,
            bg=PRIMARY_ORANGE,
            fg=WHITE,
            activebackground="#e57f4a",
            relief="flat",
            font=("Segoe UI", 12, "bold"),
            padx=30,
            pady=10,
            cursor="hand2",
        )
        convert_btn.pack(pady=10)

    def _build_log(self):
        log_frame = tk.Frame(self, bg=WHITE, bd=1, relief="solid")
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.log = tk.Text(
            log_frame,
            height=8,
            wrap="word",
            bg=WHITE,
            fg=GRAY_50,
            relief="flat",
            font=("Consolas", 10),
        )
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

    # =========================
    # Helpers
    # =========================
    def _log(self, msg: str):
        self.log.insert("end", msg)
        self.log.see("end")

    def pick_xlsx(self):
        path = filedialog.askopenfilename(
            title="Vyber Excel súbor",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.xlsx_path.set(path)
            if not self.xml_path.get() and path.lower().endswith(".xlsx"):
                self.xml_path.set(path[:-5] + ".xml")

    def pick_xml(self):
        path = filedialog.asksaveasfilename(
            title="Ulož XML ako",
            defaultextension=".xml",
            filetypes=[("XML", "*.xml"), ("All files", "*.*")],
        )
        if path:
            self.xml_path.set(path)

    # =========================
    # Run conversion
    # =========================
    def run(self):
        xlsx = self.xlsx_path.get().strip()
        xml = self.xml_path.get().strip()
        sheet_raw = self.sheet.get().strip()
        mandant = self.mandant.get().strip() or "1"

        if not xlsx:
            messagebox.showerror("Chyba", "Vyber Excel (.xlsx).")
            return
        if not xml:
            messagebox.showerror("Chyba", "Zadaj výstupný XML.")
            return

        try:
            sheet = int(sheet_raw) if sheet_raw != "" else 0
        except ValueError:
            sheet = sheet_raw

        self._log(
            f"Konvertujem:\n"
            f"- XLSX: {xlsx}\n"
            f"- XML:  {xml}\n"
            f"- Sheet: {sheet}\n"
            f"- Mandant: {mandant}\n"
        )

        try:
            docs, items = convert_xlsx_to_xml_file(
                xlsx, xml, sheet=sheet, mandant_id=mandant
            )
            self._log(f"✅ Hotovo. Doklady: {docs}, položky: {items}\n\n")
            messagebox.showinfo(
                "Hotovo",
                f"XML vytvorené.\nDoklady: {docs}\nPoložky: {items}",
            )
        except Exception as e:
            self._log(
                "❌ Chyba:\n"
                + "".join(traceback.format_exception(type(e), e, e.__traceback__))
                + "\n"
            )
            messagebox.showerror("Chyba", str(e))


if __name__ == "__main__":
    App().mainloop()
