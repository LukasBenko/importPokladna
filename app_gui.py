from __future__ import annotations

import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from converter import convert_xlsx_to_xml_file


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XLSX → XML (pokladničné doklady)")
        self.geometry("720x420")

        self.xlsx_path = tk.StringVar()
        self.xml_path = tk.StringVar()
        self.sheet = tk.StringVar(value="0")
        self.mandant = tk.StringVar(value="1")

        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}

        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Excel (.xlsx):").grid(row=0, column=0, sticky="w", **pad)
        tk.Entry(frm, textvariable=self.xlsx_path, width=70).grid(row=0, column=1, sticky="we", **pad)
        tk.Button(frm, text="Vybrať…", command=self.pick_xlsx).grid(row=0, column=2, **pad)

        tk.Label(frm, text="Výstup (.xml):").grid(row=1, column=0, sticky="w", **pad)
        tk.Entry(frm, textvariable=self.xml_path, width=70).grid(row=1, column=1, sticky="we", **pad)
        tk.Button(frm, text="Uložiť ako…", command=self.pick_xml).grid(row=1, column=2, **pad)

        opt = tk.Frame(frm)
        opt.grid(row=2, column=0, columnspan=3, sticky="we", **pad)

        tk.Label(opt, text="Sheet (index/názov):").grid(row=0, column=0, sticky="w")
        tk.Entry(opt, textvariable=self.sheet, width=15).grid(row=0, column=1, sticky="w", padx=8)

        tk.Label(opt, text="Mandant ID:").grid(row=0, column=2, sticky="w", padx=20)
        tk.Entry(opt, textvariable=self.mandant, width=10).grid(row=0, column=3, sticky="w", padx=8)

        tk.Button(frm, text="Konvertovať", command=self.run, height=2).grid(
            row=3, column=0, columnspan=3, sticky="we", padx=10, pady=10
        )

        tk.Label(frm, text="Log:").grid(row=4, column=0, sticky="nw", **pad)
        self.log = tk.Text(frm, height=12, wrap="word")
        self.log.grid(row=4, column=1, columnspan=2, sticky="nsew", **pad)

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(4, weight=1)

        self._log("Pripravené.\n")

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

    def run(self):
        xlsx = self.xlsx_path.get().strip()
        xml = self.xml_path.get().strip()
        sheet_raw = self.sheet.get().strip()
        mandant = (self.mandant.get().strip() or "1")

        if not xlsx:
            messagebox.showerror("Chyba", "Vyber Excel (.xlsx).")
            return
        if not xml:
            messagebox.showerror("Chyba", "Zadaj výstupný XML.")
            return

        # sheet: int alebo str
        try:
            sheet = int(sheet_raw) if sheet_raw != "" else 0
        except ValueError:
            sheet = sheet_raw

        self._log(f"Konvertujem:\n- XLSX: {xlsx}\n- XML:  {xml}\n- Sheet: {sheet}\n- Mandant: {mandant}\n")

        try:
            docs, items = convert_xlsx_to_xml_file(xlsx, xml, sheet=sheet, mandant_id=mandant)
            self._log(f"✅ Hotovo. Doklady: {docs}, položky: {items}\n\n")
            messagebox.showinfo("Hotovo", f"XML vytvorené.\nDoklady: {docs}\nPoložky: {items}")
        except Exception as e:
            self._log("❌ Chyba:\n" + "".join(traceback.format_exception(type(e), e, e.__traceback__)) + "\n")
            messagebox.showerror("Chyba", str(e))


if __name__ == "__main__":
    App().mainloop()
