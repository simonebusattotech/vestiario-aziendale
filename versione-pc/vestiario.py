"""
Gestione vestiario aziendale con lettore codici a barre.

Modalità:
  1. REGISTRAZIONE  - associa un capo (codice a barre) al CID di un dipendente
  2. INVIO          - segna i capi del CID come inviati in lavanderia
  3. RIENTRO        - verifica che i capi tornati dalla lavanderia siano
                      quelli del CID (segnala capi estranei e capi mancanti)

Tutti i dati vengono salvati nel file Excel  vestiario.xlsx  (stessa cartella).
Il lettore barcode USB funziona come una tastiera: legge il codice e preme Invio.
"""

import os
import sys
import datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox

try:
    import winsound
except ImportError:  # non Windows
    winsound = None

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
EXCEL_FILE = os.path.join(BASE_DIR, "vestiario.xlsx")

SHEET_CAPI = "Capi"
SHEET_MOV = "Movimenti"
CAPI_COLS = ["Codice", "CID", "Stato", "Data registrazione", "Ultimo invio", "Ultimo rientro"]
MOV_COLS = ["Data/Ora", "CID", "Codice", "Operazione", "Esito", "Note"]

STATO_SEDE = "In sede"
STATO_LAV = "In lavanderia"

MODE_REG = "REGISTRAZIONE"
MODE_INVIO = "INVIO LAVANDERIA"
MODE_RIENTRO = "RIENTRO / VERIFICA"

COL_OK = "#1e8e3e"
COL_ERR = "#c5221f"
COL_WARN = "#e37400"
COL_IDLE = "#5f6368"


def now_str():
    return dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def norm(code):
    return str(code or "").strip().upper()


# --------------------------------------------------------------------------- #
#  Archivio Excel
# --------------------------------------------------------------------------- #
class Archivio:
    """Tiene in memoria i dati e li sincronizza con il file Excel."""

    def __init__(self, path):
        self.path = path
        self.capi = {}        # codice -> dict colonne
        self.movimenti = []   # lista di liste
        self._mtime = None
        if not os.path.exists(path):
            self._save()
        self.reload()

    # -- lettura ---------------------------------------------------------- #
    def reload(self):
        wb = load_workbook(self.path)
        self.capi = {}
        if SHEET_CAPI in wb.sheetnames:
            for row in wb[SHEET_CAPI].iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                rec = dict(zip(CAPI_COLS, list(row) + [None] * len(CAPI_COLS)))
                rec["Codice"] = norm(rec["Codice"])
                rec["CID"] = norm(rec["CID"])
                self.capi[rec["Codice"]] = rec
        self.movimenti = []
        if SHEET_MOV in wb.sheetnames:
            for row in wb[SHEET_MOV].iter_rows(min_row=2, values_only=True):
                if row and any(row):
                    self.movimenti.append(list(row[: len(MOV_COLS)]))
        wb.close()
        self._mtime = os.path.getmtime(self.path)

    def refresh_if_changed(self):
        """Ricarica se il file è stato modificato da fuori (es. altra postazione)."""
        try:
            if os.path.getmtime(self.path) != self._mtime:
                self.reload()
        except FileNotFoundError:
            self._save()

    # -- scrittura -------------------------------------------------------- #
    def _save(self):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_CAPI
        self._write_sheet(ws, CAPI_COLS,
                          [[r.get(c) for c in CAPI_COLS]
                           for r in sorted(self.capi.values(), key=lambda r: (r["CID"], r["Codice"]))],
                          [22, 14, 16, 22, 22, 22])
        ws2 = wb.create_sheet(SHEET_MOV)
        self._write_sheet(ws2, MOV_COLS, self.movimenti, [22, 14, 22, 22, 18, 50])
        # colora lo stato dei capi
        for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
            c = row[0]
            if c.value == STATO_LAV:
                c.fill = PatternFill("solid", fgColor="FFF2CC")
            elif c.value == STATO_SEDE:
                c.fill = PatternFill("solid", fgColor="D9EAD3")
        for row in ws2.iter_rows(min_row=2, min_col=5, max_col=5):
            c = row[0]
            if c.value == "ERRORE":
                c.fill = PatternFill("solid", fgColor="F4CCCC")
            elif c.value == "ATTENZIONE":
                c.fill = PatternFill("solid", fgColor="FCE5CD")
        wb.save(self.path)
        self._mtime = os.path.getmtime(self.path)

    @staticmethod
    def _write_sheet(ws, cols, rows, widths):
        ws.append(cols)
        for r in rows:
            ws.append(r)
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    def save(self):
        """Salva; se il file è aperto in Excel chiede di chiuderlo e riprova."""
        while True:
            try:
                self._save()
                return True
            except PermissionError:
                if not messagebox.askretrycancel(
                        "File in uso",
                        "Impossibile salvare: il file vestiario.xlsx è aperto in Excel.\n"
                        "Chiudilo e premi Riprova."):
                    return False

    def log(self, cid, codice, operazione, esito, note=""):
        self.movimenti.append([now_str(), cid, codice, operazione, esito, note])

    def capi_di(self, cid):
        return sorted((r for r in self.capi.values() if r["CID"] == cid), key=lambda r: r["Codice"])


# --------------------------------------------------------------------------- #
#  Interfaccia
# --------------------------------------------------------------------------- #
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Vestiario aziendale - Lettura codici a barre")
        self.geometry("1100x720")
        self.minsize(900, 600)

        try:
            self.db = Archivio(EXCEL_FILE)
        except PermissionError:
            messagebox.showerror("Errore", "Il file vestiario.xlsx è aperto in Excel. Chiudilo e riavvia.")
            self.destroy()
            return

        self.cid = None
        self.mode = tk.StringVar(value=MODE_INVIO)
        self._last_mode = MODE_INVIO
        self.session_codes = set()   # codici letti nella sessione corrente

        self._build_ui()
        self._set_cid(None)

    # -- costruzione UI --------------------------------------------------- #
    def _build_ui(self):
        style = ttk.Style(self)
        style.configure("Big.TButton", font=("Segoe UI", 12), padding=8)
        style.configure("Treeview", font=("Segoe UI", 11), rowheight=26)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))

        # Modalità
        top = tk.Frame(self, padx=12, pady=8)
        top.pack(fill="x")
        tk.Label(top, text="Operazione:", font=("Segoe UI", 13, "bold")).pack(side="left")
        for m in (MODE_REG, MODE_INVIO, MODE_RIENTRO):
            tk.Radiobutton(top, text=m, variable=self.mode, value=m, indicatoron=False,
                           font=("Segoe UI", 12, "bold"), padx=14, pady=6, selectcolor="#c6dafc",
                           command=self._on_mode_change).pack(side="left", padx=4)
        tk.Label(top, text=f"File: {EXCEL_FILE}", fg=COL_IDLE, font=("Segoe UI", 9)).pack(side="right")

        # CID
        cidf = tk.LabelFrame(self, text=" 1. Utente ", font=("Segoe UI", 11, "bold"), padx=10, pady=8)
        cidf.pack(fill="x", padx=12)
        tk.Label(cidf, text="CID:", font=("Segoe UI", 14)).pack(side="left")
        self.cid_entry = tk.Entry(cidf, font=("Segoe UI", 18), width=16)
        self.cid_entry.pack(side="left", padx=8)
        self.cid_entry.bind("<Return>", lambda e: self._confirm_cid())
        self.cid_btn = ttk.Button(cidf, text="Conferma", style="Big.TButton", command=self._confirm_cid)
        self.cid_btn.pack(side="left")
        self.change_btn = ttk.Button(cidf, text="Cambia utente / Fine", style="Big.TButton",
                                     command=self._end_session)
        self.change_btn.pack(side="left", padx=8)
        self.cid_label = tk.Label(cidf, text="", font=("Segoe UI", 14, "bold"))
        self.cid_label.pack(side="left", padx=12)

        # Scansione
        scanf = tk.LabelFrame(self, text=" 2. Scansiona il codice a barre del capo ",
                              font=("Segoe UI", 11, "bold"), padx=10, pady=8)
        scanf.pack(fill="x", padx=12, pady=6)
        self.scan_entry = tk.Entry(scanf, font=("Consolas", 22), width=30)
        self.scan_entry.pack(side="left", padx=4)
        self.scan_entry.bind("<Return>", lambda e: self._on_scan())
        ttk.Button(scanf, text="Invia", style="Big.TButton", command=self._on_scan).pack(side="left", padx=6)
        self.counter = tk.Label(scanf, text="", font=("Segoe UI", 14, "bold"))
        self.counter.pack(side="right", padx=8)

        # Esito
        self.banner = tk.Label(self, text="", font=("Segoe UI", 20, "bold"), fg="white",
                               bg=COL_IDLE, pady=14, wraplength=1050)
        self.banner.pack(fill="x", padx=12, pady=4)

        # Liste
        body = tk.PanedWindow(self, orient="horizontal", sashwidth=6)
        body.pack(fill="both", expand=True, padx=12, pady=(4, 12))

        lf = tk.LabelFrame(body, text=" Capi dell'utente ", font=("Segoe UI", 11, "bold"))
        self.tree = ttk.Treeview(lf, columns=("codice", "stato", "invio", "rientro"), show="headings")
        for col, txt, w in (("codice", "Codice", 170), ("stato", "Stato", 120),
                            ("invio", "Ultimo invio", 150), ("rientro", "Ultimo rientro", 150)):
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w)
        self.tree.tag_configure("lav", background="#fff2cc")
        self.tree.tag_configure("sede", background="#d9ead3")
        self.tree.tag_configure("letto", background="#b7e1cd")
        self.tree.pack(fill="both", expand=True)
        body.add(lf, stretch="always")

        rf = tk.LabelFrame(body, text=" Letture di questa sessione ", font=("Segoe UI", 11, "bold"))
        self.log_list = tk.Listbox(rf, font=("Consolas", 11))
        self.log_list.pack(fill="both", expand=True)
        body.add(rf, stretch="always")

    # -- gestione utente -------------------------------------------------- #
    def _set_cid(self, cid):
        self.cid = cid
        self.session_codes = set()
        self.log_list.delete(0, "end")
        if cid:
            self.cid_entry.config(state="disabled")
            self.cid_btn.state(["disabled"])
            self.change_btn.state(["!disabled"])
            self.scan_entry.config(state="normal")
            self.cid_label.config(text=f"Utente attivo: {cid}", fg=COL_OK)
            self._banner(f"{self.mode.get()} - Scansiona i capi", COL_IDLE)
            self.scan_entry.focus_set()
        else:
            self.cid_entry.config(state="normal")
            self.cid_entry.delete(0, "end")
            self.cid_btn.state(["!disabled"])
            self.change_btn.state(["disabled"])
            self.scan_entry.delete(0, "end")
            self.scan_entry.config(state="disabled")
            self.cid_label.config(text="Inserisci il tuo CID e premi Invio", fg=COL_IDLE)
            self._banner("In attesa del CID", COL_IDLE)
            self.cid_entry.focus_set()
        self._refresh_tree()

    def _confirm_cid(self):
        cid = norm(self.cid_entry.get())
        if not cid:
            self._banner("Inserisci un CID valido", COL_ERR, beep=True)
            return
        self.db.refresh_if_changed()
        self._set_cid(cid)
        if self.mode.get() == MODE_RIENTRO:
            attesi = [r for r in self.db.capi_di(cid) if r["Stato"] == STATO_LAV]
            if not attesi:
                self._banner(f"Nessun capo di {cid} risulta in lavanderia", COL_WARN)

    def _end_session(self, mode=None):
        mode = mode or self.mode.get()
        if self.cid and mode == MODE_RIENTRO:
            mancanti = [r["Codice"] for r in self.db.capi_di(self.cid) if r["Stato"] == STATO_LAV]
            if mancanti:
                ok = messagebox.askyesno(
                    "Capi mancanti",
                    f"Mancano {len(mancanti)} capi di {self.cid} non ancora rientrati:\n\n"
                    + "\n".join(mancanti[:30]) + ("\n..." if len(mancanti) > 30 else "")
                    + "\n\nRegistrare la mancanza nel file Excel e chiudere?")
                if not ok:
                    self.scan_entry.focus_set()
                    return False
                for c in mancanti:
                    self.db.log(self.cid, c, MODE_RIENTRO, "ATTENZIONE", "Capo NON rientrato dalla lavanderia")
                self.db.save()
            else:
                messagebox.showinfo("Verifica completata", f"Tutti i capi di {self.cid} sono rientrati.")
        self._set_cid(None)
        return True

    def _on_mode_change(self):
        new_mode = self.mode.get()
        if self.cid:
            # cambiare operazione chiude la sessione dell'utente attuale
            if not self._end_session(mode=self._last_mode):
                self.mode.set(self._last_mode)
                return
        self._last_mode = new_mode
        self._banner(f"{new_mode} - In attesa del CID", COL_IDLE)

    # -- scansione -------------------------------------------------------- #
    def _on_scan(self):
        code = norm(self.scan_entry.get())
        self.scan_entry.delete(0, "end")
        if not code or not self.cid:
            return
        self.db.refresh_if_changed()
        mode = self.mode.get()
        if mode == MODE_REG:
            esito, msg = self._registra(code)
        elif mode == MODE_INVIO:
            esito, msg = self._invio(code)
        else:
            esito, msg = self._rientro(code)

        self.db.log(self.cid, code, mode, esito, msg)
        if not self.db.save():
            self.db.reload()   # annulla le modifiche non salvate
            self._banner("NON SALVATO - file Excel in uso", COL_ERR, beep=True)
        else:
            colore = {"OK": COL_OK, "ATTENZIONE": COL_WARN}.get(esito, COL_ERR)
            self._banner(f"{code}  →  {msg}", colore, beep=(esito != "OK"))
            self.log_list.insert(0, f"{dt.datetime.now():%H:%M:%S}  {esito:<10} {code}  {msg}")
            self.log_list.itemconfig(0, fg=colore)
        self.session_codes.add(code)
        self._refresh_tree()
        self.scan_entry.focus_set()

    def _registra(self, code):
        rec = self.db.capi.get(code)
        if rec:
            if rec["CID"] == self.cid:
                return "ATTENZIONE", "Capo già registrato a questo CID"
            return "ERRORE", f"Capo già registrato al CID {rec['CID']}"
        self.db.capi[code] = {"Codice": code, "CID": self.cid, "Stato": STATO_SEDE,
                              "Data registrazione": now_str(), "Ultimo invio": None, "Ultimo rientro": None}
        return "OK", "Capo registrato"

    def _invio(self, code):
        rec = self.db.capi.get(code)
        if rec is None:
            # primo utilizzo: registra automaticamente il capo all'utente
            self.db.capi[code] = {"Codice": code, "CID": self.cid, "Stato": STATO_LAV,
                                  "Data registrazione": now_str(), "Ultimo invio": now_str(),
                                  "Ultimo rientro": None}
            return "OK", "Nuovo capo registrato e inviato in lavanderia"
        if rec["CID"] != self.cid:
            return "ERRORE", f"Questo capo appartiene al CID {rec['CID']}!"
        if rec["Stato"] == STATO_LAV:
            return "ATTENZIONE", "Capo già segnato in lavanderia"
        rec["Stato"] = STATO_LAV
        rec["Ultimo invio"] = now_str()
        return "OK", "Inviato in lavanderia"

    def _rientro(self, code):
        rec = self.db.capi.get(code)
        if rec is None:
            return "ERRORE", "CAPO SCONOSCIUTO - non registrato!"
        if rec["CID"] != self.cid:
            return "ERRORE", f"CAPO NON TUO - appartiene al CID {rec['CID']}"
        if rec["Stato"] != STATO_LAV:
            if code in self.session_codes:
                return "ATTENZIONE", "Capo già letto"
            return "ATTENZIONE", "Capo tuo ma non risultava inviato in lavanderia"
        rec["Stato"] = STATO_SEDE
        rec["Ultimo rientro"] = now_str()
        return "OK", "Capo corretto - rientrato"

    # -- aggiornamento UI ------------------------------------------------- #
    def _banner(self, text, color, beep=False):
        self.banner.config(text=text, bg=color)
        if beep and winsound:
            winsound.MessageBeep(winsound.MB_ICONHAND)

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        if not self.cid:
            self.counter.config(text="")
            return
        capi = self.db.capi_di(self.cid)
        for r in capi:
            tag = "letto" if r["Codice"] in self.session_codes else ("lav" if r["Stato"] == STATO_LAV else "sede")
            self.tree.insert("", "end", values=(r["Codice"], r["Stato"], r["Ultimo invio"] or "",
                                                r["Ultimo rientro"] or ""), tags=(tag,))
        mode = self.mode.get()
        if mode == MODE_RIENTRO:
            mancanti = sum(1 for r in capi if r["Stato"] == STATO_LAV)
            letti = sum(1 for r in capi if r["Codice"] in self.session_codes and r["Stato"] == STATO_SEDE)
            colore = COL_OK if mancanti == 0 else COL_WARN
            self.counter.config(text=f"Rientrati: {letti}   Mancanti: {mancanti}", fg=colore)
        elif mode == MODE_INVIO:
            in_lav = sum(1 for r in capi if r["Stato"] == STATO_LAV)
            self.counter.config(text=f"In lavanderia: {in_lav} / {len(capi)}", fg=COL_IDLE)
        else:
            self.counter.config(text=f"Capi registrati: {len(capi)}", fg=COL_IDLE)


if __name__ == "__main__":
    App().mainloop()
