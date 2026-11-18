"""
MINI CONVERTISSEUR - Noir / Blanc / Or  (interface agrandie et redimensionnable)

Installation (une seule fois) :
    pip install pdf2docx docx2pdf pymupdf pillow python-docx openpyxl
    pip install tkinterdnd2          # optionnel : glisser-déposer

Lancement :
    python mini_convertisseur.py

Choix possibles (DE -> VERS) :
    PDF   -> Word, Texte, PNG, JPG, Excel (tableaux), HTML
    Word  -> PDF, Texte
    Texte -> PDF, Word
    Image -> PDF (fusion) ou PDF (séparés)

Notes : Word -> PDF demande Microsoft Word ou LibreOffice.
        PDF -> Word / Excel fonctionnent sur des PDF contenant du vrai texte (pas des scans).
"""
import os
import sys
import shutil
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    BaseTk, HAS_DND = TkinterDnD.Tk, True
except Exception:
    BaseTk, HAS_DND = tk.Tk, False

GOLD, GOLD_L, GOLD_D = "#D4AF37", "#F2D675", "#8F7420"
BLACK, PANEL, WHITE, GREY = "#0A0A0A", "#151515", "#FFFFFF", "#B5B5B5"
FONT = "Segoe UI"

# ---- Mise à jour via GitHub Releases ----
VERSION = "1.1.0"                       # à augmenter à chaque nouvelle version
GITHUB_REPO = "Delrunko/mini-convertisseur"


def parse_ver(v):
    import re
    return tuple(int(x) for x in re.findall(r"\d+", v)[:3]) or (0,)


def latest_release():
    import json
    import urllib.request
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
        headers={"User-Agent": "MiniConvertisseur"})
    with urllib.request.urlopen(req, timeout=6) as r:
        d = json.load(r)
    return d["tag_name"], d["html_url"]


# ============================================================ Conversions
def base(p):
    return os.path.splitext(os.path.basename(p))[0]


def uniq(path):
    root, ext = os.path.splitext(path)
    n, new = 1, path
    while os.path.exists(new):
        new, n = f"{root}_{n}{ext}", n + 1
    return new


def read_text(path):
    for enc in ("utf-8", "cp1252", "latin-1"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    return ""


def pdf_to_word(src, od):
    from pdf2docx import Converter
    out = uniq(os.path.join(od, base(src) + ".docx"))
    cv = Converter(src)
    try:
        try:
            cv.convert(out, multi_processing=True)  # plusieurs cœurs = bien plus rapide
        except Exception:
            cv.convert(out)
    finally:
        cv.close()
    return [out]


def pdf_to_text(src, od):
    import fitz
    out = uniq(os.path.join(od, base(src) + ".txt"))
    with fitz.open(src) as doc, open(out, "w", encoding="utf-8") as f:
        for i, pg in enumerate(doc, 1):
            f.write(f"--- Page {i} ---\n{pg.get_text()}\n")
    return [out]


def pdf_to_images(fmt):
    def run(src, od):
        import fitz
        folder = uniq(os.path.join(od, base(src) + "_images"))
        os.makedirs(folder)
        with fitz.open(src) as doc:
            for i, pg in enumerate(doc, 1):
                pg.get_pixmap(dpi=120).save(os.path.join(folder, f"page_{i:03d}.{fmt}"))
        return [folder]
    return run


def pdf_to_excel(src, od):
    import fitz
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    n = 0
    with fitz.open(src) as doc:
        for pi, pg in enumerate(doc, 1):
            for ti, tab in enumerate(pg.find_tables().tables, 1):
                ws = wb.create_sheet(f"P{pi}_T{ti}")
                for row in tab.extract():
                    ws.append(["" if c is None else c for c in row])
                n += 1
    if not n:
        raise RuntimeError("Aucun tableau détecté dans ce PDF.")
    out = uniq(os.path.join(od, base(src) + ".xlsx"))
    wb.save(out)
    return [out]


def pdf_to_html(src, od):
    import fitz
    out = uniq(os.path.join(od, base(src) + ".html"))
    with fitz.open(src) as doc:
        body = "\n".join(pg.get_text("html") for pg in doc)
    with open(out, "w", encoding="utf-8") as f:
        f.write(f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>{base(src)}</title></head>"
                f"<body>{body}</body></html>")
    return [out]


def word_to_pdf(src, od):
    out = uniq(os.path.join(od, base(src) + ".pdf"))
    try:
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        from docx2pdf import convert
        convert(src, out)
        if os.path.exists(out):
            return [out]
    except Exception:
        pass
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", od, src], check=True)
        made = os.path.join(od, base(src) + ".pdf")
        if os.path.exists(made):
            if made != out:
                os.replace(made, out)
            return [out]
    raise RuntimeError("Microsoft Word ou LibreOffice est requis pour Word -> PDF.")


def word_to_text(src, od):
    if src.lower().endswith(".doc"):
        raise RuntimeError("Format .doc non supporté ici : enregistrez en .docx.")
    from docx import Document
    out = uniq(os.path.join(od, base(src) + ".txt"))
    d = Document(src)
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(p.text for p in d.paragraphs))
    return [out]


def text_to_pdf(src, od):
    import fitz
    text = read_text(src)
    out = uniq(os.path.join(od, base(src) + ".pdf"))
    W, H, M, FS, LH = 595, 842, 50, 11, 15
    doc = fitz.open()
    page, y = doc.new_page(width=W, height=H), M
    for para in text.splitlines():
        lines, line = [], ""
        for w in para.split(" "):
            t = f"{line} {w}" if line else w
            if fitz.get_text_length(t, "helv", FS) <= W - 2 * M:
                line = t
            else:
                lines.append(line)
                line = w
        lines.append(line)
        for l in lines:
            if y > H - M:
                page, y = doc.new_page(width=W, height=H), M
            page.insert_text((M, y), l, fontname="helv", fontsize=FS)
            y += LH
    doc.save(out)
    doc.close()
    return [out]


def text_to_word(src, od):
    from docx import Document
    out = uniq(os.path.join(od, base(src) + ".docx"))
    d = Document()
    for line in read_text(src).splitlines():
        d.add_paragraph(line)
    d.save(out)
    return [out]


def images_to_pdf_merge(files, od):
    from PIL import Image
    imgs = [Image.open(f).convert("RGB") for f in files]
    out = uniq(os.path.join(od, base(files[0]) + "_fusion.pdf"))
    imgs[0].save(out, save_all=True, append_images=imgs[1:])
    return [out]


def image_to_pdf(src, od):
    from PIL import Image
    out = uniq(os.path.join(od, base(src) + ".pdf"))
    Image.open(src).convert("RGB").save(out)
    return [out]


# source -> (extensions, filtre dialogue, {cible: (fonction, fusion)})
FORMATS = {
    "PDF": ({".pdf"}, [("PDF", "*.pdf")], {
        "Word": (pdf_to_word, False), "Texte": (pdf_to_text, False),
        "PNG": (pdf_to_images("png"), False), "JPG": (pdf_to_images("jpg"), False),
        "Excel": (pdf_to_excel, False), "HTML": (pdf_to_html, False)}),
    "Word": ({".docx", ".doc"}, [("Word", "*.docx *.doc")], {
        "PDF": (word_to_pdf, False), "Texte": (word_to_text, False)}),
    "Texte": ({".txt", ".md"}, [("Texte", "*.txt *.md")], {
        "PDF": (text_to_pdf, False), "Word": (text_to_word, False)}),
    "Image": ({".png", ".jpg", ".jpeg", ".bmp", ".webp"},
              [("Images", "*.png *.jpg *.jpeg *.bmp *.webp")], {
        "PDF fusion": (images_to_pdf_merge, True), "PDF séparés": (image_to_pdf, False)}),
}


def human_size(path):
    try:
        n = os.path.getsize(path)
    except OSError:
        return ""
    for u in ("o", "Ko", "Mo", "Go"):
        if n < 1024 or u == "Go":
            return f"{n:.0f} {u}" if u == "o" else f"{n:.1f} {u}"
        n /= 1024


# ============================================================ Widgets
class Chip(tk.Label):
    def __init__(self, master, text, command):
        super().__init__(master, text=text, cursor="hand2", font=(FONT, 11, "bold"),
                         padx=8, pady=9, bd=0)
        self.cmd, self.on = command, False
        self.bind("<Button-1>", lambda e: self.cmd())
        self.bind("<Enter>", lambda e: not self.on and self.config(bg=GOLD_D, fg=WHITE))
        self.bind("<Leave>", lambda e: self.paint())
        self.paint()

    def set(self, on):
        self.on = on
        self.paint()

    def paint(self):
        if self.on:
            self.config(bg=GOLD, fg=BLACK, highlightthickness=0)
        else:
            self.config(bg=BLACK, fg=GOLD, highlightthickness=1, highlightbackground=GOLD_D)


class Btn(tk.Label):
    def __init__(self, master, text, command, big=False, outline=False):
        super().__init__(master, text=text, cursor="hand2", bd=0,
                         font=(FONT, 13 if big else 10, "bold"),
                         padx=14, pady=12 if big else 6)
        self.cmd, self.outline = command, outline
        self.bind("<Button-1>", lambda e: self.cmd())
        self.bind("<Enter>", lambda e: self.config(bg=GOLD_L if not outline else GOLD_D,
                                                   fg=BLACK if not outline else WHITE))
        self.bind("<Leave>", lambda e: self.paint())
        self.paint()

    def paint(self):
        if self.outline:
            self.config(bg=BLACK, fg=GOLD, highlightthickness=1, highlightbackground=GOLD)
        else:
            self.config(bg=GOLD, fg=BLACK, highlightthickness=0)


# ============================================================ Application
class App(BaseTk):
    def __init__(self):
        super().__init__()
        self.title(f"Mini Convertisseur v{VERSION}")
        self.geometry("960x640")
        self.minsize(800, 540)
        self.resizable(True, True)
        self.configure(bg=BLACK)
        self.src, self.dst = "PDF", "Word"
        self.files, self.busy, self.pinned = [], False, False
        self.out_dir = tk.StringVar()

        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("G.Horizontal.TProgressbar", troughcolor=PANEL, background=GOLD,
                     bordercolor=BLACK, lightcolor=GOLD, darkcolor=GOLD, thickness=12)
        st.configure("G.Vertical.TScrollbar", troughcolor=PANEL, background=GOLD_D,
                     bordercolor=PANEL, arrowcolor=GOLD, lightcolor=GOLD_D, darkcolor=GOLD_D)
        self._build()
        self.pick_src("PDF")
        self.bind("<F11>", lambda e: self.attributes("-fullscreen", not self.attributes("-fullscreen")))
        self.bind("<Escape>", lambda e: self.attributes("-fullscreen", False))
        self.after(800, lambda: threading.Thread(
            target=self.check_updates, args=(False,), daemon=True).start())

    def _build(self):
        head = tk.Frame(self, bg=GOLD)
        head.pack(fill="x")
        tk.Label(head, text="MINI CONVERTISSEUR", bg=GOLD, fg=BLACK,
                 font=(FONT, 16, "bold")).pack(side="left", padx=18, pady=12)
        tk.Label(head, text=f"v{VERSION}", bg=GOLD, fg=BLACK,
                 font=(FONT, 10)).pack(side="left", pady=(8, 0))
        up = tk.Label(head, text="⟳", bg=GOLD, fg=BLACK, cursor="hand2", font=(FONT, 16, "bold"))
        up.pack(side="right", padx=(0, 8))
        up.bind("<Button-1>", lambda e: threading.Thread(
            target=self.check_updates, args=(True,), daemon=True).start())
        self.pin = tk.Label(head, text="📌", bg=GOLD, fg=BLACK, cursor="hand2", font=(FONT, 14))
        self.pin.pack(side="right", padx=14)
        self.pin.bind("<Button-1>", lambda e: self.toggle_pin())
        tk.Frame(self, bg=GOLD_D, height=3).pack(fill="x")

        body = tk.Frame(self, bg=BLACK)
        body.pack(fill="both", expand=True, padx=20, pady=16)
        self.body = body
        self.banner = tk.Label(self, bg=GOLD, fg=BLACK, cursor="hand2",
                               font=(FONT, 10, "bold"), pady=6)

        # ---- colonne gauche : réglages
        left = tk.Frame(body, bg=BLACK, width=330)
        left.pack(side="left", fill="y", padx=(0, 20))
        left.pack_propagate(False)

        tk.Label(left, text="DE", bg=BLACK, fg=GOLD, font=(FONT, 10, "bold")).pack(anchor="w")
        self.src_row = tk.Frame(left, bg=BLACK)
        self.src_row.pack(fill="x", pady=(4, 14))
        self.src_chips = {}
        for i, s in enumerate(FORMATS):
            c = Chip(self.src_row, s, lambda s=s: self.pick_src(s))
            c.grid(row=i // 2, column=i % 2, sticky="ew", padx=2, pady=2)
            self.src_chips[s] = c
        for i in range(2):
            self.src_row.columnconfigure(i, weight=1, uniform="s")

        tk.Label(left, text="VERS", bg=BLACK, fg=GOLD, font=(FONT, 10, "bold")).pack(anchor="w")
        self.dst_row = tk.Frame(left, bg=BLACK)
        self.dst_row.pack(fill="x", pady=(4, 14))
        self.dst_chips = {}

        tk.Label(left, text="SORTIE", bg=BLACK, fg=GOLD,
                 font=(FONT, 10, "bold")).pack(anchor="w", pady=(0, 4))
        row = tk.Frame(left, bg=BLACK)
        row.pack(fill="x")
        eb = tk.Frame(row, bg=GOLD, padx=1, pady=1)
        eb.pack(side="left", fill="x", expand=True)
        tk.Entry(eb, textvariable=self.out_dir, bg=PANEL, fg=WHITE, bd=0,
                 insertbackground=GOLD, font=(FONT, 10)).pack(fill="x", ipady=7, padx=2)
        Btn(row, "…", self.choose_out, outline=True).pack(side="left", padx=(6, 0))
        tk.Label(left, text="Vide = à côté du fichier d'origine", bg=BLACK, fg=GREY,
                 font=(FONT, 9, "italic")).pack(anchor="w", pady=(3, 0))

        # bas de colonne : progression + bouton
        self.go = Btn(left, "CONVERTIR", self.start, big=True)
        self.go.pack(side="bottom", fill="x")
        self.status = tk.Label(left, text="Prêt.", bg=BLACK, fg=WHITE, font=(FONT, 10),
                               anchor="w", wraplength=320, justify="left")
        self.status.pack(side="bottom", fill="x", pady=(0, 10))
        self.pb = ttk.Progressbar(left, style="G.Horizontal.TProgressbar", maximum=100)
        self.pb.pack(side="bottom", fill="x", pady=(0, 6))

        # ---- colonne droite : fichiers (grande zone)
        right = tk.Frame(body, bg=BLACK)
        right.pack(side="left", fill="both", expand=True)

        top = tk.Frame(right, bg=BLACK)
        top.pack(fill="x")
        tk.Label(top, text="FICHIERS", bg=BLACK, fg=GOLD, font=(FONT, 10, "bold")).pack(side="left")
        self.count = tk.Label(top, text="", bg=BLACK, fg=GREY, font=(FONT, 10))
        self.count.pack(side="left", padx=10)
        Btn(top, "Vider", self.clear, outline=True).pack(side="right", padx=(6, 0))
        Btn(top, "Retirer", self.remove, outline=True).pack(side="right", padx=(6, 0))
        Btn(top, "+ Ajouter", self.add).pack(side="right")

        box = tk.Frame(right, bg=GOLD, padx=1, pady=1)
        box.pack(fill="both", expand=True, pady=(8, 6))
        inner = tk.Frame(box, bg=PANEL)
        inner.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(inner, orient="vertical", style="G.Vertical.TScrollbar")
        sb.pack(side="right", fill="y")
        self.lb = tk.Listbox(inner, bg=PANEL, fg=WHITE, bd=0, highlightthickness=0,
                             selectbackground=GOLD, selectforeground=BLACK,
                             font=(FONT, 11), activestyle="none", selectmode="extended",
                             yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        sb.config(command=self.lb.yview)
        self.lb.bind("<Delete>", lambda e: self.remove())
        self.lb.bind("<Double-Button-1>", lambda e: self.remove())
        self.lb.bind("<Control-a>", lambda e: (self.lb.select_set(0, "end"), "break")[1])

        self.hint = tk.Label(right, bg=BLACK, fg=GREY, font=(FONT, 9, "italic"), anchor="w",
                             text="Glisser-déposer ou « Ajouter » • double-clic ou Suppr pour retirer"
                             if HAS_DND else "« Ajouter » • double-clic ou Suppr pour retirer")
        self.hint.pack(fill="x")
        if HAS_DND:
            self.lb.drop_target_register(DND_FILES)
            self.lb.dnd_bind("<<Drop>>", lambda e: self._add(self.tk.splitlist(e.data)))

    # ---------- choix
    def pick_src(self, s):
        self.src = s
        for k, c in self.src_chips.items():
            c.set(k == s)
        for c in self.dst_chips.values():
            c.destroy()
        self.dst_chips = {}
        targets = list(FORMATS[s][2])
        for i, t in enumerate(targets):
            c = Chip(self.dst_row, t, lambda t=t: self.pick_dst(t))
            c.grid(row=i // 2, column=i % 2, sticky="ew", padx=2, pady=2)
            self.dst_chips[t] = c
        for i in range(2):
            self.dst_row.columnconfigure(i, weight=1, uniform="d")
        self.pick_dst(targets[0])
        exts = FORMATS[s][0]
        self.files = [f for f in self.files if os.path.splitext(f)[1].lower() in exts]
        self.refresh()

    def pick_dst(self, t):
        self.dst = t
        for k, c in self.dst_chips.items():
            c.set(k == t)

    # ---------- mises à jour
    def check_updates(self, manual):
        if GITHUB_REPO.startswith("TON_PSEUDO"):
            if manual:
                self.after(0, lambda: messagebox.showinfo(
                    "Mise à jour", "Renseignez GITHUB_REPO dans le code (ex : pseudo/depot)."))
            return
        try:
            tag, url = latest_release()
        except Exception:
            if manual:
                self.after(0, lambda: messagebox.showwarning(
                    "Mise à jour", "Impossible de vérifier (connexion ou aucune release publiée)."))
            return
        if parse_ver(tag) > parse_ver(VERSION):
            self.after(0, lambda: self.show_banner(tag, url))
        elif manual:
            self.after(0, lambda: messagebox.showinfo(
                "Mise à jour", f"Vous avez la dernière version (v{VERSION})."))

    def show_banner(self, tag, url):
        import webbrowser
        self.banner.config(text=f"⬆ Nouvelle version {tag} disponible — cliquer pour télécharger")
        self.banner.bind("<Button-1>", lambda e: webbrowser.open(url))
        self.banner.pack(side="bottom", fill="x", before=self.body)

    def toggle_pin(self):
        self.pinned = not self.pinned
        self.attributes("-topmost", self.pinned)
        self.pin.config(bg=BLACK if self.pinned else GOLD, fg=GOLD if self.pinned else BLACK)

    # ---------- fichiers
    def refresh(self):
        self.lb.delete(0, "end")
        for f in self.files:
            self.lb.insert("end", f"  {os.path.basename(f)}   —   {human_size(f)}")
        n = len(self.files)
        self.count.config(text=f"{n} fichier{'s' if n > 1 else ''}" if n else "")

    def _add(self, paths):
        exts = FORMATS[self.src][0]
        bad = 0
        for p in paths:
            if os.path.splitext(p)[1].lower() in exts:
                if p not in self.files:
                    self.files.append(p)
            else:
                bad += 1
        self.refresh()
        if bad:
            self.say(f"{bad} fichier(s) ignoré(s) : format non adapté.")

    def add(self):
        p = filedialog.askopenfilenames(title="Choisir", filetypes=FORMATS[self.src][1])
        if p:
            self._add(p)

    def remove(self):
        for i in reversed(self.lb.curselection()):
            del self.files[i]
        self.refresh()

    def clear(self):
        self.files.clear()
        self.refresh()

    def choose_out(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir.set(d)

    # ---------- conversion
    def say(self, t):
        self.after(0, lambda: self.status.config(text=t))

    def prog(self, v):
        self.after(0, lambda: self.pb.config(value=v))

    def start(self):
        if self.busy:
            return
        if not self.files:
            messagebox.showwarning("Aucun fichier", "Ajoutez au moins un fichier.")
            return
        self.busy = True
        self._outd = self.out_dir.get()
        self.go.config(text="EN COURS…")
        self.prog(0)
        threading.Thread(target=self.work, daemon=True).start()

    def work(self):
        fn, merge = FORMATS[self.src][2][self.dst]
        files, outs, errs = list(self.files), [], []
        try:
            if merge:
                od = self._outd or os.path.dirname(files[0])
                os.makedirs(od, exist_ok=True)
                self.say("Fusion en cours…")
                outs += fn(files, od)
                self.prog(100)
            else:
                heavy = fn in (word_to_pdf, pdf_to_word)  # déjà multi-cœur / dépend de Word
                done = [0]

                def job(f):
                    try:
                        od = self._outd or os.path.dirname(f)
                        os.makedirs(od, exist_ok=True)
                        return fn(f, od), None
                    except Exception as e:
                        return [], f"{os.path.basename(f)} : {e}"

                def collect(res):
                    o, e = res
                    outs.extend(o)
                    if e:
                        errs.append(e)
                    done[0] += 1
                    self.say(f"{done[0]}/{len(files)} terminé(s)")
                    self.prog(done[0] / len(files) * 100)

                if heavy or len(files) == 1:
                    for f in files:
                        self.say(f"En cours : {os.path.basename(f)}")
                        collect(job(f))
                else:
                    from concurrent.futures import ThreadPoolExecutor
                    with ThreadPoolExecutor(max_workers=min(4, os.cpu_count() or 2)) as ex:
                        for res in ex.map(job, files):
                            collect(res)
        except Exception as e:
            errs.append(str(e))
        self.after(0, lambda: self.done(outs, errs))

    def done(self, outs, errs):
        self.busy = False
        self.go.config(text="CONVERTIR")
        if errs:
            self.say(f"Terminé avec {len(errs)} erreur(s).")
            messagebox.showerror("Erreurs", "\n\n".join(errs))
        else:
            self.say(f"Terminé : {len(outs)} élément(s) créé(s).")
        if outs:
            p = outs[0] if os.path.isdir(outs[0]) else os.path.dirname(outs[0])
            try:
                if sys.platform.startswith("win"):
                    os.startfile(p)
                elif sys.platform == "darwin":
                    subprocess.run(["open", p])
                else:
                    subprocess.run(["xdg-open", p])
            except Exception:
                pass


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()  # indispensable pour le .exe
    App().mainloop()