# =====================================================================
# PEDAGOG ASISTENT – dnevnik rada za školu
# 
# Redoslijed u ovom fajlu (odozgo prema dolje):
#   1. Uvoz biblioteka i putanje do fajlova sa podacima
#   2. Zakoni i pravilnici – pomoćne funkcije (čitanje fajlova, pretraga)
#   3. Boje i čuvanje/učitavanje podataka
#   4. Izvještaji (tekst na ekranu i PDF)
#   5. class App                – glavni prozor (lista unosa, dugmad)
#   6. class PersonReportWindow – izvještaj po osobi
#   7. class ReminderWindow     – podsjetnici (lista, dodavanje, brisanje)
#   8. class TrashWindow        – kanta za smeće (vraćanje obrisanih unosa)
#   9. class LawsWindow         – zakoni i pravilnici (učitavanje + pretraga)
#  10. Pokretanje aplikacije (na samom dnu)
# 
# Podaci se čuvaju u korisničkom folderu (npr. C:\Users\Ime\):
#   school_log.json              – unosi, tipovi, statusi, podsjetnici i kanta za smeće
#   pedagog_asistent_laws.json   – tekst učitanih zakona i pravilnika
# =====================================================================

# --- 1. UVOZ BIBLIOTEKA ---
import json, os, re, sys, subprocess, unicodedata, zipfile
import xml.etree.ElementTree as ET
# --- WINDOWS HIGH-DPI / BLUR FIX ---
# Mora se pokrenuti prije kreiranja prvog Tk prozora.
# Na macOS/Linuxu se ništa ne mijenja.
if sys.platform == "win32":
    try:
        # Windows 10+: najbolji režim za oštru Tkinter aplikaciju na skaliranju
        # (125%, 150%, 175%, 200%) i na više monitora.
        ctypes = __import__("ctypes")
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        try:
            # Stariji Windows fallback.
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

import tkinter as tk
from tkinter import ttk, messagebox, colorchooser, filedialog
from datetime import date, datetime, timedelta

# pypdf je potreban samo za učitavanje PDF fajlova. Ako ga nema, aplikacija i dalje radi.
try:  # optional: only needed to import PDF files into Laws & Regulations
    import pypdf
except ImportError:
    pypdf = None

# --- PUTANJE DO FAJLOVA SA PODACIMA (oba su u korisničkom folderu) ---
# Ovdje su unosi i tipovi događaja:
FILE = os.path.join(os.path.expanduser("~"), "school_log.json")
# Ovdje je tekst zakona i pravilnika:
LAWS_FILE = os.path.join(os.path.expanduser("~"), "pedagog_asistent_laws.json")



# =====================================================================
# 2. ZAKONI I PRAVILNICI – POMOĆNE FUNKCIJE
# =====================================================================

# Otvara fajl u programu koji je na računaru podrazumijevan (npr. PDF u pregledaču).
def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    else:
        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", path])


# Pretvara tekst u oblik pogodan za pretragu: mala slova, bez kvačica (č→c, š→s, đ→d).
# Dužina teksta ostaje ista – zato tačno znamo na kojem je mjestu nešto pronađeno.
def fold(s):
    """Lower-case, strip accents (č→c, š→s, đ→d), whitespace→space. Same length as input."""
    out = []
    for ch in s:
        if ch in "đĐ":
            out.append("d")
        elif ch.isspace():
            out.append(" ")
        else:
            out.append(unicodedata.normalize("NFD", ch)[0].lower()[:1])
    return "".join(out)


# Razdvaja upit na riječi. Ono što je pod navodnicima traži se kao cijela fraza.
def parse_terms(query):
    """Words are searched separately; "words in quotes" are searched as an exact phrase."""
    terms = [fold(a or b).strip() for a, b in re.findall(r'"([^"]+)"|(\S+)', query)]
    return [t for t in terms if t]


# Nalazi mjesta u tekstu gdje se sve tražene riječi pojavljuju blizu jedna druge
# (do 300 slova razmaka). Vraća listu pozicija.
def find_hits(folded, terms, span=300):
    """Positions where all terms occur within `span` characters of each other."""
    pos = [[m.start() for m in re.finditer(re.escape(t), folded)] for t in terms]
    if not terms or not all(pos):
        return []
    hits = []
    for p in pos[0]:
        if all(any(abs(q - p) <= span for q in lst) for lst in pos[1:]):
            if not hits or p - hits[-1] > span:
                hits.append(p)
    return hits


# Čita tekst iz fajla: TXT/MD (običan tekst), DOCX (Word) ili PDF (preko pypdf).
# Ako je tip fajla nepoznat, javlja grešku.
def extract_text(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".txt", ".md"):
        raw = open(path, "rb").read()
        for enc in ("utf-8-sig", "cp1250", "cp1252"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            text = raw.decode("utf-8", "replace")
    elif ext == ".docx":
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("word/document.xml"))
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        paras = []
        for p in root.iter(ns + "p"):
            parts = []
            for el in p.iter():
                if el.tag == ns + "t":
                    parts.append(el.text or "")
                elif el.tag in (ns + "tab", ns + "br"):
                    parts.append(" ")
            paras.append("".join(parts))
        text = "\n".join(paras)
    elif ext == ".pdf":
        if pypdf is None:
            raise ImportError("pypdf")
        text = "\n".join((pg.extract_text() or "") for pg in pypdf.PdfReader(path).pages)
    else:
        raise ValueError("Unsupported file type. Use PDF, DOCX or TXT.")
    # NFC: slovo sa kvačicom koje je u fajlu rastavljeno na "c" + kvačicu spaja se u jedno slovo (č),
    # da pretraga radi. Novi redovi se ujednačavaju.
    return unicodedata.normalize("NFC", text.replace("\r\n", "\n").replace("\r", "\n").replace("\x00", ""))


# Učitava listu sačuvanih zakona i pravilnika (prazna lista ako fajl ne postoji).
def load_laws():
    try:
        with open(LAWS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

# =====================================================================
# 3. BOJE I ČUVANJE PODATAKA
# =====================================================================

# Šest pastelnih boja koje se nude za tipove događaja:
PALETTE = ["#f0d9ef", "#fcdce1", "#ffe6bb", "#e9ecce", "#cde9dc", "#c4dfe5"]
# Boje aplikacije: pozadina, bijela "kartica", boja teksta i boja označenog reda.
BG, CARD, TEXT, SEL = "#fff8fa", "#ffffff", "#5e4458", "#e3b9dc"
# Početni tipovi događaja (koriste se kad aplikacija prvi put krene).
DEFAULT = {
    "types": {"Meeting": "#f0d9ef", "Parent meeting": "#fcdce1", "Class": "#ffe6bb",
              "Paperwork": "#e9ecce", "Student support": "#cde9dc", "Other": "#c4dfe5"},
    "entries": [],
}

# Početni statusi osoba: ime -> da li se za taj status upisuju razred i odjeljenje.
# Nove statuse dodajete u prozoru za unos dugmetom "+ New status".
DEFAULT_STATUSES = {"Parent": False, "Student": True, "Teacher": False, "Other": False}

# Koliko dana obrisani unosi ostaju u kanti prije nego što se trajno uklone (promijenite broj po želji).
TRASH_DAYS = 30

# --- PREČICE NA TASTATURI ---
# Svaki red: (kombinacija tastera, tekst za prikaz, opis, radnja). "Mod" je Ctrl na Windowsu (Cmd na Macu).
# Kombinaciju (prvu kolonu) slobodno promijenite; radnja je naziv metode u klasi App.
# Prečice rade iz svih prozora aplikacije, ali ne dok je otvoren prozor za unos.
SHORTCUTS = [
    ("<Mod-n>", "Ctrl+N", "Add a new entry", "entry_dialog"),
    ("<Mod-r>", "Ctrl+R", "Add a new reminder", "new_reminder"),
    ("<Mod-Shift-R>", "Ctrl+Shift+R", "Open the reminders list", "open_reminders"),
    ("<Mod-l>", "Ctrl+L", "Open Laws & Regulations", "open_laws"),
    ("<Mod-p>", "Ctrl+P", "Open Person report", "open_person"),
    ("<Mod-s>", "Ctrl+S", "Save the end-of-day report as PDF", "export_pdf"),
    ("<Mod-Shift-T>", "Ctrl+Shift+T", "Open the trash", "open_trash"),
    ("<Mod-f>", "Ctrl+F", "Search in the log", "focus_search"),
    ("<F5>", "F5", "Jump to today", "go_today"),
    ("<Alt-Left>", "Alt+Left", "Previous day (or week / month)", "go_prev"),
    ("<Alt-Right>", "Alt+Right", "Next day (or week / month)", "go_next"),
    ("<F1>", "F1", "Show this list of shortcuts", "show_shortcuts"),
]

# Prečice koje važe samo na određenom mjestu (prikazuju se u listi na F1):
EXTRA_SHORTCUTS = [
    ("Delete", "Delete the selected entry (it goes to the trash)"),
    ("Enter", "Edit the selected entry (when the list is selected)"),
    ("Ctrl+Enter", "Save an entry or reminder, even while typing in the notes box"),
]


# Učitava unose i tipove događaja iz school_log.json (ili početne vrijednosti).
def load():
    try:
        with open(FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return json.loads(json.dumps(DEFAULT))



# =====================================================================
# 4. IZVJEŠTAJI (tekst na ekranu i PDF)
# =====================================================================

# Pretvara datum iz 2026-10-06 u 06.10.2026.
def fmt_date(iso):
    return datetime.strptime(iso, "%Y-%m-%d").strftime("%d.%m.%Y")


# Vraća trajanje unosa u minutima (0 ako nije upisano).
def entry_minutes(e):
    m = str(e.get("minutes", ""))
    return int(m) if m.isdigit() else 0


# Razdvaja polje "People" na pojedinačna imena (odvojena zarezom ili tačka-zarezom),
# npr. "Marko Marković, Ana Marković" -> ["Marko Marković", "Ana Marković"].
def split_people(text):
    return [n.strip() for n in re.split(r"[,;]", text or "") if n.strip()]


# Sva imena koja su ikad upisana u unose (bez duplikata, po abecedi) – za padajuće menije.
def all_people(entries):
    return sorted({n for e in entries for n in split_people(e.get("person", ""))}, key=str.lower)


# Kratak tekst za kolonu Status u tabeli, npr. "Student (7-2)" ili "Parent".
def status_text(e):
    cls = "-".join(x for x in (e.get("grade", ""), e.get("section", "")) if x)
    return (e.get("status", "") + (f" ({cls})" if cls else "")).strip()


# Pravi liniju sa detaljima unosa, npr. "People: Marko, Ana · Status: Parent · Type: Meeting · 30 min".
def entry_meta(e):
    bits = []
    people = split_people(e.get("person", ""))
    if people:
        bits.append(("People: " if len(people) > 1 else "Person: ") + ", ".join(people))
    if e.get("status"):
        bits.append("Status: " + e["status"])
    if e.get("grade"):
        bits.append("Grade: " + e["grade"])
    if e.get("section"):
        bits.append("Section: " + e["section"])
    bits.append("Type: " + e["type"])
    if entry_minutes(e):
        bits.append(f"{entry_minutes(e)} min")
    return " · ".join(bits)


# Pravi završni red izvještaja: ukupan broj unosa i ukupno vrijeme.
def report_footer(entries):
    total = sum(entry_minutes(e) for e in entries)
    n = len(entries)
    return f"Total: {n} {'entry' if n == 1 else 'entries'}" + (
        f" · {total // 60} h {total % 60} min logged" if total else "")


# Sastavlja podatke za izvještaj jednog dana (naslov, zaglavlje, unosi, zbir).
def day_report_data(day, entries):
    entries = sorted(entries, key=lambda e: e["time"])
    items = [{"when": e["time"], "title": e["title"], "meta": entry_meta(e),
              "notes": e.get("notes", "")} for e in entries]
    return "Daily Work Report", [day.strftime("%A, %d %B %Y")], items, report_footer(entries)


# Da li se osoba pominje u unosu? Gleda se svako ime iz polja People posebno
# (unos sa više ljudi se nađe preko bilo kojeg od njih), a zatim naziv i bilješke.
# Moraju se naći sve upisane riječi, a kvačice se ignorišu.
def person_matches(e, terms):
    for name in split_people(e.get("person", "")):
        f = fold(name)
        if all(t in f for t in terms):
            return True
    hay = fold(" ".join([e["title"], e.get("notes", "")]))
    return all(t in hay for t in terms)


# Sastavlja podatke za izvještaj po osobi: svi unosi te osobe u izabranom periodu.
def person_report_data(name, start, end, entries):
    terms = [fold(w) for w in name.split()]
    lo, hi = start or "0000-00-00", end or "9999-99-99"
    found = sorted((e for e in entries if person_matches(e, terms) and lo <= e["date"] <= hi),
                   key=lambda e: (e["date"], e["time"]))
    items = [{"when": f'{fmt_date(e["date"])} {e["time"]}', "title": e["title"],
              "meta": entry_meta(e), "notes": e.get("notes", "")} for e in found]
    if start or end:
        period = f"{fmt_date(start) if start else 'beginning'} – {fmt_date(end) if end else 'now'}"
    else:
        period = "all dates"
    return "Person Report", [f"Person: {name}", f"Period: {period}"], items, report_footer(found)


# Pretvara podatke izvještaja u običan tekst koji se prikazuje na ekranu.
def report_text(title, header, items, footer):
    """The plain text shown on screen (same content as the PDF)."""
    out = [title, ""] + header + [""]
    for it in items:
        out.append(f'{it["when"]}   {it["title"]}')
        out.append("      " + it["meta"])
        if it["notes"]:
            out.append("      Notes: " + it["notes"].replace("\n", "\n             "))
        out.append("")
    if not items:
        out += ["No entries found.", ""]
    out.append(footer)
    return "\n".join(out)


# Pravi PDF izvještaj: samo crni tekst na bijeloj pozadini. Potrebno: pip install reportlab
def build_pdf(path, title, header, items, footer):
    """Plain black-on-white text report. Needs: pip install reportlab"""
    from xml.sax.saxutils import escape
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

    # Font: traži font koji ima slova č, ć, š, ž, đ (prvo Arial, pa slični fontovi na Windowsu/Macu).
    # Standardni PDF font (Helvetica) nema Č, Ć i Đ – pokvarila bi se u "kvadratiće" – pa se ne koristi.
    # Ako nijedan font nije nađen, javlja se jasna greška umjesto pogrešnog PDF-a.
    win = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    mac = "/System/Library/Fonts/Supplemental"
    lin = "/usr/share/fonts/truetype"
    candidates = [
        (os.path.join(win, "arial.ttf"), os.path.join(win, "arialbd.ttf")),
        (os.path.join(win, "segoeui.ttf"), os.path.join(win, "segoeuib.ttf")),
        (os.path.join(win, "calibri.ttf"), os.path.join(win, "calibrib.ttf")),
        (os.path.join(win, "tahoma.ttf"), os.path.join(win, "tahomabd.ttf")),
        (os.path.join(win, "verdana.ttf"), os.path.join(win, "verdanab.ttf")),
        (os.path.join(mac, "Arial.ttf"), os.path.join(mac, "Arial Bold.ttf")),
        (lin + "/dejavu/DejaVuSans.ttf", lin + "/dejavu/DejaVuSans-Bold.ttf"),
        (lin + "/liberation/LiberationSans-Regular.ttf", lin + "/liberation/LiberationSans-Bold.ttf"),
    ]
    font = bold = None
    for reg, bld in candidates:
        try:
            pdfmetrics.registerFont(TTFont("RptFont", reg))
            pdfmetrics.registerFont(TTFont("RptFontBold", bld))
            font, bold = "RptFont", "RptFontBold"
            break
        except Exception:
            continue
    if font is None:
        raise RuntimeError("No font with the letters Č, Ć, Ž, Š, Đ was found on this computer "
                           "(Arial is normally installed with Windows).")

    # Stilovi teksta – sve je crno, bez boja.
    k = colors.black
    st_title = ParagraphStyle("t", fontName=bold, fontSize=18, leading=22, textColor=k, spaceAfter=6)
    st_head = ParagraphStyle("h", fontName=font, fontSize=11, leading=15, textColor=k)
    st_cell = ParagraphStyle("c", fontName=font, fontSize=10, leading=14, textColor=k)
    st_bold = ParagraphStyle("b", parent=st_cell, fontName=bold)

    def para(text, style):
        return Paragraph(escape(text).replace("\n", "<br/>"), style)

    # Sastavljanje stranice: naslov, zaglavlje, pa tabela sa unosima.
    story = [para(title, st_title)] + [para(h, st_head) for h in header] + [Spacer(1, 12)]
    rows = []
    for it in items:
        body = [para(it["title"], st_bold), para(it["meta"], st_cell)]
        if it["notes"]:
            body.append(para("Notes: " + it["notes"], st_cell))
        rows.append([para(it["when"], st_cell), body])
    if rows:
        table = Table(rows, colWidths=[38 * mm, 136 * mm])
        table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))
        story.append(table)
    else:
        story.append(para("No entries found.", st_cell))
    story += [Spacer(1, 8), para(footer, st_bold)]
    # Pravljenje PDF fajla (A4, margine 18 mm).
    SimpleDocTemplate(path, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                      topMargin=18 * mm, bottomMargin=18 * mm, title=title).build(story)



# =====================================================================
# 5. GLAVNI PROZOR
# =====================================================================

# Glavni prozor: lista unosa za izabrani dan (ili sedmicu/mjesec) i sva dugmad.
class App(tk.Tk):
    # Pravi prozor: učitava podatke, postavlja stil, gradi izgled i prikazuje unose.
    def __init__(self):
        super().__init__()
        self.title("Pedagog Asistent")
        self.geometry("1120x620")
        self.configure(bg=BG)
        self.data = load()
        self.data.setdefault("reminders", [])  # stariji fajlovi još nemaju podsjetnike
        self.data.setdefault("statuses", dict(DEFAULT_STATUSES))  # ni statuse osoba
        self.data.setdefault("trash", [])  # ni kantu za smeće
        self.purge_trash()  # trajno ukloni unose koji su u kanti predugo
        self.cur = date.today()
        self.style_setup()
        self.build()
        self.refresh()
        self.after(1500, self.check_reminders)  # počni provjeru podsjetnika
        self.setup_shortcuts()  # poveži prečice na tastaturi

    # Čuva sve unose i tipove u school_log.json (poziva se nakon svake promjene).
    def save(self):
        with open(FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    # Pravi sigurnosnu kopiju dnevnika i zakona/pravilnika u izabrani folder.
    def backup_data(self):
        folder = filedialog.askdirectory(title="Choose folder for backup", parent=self)
        if not folder:
            return
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_dir = os.path.join(folder, f"Pedagog_Asistent_backup_{stamp}")
        os.makedirs(backup_dir, exist_ok=True)

        copied = []
        for path in (FILE, LAWS_FILE):
            if os.path.exists(path):
                shutil.copy2(path, os.path.join(backup_dir, os.path.basename(path)))
                copied.append(os.path.basename(path))

        if copied:
            messagebox.showinfo(
                "Backup created",
                "Backup je uspješno napravljen.\\n\\nLokacija:\\n" + backup_dir,
                parent=self
            )
        else:
            messagebox.showwarning("Backup", "Nema podataka za backup.", parent=self)

    # Vraća dnevnik iz ranije sigurnosne kopije; prije toga čuva trenutni fajl.
    def restore_data(self):
        path = filedialog.askopenfilename(
            title="Choose school_log.json backup",
            filetypes=[("Pedagog Asistent data", "school_log.json"), ("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                restored = json.load(f)
            if not isinstance(restored, dict) or "entries" not in restored:
                raise ValueError("Selected file is not a valid Pedagog Asistent backup.")
        except Exception as exc:
            messagebox.showerror("Restore", f"Backup nije moguće učitati.\\n\\n{exc}", parent=self)
            return

        if not messagebox.askyesno(
            "Restore backup",
            "Ovo će zamijeniti trenutni dnevnik podacima iz backup-a.\\n\\nNastaviti?",
            parent=self
        ):
            return

        # Automatski sačuvaj trenutni dnevnik prije vraćanja starog.
        if os.path.exists(FILE):
            stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            shutil.copy2(FILE, f"{FILE}.before_restore_{stamp}.bak")

        self.data = restored
        self.data.setdefault("reminders", [])
        self.data.setdefault("statuses", dict(DEFAULT_STATUSES))
        self.data.setdefault("trash", [])
        self.save()
        self.refresh()
        messagebox.showinfo("Restore", "Dnevnik je uspješno vraćen iz backup-a.", parent=self)

    # Boje i izgled tabele (naslovi kolona, označeni red).
    def style_setup(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                    rowheight=32, font=("Segoe UI", 10), borderwidth=0)
        s.configure("Treeview.Heading", background="#c4dfe5", foreground=TEXT,
                    font=("Segoe UI", 10, "bold"), relief="flat")
        s.map("Treeview", background=[("selected", SEL)], foreground=[("selected", TEXT)])

    # Pravi jedno obojeno dugme i stavlja ga u zadati dio prozora.
    def btn(self, parent, text, cmd, color="#f0d9ef", side="left"):
        b = tk.Button(parent, text=text, command=cmd, bg=color, fg=TEXT, bd=0,
                      font=("Segoe UI", 10, "bold"), padx=14, pady=6, cursor="hand2",
                      activebackground=SEL, activeforeground=TEXT)
        b.pack(side=side, padx=4)
        return b

    # Gradi izgled glavnog prozora, odozgo prema dolje.
    def build(self):
        # Naslov aplikacije i datum koji je trenutno otvoren
        tk.Label(self, text="✿ Pedagog Asistent ✿", bg=BG, fg=TEXT,
                 font=("Segoe UI", 20, "bold")).pack(pady=(14, 2))
        self.day_lbl = tk.Label(self, bg=BG, fg=TEXT, font=("Segoe UI", 12))
        self.day_lbl.pack()

        # Dugmad za kretanje kroz dane
        nav = tk.Frame(self, bg=BG)
        nav.pack(pady=6)
        self.btn(nav, "◀", lambda: self.move(-1), "#fcdce1")
        self.btn(nav, "Today", lambda: self.goto(date.today()), "#ffe6bb")
        self.btn(nav, "▶", lambda: self.move(1), "#fcdce1")
        self.btn(nav, "Go to date…", self.pick_date, "#e9ecce")

        # Traka sa dugmadima (novi unos, tip, zakoni, osoba), izborom prikaza i pretragom
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=16, pady=6)
        self.btn(bar, "+ Add entry", lambda: self.entry_dialog(), "#cde9dc")
        self.btn(bar, "+ New type", self.type_dialog, "#f0d9ef")
        self.btn(bar, "Delete", self.delete, "#fcdce1")
        self.btn(bar, "⚖ Laws & Regulations", self.open_laws, "#c4dfe5")
        self.btn(bar, "👤 Person report", self.open_person, "#ffe6bb")
        self.view = tk.StringVar(value="Day")
        cb = ttk.Combobox(bar, textvariable=self.view, state="readonly", width=8,
                          values=["Day", "Week", "Month", "All"])
        cb.pack(side="right", padx=4)
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh())
        tk.Label(bar, text="Show:", bg=BG, fg=TEXT).pack(side="right")
        self.search = tk.StringVar()
        self.search.trace_add("write", lambda *a: self.refresh())
        self.search_entry = tk.Entry(bar, textvariable=self.search, width=16, relief="flat")
        self.search_entry.pack(side="right", padx=(0, 14))
        tk.Label(bar, text="Search:", bg=BG, fg=TEXT).pack(side="right")

        # Tabela sa unosima
        cols = ("date", "time", "activity", "person", "status", "type", "min")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for c, t, w in zip(cols, ("Date", "Time", "What I did", "People", "Status", "Type", "Minutes"),
                           (95, 65, 260, 150, 130, 120, 65)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=16, pady=(4, 6))
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.tree.bind("<Delete>", lambda e: self.delete())          # taster Delete: u kantu
        self.tree.bind("<Return>", lambda e: self.edit_selected())   # taster Enter: izmjena
        # Red ispod tabele: broj unosa i ukupno vrijeme
        self.total = tk.Label(self, bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold"))
        self.total.pack(pady=(0, 6))
        # Dugmad na dnu: izvještaj za kraj dana u PDF-u i podsjetnici
        foot = tk.Frame(self, bg=BG)
        foot.pack(pady=(0, 4))
        self.btn(foot, "📄  End of day – save report as PDF", self.export_pdf, "#ffe6bb")
        self.btn(foot, "🔔 Reminders", self.open_reminders, "#f0d9ef")
        self.btn(foot, "🗑 Trash", self.open_trash, "#c4dfe5")
        self.btn(foot, "💾 Backup", self.backup_data, "#cde9dc")
        self.btn(foot, "↩ Restore", self.restore_data, "#e9ecce")

        # Mali podsjetnik na najvažnije prečice (cijela lista: F1)
        k = "Cmd" if sys.platform == "darwin" else "Ctrl"
        tk.Label(self, text=f"Shortcuts:  {k}+N new entry  ·  {k}+R reminder  ·  {k}+S PDF report  ·  "
                 "F1 all shortcuts", bg=BG, fg="#a58a9c", font=("Segoe UI", 9)).pack(pady=(0, 8))

    # Povezuje tastere iz SHORTCUTS sa radnjama. Za slova se isto radi i kad je uključen Caps Lock.
    def setup_shortcuts(self):
        mod = "Command" if sys.platform == "darwin" else "Control"
        for seq, _, _, method in SHORTCUTS:
            seq = seq.replace("Mod", mod)
            variants = [seq]
            m = re.fullmatch(r"<(\w+)-([a-z])>", seq)
            if m:
                variants.append(f"<{m.group(1)}-{m.group(2).upper()}>")
            for v in variants:
                self.bind_all(v, lambda ev, name=method: self.run_shortcut(name))

    # Izvršava radnju prečice, osim ako je otvoren prozor za unos (da se ne otvori novi usred pisanja).
    def run_shortcut(self, method):
        if not self.grab_current():
            getattr(self, method)()
        return "break"

    # Ctrl+R: odmah otvara formu za novi podsjetnik.
    def new_reminder(self):
        self.open_reminders()
        self.rem_win.edit(None)

    # Ctrl+F: prebacuje kursor u polje za pretragu u glavnom prozoru.
    def focus_search(self):
        self.deiconify(); self.lift(); self.focus_force()
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")

    # F5, Alt+Lijevo, Alt+Desno: skok na danas, dan unazad, dan naprijed.
    def go_today(self):
        self.goto(date.today())

    def go_prev(self):
        self.move(-1)

    def go_next(self):
        self.move(1)

    # F1: prozor sa spiskom svih prečica.
    def show_shortcuts(self):
        w = self.popup("Keyboard shortcuts")
        tk.Label(w, text="Keyboard shortcuts", bg=BG, fg=TEXT, font=("Segoe UI", 16, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 10))
        mac = sys.platform == "darwin"
        rows = [(label.replace("Ctrl", "Cmd") if mac else label, desc) for _, label, desc, _ in SHORTCUTS]
        rows += EXTRA_SHORTCUTS
        for r, (key, desc) in enumerate(rows, start=1):
            tk.Label(w, text=key, bg="#f0d9ef", fg=TEXT, font=("Segoe UI", 10, "bold"), padx=8,
                     pady=2).grid(row=r, column=0, sticky="ew", padx=(0, 14), pady=3)
            tk.Label(w, text=desc, bg=BG, fg=TEXT, anchor="w").grid(row=r, column=1, sticky="w")
        bf = tk.Frame(w, bg=BG)
        bf.grid(row=len(rows) + 1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        self.btn(bf, "Close", w.destroy, "#cde9dc")

    # Trajno uklanja unose koji su u kanti duže od TRASH_DAYS dana (poziva se pri pokretanju).
    def purge_trash(self):
        cutoff = datetime.now() - timedelta(days=TRASH_DAYS)
        keep = []
        for t in self.data["trash"]:
            try:
                if datetime.strptime(t["deleted"], "%Y-%m-%d %H:%M") >= cutoff:
                    keep.append(t)
            except (KeyError, ValueError):
                keep.append(t)  # ako datum nije jasan, radije sačuvaj nego izgubi
        if len(keep) != len(self.data["trash"]):
            self.data["trash"] = keep
            self.save()

    # Otvara kantu za smeće (ako je već otvorena, samo je vraća naprijed).
    def open_trash(self):
        w = getattr(self, "trash_win", None)
        if w is not None and w.winfo_exists():
            w.deiconify(); w.lift(); w.focus_force()
        else:
            self.trash_win = TrashWindow(self)

    # Ako je kanta otvorena, osvježi njenu listu.
    def refresh_trash_window(self):
        w = getattr(self, "trash_win", None)
        if w is not None and w.winfo_exists():
            w.refresh()

    # Svakih 20 sekundi provjerava da li je neki podsjetnik stigao na red i prikazuje ga.
    # Podsjetnik se prikaže samo jednom (označi se kao "fired"), čak i kad se aplikacija ponovo pokrene.
    def check_reminders(self):
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        shown = 0
        for r in self.data["reminders"]:
            if not r.get("fired") and not r.get("done") and f'{r["date"]} {r["time"]}' <= now:
                r["fired"] = True
                self.show_reminder(r, shown)
                shown += 1
        if shown:
            self.save()
            self.refresh_reminder_window()
        self.after(20000, self.check_reminders)

    # Mali prozor koji iskoči na ekranu (dolje desno, uvijek preko ostalih) kad dođe vrijeme podsjetnika.
    # Ako ih iskoči više odjednom, svaki je malo iznad prethodnog.
    def show_reminder(self, r, n=0):
        pink = "#fcdce1"
        w = tk.Toplevel(self)
        w.title("Reminder")
        w.configure(bg=pink, padx=22, pady=16)
        w.attributes("-topmost", True)
        tk.Label(w, text="🔔 Reminder", bg=pink, fg=TEXT, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Label(w, text=r["title"], bg=pink, fg=TEXT, font=("Segoe UI", 15, "bold"),
                 wraplength=380, justify="left").pack(anchor="w", pady=(6, 2))
        tk.Label(w, text=f'{fmt_date(r["date"])}  {r["time"]}', bg=pink, fg=TEXT,
                 font=("Segoe UI", 10)).pack(anchor="w")
        if r.get("notes"):
            note = r["notes"] if len(r["notes"]) <= 500 else r["notes"][:500] + "…"
            tk.Label(w, text=note, bg=pink, fg=TEXT, font=("Segoe UI", 11), wraplength=380,
                     justify="left").pack(anchor="w", pady=(8, 0))
        row = tk.Frame(w, bg=pink)
        row.pack(anchor="e", pady=(14, 0))

        # "Remind me in 10 min": isti podsjetnik se vraća za 10 minuta
        def snooze():
            t = datetime.now() + timedelta(minutes=10)
            r["date"], r["time"], r["fired"] = t.strftime("%Y-%m-%d"), t.strftime("%H:%M"), False
            self.save()
            self.refresh_reminder_window()
            w.destroy()

        # "Mark as done": podsjetnik je urađen i dobija kvačicu u listi podsjetnika
        def mark_done():
            r["done"] = True
            self.save()
            self.refresh_reminder_window()
            w.destroy()

        self.btn(row, "Remind me in 10 min", snooze, "#ffe6bb")
        self.btn(row, "Mark as done", mark_done, "#cde9dc")
        self.btn(row, "OK ✓", w.destroy, "#f0d9ef")
        w.update_idletasks()
        x = w.winfo_screenwidth() - w.winfo_reqwidth() - 30
        y = w.winfo_screenheight() - w.winfo_reqheight() - 90 - 40 * (n % 5)
        w.geometry(f"+{x}+{y}")
        w.lift()
        self.bell()

    # Otvara prozor sa spiskom podsjetnika (ako je već otvoren, samo ga vraća naprijed).
    def open_reminders(self):
        w = getattr(self, "rem_win", None)
        if w is not None and w.winfo_exists():
            w.deiconify(); w.lift(); w.focus_force()
        else:
            self.rem_win = ReminderWindow(self)

    # Ako je prozor sa podsjetnicima otvoren, osvježi njegovu listu.
    def refresh_reminder_window(self):
        w = getattr(self, "rem_win", None)
        if w is not None and w.winfo_exists():
            w.refresh()

    # Otvara prozor "Person report" (izvještaj po osobi).
    def open_person(self):
        w = getattr(self, "person_win", None)
        if w is not None and w.winfo_exists():
            w.destroy()  # reopen fresh so names and dates are up to date
        self.person_win = PersonReportWindow(self)

    # Otvara prozor sa zakonima i pravilnicima (ako je već otvoren, samo ga vraća naprijed).
    def open_laws(self):
        w = getattr(self, "laws", None)
        if w is not None and w.winfo_exists():
            w.deiconify(); w.lift(); w.focus_force()
        else:
            self.laws = LawsWindow(self)

    # Dugme "End of day": pravi PDF izvještaj za dan koji je otvoren.
    def export_pdf(self):
        day = self.cur.isoformat()
        entries = [e for e in self.data["entries"] if e["date"] == day]
        if not entries:
            return messagebox.showinfo("Nothing to export", f"There are no entries for {day}.")
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile=f"Pedagog_Asistent_{day}.pdf", title="Save daily report")
        if not path:
            return
        try:
            build_pdf(path, *day_report_data(self.cur, entries))
        except ImportError:
            return messagebox.showerror("Missing package",
                                        "Please run this once in a terminal:\n\npip install reportlab")
        except Exception as ex:
            return messagebox.showerror("Could not save PDF", str(ex))
        if messagebox.askyesno("Saved", "Report saved. Open it now?"):
            open_file(path)

    # Pomjera prikaz naprijed/nazad (za 1 dan, 7 dana ili 30 dana, zavisno od prikaza).
    def move(self, n):
        step = {"Day": 1, "Week": 7, "Month": 30, "All": 1}[self.view.get()]
        self.goto(self.cur + timedelta(days=n * step))

    # Prelazi na izabrani datum i osvježava listu.
    def goto(self, d):
        self.cur = d
        self.refresh()

    # Mali prozor za upis datuma na koji želite da pređete.
    def pick_date(self):
        w = self.popup("Go to date")
        v = tk.StringVar(value=self.cur.isoformat())
        tk.Label(w, text="Date (YYYY-MM-DD)", bg=BG, fg=TEXT).pack()
        tk.Entry(w, textvariable=v, relief="flat").pack(pady=6)

        # Prelazi na upisani datum (javlja grešku ako format nije dobar)
        def ok():
            try:
                self.goto(datetime.strptime(v.get(), "%Y-%m-%d").date())
                w.destroy()
            except ValueError:
                messagebox.showerror("Oops", "Please use YYYY-MM-DD.")
        self.btn(w, "Go", ok, "#cde9dc")

    # Da li datum spada u trenutni prikaz (Day / Week / Month / All)?
    def in_range(self, d):
        v = self.view.get()
        if v == "All":
            return True
        if v == "Day":
            return d == self.cur
        if v == "Week":
            start = self.cur - timedelta(days=self.cur.weekday())
            return start <= d <= start + timedelta(days=6)
        return (d.year, d.month) == (self.cur.year, self.cur.month)

    # Ponovo puni tabelu: uzima unose za izabrani period, primjenjuje pretragu i boji redove.
    def refresh(self):
        self.day_lbl.config(text=self.cur.strftime("%A, %d %B %Y"))
        self.tree.delete(*self.tree.get_children())
        for c in set(self.data["types"].values()):
            self.tree.tag_configure("c" + c[1:], background=c)
        # Riječ iz polja za pretragu
        q = self.search.get().lower().strip()
        # Unosi poredani po datumu i vremenu
        rows = sorted(enumerate(self.data["entries"]), key=lambda x: (x[1]["date"], x[1]["time"]))
        count = mins = 0
        for i, e in rows:
            try:
                d = datetime.strptime(e["date"], "%Y-%m-%d").date()
            except ValueError:
                continue
            if not self.in_range(d):
                continue
            if q and q not in (e["title"] + e["type"] + e.get("notes", "") + e.get("person", "")
                               + " " + status_text(e)).lower():
                continue
            color = self.data["types"].get(e["type"], "#ffffff")
            self.tree.insert("", "end", iid=str(i), tags=("c" + color[1:],),
                             values=(e["date"], e["time"], e["title"], e.get("person", ""), status_text(e),
                                    e["type"], e.get("minutes", "")))
            count += 1
            mins += int(e["minutes"]) if str(e.get("minutes", "")).isdigit() else 0
        self.total.config(text=f"{count} entries · {mins // 60} h {mins % 60} min logged")

    # Vraća redni broj označenog unosa (ili None ako ništa nije označeno).
    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    # Briše označeni unos: pita za potvrdu, a unos se ne briše trajno nego ide u kantu (dugme "Trash").
    def delete(self):
        i = self.selected()
        if i is not None and messagebox.askyesno(
                "Delete", f"Move this entry to the trash? You can restore it for {TRASH_DAYS} days."):
            entry = self.data["entries"].pop(i)
            self.data["trash"].append({"entry": entry, "deleted": datetime.now().strftime("%Y-%m-%d %H:%M")})
            self.save(); self.refresh()
            self.refresh_trash_window()

    # Dvoklik na unos otvara prozor za izmjenu.
    def edit_selected(self):
        i = self.selected()
        if i is not None:
            self.entry_dialog(i)

    # Pravi mali prozor (dijalog) u bojama aplikacije.
    def popup(self, title):
        w = tk.Toplevel(self)
        w.title(title); w.configure(bg=BG, padx=20, pady=16)
        w.transient(self); w.grab_set()
        return w

    # Prozor za dodavanje novog unosa ili izmjenu postojećeg (kad je zadat idx).
    def entry_dialog(self, idx=None):
        # e = unos koji se mijenja (prazno ako dodajemo novi)
        e = self.data["entries"][idx] if idx is not None else {}
        w = self.popup("Edit entry" if e else "Add entry")
        now = datetime.now().strftime("%H:%M") if self.cur == date.today() else "08:00"
        # v = vrijednosti polja u prozoru (naziv, datum, vrijeme, trajanje, osoba, status, razred, tip)
        v = {
            "title": tk.StringVar(value=e.get("title", "")),
            "date": tk.StringVar(value=e.get("date", self.cur.isoformat())),
            "time": tk.StringVar(value=e.get("time", now)),
            "minutes": tk.StringVar(value=str(e.get("minutes", ""))),
            "person": tk.StringVar(value=e.get("person", "")),
            "status": tk.StringVar(value=e.get("status") or "(none)"),
            "grade": tk.StringVar(value=e.get("grade", "")),
            "section": tk.StringVar(value=e.get("section", "")),
            "type": tk.StringVar(value=e.get("type", list(self.data["types"])[0])),
        }
        # Prozor je velik i može se razvlačiti; polje za bilješke zauzima sav slobodan prostor
        w.geometry("780x790")
        w.minsize(560, 560)
        w.columnconfigure(1, weight=1)
        w.rowconfigure(8, weight=1)
        # Već upisana imena osoba za padajući meni "+ add known"
        known = all_people(self.data["entries"])
        # (naziv, ključ, red u prozoru) – redovi 2 i 3 su status i razred, a 7 i 8 tip i bilješke
        fields = [("What I did", "title", 0), ("People (separate with commas)", "person", 1),
                  ("Date (YYYY-MM-DD)", "date", 4), ("Time (HH:MM)", "time", 5),
                  ("Duration (minutes, optional)", "minutes", 6)]
        for label, key, r in fields:
            tk.Label(w, text=label, bg=BG, fg=TEXT).grid(row=r, column=0, sticky="w", pady=4, padx=(0, 10))
            if key == "person":
                # Polje za imena (može ih biti više, npr. oba roditelja) + izbor već upisanih imena
                pf = tk.Frame(w, bg=BG)
                pf.grid(row=r, column=1, sticky="ew", pady=4)
                tk.Entry(pf, textvariable=v[key], relief="flat").pack(side="left", fill="x", expand=True)
                pick = tk.StringVar(value="+ add known")
                pc = ttk.Combobox(pf, textvariable=pick, values=known, state="readonly", width=16)
                pc.pack(side="left", padx=(6, 0))

                # Izabrano ime se dodaje na kraj polja (ako već nije upisano)
                def add_person(ev=None, pick=pick):
                    cur = split_people(v["person"].get())
                    if pick.get() in known and pick.get() not in cur:
                        v["person"].set(", ".join(cur + [pick.get()]))
                    pick.set("+ add known")

                pc.bind("<<ComboboxSelected>>", add_person)
            else:
                tk.Entry(w, textvariable=v[key], relief="flat").grid(row=r, column=1, sticky="ew", pady=4)

        # Status osobe: izbor iz liste + dugme za dodavanje novog statusa (Parent, Student, Teacher...)
        tk.Label(w, text="Status of the person", bg=BG, fg=TEXT).grid(row=2, column=0, sticky="w", pady=4, padx=(0, 10))
        sf = tk.Frame(w, bg=BG)
        sf.grid(row=2, column=1, sticky="ew", pady=4)
        sc = ttk.Combobox(sf, textvariable=v["status"], values=["(none)"] + list(self.data["statuses"]),
                          state="readonly")
        sc.pack(side="left", fill="x", expand=True)

        # Razred i odjeljenje: vide se samo kad izabrani status to traži (npr. Student)
        cl = tk.Label(w, text="Class (razred / odjeljenje)", bg=BG, fg=TEXT)
        cl.grid(row=3, column=0, sticky="w", pady=4, padx=(0, 10))
        cf = tk.Frame(w, bg=BG)
        cf.grid(row=3, column=1, sticky="w", pady=4)
        tk.Label(cf, text="Grade", bg=BG, fg=TEXT).pack(side="left")
        tk.Entry(cf, textvariable=v["grade"], width=6, relief="flat").pack(side="left", padx=(6, 18))
        tk.Label(cf, text="Section", bg=BG, fg=TEXT).pack(side="left")
        tk.Entry(cf, textvariable=v["section"], width=6, relief="flat").pack(side="left", padx=6)

        # Prikaži ili sakrij red sa razredom i odjeljenjem, zavisno od izabranog statusa
        def update_class(ev=None):
            if self.data["statuses"].get(v["status"].get()):
                cl.grid()
                cf.grid()
            else:
                cl.grid_remove()
                cf.grid_remove()

        # Dugme "+ New status": otvara mali prozor, a novi status se odmah izabere u listi
        def new_status():
            def done(name):
                sc.configure(values=["(none)"] + list(self.data["statuses"]))
                v["status"].set(name)
                update_class()
            self.status_dialog(done, w)

        sc.bind("<<ComboboxSelected>>", update_class)
        self.btn(sf, "+ New status", new_status, "#f0d9ef")
        update_class()

        tk.Label(w, text="Type", bg=BG, fg=TEXT).grid(row=7, column=0, sticky="w", pady=4)
        ttk.Combobox(w, textvariable=v["type"], values=list(self.data["types"]),
                     state="readonly").grid(row=7, column=1, sticky="ew", pady=4)
        # Veliko polje za bilješke sa klizačem
        tk.Label(w, text="Notes", bg=BG, fg=TEXT).grid(row=8, column=0, sticky="nw", pady=4)
        nf = tk.Frame(w, bg=BG)
        nf.grid(row=8, column=1, sticky="nsew", pady=4)
        sb = tk.Scrollbar(nf)
        sb.pack(side="right", fill="y")
        notes = tk.Text(nf, wrap="word", relief="flat", font=("Segoe UI", 11), padx=8, pady=6,
                        yscrollcommand=sb.set)
        notes.pack(side="left", fill="both", expand=True)
        sb.config(command=notes.yview)
        notes.insert("1.0", e.get("notes", ""))
        # Dugmad sa slovima š đ č ć ž ispod bilješki (ubacuju slovo u polje u kojem ste zadnje pisali)
        target = [notes]
        self.track_focus(w, target)
        self.letters_bar(w, target).grid(row=9, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Klik na "Done ✓": provjeri unos, sačuvaj ga i osvježi listu
        def ok():
            try:
                d = datetime.strptime(v["date"].get(), "%Y-%m-%d").date()
                datetime.strptime(v["time"].get(), "%H:%M")
            except ValueError:
                return messagebox.showerror("Oops", "Please use date YYYY-MM-DD and time HH:MM.")
            if not v["title"].get().strip():
                return messagebox.showerror("Oops", "Please write what you did.")
            m = v["minutes"].get().strip()
            if m and not m.isdigit():
                return messagebox.showerror("Oops", "Minutes must be a number.")
            new = {k: x.get().strip() for k, x in v.items()}
            new["notes"] = notes.get("1.0", "end").strip()
            new["person"] = ", ".join(split_people(new["person"]))  # uredno: "Ime, Ime"
            if new["status"] == "(none)":
                new["status"] = ""
            if not self.data["statuses"].get(new["status"]):  # razred/odjeljenje samo za status koji to traži
                new["grade"] = new["section"] = ""
            if idx is None:
                self.data["entries"].append(new)
            else:
                self.data["entries"][idx] = new
            self.cur = d
            self.save(); self.refresh(); w.destroy()

        # Dugme "Done ✓" (i taster Enter)
        bf = tk.Frame(w, bg=BG)
        bf.grid(row=10, column=0, columnspan=2, sticky="e", pady=(10, 0))
        self.btn(bf, "Done ✓", ok, "#cde9dc")
        w.bind("<Return>", lambda ev: ok() if ev.widget is not notes else None)
        w.bind("<Control-Return>", lambda ev: ok())  # Ctrl+Enter čuva i dok pišete bilješke

    # Pamti u koje polje (Entry ili Text) je zadnje kliknuto u prozoru w; to polje primaju dugmad sa slovima.
    # target je lista sa jednim elementom (polje u koje se ubacuje slovo).
    def track_focus(self, w, target):
        def on_focus(ev):
            if isinstance(ev.widget, (tk.Entry, tk.Text)):
                target[0] = ev.widget
        w.bind("<FocusIn>", on_focus)

    # Red malih dugmadi sa slovima š đ č ć ž (i velikim). Klik ubacuje slovo u polje u kojem ste zadnje pisali.
    # Radi bez obzira na tastaturu: korisno ako Windows tastatura nema ta slova (npr. engleski raspored).
    # Mali prozor za novi status osobe (npr. "Psychologist"): ime + da li se za njega upisuju razred
    # i odjeljenje. Poslije dodavanja poziva on_done(ime); parent je prozor iz kojeg je otvoren.
    def status_dialog(self, on_done, parent=None):
        w = self.popup("New status")
        name, child = tk.StringVar(), tk.BooleanVar(value=False)
        tk.Label(w, text="Status name", bg=BG, fg=TEXT).grid(row=0, column=0, sticky="w", padx=(0, 10))
        tk.Entry(w, textvariable=name, width=28, relief="flat").grid(row=0, column=1, pady=4)
        tk.Checkbutton(w, text="Ask for grade and section (for students / children)", variable=child,
                       bg=BG, fg=TEXT, activebackground=BG, selectcolor=CARD).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=6)

        # Zatvara mali prozor i vraća fokus na prozor za unos
        def close():
            w.destroy()
            if parent is not None:
                parent.grab_set()

        # Dodaje novi status i čuva ga
        def ok():
            n = name.get().strip()
            if not n:
                return messagebox.showerror("Oops", "Please add a name.", parent=w)
            self.data["statuses"][n] = bool(child.get())
            self.save()
            close()
            on_done(n)

        bf = tk.Frame(w, bg=BG)
        bf.grid(row=2, column=0, columnspan=2, sticky="e", pady=(10, 0))
        self.btn(bf, "Add status", ok, "#cde9dc")

    # Prozor za pravljenje novog tipa događaja: ime + boja.
    def type_dialog(self):
        w = self.popup("New type")
        name, color = tk.StringVar(), tk.StringVar(value=PALETTE[0])
        tk.Label(w, text="Type name", bg=BG, fg=TEXT).grid(row=0, column=0, sticky="w")
        tk.Entry(w, textvariable=name, width=28, relief="flat").grid(row=0, column=1, columnspan=3, pady=4)
        tk.Label(w, text="Color", bg=BG, fg=TEXT).grid(row=1, column=0, sticky="w", pady=6)
        preview = tk.Label(w, text="   preview   ", bg=color.get(), fg=TEXT)
        preview.grid(row=3, column=1, columnspan=2, sticky="w", pady=6)

        # Izbor boje iz ponuđene palete
        def pick(c):
            color.set(c); preview.config(bg=c)

        for n, c in enumerate(PALETTE):
            tk.Button(w, bg=c, width=4, bd=0, command=lambda c=c: pick(c)).grid(
                row=1 + n // 3, column=1 + n % 3, padx=2, pady=2)

        # Izbor bilo koje druge boje (prozor za izbor boje)
        def custom():
            c = colorchooser.askcolor(parent=w)[1]
            if c:
                pick(c)

        tk.Button(w, text="Custom…", command=custom, bd=0, bg=CARD, fg=TEXT).grid(row=3, column=3)

        # Dodaje novi tip i čuva ga
        def ok():
            n = name.get().strip()
            if not n:
                return messagebox.showerror("Oops", "Please add a name.")
            self.data["types"][n] = color.get()
            self.save(); self.refresh(); w.destroy()

        bf = tk.Frame(w, bg=BG)
        bf.grid(row=4, column=0, columnspan=4, sticky="e", pady=(10, 0))
        self.btn(bf, "Add type", ok, "#cde9dc")



# =====================================================================
# 6. IZVJEŠTAJ PO OSOBI
# =====================================================================

# Prozor u kojem upišete osobu i period, vidite sve unose sa tom osobom i čuvate ih kao PDF.
class PersonReportWindow(tk.Toplevel):
    """Every entry involving one person, within a date range, as a plain-text report."""

    # Gradi prozor: polje za ime, datumi "From"/"To", dugmad i prostor za izvještaj.
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Person report")
        self.geometry("780x640")
        self.configure(bg=BG)
        entries = app.data["entries"]
        names = all_people(entries)
        self.name = tk.StringVar()
        self.start = tk.StringVar(value=min((x["date"] for x in entries), default=""))
        self.end = tk.StringVar(value=date.today().isoformat())
        self.report = None

        tk.Label(self, text="👤 Person report", bg=BG, fg=TEXT,
                 font=("Segoe UI", 18, "bold")).pack(pady=(12, 4))
        # Polja: osoba, datum od, datum do
        form = tk.Frame(self, bg=BG)
        form.pack(padx=14, pady=4)
        tk.Label(form, text="Person", bg=BG, fg=TEXT).grid(row=0, column=0, sticky="w", pady=4)
        cb = ttk.Combobox(form, textvariable=self.name, values=names, width=36)
        cb.grid(row=0, column=1, columnspan=3, sticky="w", pady=4)
        cb.focus_set()
        cb.bind("<Return>", lambda ev: self.show())
        tk.Label(form, text="From (YYYY-MM-DD)", bg=BG, fg=TEXT).grid(row=1, column=0, sticky="w", pady=4)
        tk.Entry(form, textvariable=self.start, width=13, relief="flat").grid(row=1, column=1, padx=(0, 16))
        tk.Label(form, text="To", bg=BG, fg=TEXT).grid(row=1, column=2, sticky="e", padx=(0, 6))
        tk.Entry(form, textvariable=self.end, width=13, relief="flat").grid(row=1, column=3)
        # Dugmad Show i Save as PDF
        row = tk.Frame(self, bg=BG)
        row.pack(pady=6)
        app.btn(row, "Show", self.show, "#cde9dc")
        app.btn(row, "Save as PDF", self.save_pdf, "#ffe6bb")
        tk.Label(self, bg=BG, fg=TEXT, text="Leave a date empty for no limit. The person is found in the "
                 "Person field and also in the text of an entry.").pack()
        # Prostor u kojem se prikazuje izvještaj (običan tekst)
        vf = tk.Frame(self)
        vf.pack(fill="both", expand=True, padx=14, pady=10)
        sb = tk.Scrollbar(vf)
        sb.pack(side="right", fill="y")
        self.out = tk.Text(vf, wrap="word", bg=CARD, fg="black", relief="flat", padx=12, pady=10,
                           font="TkFixedFont", yscrollcommand=sb.set, state="disabled")
        self.out.pack(fill="both", expand=True)
        sb.config(command=self.out.yview)

    # Provjerava unos i prikazuje izvještaj na ekranu. Vraća False ako nešto nije u redu.
    def show(self):
        name = self.name.get().strip()
        if not name:
            messagebox.showinfo("Person report", "Please write a person's name.", parent=self)
            return False
        start, end = self.start.get().strip(), self.end.get().strip()
        try:
            for d in (start, end):
                if d:
                    datetime.strptime(d, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Oops", "Please use dates as YYYY-MM-DD, or leave them empty.", parent=self)
            return False
        self.report = person_report_data(name, start, end, self.app.data["entries"])
        self.out.config(state="normal")
        self.out.delete("1.0", "end")
        self.out.insert("1.0", report_text(*self.report))
        self.out.config(state="disabled")
        return True

    # Prikazuje izvještaj, pita gdje da se sačuva i pravi PDF.
    def save_pdf(self):
        if not self.show():
            return
        title, header, items, footer = self.report
        if not items:
            return messagebox.showinfo("Nothing to save", "No entries found for this person in that period.",
                                       parent=self)
        safe = re.sub(r"[^\w\-]+", "_", self.name.get().strip())
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".pdf",
                                            filetypes=[("PDF", "*.pdf")], initialfile=f"Report_{safe}.pdf",
                                            title="Save person report")
        if not path:
            return
        try:
            build_pdf(path, title, header, items, footer)
        except ImportError:
            return messagebox.showerror("Missing package",
                                        "Please run this once in a terminal:\n\npip install reportlab", parent=self)
        except Exception as ex:
            return messagebox.showerror("Could not save PDF", str(ex), parent=self)
        if messagebox.askyesno("Saved", "Report saved. Open it now?", parent=self):
            open_file(path)



# =====================================================================
# 7. PODSJETNICI
# =====================================================================

# Prozor sa spiskom podsjetnika: dodavanje, izmjena (dvoklik) i brisanje.
# Kad dođe vrijeme, glavni prozor prikaže mali prozorčić (App.show_reminder).
class ReminderWindow(tk.Toplevel):
    # Gradi prozor: naslov, dugmad i tabela sa podsjetnicima.
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Reminders")
        self.geometry("780x520")
        self.configure(bg=BG)
        tk.Label(self, text="🔔 Reminders", bg=BG, fg=TEXT, font=("Segoe UI", 18, "bold")).pack(pady=(12, 2))
        tk.Label(self, bg=BG, fg=TEXT, text="When the time comes, a small window pops up on the screen. "
                 "The app must be open (it can be minimized).\n"
                 "Finished something early? Click the box in the first column to tick it - it will not pop up.").pack()
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=14, pady=8)
        app.btn(bar, "+ Add reminder", lambda: self.edit(None), "#cde9dc")
        app.btn(bar, "✓ Mark done / undo", self.toggle_selected, "#f0d9ef")
        app.btn(bar, "Delete", self.delete, "#fcdce1")
        cols = ("check", "when", "title", "notes", "status")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for c, t, w in zip(cols, ("✓", "When", "Reminder", "Notes", "Status"), (40, 140, 210, 290, 80)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.tag_configure("waiting", background="#ffe6bb")   # čeka da dođe vrijeme
        self.tree.tag_configure("done", foreground="#b8a3af")      # urađen (sivo)
        self.tree.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.tree.bind("<Button-1>", self.on_click)    # klik na kvadratić = urađeno / nije urađeno
        self.tree.bind("<Double-1>", self.on_double)   # dvoklik = izmjena
        self.refresh()

    # Ponovo puni tabelu (poredano po datumu i vremenu).
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        rows = sorted(enumerate(self.app.data["reminders"]), key=lambda x: (x[1]["date"], x[1]["time"]))
        for i, r in rows:
            note = " ".join(r.get("notes", "").split())
            if len(note) > 60:
                note = note[:60] + "…"
            done = bool(r.get("done"))
            status = "done" if done else ("shown" if r.get("fired") else "waiting")
            self.tree.insert("", "end", iid=str(i), tags=(status,),
                             values=("☑" if done else "☐", f'{fmt_date(r["date"])} {r["time"]}',
                                     r["title"], note, status))

    # Označi / skini kvačicu "urađeno" na podsjetniku (urađeni podsjetnici se ne prikazuju na ekranu).
    def toggle(self, i):
        r = self.app.data["reminders"][i]
        r["done"] = not r.get("done")
        self.app.save()
        self.refresh()
        self.tree.selection_set(str(i))

    # Dugme "✓ Mark done / undo": isto za označeni podsjetnik.
    def toggle_selected(self):
        i = self.selected()
        if i is not None:
            self.toggle(i)

    # Klik na kvadratić u prvoj koloni označava podsjetnik kao urađen (ili vraća nazad).
    def on_click(self, ev):
        if self.tree.identify_region(ev.x, ev.y) == "cell" and self.tree.identify_column(ev.x) == "#1":
            row = self.tree.identify_row(ev.y)
            if row:
                self.toggle(int(row))
                return "break"

    # Dvoklik otvara izmjenu (osim na kvadratiću, da se ne otvori slučajno).
    def on_double(self, ev):
        if self.tree.identify_column(ev.x) != "#1":
            self.edit(self.selected())

    # Redni broj označenog podsjetnika (ili None).
    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    # Briše označeni podsjetnik (prvo pita za potvrdu).
    def delete(self):
        i = self.selected()
        if i is not None and messagebox.askyesno("Delete", "Delete this reminder?", parent=self):
            del self.app.data["reminders"][i]
            self.app.save()
            self.refresh()

    # Prozor za novi podsjetnik ili izmjenu postojećeg (kad je zadat idx): naziv, datum, vrijeme, bilješke.
    def edit(self, idx):
        r = self.app.data["reminders"][idx] if idx is not None else {}
        w = self.app.popup("Edit reminder" if r else "New reminder")
        w.geometry("640x560")
        w.minsize(480, 400)
        w.columnconfigure(1, weight=1)
        w.rowconfigure(3, weight=1)
        soon = datetime.now() + timedelta(hours=1)   # početno: za sat vremena, na puni sat
        v = {"title": tk.StringVar(value=r.get("title", "")),
             "date": tk.StringVar(value=r.get("date", soon.strftime("%Y-%m-%d"))),
             "time": tk.StringVar(value=r.get("time", soon.strftime("%H:00"))),
             "done": tk.BooleanVar(value=bool(r.get("done")))}
        for n, (label, key) in enumerate([("Reminder", "title"), ("Date (YYYY-MM-DD)", "date"),
                                          ("Time (HH:MM)", "time")]):
            tk.Label(w, text=label, bg=BG, fg=TEXT).grid(row=n, column=0, sticky="w", pady=4, padx=(0, 10))
            tk.Entry(w, textvariable=v[key], relief="flat").grid(row=n, column=1, sticky="ew", pady=4)
        tk.Label(w, text="Notes", bg=BG, fg=TEXT).grid(row=3, column=0, sticky="nw", pady=4)
        nf = tk.Frame(w, bg=BG)
        nf.grid(row=3, column=1, sticky="nsew", pady=4)
        sb = tk.Scrollbar(nf)
        sb.pack(side="right", fill="y")
        notes = tk.Text(nf, wrap="word", relief="flat", font=("Segoe UI", 11), padx=8, pady=6,
                        yscrollcommand=sb.set)
        notes.pack(side="left", fill="both", expand=True)
        sb.config(command=notes.yview)
        notes.insert("1.0", r.get("notes", ""))
        # Kvačica: urađeno unaprijed – takav podsjetnik se neće prikazati na ekranu
        tk.Checkbutton(w, text="Already done (it will not pop up)", variable=v["done"], bg=BG, fg=TEXT,
                       activebackground=BG, selectcolor=CARD).grid(row=4, column=1, sticky="w", pady=4)
        # Dugmad sa slovima š đ č ć ž (ubacuju slovo u polje u kojem ste zadnje pisali)
        target = [notes]
        self.app.track_focus(w, target)
        self.app.letters_bar(w, target).grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Klik na "Save ✓": provjeri unos i sačuvaj podsjetnik (ponovo se "naoruža", osim ako je urađen)
        def ok():
            try:
                when = datetime.strptime(v["date"].get().strip() + " " + v["time"].get().strip(),
                                         "%Y-%m-%d %H:%M")
            except ValueError:
                return messagebox.showerror("Oops", "Please use date YYYY-MM-DD and time HH:MM.", parent=w)
            if not v["title"].get().strip():
                return messagebox.showerror("Oops", "Please write what you want to be reminded of.", parent=w)
            if when < datetime.now() and not v["done"].get() and not messagebox.askyesno(
                    "Time has passed", "That time has already passed, so the reminder will pop up "
                    "right away. Save anyway?", parent=w):
                return
            new = {"title": v["title"].get().strip(), "date": when.strftime("%Y-%m-%d"),
                   "time": when.strftime("%H:%M"), "notes": notes.get("1.0", "end").strip(),
                   "fired": False, "done": v["done"].get()}
            if idx is None:
                self.app.data["reminders"].append(new)
            else:
                self.app.data["reminders"][idx] = new
            self.app.save()
            self.refresh()
            w.destroy()

        bf = tk.Frame(w, bg=BG)
        bf.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10, 0))
        self.app.btn(bf, "Save ✓", ok, "#cde9dc")
        w.bind("<Return>", lambda ev: ok() if ev.widget is not notes else None)
        w.bind("<Control-Return>", lambda ev: ok())  # Ctrl+Enter čuva i dok pišete bilješke


# =====================================================================
# 8. KANTA ZA SMEĆE
# =====================================================================

# Prozor sa obrisanim unosima: vraćanje (Restore), trajno brisanje i pražnjenje kante.
# Unosi ovdje ostaju TRASH_DAYS dana, a zatim se sami trajno uklone (App.purge_trash).
class TrashWindow(tk.Toplevel):
    # Gradi prozor: naslov, dugmad i tabela sa obrisanim unosima.
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Trash")
        self.geometry("900x520")
        self.configure(bg=BG)
        tk.Label(self, text="🗑 Trash", bg=BG, fg=TEXT, font=("Segoe UI", 18, "bold")).pack(pady=(12, 2))
        tk.Label(self, bg=BG, fg=TEXT, text=f"Deleted entries stay here for {TRASH_DAYS} days, "
                 "then they are removed for good.").pack()
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=14, pady=8)
        app.btn(bar, "↩ Restore", self.restore, "#cde9dc")
        app.btn(bar, "Delete forever", self.forever, "#fcdce1")
        app.btn(bar, "Empty trash", self.empty, "#ffe6bb")
        cols = ("deleted", "date", "time", "title", "people")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for c, t, w in zip(cols, ("Deleted", "Date", "Time", "What I did", "People"), (130, 100, 60, 380, 180)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.refresh()

    # Ponovo puni tabelu (najskorije obrisani su na vrhu).
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        rows = sorted(enumerate(self.app.data["trash"]), key=lambda x: x[1].get("deleted", ""), reverse=True)
        for i, t in rows:
            e = t["entry"]
            self.tree.insert("", "end", iid=str(i), values=(t.get("deleted", ""), e["date"], e["time"],
                                                              e["title"], e.get("person", "")))

    # Redni broj označenog unosa u kanti (ili None).
    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    # Vraća označeni unos u dnevnik (na njegov originalni datum).
    def restore(self):
        i = self.selected()
        if i is None:
            return messagebox.showinfo("Restore", "Select an entry in the list first.", parent=self)
        e = self.app.data["trash"].pop(i)["entry"]
        self.app.data["entries"].append(e)
        self.app.save()
        self.app.refresh()
        self.refresh()
        messagebox.showinfo("Restored", f'The entry is back in your log on {fmt_date(e["date"])}.', parent=self)

    # Trajno briše označeni unos (prvo pita za potvrdu).
    def forever(self):
        i = self.selected()
        if i is None:
            return messagebox.showinfo("Delete forever", "Select an entry in the list first.", parent=self)
        if messagebox.askyesno("Delete forever", "Delete this entry for good? This cannot be undone.", parent=self):
            del self.app.data["trash"][i]
            self.app.save()
            self.refresh()

    # Prazni cijelu kantu (prvo pita za potvrdu).
    def empty(self):
        if not self.app.data["trash"]:
            return messagebox.showinfo("Trash", "The trash is already empty.", parent=self)
        if messagebox.askyesno("Empty trash", "Delete everything in the trash for good? This cannot be undone.",
                               parent=self):
            self.app.data["trash"] = []
            self.app.save()
            self.refresh()


# =====================================================================
# 9. ZAKONI I PRAVILNICI (poseban prozor)
# =====================================================================

# Prozor za učitavanje zakona/pravilnika (PDF, DOCX, TXT) i pretragu teksta u njima.
class LawsWindow(tk.Toplevel):
    """Separate window: import laws/regulations and search inside them."""

    # Učitava sačuvane dokumente i gradi prozor.
    def __init__(self, master):
        super().__init__(master)
        self.title("Laws & Regulations")
        self.geometry("1040x660")
        self.configure(bg=BG)
        self.docs = load_laws()
        self.results, self.terms, self.cache = [], [], {}
        self.query = tk.StringVar()
        self.only_sel = tk.BooleanVar(value=False)
        self.build()
        self.refresh_docs()

    # Čuva tekst svih dokumenata u pedagog_asistent_laws.json.
    def save(self):
        with open(LAWS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.docs, f, ensure_ascii=False)

    # Gradi izgled: gore pretraga, lijevo lista dokumenata, desno rezultati i tekst.
    def build(self):
        btn = self.master.btn
        tk.Label(self, text="⚖ Laws & Regulations", bg=BG, fg=TEXT,
                 font=("Segoe UI", 18, "bold")).pack(pady=(12, 4))
        # Gornja traka: polje za pretragu, dugme Search, "Only selected document"
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=14, pady=4)
        e = tk.Entry(top, textvariable=self.query, relief="flat", font=("Segoe UI", 11))
        e.pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=4)
        e.bind("<Return>", lambda ev: self.search())
        e.focus_set()
        btn(top, "Search", self.search, "#fcdce1")
        tk.Checkbutton(top, text="Only selected document", variable=self.only_sel, bg=BG, fg=TEXT,
                       activebackground=BG, selectcolor=CARD).pack(side="left", padx=8)

        # Ispod: lijeva i desna strana
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=6)
        # Lijeva strana: lista dokumenata i dugmad za dodavanje/uklanjanje
        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="y")
        tk.Label(left, text="Documents", bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.lb = tk.Listbox(left, width=30, relief="flat", bg=CARD, fg=TEXT, selectbackground=SEL,
                             selectforeground=TEXT, exportselection=False, activestyle="none",
                             font=("Segoe UI", 10))
        self.lb.pack(fill="y", expand=True, pady=4)
        self.lb.bind("<<ListboxSelect>>", self.on_doc)
        for text, cmd, color in (("+ Add files…", self.add_files, "#cde9dc"),
                                 ("Remove selected", self.remove_doc, "#fcdce1")):
            row = tk.Frame(left, bg=BG)
            row.pack(pady=2)
            btn(row, text, cmd, color)

        # Desna strana: lista pronađenih mjesta (gore) i tekst dokumenta (dolje)
        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.status = tk.Label(right, bg=BG, fg=TEXT, anchor="w",
                               text="Add documents, then search. Use \"quotes\" for an exact phrase.")
        self.status.pack(fill="x")
        self.res = ttk.Treeview(right, columns=("doc", "text"), show="headings", height=8,
                                selectmode="browse")
        self.res.heading("doc", text="Document")
        self.res.heading("text", text="Found text")
        self.res.column("doc", width=200, anchor="w")
        self.res.column("text", width=560, anchor="w")
        self.res.pack(fill="x", pady=4)
        self.res.bind("<<TreeviewSelect>>", self.on_result)

        # Tekst dokumenta sa klizačem
        vf = tk.Frame(right)
        vf.pack(fill="both", expand=True)
        sb = tk.Scrollbar(vf)
        sb.pack(side="right", fill="y")
        self.view = tk.Text(vf, wrap="word", bg=CARD, fg=TEXT, relief="flat", padx=10, pady=8,
                            font=("Segoe UI", 11), yscrollcommand=sb.set, state="disabled")
        self.view.pack(fill="both", expand=True)
        sb.config(command=self.view.yview)
        self.view.tag_configure("hit", background="#ffe6bb")
        self.view.tag_configure("cur", background="#f3a6c8")

    # Ponovo puni listu dokumenata sa lijeve strane i briše stare rezultate.
    def refresh_docs(self):
        self.lb.delete(0, "end")
        for d in self.docs:
            self.lb.insert("end", d["name"])
        self.res.delete(*self.res.get_children())
        self.results = []

    # Vraća redni broj dokumenta označenog u listi (ili None).
    def sel_index(self):
        s = self.lb.curselection()
        return s[0] if s else None

    # Dugme "+ Add files…": bira fajlove, vadi iz njih tekst i čuva ga.
    # Fajlovi koji se ne mogu pročitati se preskaču, a na kraju se prikaže spisak problema.
    def add_files(self):
        paths = filedialog.askopenfilenames(
            parent=self, title="Choose laws / regulations",
            filetypes=[("Documents", "*.pdf *.docx *.txt *.md"), ("All files", "*.*")])
        if not paths:
            return
        problems = []
        for p in paths:
            name = os.path.basename(p)
            try:
                text = extract_text(p)
            except ImportError:
                problems.append(f"{name}: PDF files need a one-time install:  pip install pypdf")
                continue
            except Exception as ex:
                problems.append(f"{name}: {ex}")
                continue
            if not text.strip():
                problems.append(f"{name}: no text found (a scanned PDF/image can't be searched).")
                continue
            self.docs = [d for d in self.docs if d["name"] != name]
            self.docs.append({"name": name, "text": text, "added": date.today().isoformat()})
        self.docs.sort(key=lambda d: d["name"].lower())
        self.cache = {}
        self.save()
        self.refresh_docs()
        if problems:
            messagebox.showwarning("Some files were skipped", "\n\n".join(problems), parent=self)

    # Briše označeni dokument iz aplikacije (originalni fajl na računaru ostaje).
    def remove_doc(self):
        i = self.sel_index()
        if i is None:
            return messagebox.showinfo("Remove", "Select a document in the list first.", parent=self)
        if messagebox.askyesno("Remove", f"Remove '{self.docs[i]['name']}' from the app?", parent=self):
            del self.docs[i]
            self.cache = {}
            self.save()
            self.refresh_docs()
            self.view.config(state="normal")
            self.view.delete("1.0", "end")
            self.view.config(state="disabled")

    # Vraća "uporedivu" verziju teksta dokumenta (bez kvačica). Pamti se da se ne računa svaki put.
    def folded(self, i):
        d = self.docs[i]
        key = (d["name"], len(d["text"]))
        if key not in self.cache:
            self.cache[key] = fold(d["text"])
        return self.cache[key]

    # Prikazuje tekst dokumenta i boji tražene riječi; ako je zadata pozicija, skače na nju.
    def show_doc(self, i, pos=None):
        text = self.docs[i]["text"]
        v = self.view
        v.config(state="normal")
        v.delete("1.0", "end")
        v.insert("1.0", text)
        # Žuto označi sve tražene riječi
        if self.terms:
            f = self.folded(i)
            for t in self.terms:
                for m in re.finditer(re.escape(t), f):
                    v.tag_add("hit", f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        # Roze označi izabrano mjesto i pomjeri tekst do njega
        if pos is not None:
            v.tag_add("cur", f"1.0+{pos}c", f"1.0+{pos + len(self.terms[0])}c")
            v.yview(v.index(f"1.0+{pos}c linestart -3 lines"))
        else:
            v.yview("1.0")
        v.config(state="disabled")

    # Klik na dokument u listi: prikaži njegov tekst.
    def on_doc(self, _=None):
        i = self.sel_index()
        if i is not None:
            self.show_doc(i)

    # Klik na rezultat: otvori dokument na tom mjestu.
    def on_result(self, _=None):
        sel = self.res.selection()
        if sel:
            i, pos = self.results[int(sel[0])]
            self.show_doc(i, pos)

    # Pretraga: nalazi sva mjesta gdje su tražene riječi i puni listu rezultata (najviše 300).
    def search(self):
        terms = parse_terms(self.query.get())
        if not terms:
            return
        if not self.docs:
            return messagebox.showinfo("Search", "Add a document first (+ Add files…).", parent=self)
        self.terms = terms
        sel = self.sel_index()
        indexes = [sel] if (self.only_sel.get() and sel is not None) else range(len(self.docs))
        self.res.delete(*self.res.get_children())
        self.results = []
        for i in indexes:
            text = self.docs[i]["text"]
            # Svako pronađeno mjesto postaje jedan red u rezultatima
            for pos in find_hits(self.folded(i), terms):
                if len(self.results) >= 300:
                    break
                # Kratak isječak teksta oko pronađenog mjesta
                snippet = " ".join(text[max(0, pos - 70):pos + 160].split())
                self.res.insert("", "end", iid=str(len(self.results)),
                                values=(self.docs[i]["name"], "… " + snippet + " …"))
                self.results.append((i, pos))
        n = len(self.results)
        self.status.config(text=f"{n} result{'s' if n != 1 else ''}" +
                           (" (showing the first 300)" if n >= 300 else "") +
                           ". Click a result to see it in the document.")
        if n:
            self.res.selection_set("0")



# =====================================================================
# 10. POKRETANJE
# =====================================================================

# Kad se fajl pokrene, otvara se glavni prozor.
if __name__ == "__main__":
    App().mainloop()
