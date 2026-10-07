import json, os, re, sys, subprocess, unicodedata, zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox, colorchooser, filedialog
from datetime import date, datetime, timedelta

try:  # optional: only needed to import PDF files into Laws & Regulations
    import pypdf
except ImportError:
    pypdf = None

FILE = os.path.join(os.path.expanduser("~"), "school_log.json")
LAWS_FILE = os.path.join(os.path.expanduser("~"), "pedagog_asistent_laws.json")


def open_file(path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    else:
        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", path])


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


def parse_terms(query):
    """Words are searched separately; "words in quotes" are searched as an exact phrase."""
    terms = [fold(a or b).strip() for a, b in re.findall(r'"([^"]+)"|(\S+)', query)]
    return [t for t in terms if t]


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
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\x00", "")


def load_laws():
    try:
        with open(LAWS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []
PALETTE = ["#f0d9ef", "#fcdce1", "#ffe6bb", "#e9ecce", "#cde9dc", "#c4dfe5"]
BG, CARD, TEXT, SEL = "#fff8fa", "#ffffff", "#5e4458", "#e3b9dc"
DEFAULT = {
    "types": {"Meeting": "#f0d9ef", "Parent meeting": "#fcdce1", "Class": "#ffe6bb",
              "Paperwork": "#e9ecce", "Student support": "#cde9dc", "Other": "#c4dfe5"},
    "entries": [],
}


def load():
    try:
        with open(FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return json.loads(json.dumps(DEFAULT))


def fmt_date(iso):
    return datetime.strptime(iso, "%Y-%m-%d").strftime("%d.%m.%Y")


def entry_minutes(e):
    m = str(e.get("minutes", ""))
    return int(m) if m.isdigit() else 0


def entry_meta(e):
    bits = []
    if e.get("person"):
        bits.append("Person: " + e["person"])
    bits.append("Type: " + e["type"])
    if entry_minutes(e):
        bits.append(f"{entry_minutes(e)} min")
    return " · ".join(bits)


def report_footer(entries):
    total = sum(entry_minutes(e) for e in entries)
    n = len(entries)
    return f"Total: {n} {'entry' if n == 1 else 'entries'}" + (
        f" · {total // 60} h {total % 60} min logged" if total else "")


def day_report_data(day, entries):
    entries = sorted(entries, key=lambda e: e["time"])
    items = [{"when": e["time"], "title": e["title"], "meta": entry_meta(e),
              "notes": e.get("notes", "")} for e in entries]
    return "Daily Work Report", [day.strftime("%A, %d %B %Y")], items, report_footer(entries)


def person_matches(e, terms):
    hay = fold(" ".join([e.get("person", ""), e["title"], e.get("notes", "")]))
    return all(t in hay for t in terms)


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

    # Arial supports Serbian letters (č, ć, š, ž, đ); fall back to Helvetica if missing
    font, bold = "Helvetica", "Helvetica-Bold"
    win = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    mac = "/System/Library/Fonts/Supplemental"
    for reg, bld in [(os.path.join(win, "arial.ttf"), os.path.join(win, "arialbd.ttf")),
                     (os.path.join(mac, "Arial.ttf"), os.path.join(mac, "Arial Bold.ttf"))]:
        try:
            pdfmetrics.registerFont(TTFont("RptFont", reg))
            pdfmetrics.registerFont(TTFont("RptFontBold", bld))
            font, bold = "RptFont", "RptFontBold"
            break
        except Exception:
            continue

    k = colors.black
    st_title = ParagraphStyle("t", fontName=bold, fontSize=18, leading=22, textColor=k, spaceAfter=6)
    st_head = ParagraphStyle("h", fontName=font, fontSize=11, leading=15, textColor=k)
    st_cell = ParagraphStyle("c", fontName=font, fontSize=10, leading=14, textColor=k)
    st_bold = ParagraphStyle("b", parent=st_cell, fontName=bold)

    def para(text, style):
        return Paragraph(escape(text).replace("\n", "<br/>"), style)

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
    SimpleDocTemplate(path, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                      topMargin=18 * mm, bottomMargin=18 * mm, title=title).build(story)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pedagog Asistent")
        self.geometry("1120x620")
        self.configure(bg=BG)
        self.data = load()
        self.cur = date.today()
        self.style_setup()
        self.build()
        self.refresh()

    def save(self):
        with open(FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def style_setup(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                    rowheight=32, font=("Segoe UI", 10), borderwidth=0)
        s.configure("Treeview.Heading", background="#c4dfe5", foreground=TEXT,
                    font=("Segoe UI", 10, "bold"), relief="flat")
        s.map("Treeview", background=[("selected", SEL)], foreground=[("selected", TEXT)])

    def btn(self, parent, text, cmd, color="#f0d9ef", side="left"):
        b = tk.Button(parent, text=text, command=cmd, bg=color, fg=TEXT, bd=0,
                      font=("Segoe UI", 10, "bold"), padx=14, pady=6, cursor="hand2",
                      activebackground=SEL, activeforeground=TEXT)
        b.pack(side=side, padx=4)
        return b

    def build(self):
        tk.Label(self, text="✿ Pedagog Asistent ✿", bg=BG, fg=TEXT,
                 font=("Segoe UI", 20, "bold")).pack(pady=(14, 2))
        self.day_lbl = tk.Label(self, bg=BG, fg=TEXT, font=("Segoe UI", 12))
        self.day_lbl.pack()

        nav = tk.Frame(self, bg=BG)
        nav.pack(pady=6)
        self.btn(nav, "◀", lambda: self.move(-1), "#fcdce1")
        self.btn(nav, "Today", lambda: self.goto(date.today()), "#ffe6bb")
        self.btn(nav, "▶", lambda: self.move(1), "#fcdce1")
        self.btn(nav, "Go to date…", self.pick_date, "#e9ecce")

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
        tk.Entry(bar, textvariable=self.search, width=16, relief="flat").pack(side="right", padx=(0, 14))
        tk.Label(bar, text="Search:", bg=BG, fg=TEXT).pack(side="right")

        cols = ("date", "time", "activity", "person", "type", "min")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for c, t, w in zip(cols, ("Date", "Time", "What I did", "Person", "Type", "Minutes"),
                           (100, 70, 340, 130, 150, 70)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=16, pady=(4, 6))
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.total = tk.Label(self, bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold"))
        self.total.pack(pady=(0, 6))
        foot = tk.Frame(self, bg=BG)
        foot.pack(pady=(0, 14))
        self.btn(foot, "📄  End of day – save report as PDF", self.export_pdf, "#ffe6bb")

    def open_person(self):
        w = getattr(self, "person_win", None)
        if w is not None and w.winfo_exists():
            w.destroy()  # reopen fresh so names and dates are up to date
        self.person_win = PersonReportWindow(self)

    def open_laws(self):
        w = getattr(self, "laws", None)
        if w is not None and w.winfo_exists():
            w.deiconify(); w.lift(); w.focus_force()
        else:
            self.laws = LawsWindow(self)

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

    def move(self, n):
        step = {"Day": 1, "Week": 7, "Month": 30, "All": 1}[self.view.get()]
        self.goto(self.cur + timedelta(days=n * step))

    def goto(self, d):
        self.cur = d
        self.refresh()

    def pick_date(self):
        w = self.popup("Go to date")
        v = tk.StringVar(value=self.cur.isoformat())
        tk.Label(w, text="Date (YYYY-MM-DD)", bg=BG, fg=TEXT).pack()
        tk.Entry(w, textvariable=v, relief="flat").pack(pady=6)

        def ok():
            try:
                self.goto(datetime.strptime(v.get(), "%Y-%m-%d").date())
                w.destroy()
            except ValueError:
                messagebox.showerror("Oops", "Please use YYYY-MM-DD.")
        self.btn(w, "Go", ok, "#cde9dc")

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

    def refresh(self):
        self.day_lbl.config(text=self.cur.strftime("%A, %d %B %Y"))
        self.tree.delete(*self.tree.get_children())
        for c in set(self.data["types"].values()):
            self.tree.tag_configure("c" + c[1:], background=c)
        q = self.search.get().lower().strip()
        rows = sorted(enumerate(self.data["entries"]), key=lambda x: (x[1]["date"], x[1]["time"]))
        count = mins = 0
        for i, e in rows:
            try:
                d = datetime.strptime(e["date"], "%Y-%m-%d").date()
            except ValueError:
                continue
            if not self.in_range(d):
                continue
            if q and q not in (e["title"] + e["type"] + e.get("notes", "") + e.get("person", "")).lower():
                continue
            color = self.data["types"].get(e["type"], "#ffffff")
            self.tree.insert("", "end", iid=str(i), tags=("c" + color[1:],),
                             values=(e["date"], e["time"], e["title"], e.get("person", ""), e["type"],
                                    e.get("minutes", "")))
            count += 1
            mins += int(e["minutes"]) if str(e.get("minutes", "")).isdigit() else 0
        self.total.config(text=f"{count} entries · {mins // 60} h {mins % 60} min logged")

    def selected(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def delete(self):
        i = self.selected()
        if i is not None and messagebox.askyesno("Delete", "Delete this entry?"):
            del self.data["entries"][i]
            self.save(); self.refresh()

    def edit_selected(self):
        i = self.selected()
        if i is not None:
            self.entry_dialog(i)

    def popup(self, title):
        w = tk.Toplevel(self)
        w.title(title); w.configure(bg=BG, padx=20, pady=16)
        w.transient(self); w.grab_set()
        return w

    def entry_dialog(self, idx=None):
        e = self.data["entries"][idx] if idx is not None else {}
        w = self.popup("Edit entry" if e else "Add entry")
        now = datetime.now().strftime("%H:%M") if self.cur == date.today() else "08:00"
        v = {
            "title": tk.StringVar(value=e.get("title", "")),
            "date": tk.StringVar(value=e.get("date", self.cur.isoformat())),
            "time": tk.StringVar(value=e.get("time", now)),
            "minutes": tk.StringVar(value=str(e.get("minutes", ""))),
            "person": tk.StringVar(value=e.get("person", "")),
            "type": tk.StringVar(value=e.get("type", list(self.data["types"])[0])),
        }
        persons = sorted({x["person"] for x in self.data["entries"] if x.get("person")}, key=str.lower)
        fields = [("What I did", "title"), ("Person (optional)", "person"),
                  ("Date (YYYY-MM-DD)", "date"), ("Time (HH:MM)", "time"),
                  ("Duration (minutes, optional)", "minutes")]
        for r, (label, key) in enumerate(fields):
            tk.Label(w, text=label, bg=BG, fg=TEXT).grid(row=r, column=0, sticky="w", pady=4)
            if key == "person":
                ttk.Combobox(w, textvariable=v[key], values=persons, width=31).grid(row=r, column=1, pady=4)
            else:
                tk.Entry(w, textvariable=v[key], width=34, relief="flat").grid(row=r, column=1, pady=4)
        tk.Label(w, text="Type", bg=BG, fg=TEXT).grid(row=5, column=0, sticky="w", pady=4)
        ttk.Combobox(w, textvariable=v["type"], values=list(self.data["types"]),
                     state="readonly", width=31).grid(row=5, column=1, pady=4)
        tk.Label(w, text="Notes", bg=BG, fg=TEXT).grid(row=6, column=0, sticky="nw", pady=4)
        notes = tk.Text(w, width=26, height=4, relief="flat")
        notes.insert("1.0", e.get("notes", ""))
        notes.grid(row=6, column=1, pady=4)

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
            if idx is None:
                self.data["entries"].append(new)
            else:
                self.data["entries"][idx] = new
            self.cur = d
            self.save(); self.refresh(); w.destroy()

        bf = tk.Frame(w, bg=BG)
        bf.grid(row=7, column=0, columnspan=2, sticky="e", pady=(10, 0))
        self.btn(bf, "Done ✓", ok, "#cde9dc")
        w.bind("<Return>", lambda ev: ok() if ev.widget is not notes else None)

    def type_dialog(self):
        w = self.popup("New type")
        name, color = tk.StringVar(), tk.StringVar(value=PALETTE[0])
        tk.Label(w, text="Type name", bg=BG, fg=TEXT).grid(row=0, column=0, sticky="w")
        tk.Entry(w, textvariable=name, width=28, relief="flat").grid(row=0, column=1, columnspan=3, pady=4)
        tk.Label(w, text="Color", bg=BG, fg=TEXT).grid(row=1, column=0, sticky="w", pady=6)
        preview = tk.Label(w, text="   preview   ", bg=color.get(), fg=TEXT)
        preview.grid(row=3, column=1, columnspan=2, sticky="w", pady=6)

        def pick(c):
            color.set(c); preview.config(bg=c)

        for n, c in enumerate(PALETTE):
            tk.Button(w, bg=c, width=4, bd=0, command=lambda c=c: pick(c)).grid(
                row=1 + n // 3, column=1 + n % 3, padx=2, pady=2)

        def custom():
            c = colorchooser.askcolor(parent=w)[1]
            if c:
                pick(c)

        tk.Button(w, text="Custom…", command=custom, bd=0, bg=CARD, fg=TEXT).grid(row=3, column=3)

        def ok():
            n = name.get().strip()
            if not n:
                return messagebox.showerror("Oops", "Please add a name.")
            self.data["types"][n] = color.get()
            self.save(); self.refresh(); w.destroy()

        bf = tk.Frame(w, bg=BG)
        bf.grid(row=4, column=0, columnspan=4, sticky="e", pady=(10, 0))
        self.btn(bf, "Add type", ok, "#cde9dc")


class PersonReportWindow(tk.Toplevel):
    """Every entry involving one person, within a date range, as a plain-text report."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Person report")
        self.geometry("780x640")
        self.configure(bg=BG)
        entries = app.data["entries"]
        names = sorted({x["person"] for x in entries if x.get("person")}, key=str.lower)
        self.name = tk.StringVar()
        self.start = tk.StringVar(value=min((x["date"] for x in entries), default=""))
        self.end = tk.StringVar(value=date.today().isoformat())
        self.report = None

        tk.Label(self, text="👤 Person report", bg=BG, fg=TEXT,
                 font=("Segoe UI", 18, "bold")).pack(pady=(12, 4))
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
        row = tk.Frame(self, bg=BG)
        row.pack(pady=6)
        app.btn(row, "Show", self.show, "#cde9dc")
        app.btn(row, "Save as PDF", self.save_pdf, "#ffe6bb")
        tk.Label(self, bg=BG, fg=TEXT, text="Leave a date empty for no limit. The person is found in the "
                 "Person field and also in the text of an entry.").pack()
        vf = tk.Frame(self)
        vf.pack(fill="both", expand=True, padx=14, pady=10)
        sb = tk.Scrollbar(vf)
        sb.pack(side="right", fill="y")
        self.out = tk.Text(vf, wrap="word", bg=CARD, fg="black", relief="flat", padx=12, pady=10,
                           font="TkFixedFont", yscrollcommand=sb.set, state="disabled")
        self.out.pack(fill="both", expand=True)
        sb.config(command=self.out.yview)

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


class LawsWindow(tk.Toplevel):
    """Separate window: import laws/regulations and search inside them."""

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

    def save(self):
        with open(LAWS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.docs, f, ensure_ascii=False)

    def build(self):
        btn = self.master.btn
        tk.Label(self, text="⚖ Laws & Regulations", bg=BG, fg=TEXT,
                 font=("Segoe UI", 18, "bold")).pack(pady=(12, 4))
        top = tk.Frame(self, bg=BG)
        top.pack(fill="x", padx=14, pady=4)
        e = tk.Entry(top, textvariable=self.query, relief="flat", font=("Segoe UI", 11))
        e.pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=4)
        e.bind("<Return>", lambda ev: self.search())
        e.focus_set()
        btn(top, "Search", self.search, "#fcdce1")
        tk.Checkbutton(top, text="Only selected document", variable=self.only_sel, bg=BG, fg=TEXT,
                       activebackground=BG, selectcolor=CARD).pack(side="left", padx=8)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=14, pady=6)
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

    def refresh_docs(self):
        self.lb.delete(0, "end")
        for d in self.docs:
            self.lb.insert("end", d["name"])
        self.res.delete(*self.res.get_children())
        self.results = []

    def sel_index(self):
        s = self.lb.curselection()
        return s[0] if s else None

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

    def folded(self, i):
        d = self.docs[i]
        key = (d["name"], len(d["text"]))
        if key not in self.cache:
            self.cache[key] = fold(d["text"])
        return self.cache[key]

    def show_doc(self, i, pos=None):
        text = self.docs[i]["text"]
        v = self.view
        v.config(state="normal")
        v.delete("1.0", "end")
        v.insert("1.0", text)
        if self.terms:
            f = self.folded(i)
            for t in self.terms:
                for m in re.finditer(re.escape(t), f):
                    v.tag_add("hit", f"1.0+{m.start()}c", f"1.0+{m.end()}c")
        if pos is not None:
            v.tag_add("cur", f"1.0+{pos}c", f"1.0+{pos + len(self.terms[0])}c")
            v.yview(v.index(f"1.0+{pos}c linestart -3 lines"))
        else:
            v.yview("1.0")
        v.config(state="disabled")

    def on_doc(self, _=None):
        i = self.sel_index()
        if i is not None:
            self.show_doc(i)

    def on_result(self, _=None):
        sel = self.res.selection()
        if sel:
            i, pos = self.results[int(sel[0])]
            self.show_doc(i, pos)

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
            for pos in find_hits(self.folded(i), terms):
                if len(self.results) >= 300:
                    break
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


if __name__ == "__main__":
    App().mainloop()