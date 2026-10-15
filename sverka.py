"""Сверка «Получено от плательщиков» (HolliHop, expertonline.t8s.ru) с таблицей учеников (xlsx).

Вход: xlsx с листами «<ПРЕДМЕТ> ГК» и «<ПРЕДМЕТ> спецкурсы» + PDF со ссылками на группы.
Выход: docx с расхождениями. На сайте только входит и читает страницы групп (GET), ничего не меняет.

Без аргументов открывается окно. Консоль:
    python sverka.py --xlsx table.xlsx --pdf courses.pdf [--out report.docx]
    логин/пароль — из SVERKA_LOGIN / SVERKA_PASSWORD или спросит в консоли.
"""
import argparse
import datetime as dt
import getpass
import html
import http.cookiejar
import os
import queue
import re
import ssl
import subprocess
import sys
import threading
import urllib.parse
import urllib.request
import warnings
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import docx
import openpyxl
import pypdf
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.shared import Pt, RGBColor
from openpyxl.utils import get_column_letter

MONTHS = [("сентяб", "Сентябрь"), ("октяб", "Октябрь"), ("ноябр", "Ноябрь"), ("декаб", "Декабрь"),
          ("январ", "Январь"), ("феврал", "Февраль"), ("март", "Март"), ("апрел", "Апрель"),
          ("май", "Май"), ("мая", "Май"), ("июн", "Июнь"), ("июл", "Июль"), ("август", "Август")]
SUBJECTS = {"общ": "ОБЩ", "ист": "ИСТ", "рус": "РУС", "мат": "МАТ"}
SUBJECT_WORDS = {"обществознан": "ОБЩ", "истори": "ИСТ", "русск": "РУС", "математик": "МАТ"}
ALIASES = {"с-п": "стартер-пак"}  # сокращение в названии группы на сайте → подпись в xlsx
LATIN_LOOKALIKES = str.maketrans("aceopxykmthb", "асеорхукмтнв")
SITE = "https://expertonline.t8s.ru"  # пароль уходит только сюда; ссылки PDF на другие адреса пропускаются
GROUP_PATH = re.compile(r"/Learner/(?:Group/)?(\d+)/?")
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
ZERO = Decimal(0)


class LoginError(Exception):
    pass


# ---------- входные файлы ----------

def pdf_links(path, site=SITE):
    """Ссылки на группы сайта из PDF в порядке появления, по одной на группу; остальные — в skipped."""
    origin = urllib.parse.urlsplit(site)
    links, ids, skipped = [], set(), []
    for page in pypdf.PdfReader(path).pages:
        annots = page.get("/Annots")
        for a in (annots.get_object() if annots else []):
            act = a.get_object().get("/A")
            uri = act.get_object().get("/URI") if act else None
            if not uri:
                continue
            uri = str(uri).strip()
            u = urllib.parse.urlsplit(uri)
            m = GROUP_PATH.fullmatch(u.path)
            if m and not u.query and (u.scheme, u.netloc.casefold()) == (origin.scheme, origin.netloc.casefold()):
                if m[1] not in ids:  # /Learner/387 и /Learner/Group/387 — одна группа
                    ids.add(m[1])
                    links.append(uri)
            elif uri not in skipped:
                skipped.append(uri)
    return links, skipped


def norm(s):
    return re.sub(r"\s+", "", str(s).casefold().translate(LATIN_LOOKALIKES))


def month_of(token):
    return next((m for stem, m in MONTHS if token.startswith(stem)), None)


def subject_of(token):
    return SUBJECTS.get(token) or next((v for k, v in SUBJECT_WORDS.items() if token.startswith(k)), None)


@dataclass
class Entry:  # одна итоговая ячейка xlsx
    sheet: str
    subject: str
    month: str | None
    label: str  # нормализованная подпись, напр. «профи2»
    value: Decimal
    cell: str


def load_xlsx(path):
    """Итоговые ячейки из сводок наверху листов «… ГК» (месяц × тариф) и «… спецкурсы» (курс)
    + замечания о строках сводки без суммы."""
    entries, problems = [], []
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    for ws in wb.worksheets:
        tokens = re.findall(r"[\wё-]+", ws.title.casefold())
        subject = next(filter(None, map(subject_of, tokens)), None)
        if not subject:
            continue
        rows = [list(r) for r in ws.iter_rows(max_row=60, values_only=True)]
        width = max((len(r) for r in rows), default=0)
        rows = [r + [None] * (width - len(r)) for r in rows]
        if not rows:
            continue
        if "гк" in tokens:
            # Шапка: месяцы дважды — сначала оплаты, потом тарифы. Сводка «подпись / сумма под ней»
            # стоит в колонках второго набора месяцев.
            cols = {}
            for c, h in enumerate(rows[0]):
                m = month_of(norm(h)) if isinstance(h, str) else None
                if m:
                    cols.setdefault(m, []).append(c)
            for m, cs in cols.items():
                if len(cs) < 2:
                    continue
                c = cs[-1]
                for r in range(1, len(rows) - 1):
                    label, value = rows[r][c], rows[r + 1][c]
                    if isinstance(label, str) and label.strip() and _is_num(value):
                        lab = norm(label).removeprefix("гк")
                        entries.append(Entry(ws.title, subject, m, lab, _dec(value), f"{get_column_letter(c + 1)}{r + 2}"))
        elif any(t.startswith("спецкурс") for t in tokens):
            # Строки «Количество | N | сумма | Курс»
            for r in range(1, len(rows)):
                a, c, d = rows[r][0], rows[r][2] if width > 2 else None, rows[r][3] if width > 3 else None
                if not (isinstance(a, str) and a.strip().casefold() == "количество"):
                    continue
                if not _is_num(c):
                    problems.append(f"«{ws.title}», строка {r + 1}: в C вместо суммы {f'«{c}»' if c is not None else 'пусто'}")
                elif isinstance(d, str) and norm(d) != "все":
                    entries.append(Entry(ws.title, subject, None, norm(d), _dec(c), f"C{r + 1}"))
    wb.close()
    return entries, problems


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _dec(v):
    return Decimal(str(v)).quantize(Decimal("0.01"))


# ---------- сайт (только чтение) ----------

def _tls():
    # Собранное .app ищет сертификаты там, где они лежали у Python на машине сборки; на macOS берём системные.
    if sys.platform == "darwin" and os.path.exists("/etc/ssl/cert.pem"):
        return ssl.create_default_context(cafile="/etc/ssl/cert.pem")
    return ssl.create_default_context()


class Site:
    def __init__(self, base):
        self.base = base.rstrip("/")
        self.opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=_tls()),
                                                  urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
        self.opener.addheaders = [("User-Agent", UA)]

    def _open(self, path, data=None):
        with self.opener.open(self.base + path, data=data, timeout=60) as r:
            return r.geturl(), r.read().decode(r.headers.get_content_charset() or "utf-8", "replace")

    def login(self, login, password):
        self._open("/Account/LogOn")
        form = urllib.parse.urlencode({"LogLogin": login, "LogPassword": password, "LogRememberMe": "false"})
        url, _ = self._open("/Account/LogOn", form.encode())
        if "/account/logon" in url.casefold():
            raise LoginError("Не удалось войти на сайт: проверьте логин и пароль.")

    def get(self, url):
        path = urllib.parse.urlsplit(url).path
        if not GROUP_PATH.fullmatch(path):  # читаем только страницы групп
            raise ValueError(f"не страница группы: {url}")
        for attempt in (1, 2):
            try:
                final, text = self._open(path)
                break
            except OSError:
                if attempt == 2:
                    raise
        if "/account/logon" in final.casefold():
            raise LoginError("Сайт попросил войти заново (сессия закончилась).")
        return text


def parse_group_page(text):
    """(название группы, «Получено от плательщиков») со страницы группы."""
    m = re.search(r"<x-caption[^>]*>\s*([^<]*?)\s*<", text) or re.search(r"<title>\s*(?:Группа\s+)?([^<]*?)\s*</title>", text)
    name = html.unescape(m.group(1)).strip() if m else ""
    m = re.search(r"Получено от плательщиков:\s*</td>\s*<td[^>]*>(.*?)</td>", text, re.S)
    if not m:
        raise ValueError("на странице нет «Получено от плательщиков»")
    return name, parse_money(html.unescape(re.sub(r"<[^>]+>", "", m.group(1))))


def parse_money(s):
    m = re.search(r"-?\d+(?:,\d+)?", re.sub(r"\s", "", s).replace("−", "-"))
    if not m:
        raise ValueError(f"не число: {s!r}")
    return Decimal(m.group().replace(",", "."))


# ---------- сопоставление ----------

def parse_group_name(name):
    """«26/27 октябрь ГК ОБЩ ПРОФИ 2» → ('ОБЩ', 'Октябрь', 'профи2')."""
    s = re.sub(r"\b\d{2}/\d{2}\b", " ", name.casefold().translate(LATIN_LOOKALIKES))
    subject = next(filter(None, map(subject_of, re.findall(r"[\wё+-]+", s))), None)
    head = re.split(r"[:/]", s, maxsplit=1)[0]  # «на пересдачу: математика / базовый минимум» → «на пересдачу»
    month, label = None, []
    for t in re.findall(r"[\wё+-]+", head):
        if t == "гк" or subject_of(t):
            continue
        if not month and month_of(t):
            month = month_of(t)
            continue
        label.append(t)
    label = "".join(label)
    return subject, month, ALIASES.get(label, label)


@dataclass
class Group:
    url: str
    name: str = ""
    site: Decimal | None = None
    error: str | None = None


@dataclass
class Row:  # одна строка отчёта: группа (или несколько групп на одну ячейку xlsx)
    groups: list
    status: str  # ключ из STATUS
    xlsx: Decimal | None = None
    cells: list = field(default_factory=list)
    note: str = ""

    @property
    def site(self):
        vals = [g.site for g in self.groups if g.site is not None]
        return sum(vals, ZERO) if vals else None

    @property
    def diff(self):
        return None if self.site is None or self.xlsx is None else self.site - self.xlsx


def match(groups, entries):
    rows, by_cells = [], {}
    for g in groups:
        if g.error:
            rows.append(Row([g], "error", note=g.error))
            continue
        subject, month, label = parse_group_name(g.name)
        hits = [e for e in entries if e.label == label
                and (subject is None or e.subject == subject)
                and (month is None or e.month == month)]
        sheets_by_subject = {}
        for e in hits:
            sheets_by_subject.setdefault(e.subject, set()).add(e.sheet)
        if not label or not hits:
            where = ", ".join(filter(None, [subject, month])) or "все листы"
            rows.append(Row([g], "missing" if g.site else "empty", note=f"в сводке xlsx ({where}) нет строки для этой группы"))
        elif any(len(s) > 1 for s in sheets_by_subject.values()):
            sheets = ", ".join(f"«{s}»" for ss in sheets_by_subject.values() for s in sorted(ss))
            rows.append(Row([g], "ambiguous", note=f"подходит несколько листов: {sheets}"))
        else:
            key = tuple(sorted((e.sheet, e.cell) for e in hits))
            if key in by_cells:  # несколько групп на сайте → одна ячейка xlsx: сравниваем сумму
                by_cells[key].groups.append(g)
                continue
            row = Row([g], "", xlsx=sum((e.value for e in hits), ZERO), cells=[f"«{s}»!{c}" for s, c in key])
            by_cells[key] = row
            rows.append(row)
    for row in by_cells.values():
        row.status = "ok" if abs(row.diff) < Decimal("0.01") else "diff"
        if len(row.groups) > 1:
            row.note = "на сайте несколько групп, в xlsx одна сумма — сравнивается их сумма"
    return rows


def reconcile(xlsx, pdf, fetch, progress=lambda i, n, msg: None, site=SITE):
    links, skipped = pdf_links(pdf, site)
    entries, problems = load_xlsx(xlsx)
    groups = []
    for i, url in enumerate(links, 1):
        g = Group(url)
        try:
            g.name, g.site = parse_group_page(fetch(url))
        except LoginError:
            raise
        except Exception as e:  # одна битая страница не должна ронять всю сверку
            g.error = f"{type(e).__name__}: {e}"
        progress(i, len(links), f"{g.name or url}: {fmt(g.site) if g.error is None else g.error}")
        groups.append(g)
    return match(groups, entries), links, skipped, problems


# ---------- отчёт ----------

STATUS = {"ok": "совпадает", "diff": "РАСХОЖДЕНИЕ", "missing": "РАСХОЖДЕНИЕ: нет в xlsx",
          "empty": "нет в xlsx, на сайте 0", "ambiguous": "неоднозначно", "error": "ошибка загрузки"}
PROBLEM = ("diff", "missing")
RED = RGBColor(0xC0, 0x00, 0x00)


def fmt(d):
    if d is None:
        return "—"
    return f"{d:,.2f}".replace(",", " ").replace(".", ",")


def _link(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    for tag, attr in (("w:color", "0563C1"), ("w:u", "single")):
        el = OxmlElement(tag)
        el.set(qn("w:val"), attr)
        rpr.append(el)
    run.append(rpr)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    run.append(t)
    h.append(run)
    paragraph._p.append(h)


def _table(doc, header, rows, red=lambda row: False):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    for cell, text in zip(t.rows[0].cells, header):
        cell.paragraphs[0].add_run(text).bold = True
    for row in rows:
        cells = t.add_row().cells
        for cell, value in zip(cells, row):
            p = cell.paragraphs[0]
            if isinstance(value, list):  # [(url, text), …] — ссылки на группы
                for j, (url, text) in enumerate(value):
                    if j:
                        p = cell.add_paragraph()
                    _link(p, url, text)
            else:
                run = p.add_run(str(value))
                if red(row):
                    run.font.color.rgb = RED
    return t


def write_report(path, rows, links, skipped, problems, xlsx, pdf):
    doc = docx.Document()
    sec = doc.sections[0]
    sec.orientation, sec.page_width, sec.page_height = WD_ORIENT.LANDSCAPE, sec.page_height, sec.page_width
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Pt(36))
    doc.styles["Normal"].font.size = Pt(9)

    doc.add_heading("Сверка «Получено от плательщиков»: сайт ↔ xlsx", level=1)
    doc.add_paragraph(f"Сформировано: {dt.datetime.now():%d.%m.%Y %H:%M}\n"
                      f"Таблица: {Path(xlsx).name}\nPDF: {Path(pdf).name} — ссылок на группы: {len(links)}")
    doc.add_paragraph(summary(rows))

    def cells_cell(r):
        c = r.cells
        return ", ".join(c) if len(c) <= 3 else f"{c[0]} … {c[-1]} (сумма {len(c)} ячеек)"

    def groups_cell(r):
        return [(g.url, g.name or g.url) for g in r.groups]

    def site_cell(r):
        if len(r.groups) == 1:
            return fmt(r.site)
        return " + ".join(fmt(g.site) for g in r.groups) + f" = {fmt(r.site)}"

    head = ["Группа на сайте", "Сайт, руб.", "xlsx, руб.", "Разница (сайт − xlsx)", "Ячейка xlsx", "Примечание"]
    doc.add_heading("Расхождения", level=2)
    diffs = [r for r in rows if r.status in PROBLEM]
    if diffs:
        _table(doc, head, [[groups_cell(r), site_cell(r), fmt(r.xlsx), fmt(r.diff), cells_cell(r), r.note] for r in diffs])
    else:
        doc.add_paragraph("Расхождений не найдено.")

    other = [r for r in rows if r.status in ("ambiguous", "error")]
    if other:
        doc.add_heading("Не удалось сверить", level=2)
        _table(doc, ["Группа на сайте", "Сайт, руб.", "Статус", "Причина"],
               [[groups_cell(r), site_cell(r), STATUS[r.status], r.note] for r in other])
    if problems:
        doc.add_heading("Замечания по сводке xlsx", level=2)
        for p in problems:
            doc.add_paragraph(p, style="List Bullet")
    if skipped:
        doc.add_heading("Пропущенные ссылки PDF (не группы или не " + urllib.parse.urlsplit(SITE).netloc + ")", level=2)
        for u in skipped:
            doc.add_paragraph(u, style="List Bullet")

    doc.add_heading("Все группы", level=2)
    _table(doc, ["Группа на сайте", "Сайт, руб.", "xlsx, руб.", "Разница", "Статус", "Ячейка xlsx"],
           [[groups_cell(r), site_cell(r), fmt(r.xlsx), fmt(r.diff), STATUS[r.status], cells_cell(r)] for r in rows],
           red=lambda row: row[4] in (STATUS[s] for s in PROBLEM))
    doc.save(path)


def run(xlsx, pdf, login, password, out=None, progress=lambda i, n, msg: None, site=SITE):
    links, _ = pdf_links(pdf, site)
    if not links:
        raise ValueError(f"В PDF нет ссылок на группы {site}/Learner/Group/<номер>.")
    web = Site(site)
    progress(0, len(links), f"Вход на {urllib.parse.urlsplit(site).netloc}…")
    web.login(login, password)
    rows, links, skipped, problems = reconcile(xlsx, pdf, web.get, progress, site)
    out = out or str(Path(xlsx).with_name(f"Расхождения {dt.datetime.now():%Y-%m-%d %H-%M}.docx"))
    write_report(out, rows, links, skipped, problems, xlsx, pdf)
    return out, rows


def summary(rows):
    n = {s: sum(len(r.groups) for r in rows if r.status == s) for s in STATUS}
    return (f"Групп: {sum(n.values())}. Совпадает: {n['ok']}. Расхождений: {n['diff'] + n['missing']}"
            f" (из них нет в xlsx: {n['missing']}). Нет в xlsx при нуле на сайте: {n['empty']}."
            f" Неоднозначно: {n['ambiguous']}. Ошибок загрузки: {n['error']}.")


def open_file(path):
    if sys.platform == "win32":
        os.startfile(path)
    else:
        subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", path], check=False)


# ---------- окно ----------

def gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Сверка ExpertOnline")
    f = ttk.Frame(root, padding=12)
    f.grid(sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    f.columnconfigure(1, weight=1)
    xlsx, pdf, login, password = (tk.StringVar() for _ in range(4))

    def pick(var, kind, ext):
        p = filedialog.askopenfilename(filetypes=[(kind, ext), ("Все файлы", "*.*")])
        if p:
            var.set(p)

    for r, (text, var, kind, ext) in enumerate([("Таблица xlsx", xlsx, "Excel", "*.xlsx"), ("PDF со ссылками", pdf, "PDF", "*.pdf")]):
        ttk.Label(f, text=text).grid(row=r, column=0, sticky="w", pady=3)
        ttk.Entry(f, textvariable=var, width=60).grid(row=r, column=1, sticky="ew", padx=6)
        ttk.Button(f, text="Выбрать…", command=lambda v=var, k=kind, e=ext: pick(v, k, e)).grid(row=r, column=2)
    ttk.Label(f, text="Логин").grid(row=2, column=0, sticky="w", pady=3)
    ttk.Entry(f, textvariable=login).grid(row=2, column=1, sticky="ew", padx=6)
    ttk.Label(f, text="Пароль").grid(row=3, column=0, sticky="w", pady=3)
    ttk.Entry(f, textvariable=password, show="•").grid(row=3, column=1, sticky="ew", padx=6)
    btn = ttk.Button(f, text="Сверить")
    btn.grid(row=4, column=1, sticky="w", padx=6, pady=8)
    bar = ttk.Progressbar(f, mode="determinate")
    bar.grid(row=5, column=0, columnspan=3, sticky="ew")
    log = tk.Text(f, height=16, width=90, state="disabled", wrap="word")
    log.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
    f.rowconfigure(6, weight=1)
    events = queue.Queue()

    def say(msg):
        log.configure(state="normal")
        log.insert("end", msg + "\n")
        log.see("end")
        log.configure(state="disabled")

    def work(args):
        try:
            out, rows = run(*args, progress=lambda i, n, msg: events.put(("p", i, n, msg)))
            events.put(("done", out, summary(rows)))
        except Exception as e:
            events.put(("err", f"{e}" if isinstance(e, (LoginError, ValueError)) else f"{type(e).__name__}: {e}"))

    def poll():
        try:
            while True:
                ev = events.get_nowait()
                if ev[0] == "p":
                    bar.configure(maximum=max(ev[2], 1), value=ev[1])
                    say(f"[{ev[1]}/{ev[2]}] {ev[3]}" if ev[1] else ev[3])
                elif ev[0] == "done":
                    btn.state(["!disabled"])
                    say(f"\n{ev[2]}\nОтчёт: {ev[1]}")
                    try:
                        open_file(ev[1])
                    except OSError as e:
                        say(f"Не удалось открыть отчёт: {e}")
                else:
                    btn.state(["!disabled"])
                    say(f"\nОшибка: {ev[1]}")
                    messagebox.showerror("Сверка", ev[1])
        except queue.Empty:
            pass
        root.after(100, poll)

    def start():
        paths = [xlsx.get().strip(), pdf.get().strip()]
        if not all(Path(p).is_file() for p in paths):
            return messagebox.showwarning("Сверка", "Выберите файлы xlsx и PDF.")
        if not login.get().strip() or not password.get():
            return messagebox.showwarning("Сверка", "Введите логин и пароль от сайта.")
        btn.state(["disabled"])
        log.configure(state="normal")
        log.delete("1.0", "end")
        log.configure(state="disabled")
        threading.Thread(target=work, args=((*paths, login.get().strip(), password.get()),), daemon=True).start()

    btn.configure(command=start)
    root.after(100, poll)
    root.mainloop()


def main():
    if len(sys.argv) == 1:
        return gui()
    ap = argparse.ArgumentParser(description="Сверка «Получено от плательщиков» с xlsx")
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out")
    ap.add_argument("--site", default=SITE, help=argparse.SUPPRESS)  # для проверки на локальной копии
    a = ap.parse_args()
    login = os.environ.get("SVERKA_LOGIN") or input("Логин: ")
    password = os.environ.get("SVERKA_PASSWORD") or getpass.getpass("Пароль: ")
    out, rows = run(a.xlsx, a.pdf, login, password, a.out,
                    progress=lambda i, n, msg: print(f"[{i}/{n}] {msg}" if i else msg, flush=True), site=a.site)
    print(summary(rows))
    print("Отчёт:", out)


if __name__ == "__main__":
    main()
