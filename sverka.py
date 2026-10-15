"""Сверка «Получено от плательщиков» (HolliHop, expertonline.t8s.ru) с таблицей учеников (xlsx).

Вход: xlsx с листами «<ПРЕДМЕТ> ГК» и «<ПРЕДМЕТ> спецкурсы» + PDF со ссылками на группы.
Выход: docx с расхождениями. Вход на сайт человек делает сам в окне приложения (обычная страница входа),
дальше программа только читает страницы групп из PDF (GET) этой сессией и ничего на сайте не меняет.

Без аргументов открывается окно. Консоль (окно сайта для входа всё равно откроется):
    python sverka.py --xlsx table.xlsx --pdf courses.pdf [--out report.docx]
"""
import argparse
import datetime as dt
import html
import json
import os
import re
import subprocess
import sys
import threading
import urllib.parse
import warnings
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import docx
import openpyxl
import pypdf
import webview
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

def group_path(url, site=SITE):
    """Путь страницы группы на нашем сайте; всё остальное читать отказываемся."""
    u, o = urllib.parse.urlsplit(url), urllib.parse.urlsplit(site)
    if (u.scheme, u.netloc.casefold()) != (o.scheme, o.netloc.casefold()) or not GROUP_PATH.fullmatch(u.path):
        raise ValueError(f"не страница группы {o.netloc}: {url}")
    return u.path


class WebSession:
    """Окно сайта: человек сам входит на обычной странице входа, дальше читаем страницы групп его сессией."""

    def __init__(self, site, return_path):
        self.site = site.rstrip("/")
        self.closed = threading.Event()
        self.win = webview.create_window(
            "Вход на сайт — войдите своим логином и паролем",
            f"{self.site}/Account/LogOn?ReturnUrl={urllib.parse.quote(return_path)}", width=1000, height=760)
        self.win.events.closed += self.closed.set

    def wait_login(self):
        host = urllib.parse.urlsplit(self.site).netloc.casefold()
        while not self.closed.wait(0.5):
            u = urllib.parse.urlsplit(self.win.get_current_url() or "")
            if u.netloc.casefold() == host and not u.path.casefold().startswith("/account/"):
                try:
                    if self.win.evaluate_js("document.readyState") == "complete":
                        return
                except Exception:  # страница ещё грузится
                    pass
        raise LoginError("Окно сайта закрыто — сверка отменена.")

    def get(self, url):
        path = group_path(url, self.site)
        if self.closed.is_set():
            raise LoginError("Окно сайта закрыто — сверка прервана.")
        done, box = threading.Event(), {}
        self.win.evaluate_js(  # обычный GET из окна сайта, с его же сессией
            f"fetch({json.dumps(path)}, {{credentials: 'same-origin', cache: 'no-store'}})"
            ".then(r => r.text().then(t => ({url: r.url, status: r.status, text: t})))",
            callback=lambda r: (box.setdefault("r", r), done.set()))
        if not done.wait(120):
            raise TimeoutError("сайт не ответил за 2 минуты")
        r = box["r"]
        if not isinstance(r, dict) or "text" not in r:
            raise ValueError(f"страница не загрузилась: {r}")
        if "/account/logon" in r["url"].casefold():
            raise LoginError("Сайт попросил войти заново (сессия закончилась).")
        if r["status"] != 200:
            raise ValueError(f"HTTP {r['status']}")
        return r["text"]

    def close(self):
        if not self.closed.is_set():
            self.win.destroy()


def login_window(pdf, site=SITE):
    links, _ = pdf_links(pdf, site)
    if not links:
        raise ValueError(f"В PDF нет ссылок на группы {site}/Learner/Group/<номер>.")
    return WebSession(site, group_path(links[0], site))


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


def run(xlsx, pdf, web, out=None, progress=lambda i, n, msg: None, site=SITE):
    progress(0, 0, "Войдите на сайт в открывшемся окне — логин и пароль вводите сами.")
    web.wait_login()
    progress(0, 0, "Вход выполнен, читаю страницы групп…")
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

APP_HTML = """<!doctype html><meta charset="utf-8"><style>
body { font: 14px -apple-system, "Segoe UI", sans-serif; margin: 16px; background: #f4f4f4; color: #222 }
.row { display: flex; gap: 10px; align-items: center; margin: 8px 0 }
.row span { color: #555; overflow: hidden; text-overflow: ellipsis; white-space: nowrap }
button { padding: 6px 14px; font: inherit }
progress { width: 100%; margin: 10px 0 }
pre { background: #fff; border: 1px solid #ccc; height: 280px; overflow: auto; padding: 8px; margin: 0; white-space: pre-wrap }
</style>
<div class="row"><button onclick="pick('xlsx')">Таблица xlsx…</button><span id="xlsx">не выбрана</span></div>
<div class="row"><button onclick="pick('pdf')">PDF со ссылками…</button><span id="pdf">не выбран</span></div>
<div class="row"><button id="go" onclick="go()">Сверить</button>
  <span>Откроется окно сайта — войдите в нём сами, дальше программа только читает страницы групп.</span></div>
<progress id="bar" value="0" max="1"></progress>
<pre id="log"></pre>
<script>
const files = {};
async function pick(kind) {
  const path = await pywebview.api.pick(kind);
  if (path) { files[kind] = path; document.getElementById(kind).textContent = path; }
}
function go() {
  if (!files.xlsx || !files.pdf) { say("Выберите таблицу xlsx и PDF."); return; }
  document.getElementById("go").disabled = true;
  document.getElementById("log").textContent = "";
  pywebview.api.start(files.xlsx, files.pdf);
}
function say(msg) { const l = document.getElementById("log"); l.textContent += msg + "\\n"; l.scrollTop = l.scrollHeight; }
function progress(i, n) { const b = document.getElementById("bar"); b.max = Math.max(n, 1); b.value = i; }
function done() { document.getElementById("go").disabled = false; }
</script>"""


class Api:  # методы доступны странице приложения как pywebview.api.*
    _win = None

    def pick(self, kind):
        types = ("Excel (*.xlsx)",) if kind == "xlsx" else ("PDF (*.pdf)",)
        paths = self._win.create_file_dialog(webview.FileDialog.OPEN, file_types=types)
        return paths[0] if paths else None

    def start(self, xlsx, pdf):
        threading.Thread(target=self._work, args=(xlsx, pdf), daemon=True).start()

    def _js(self, fn, *args):
        self._win.evaluate_js(f"{fn}({', '.join(json.dumps(a, ensure_ascii=False) for a in args)})")

    def _progress(self, i, n, msg):
        self._js("progress", i, n)
        self._js("say", f"[{i}/{n}] {msg}" if i else msg)

    def _work(self, xlsx, pdf):
        web = None
        try:
            web = login_window(pdf)
            out, rows = run(xlsx, pdf, web, progress=self._progress)
            self._js("say", f"\n{summary(rows)}\nОтчёт: {out}")
            try:
                open_file(out)
            except OSError as e:
                self._js("say", f"Не удалось открыть отчёт: {e}")
        except Exception as e:
            self._js("say", f"\nОшибка: {e if isinstance(e, (LoginError, ValueError)) else f'{type(e).__name__}: {e}'}")
        finally:
            if web:
                web.close()
            self._js("done")


WEBVIEW2_URL = "https://developer.microsoft.com/microsoft-edge/webview2/"


def webview2_missing():
    """На Windows без WebView2 pywebview молча берёт движок IE, где окна не работают — лучше сказать сразу."""
    if sys.platform != "win32":
        return False
    from webview.platforms import winforms  # тот же выбор движка, что сделает webview.start()
    return winforms.renderer != "edgechromium"


def gui():
    if webview2_missing():
        import ctypes
        import webbrowser
        ctypes.windll.user32.MessageBoxW(
            0, "Для окна сайта нужен Microsoft Edge WebView2 Runtime. Сейчас откроется страница загрузки — "
               "установите «Evergreen Bootstrapper» и запустите программу снова.", "Сверка", 0x30)
        return webbrowser.open(WEBVIEW2_URL)
    api = Api()
    api._win = webview.create_window("Сверка ExpertOnline", html=APP_HTML, js_api=api, width=820, height=560)
    webview.start(private_mode=True)


def main():
    if len(sys.argv) == 1:
        return gui()
    ap = argparse.ArgumentParser(description="Сверка «Получено от плательщиков» с xlsx")
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out")
    ap.add_argument("--site", default=SITE, help=argparse.SUPPRESS)  # для проверки на локальной копии
    a = ap.parse_args()
    for stream in (sys.stdout, sys.stderr):  # собранный .exe пишет в файл/pipe в cp1252 и падает на кириллице
        if stream is not None:
            stream.reconfigure(encoding="utf-8", errors="replace")
    if webview2_missing():
        sys.exit(f"Нужен Microsoft Edge WebView2 Runtime: {WEBVIEW2_URL}")
    web = login_window(a.pdf, a.site)

    def job():
        code = 1
        try:
            out, rows = run(a.xlsx, a.pdf, web, a.out, site=a.site,
                            progress=lambda i, n, msg: print(f"[{i}/{n}] {msg}" if i else msg, flush=True))
            print(summary(rows), "\nОтчёт:", out, flush=True)
            code = 0
        except Exception as e:
            print("Ошибка:", e, file=sys.stderr, flush=True)
        finally:
            web.close()
            # ponytail: выходим сразу — на Windows destroy() окна, которое ещё грузится, теряется и start() не вернётся
            os._exit(code)

    webview.start(job, private_mode=True)


if __name__ == "__main__":
    main()
