"""Самопроверка без настоящего сайта: синтетические xlsx/PDF + локальный сервер, похожий на HolliHop.

    python test_sverka.py
"""
import html
import tempfile
import threading
import urllib.parse
from decimal import Decimal
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import docx
import openpyxl
from pypdf import PdfWriter
from pypdf.annotations import Link

import sverka

MONTHS = ["Сентябрь", "Октябрь"]
PAGES = {  # id группы → (название, «Получено от плательщиков»)
    1: ("26/27 сентябрь ГК ОБЩ БАЗА", "233&#160;596,00 руб."),       # совпадает
    2: ("26/27 октябрь ГК ОБЩ ПРОФИ 2", "75&#160;000,00 руб."),       # расхождение: в xlsx 75 105
    3: ("26/27 СВУ МВД ОБЩ", "275&#160;930,00 руб."),                 # сумма по месяцам
    4: ("26/27 октябрь ГК МАТ ТВИКС", "0,00"),                        # нет в xlsx, на сайте 0
    5: ("26/27 сентябрь ГК МАТ ТВИКС", "1&#160;000,00 руб."),         # нет в xlsx, на сайте деньги
    6: ("25/26 на пересдачу: математика / базовый минимум", "0,00"),   # две группы → одна ячейка
    7: ("25/26 на пересдачу: математика / роскошный максимум", "43&#160;330,00 руб."),
    8: ("С-П", "10&#160;000,00 руб."),                                 # алиас → «Стартер-пак»
    9: ("26/27 сентябрь ГК ОБЩ ПРО+ПА", "98&#160;671,00 руб."),
}
BROKEN = 10  # страница без «Получено от плательщиков»
LOGIN, PASSWORD = "test", "test-pass"
FOREIGN = "http://evil.example/Learner/Group/1"  # первая ссылка PDF на чужой хост — пароль туда уходить не должен


def make_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ОБЩ ГК "  # 26/27, с пробелом в конце, как в настоящем файле
    ws.append(["Дата", "ФИО", *MONTHS, *MONTHS, "Тариф"])
    summary = [("ГК база", 233596, 174291), ("ГК профи 2", 0, 75105),
               ("ГК про+ПА", 98671, 86181), ("СВУ МВД", 275930, 0)]
    for label, sep, oct_ in summary:
        ws.append(["ФАКТ", None, 1, 2, label, label])
        ws.append([None, None, "оплат", "оплат", sep, oct_])  # «оплат» над числом не должно стать подписью
    old = wb.create_sheet("ОБЩ ГК")  # прошлый год, другие тарифы — не должен мешать
    old.append(["Дата", "ФИО", *MONTHS, *MONTHS])
    old.append([None, None, None, None, "ГК pro+ПА", "ГК pro+ПА"])
    old.append([None, None, None, None, 1, 1])
    mat = wb.create_sheet("МАТ ГК")
    mat.append(["Дата", "ФИО", *MONTHS, *MONTHS])
    mat.append([None, None, None, None, "ГК база", "ГК база"])
    mat.append([None, None, None, None, 0, 12502])
    for name, rows in [("МАТ спецкурсы", [("Все", 43330), ("На пересдачу", 43330)]),
                       ("ОБЩ спецкурсы", [("Все", 9500), ("Стартер-пак", 9500)]),
                       ("ИСТ спецкурсы", [("Все", 500), ("для него в тарифе пустым оставлять", "Стартер-пак")])]:
        s = wb.create_sheet(name)
        s.append(["Дата", "ФИО", "Сумма", "Курс"])
        for course, total in rows:
            s.append(["Количество", 1, total, course])  # у ИСТ сумма съехала — как в настоящем файле
    wb.create_sheet("ЗП 26-27").append(["факт", 1, 2, "ГК база"])
    wb.save(path)


def make_pdf(path, base):
    w = PdfWriter()
    w.add_blank_page(600, 800)
    urls = [f"{base}/Learner/Group/{i}" if i != 8 else f"{base}/Learner/8"  # короткая: сайт редиректит на Group/8
            for i in [*PAGES, BROKEN]]
    urls.insert(1, urls[0])  # повтор ссылки — должна посчитаться один раз
    urls.append(f"{base}/Learner/Group/8")  # та же группа 8 в другом написании
    urls.insert(0, FOREIGN)
    urls.append("https://example.com/not-a-group")
    for i, url in enumerate(urls):
        w.add_annotation(0, Link(rect=(10, 10 + 15 * i, 200, 22 + 15 * i), url=url))
    w.write(path)


class FakeHolliHop(BaseHTTPRequestHandler):
    log = []

    def log_message(self, *a):
        pass

    def _send(self, code, body="", headers=()):
        self.send_response(code)
        for k, v in headers:
            self.send_header(k, v)
        data = body.encode()
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        self.log.append(("POST", self.path))
        form = urllib.parse.parse_qs(self.rfile.read(int(self.headers["Content-Length"])).decode())
        if self.path == "/Account/LogOn" and form.get("LogLogin") == [LOGIN] and form.get("LogPassword") == [PASSWORD]:
            return self._send(302, headers=[("Location", "/"), ("Set-Cookie", ".ASPXAUTH=ok; path=/; HttpOnly")])
        self._send(200, "<form>login</form>")

    def do_GET(self):
        self.log.append(("GET", self.path))
        path = urllib.parse.urlsplit(self.path).path
        if path in ("/", "/Account/LogOn"):
            return self._send(200, "<form>login</form>")
        if ".ASPXAUTH=ok" not in (self.headers.get("Cookie") or ""):
            return self._send(302, headers=[("Location", "/Account/LogOn?ReturnUrl=" + urllib.parse.quote(path))])
        parts = path.strip("/").split("/")
        if parts[:1] == ["Learner"] and parts[-1].isdigit() and len(parts) == 2:
            return self._send(302, headers=[("Location", f"/Learner/Group/{parts[-1]}")])
        gid = int(parts[-1])
        if gid == BROKEN:
            return self._send(200, "<title>Группа без статистики</title>")
        name, received = PAGES[gid]
        self._send(200, f"<title>Группа {html.escape(name)}</title><header><x-caption class=\"col-sm\">\r\n\r\n"
                        f"{html.escape(name)}                    <x-status>в работе</x-status></x-caption></header>"
                        f"<table><tr class=\"TrNoTopBorders\">\r\n<td>Получено от плательщиков:</td>\r\n"
                        f"<td><a href=\"/Payment/SearchLearnerPayments?LearnerId={gid}\">{received}</a></td></tr></table>")


def main():
    server = ThreadingHTTPServer(("127.0.0.1", 0), FakeHolliHop)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    base = f"http://127.0.0.1:{server.server_port}"
    tmp = Path(tempfile.mkdtemp())
    xlsx, pdf, out = tmp / "t.xlsx", tmp / "t.pdf", tmp / "r.docx"
    make_xlsx(xlsx)
    make_pdf(pdf, base)

    try:
        sverka.run(xlsx, pdf, LOGIN, "wrong", out, site=base)
        raise AssertionError("неверный пароль должен давать LoginError")
    except sverka.LoginError:
        pass

    FakeHolliHop.log.clear()
    _, rows = sverka.run(xlsx, pdf, LOGIN, PASSWORD, out, site=base)
    by_name = {g.name or g.url: r for r in rows for g in r.groups}
    status = {name: r.status for name, r in by_name.items()}
    assert status["26/27 сентябрь ГК ОБЩ БАЗА"] == "ok"
    assert status["26/27 сентябрь ГК ОБЩ ПРО+ПА"] == "ok", "старый лист «ОБЩ ГК» не должен мешать"
    r = by_name["26/27 октябрь ГК ОБЩ ПРОФИ 2"]
    assert (r.status, r.site, r.xlsx, r.diff, r.cells) == ("diff", Decimal("75000"), Decimal("75105.00"), Decimal("-105.00"), ["«ОБЩ ГК »!F5"])
    assert by_name["26/27 СВУ МВД ОБЩ"].xlsx == Decimal("275930") and status["26/27 СВУ МВД ОБЩ"] == "ok"
    assert status["26/27 октябрь ГК МАТ ТВИКС"] == "empty"
    assert status["26/27 сентябрь ГК МАТ ТВИКС"] == "missing"
    retake = by_name["25/26 на пересдачу: математика / роскошный максимум"]
    assert retake is by_name["25/26 на пересдачу: математика / базовый минимум"] and retake.status == "ok"
    sp = by_name["С-П"]
    assert (sp.status, sp.diff, sp.cells) == ("diff", Decimal("500.00"), ["«ОБЩ спецкурсы»!C3"])
    assert status[f"{base}/Learner/Group/{BROKEN}"] == "error"
    assert sum(len(r.groups) for r in rows) == len(PAGES) + 1  # + битая; повтор ссылки — один раз

    # только чтение: GET страниц групп и формы входа, POST — только сам вход
    assert {m for m, _ in FakeHolliHop.log} <= {"GET", "POST"}
    assert [p for m, p in FakeHolliHop.log if m == "POST"] == ["/Account/LogOn"]
    assert all(p in ("/", "/Account/LogOn") or p.startswith("/Learner/") for m, p in FakeHolliHop.log if m == "GET")
    try:
        sverka.Site(base).get(base + "/Payment/Delete/1")
        raise AssertionError("Site.get должен отказываться от не-групповых адресов")
    except ValueError:
        pass

    doc = docx.Document(out)
    text = "\n".join(c.text for t in doc.tables for row in t.rows for c in row.cells)
    assert "РАСХОЖДЕНИЕ" in text and "75\u00a0000,00" in text and "-105,00" in text, text[:500]
    paras = "\n".join(p.text for p in doc.paragraphs)
    assert FOREIGN in paras and "«ИСТ спецкурсы», строка 3: в C вместо суммы «Стартер-пак»" in paras, paras
    assert sverka.parse_money("−1 234,50 руб.") == Decimal("-1234.50")
    assert sverka.parse_group_name("26/27 октябрьГК ОБЩ ПРО + ВГ") == ("ОБЩ", "Октябрь", "про+вг")
    server.shutdown()
    print("ok:", sverka.summary(rows))


if __name__ == "__main__":
    main()
