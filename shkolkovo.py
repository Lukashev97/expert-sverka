"""Перенос заданий с Школково в тесты worldofege.ru / egespace.ru.

Логика целиком взята из скрипта test_creator (test_creator.py + create_lessons.py);
здесь она отвязана от командной строки и config.json: логин с паролем приходят
аргументами, а статус идёт в Log, как у бэкапа предметов.
"""
import json
import re
import urllib.error
import urllib.parse
import urllib.request
from html import escape, unescape

from egespace import ApiError, Log  # noqa: F401  (Log нужен вызывающей стороне)

PLATFORMS = ("worldofege.ru", "egespace.ru")  # один и тот же движок, разные сайты и аккаунты
SUBTYPES = ("directComparison", "sequence", "set")


class ShkolkovoError(ApiError):
    pass


def _fail(msg):
    raise ShkolkovoError(msg)


def http(url, data=None, headers=None):
    req = urllib.request.Request(url, data=data, headers={"User-Agent": "Mozilla/5.0", **(headers or {})})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read().decode("utf-8")


def parse_range(s):
    """'1-15' / '3' / '1-5,8,10-12' -> [1, 2, ...]"""
    nums = []
    for part in s.split(","):
        a, _, b = part.strip().partition("-")
        lo, hi = int(a), int(b or a)
        if not 1 <= lo <= hi:
            raise ValueError(s)
        nums += range(lo, hi + 1)
    return list(dict.fromkeys(nums))


def parse_answers(text):
    """Школково хранит допустимые варианты ответа по одному на строку."""
    return list(dict.fromkeys(a.strip() for a in text.replace("\xad", "").splitlines() if a.strip()))


def answer_variants(q):
    text = (q.get("Answer") or {}).get("text") or ""
    if q.get("AnswerProperty"):  # ответ-таблица: ячейки через запятую, строки через \n — это один ответ, а не варианты
        text = " ".join(c.strip() for c in re.split(r"[,\n]", text) if c.strip())
    return parse_answers(text)


def tex_html(session, origin):
    if not session:
        return ""
    # Формулы из LaTeX лежат относительными svg (index-<hash>.svg) рядом с сессией, картинки — /st/...
    base = f"{origin}/api/latex-service/v1/GetSession/{session['Id']}/"
    html = re.sub(r'\b(src|href)="([^"]*)"', lambda m: f'{m[1]}="{urllib.parse.urljoin(base, m[2])}"', session["Html"])
    # Редактор платформы (TipTap) понимает цветовую подсветку только как <mark>, span с background-color теряется
    html = re.sub(r'<span style="background-color:\s*([^;"]+);?">((?:(?!</?span).)*?)</span>',
                  r'<mark data-color="\1">\2</mark>', html, flags=re.S)
    return html.strip()


# Подпись об источнике в конце задания: «(Взято из сборника Дощинского-2026, 50 вар. (вар. 3))», «(ЕГКР 17.12.2024, вариант 953)»,
# «(Реальный ЕГЭ-2026; Банк ФИПИ)»… Авторство текста вроде «(По Л.Н. Толстому)» — часть задания, его не трогаем.
SOURCE_WORDS = re.compile(r"взят|егэ|егкр|фипи|сборник|дощинск|цыбульк|демоверс|досрочн|резерв|школково", re.I)
TRAILING_PAREN = re.compile(r"(?:\s|&nbsp;)*\(((?:[^()]|\([^()]*\))*)\)((?:\s|&nbsp;|</\w+>)*)$")
TRAILING_JUNK = re.compile(r"(?:\s|<!--(?:(?!-->).)*-->|<p[^>]*>(?:\s|&nbsp;|<(?!img)[^>]+>)*</p>)+$", re.S)  # пустые абзацы, комментарии


def strip_source(html):
    html = TRAILING_JUNK.sub("", html)
    m = TRAILING_PAREN.search(html)
    if m and SOURCE_WORDS.search(re.sub(r"<[^>]+>", "", m[1])):
        html = TRAILING_JUNK.sub("", html[:m.start()] + m[2])
    # скобки в подписи бывают непарными («…(вар. 21)))», «Взято из … (вар. 3))») — тогда снимаем последний абзац целиком
    last = re.search(r"<p[^>]*>((?:(?!<p[\s>]).)*)</p>$", html, re.S)
    if last and re.match(r"\(?\s*Взято из", unescape(re.sub(r"<[^>]+>", "", last[1])).strip()):
        html = TRAILING_JUNK.sub("", html[:last.start()])
    return html.strip()


def element(q, deps, origin, max_points, subtype):
    tex = deps.get("TexSessionsMap") or {}
    content = (deps.get("QuestionContent") or {}).get(str(q.get("RelatesToQuestionContentId")))  # общий текст к нескольким задачам
    files = [(deps.get("QuestionFiles") or {}).get(str(i)) for i in q.get("QuestionFiles") or []]
    text = tex_html(content and content["TexSession"], origin) + strip_source(tex_html(tex.get(str(q["QuestionTexSessionId"])), origin))
    text += "".join(f'<p><a href="{origin}{f["LocationOriginalSrc"]}">{escape(f["OriginalFilename"])}</a></p>' for f in files if f)
    solution = tex_html(tex.get(str(q["SolutionTexSessionId"])), origin)
    answers = answer_variants(q)
    return {
        "type": "simpleQuestion",
        "data": {
            "text": text,
            "clarification": {"text": solution, "attachments": []},
            "maxPoints": max_points,
            "subtype": subtype,
            # для set/sequence платформа хранит один ответ строкой, список вариантов — только у directComparison
            "answer": answers if subtype == "directComparison" else "".join(answers[:1]),
            "attachments": [],
        },
    }


def next_data(url):
    """Данные страницы Школково (Next.js кладёт их в __NEXT_DATA__); None — если такой страницы нет."""
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', http(url), re.S)
    return m and json.loads(m[1])["props"]["pageProps"]


def scrape(url, nums=None):
    """Номер N = «Задача N» на сайте: сквозная нумерация по страницам темы. nums=None — все задачи темы."""
    parts = urllib.parse.urlsplit(url)
    if not (parts.scheme and parts.netloc and re.search(r"/catalog/\d+", parts.path)):
        _fail(f"Ожидалась ссылка вида https://3.shkolkovo.online/catalog/13539, получено: {url}")
    origin = f"{parts.scheme}://{parts.netloc}"
    query = dict(urllib.parse.parse_qsl(parts.query))  # сохраняем ?SubTheme= и т.п.
    pages = {}

    def page(p):
        if p not in pages:
            pages[p] = next_data(f"{origin}{parts.path.rstrip('/')}?{urllib.parse.urlencode({**query, 'Page': p})}")
            if not pages[p]:
                _fail("Тема не найдена — откройте её в браузере и скопируйте адрес из адресной строки")
            if (pages[p].get("pagination") or {}).get("Page") != p:  # раздел редиректит на подтему и теряет ?Page
                _fail(f"Ссылка ведёт на раздел, а не на тему (Школково перенаправил на /catalog/{pages[p].get('themeId')}). "
                         "Откройте нужную тему и скопируйте адрес из адресной строки")
        return pages[p]

    pagination = page(1).get("pagination") or {}
    per, total = pagination.get("PerPage") or 20, pagination.get("TotalRecords") or 0
    nums = range(1, total + 1) if nums is None else nums
    missing = [n for n in nums if n > total]
    if missing:
        _fail(f"В теме {total} заданий, нет номеров: {missing}")
    for n in nums:
        ql = page((n - 1) // per + 1)["questionList"]
        yield n, ql["questions"][(n - 1) % per], ql["dependencies"], origin


def dedupe_key(html):
    """Сравниваем по тексту и картинкам: редактор платформы переписывает HTML задания при открытии."""
    html = re.sub(r'<img[^>]*?\ssrc="([^"]*)"[^>]*>', r" \1 ", html or "")
    return re.sub(r"\s+", "", unescape(re.sub(r"<[^>]+>", "", html)))


def parse_target(url):
    """Ссылка на тест -> (сайт, id теста). Хосты зашиты: пароль уходит только на свою платформу и только по https."""
    hosts = "|".join(map(re.escape, PLATFORMS))
    m = re.fullmatch(rf"https://(?:www\.)?({hosts})/extendedTests/\d+/\d+/\d+/(\d+)/?", url.split("#")[0].split("?")[0])
    if not m:
        _fail(f"Ожидалась ссылка на тест вида https://worldofege.ru/extendedTests/598/3761/19634/40245 "
                 f"(или то же на egespace.ru), получено: {url}")
    return m[1], m[2]


def login(host, user, password):
    """Заголовки с JWT для API платформы."""
    if not (user and password):
        _fail("Укажите логин и пароль для %s" % host)
    creds = urllib.parse.urlencode({"login": user, "password": password})
    token = http(f"https://{host}/api/validation/log_in?{creds}", data=b"").strip().strip('"')
    return {"Authorization": f"Bearer {token}"}


def post_json(url, body, auth):
    return http(url, data=json.dumps(body).encode("utf-8"), headers={**auth, "Content-Type": "application/json"})


def add_to_test(api, auth, test_id, elements):
    """Дописывает в тест задания, которых в нём ещё нет. -> (добавлено, пропущено, тест)"""
    test = json.loads(http(f"{api}/extended-tests/{test_id}", headers=auth))
    have = {dedupe_key((e.get("data") or {}).get("text")) for e in test["elements"]}
    new = [e for e in elements if dedupe_key(e["data"]["text"]) not in have]
    if new:
        test["elements"] += new
        test["autoUncompleted"] = bool((test.get("view") or {}).get("autoUncompleted"))  # так шлёт фронт платформы
        post_json(f"{api}/extended-tests/{test_id}/update", test, auth)
    return len(new), len(elements) - len(new), test


HOST = "egespace.ru"
API = f"https://{HOST}/api"
COURSE_ID = 35  # «Дополнительные материалы по русскому языку»
SUBJECT_ID = 104  # раздел «Задания» в этом курсе, сюда создаются уроки
CATALOG = "https://3.shkolkovo.online/catalog?SubjectId=2"  # все задания ЕГЭ по русскому
COLLECTION = re.compile(r"\d+\. Дощинский-2026 \(36 вар\.\)")
PART_1_VARIANTS = 18  # Часть 1 — варианты 1–18, Часть 2 — 19–36

# № задания ЕГЭ -> (тип проверки, балл)
EXACT = "directComparison"  # прямое сравнение
ORDERED = "sequence"  # символы в определённой последовательности
ANY_ORDER = "set"  # символы в произвольной последовательности
TASKS = {
    1: (EXACT, 1), 2: (ANY_ORDER, 1), 3: (ANY_ORDER, 1), 4: (ANY_ORDER, 1), 5: (EXACT, 1), 6: (EXACT, 1),
    7: (EXACT, 1), 8: (ORDERED, 2), 9: (ANY_ORDER, 1), 10: (ANY_ORDER, 1), 11: (ANY_ORDER, 1), 12: (ANY_ORDER, 1),
    13: (ANY_ORDER, 1), 14: (ANY_ORDER, 1), 15: (ANY_ORDER, 1), 16: (ANY_ORDER, 1), 17: (ANY_ORDER, 1), 18: (ANY_ORDER, 1),
    19: (ANY_ORDER, 1), 20: (ANY_ORDER, 1), 21: (ANY_ORDER, 1), 22: (ORDERED, 2), 23: (ANY_ORDER, 1), 24: (ANY_ORDER, 1),
    25: (EXACT, 1), 26: (EXACT, 1),
}


def variant(q):
    """Номер варианта сборника из названия задачи на Школково: «1.17», «8. 34», «№21.6» -> 17, 34, 6."""
    m = re.search(r"(\d+)\s*$", q.get("Name") or "")
    return int(m[1]) if m else 999  # без номера — в конец, во вторую часть


def collections(tasks):
    """№ задания -> (тема на Школково, подтема со сборником)."""
    pp = next_data(CATALOG) or _fail(f"Не открылся каталог Школково {CATALOG}")
    found = {}
    for parent in pp["firstColumn"] + pp["secondColumn"]:
        n = parent["ExamPosition"]
        if n in tasks:
            kids = [t for t in pp["themeList"] if t["ParentId"] == parent["Id"] and COLLECTION.fullmatch(t["Name"])]
            found[n] = (parent, kids[0] if kids else None)
    return found


def subject_lessons(auth):
    """Уроки раздела SUBJECT_ID: название -> id."""
    def find(node):
        if node.get("type") == "subject" and node.get("id") == SUBJECT_ID:
            return node
        return next(filter(None, map(find, node.get("contents") or [])), None)

    subject = find(json.loads(http(f"{API}/content?course_id={COURSE_ID}", headers=auth)))
    if not subject:
        _fail(f"В курсе {COURSE_ID} на {HOST} не найден раздел {SUBJECT_ID}")
    lessons = {}
    for lesson in subject.get("contents") or []:
        lessons.setdefault(lesson["name"].strip(), lesson["id"])
    return lessons


def lesson_tests(auth, lesson_id):
    """Тесты урока: название -> id. У урока без тестов API отвечает 404 «Tests are not found»."""
    try:
        data = json.loads(http(f"{API}/extended-tests/all?lessonId={lesson_id}", headers=auth))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return {}
        raise
    return {t["name"].strip(): t["id"] for t in data.get("tests") or []}




# ---------- сценарии ----------

def add_questions(user, password, source, target, nums, max_points, subtype, log):
    """Задачи темы Школково -> существующий тест платформы."""
    if subtype not in SUBTYPES:
        _fail("Неизвестный тип проверки: %s" % subtype)
    host, test_id = parse_target(target)
    elements = []
    for n, q, deps, origin in scrape(source, parse_range(nums) if nums else None):
        variants = answer_variants(q)
        log("Задача %s (#%s): ответ %s" % (n, q["Id"], " | ".join(variants) or "—"))
        if not variants:
            log.fail("у задачи %s нет текстового ответа — впишите его вручную" % n)
        elif len(variants) > 1 and subtype != "directComparison":
            log.fail("для %s платформа хранит один ответ — взят «%s», остальные потеряны" % (subtype, variants[0]))
        elements.append(element(q, deps, origin, max_points, subtype))
    added, skipped, test = add_to_test("https://%s/api" % host, login(host, user, password), test_id, elements)
    log("Добавлено %d, пропущено (уже были) %d; в тесте «%s» теперь %d заданий"
        % (added, skipped, test.get("name"), len(test["elements"])), count="вопросы", n=added)
    log.total("Готово")


def upload_lessons(user, password, tasks, log):
    """Сборник «Дощинский-2026 (36 вар.)» -> уроки и тесты курса COURSE_ID на egespace.ru."""
    wanted = set(parse_range(tasks)) if tasks else set(TASKS)
    for n in sorted(wanted - set(TASKS)):
        log("№%s нет в таблице заданий — пропускаю" % n)
    found = collections(wanted & set(TASKS))
    for n in sorted((wanted & set(TASKS)) - set(found)):
        log("№%s: такого задания ЕГЭ по русскому на Школково нет — пропускаю" % n)

    plan = []
    for n, (parent, coll) in sorted(found.items()):
        title = re.sub(r"^№\s*\d+\.\s*", "", parent["Name"]).strip()
        if not coll:
            log("№%s %s: нет сборника «Дощинский-2026 (36 вар.)» — пропускаю" % (n, title))
            continue
        # Школково отдаёт задачи в порядке добавления — раскладываем как в книге, по номеру варианта
        rows = sorted(scrape("https://3.shkolkovo.online/catalog/%s" % coll["Id"]), key=lambda r: variant(r[1]))
        if not rows:
            log("№%s %s: в сборнике нет задач — пропускаю" % (n, title))
            continue
        subtype, points = TASKS[n]
        elements = [element(q, deps, origin, points, subtype) for _, q, deps, origin in rows]
        half = sum(variant(q) <= PART_1_VARIANTS for _, q, _, _ in rows)
        parts = [("Задание №%s Часть %s" % (n, i), chunk)
                 for i, chunk in enumerate((elements[:half], elements[half:]), 1) if chunk]
        log("№%s %s: %d задач, %s, %s балл(а) → %s"
            % (n, title, len(elements), subtype, points,
               ", ".join("«%s» (%d)" % (name, len(chunk)) for name, chunk in parts)))
        variants = [answer_variants(q) for _, q, _, _ in rows]
        blank = sum(not v for v in variants)
        if blank:
            log.fail("у %d задач нет текстового ответа — впишите вручную" % blank)
        if subtype != "directComparison":
            lost = sum(len(v) > 1 for v in variants)
            if lost:
                log.fail("у %d задач несколько ответов, взят первый" % lost)
            repeats = sum(bool(v) and len(set(v[0])) < len(v[0]) for v in variants)
            if subtype == "set" and repeats:
                log.fail("у %d задач в ответе повторяются символы — редактор платформы считает такой ответ неверным" % repeats)
        plan.append(("Задание №%s - %s" % (n, title), parts))

    if not plan:
        _fail("Нечего выгружать")
    auth = login(HOST, user, password)
    lessons = subject_lessons(auth)
    for name, parts in plan:
        if name not in lessons:
            post_json(f"{API}/structure/lesson/create", {"name": name, "parentId": SUBJECT_ID}, auth)
            lessons = subject_lessons(auth)  # id нового урока ответ не отдаёт — перечитываем структуру, как фронт
            if name not in lessons:
                _fail("Урок «%s» не появился после создания" % name)
            log("Урок «%s» создан" % name, count="уроки")
        tests = lesson_tests(auth, lessons[name])
        for test_name, elements in parts:
            if test_name in tests:
                added, skipped, _ = add_to_test(API, auth, tests[test_name], elements)
                log("   «%s»: уже был, добавлено %d, пропущено %d" % (test_name, added, skipped),
                    count="вопросы", n=added)
            else:
                post_json(f"{API}/extended-tests/create", {
                    "name": test_name, "elements": elements, "view": {"autoUncompleted": False},
                    "autoUncompleted": False, "lessonId": str(lessons[name]),
                }, auth)
                log("   «%s» создан, %d задач" % (test_name, len(elements)), count="тесты")
                log.counts["вопросы"] = log.counts.get("вопросы", 0) + len(elements)
    log.total("Готово")


def _selftest():
    assert parse_range("1-3,7") == [1, 2, 3, 7]
    assert parse_answers("а\nб\n\nа") == ["а", "б"]
    assert strip_source("<p>Текст</p><p>(Взято из сборника Дощинского-2026, 50 вар. (вар. 3))</p>") == "<p>Текст</p>"
    assert strip_source("<p>Текст (По Л.Н. Толстому)</p>") == "<p>Текст (По Л.Н. Толстому)</p>"
    assert dedupe_key("<p>а б</p>") == dedupe_key("<p><b>а</b> б</p>")
    print("selftest ok")


if __name__ == "__main__":
    _selftest()
