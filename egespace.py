"""Бэкап и восстановление предметов egespace.ru через её же API.

backup()  — качает дерево курсов, медиа-файлы и ссылки в папку.
restore() — заливает такую папку обратно: курс → раздел → урок → медиа.
Тесты и вопросы не восстанавливаются (в бэкапе лежат только их метаданные).
"""
import json
import os
import re
import time
import urllib.parse
import urllib.request

API = "https://egespace.ru/api"
MEDIA_TYPES = ("article", "podcast", "video", "link")
FILE_MEDIA = ("article", "podcast")  # качаются как файлы, остальные — внешние ссылки
TEST_TYPES = ("exam", "exercise")


class ApiError(Exception):
    pass


class Log:
    """Печатает статус в окно и дублирует его в файл рядом с копией."""

    def __init__(self, say, path=None):
        self.say, self.path, self.counts = say, path, {}
        if path:
            os.makedirs(os.path.dirname(path), exist_ok=True)

    def __call__(self, msg, *, count=None, n=1):
        if count:
            self.counts[count] = self.counts.get(count, 0) + n
        line = "%s  %s" % (time.strftime("%H:%M:%S"), msg)
        self.say(line)
        if self.path:
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(line + "\n")

    def fail(self, msg):
        self(("! " + msg), count="ошибки")

    def total(self, head):
        order = ["предметы", "разделы", "уроки", "медиа", "файлы", "вопросы", "тесты", "ошибки"]
        parts = ["%s: %s" % (k, self.counts[k]) for k in order if self.counts.get(k)]
        self("%s — %s" % (head, ", ".join(parts) if parts else "ничего не сделано"))


def _req(path, token=None, params=None, method="GET", body=None, binary=False, retries=3):
    url = API + path + ("?" + urllib.parse.urlencode(params) if params else "")
    data = None
    headers = {}
    if body is not None:
        if isinstance(body, bytes):
            data, headers["Content-Type"] = body, "application/octet-stream"
        else:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")
            headers["Content-Type"] = "application/json"
    if token:
        headers["Authorization"] = "Bearer " + token
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, data=data, method=method, headers=headers)
            with urllib.request.urlopen(req, timeout=180) as resp:
                raw = resp.read()
            if binary:
                return raw
            if not raw:
                return None
            try:
                return json.loads(raw.decode("utf-8"))
            except ValueError:  # часть ручек отвечает обычным текстом («Вопрос успешно создан.»)
                return raw.decode("utf-8", "replace")
        except Exception as e:  # сеть моргает чаще, чем ломается API
            if attempt == retries - 1:
                raise ApiError("%s %s: %s" % (method, path, e)) from e
            time.sleep(2 * (attempt + 1))


def login(user, password):
    token = _req("/validation/log_in", params={"login": user, "password": password},
                 method="POST", binary=True)
    return token.decode("utf-8").strip()


def slug(name, limit=60):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(name)).strip(" .")[:limit] or "unnamed"


def walk(node, kind):
    """Все узлы указанного типа в дереве /content."""
    if isinstance(node, list):
        for child in node:
            yield from walk(child, kind)
    elif isinstance(node, dict):
        if node.get("type") == kind:
            yield node
        yield from walk(node.get("contents") or [], kind)


def _save(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


# ---------- бэкап ----------

def default_out():
    """Куда класть копию, если папку не выбрали: Документы, а не случайный CWD."""
    docs = os.path.join(os.path.expanduser("~"), "Documents")
    return os.path.join(docs if os.path.isdir(docs) else os.path.expanduser("~"), "egespace_backup")


def backup(user, password, course_ids=(), out=None, say=print):
    """course_ids пустой — выгружаются все предметы. Возвращает путь к папке бэкапа."""
    out = out or default_out()
    token = login(user, password)
    meta = _req("/metadata", token) or {}
    courses = [c for c in meta.get("contents", []) if c.get("type") == "course"]
    if course_ids:
        wanted = set(course_ids)
        courses = [c for c in courses if c["id"] in wanted]
        for cid in sorted(wanted - {c["id"] for c in courses}):
            # /metadata показывает не всё (например, только что созданные предметы) —
            # спрашиваем про такой id напрямую
            node = next(iter(walk(_req("/content", token, {"course_id": cid}), "course")), None)
            if node:
                courses.append(node)
            else:
                say("Нет такого предмета: %s" % cid)
    if not courses:
        raise ApiError("нечего выгружать")

    root = os.path.join(out, time.strftime("%Y%m%d-%H%M%S"))
    log = Log(say, os.path.join(root, "backup.log"))
    log("Копия в %s" % os.path.abspath(root))
    _save(os.path.join(root, "metadata.json"), meta)
    for n, course in enumerate(courses, 1):
        log("[%d/%d] предмет «%s» (id %s)" % (n, len(courses), course.get("name"), course["id"]),
            count="предметы")
        try:
            _backup_course(token, course, root, log)
        except ApiError as e:
            log.fail("предмет %s не выгружен: %s" % (course["id"], e))
    log.total("Готово")
    return os.path.abspath(root)


def _description(token, kind, struct_id):
    try:
        return _req("/structure/description/%s/%s/exists" % (kind, struct_id), token, retries=1)
    except ApiError:  # описания может не быть — тогда ручка отдаёт Not Found
        return None


def _backup_course(token, course, root, log):
    tree = _req("/content", token, {"course_id": course["id"]})
    dest = os.path.join(root, "course_%s_%s" % (course["id"], slug(course.get("name"))))
    _save(os.path.join(dest, "course.json"), tree)
    # картинка, цвет, заголовок и преподаватель курса лежат в самом дереве,
    # отдельно нужны только тексты описаний и список преподавателей
    extras = {"descriptions": {}, "teachers": {}}
    for kind in ("course", "subject", "lesson"):
        for node in walk(tree, kind):
            desc = _description(token, kind, node["id"])
            if desc:
                extras["descriptions"]["%s/%s" % (kind, node["id"])] = desc
    for node in walk(tree, "course"):
        try:  # у курса без преподавателей ручка отвечает 404 — это не ошибка
            extras["teachers"][str(node["id"])] = _req(
                "/structure/course/%s/teachers/get" % node["id"], token, retries=1)
        except ApiError:
            extras["teachers"][str(node["id"])] = []
    _save(os.path.join(dest, "extras.json"), extras)
    log("   оформление: описаний %d, разделов %d" % (len(extras["descriptions"]), len(list(walk(tree, "subject")))),
        count="разделы", n=len(list(walk(tree, "subject"))))
    lessons = list(walk(tree, "lesson"))
    for i, lesson in enumerate(lessons, 1):
        try:
            got = _backup_lesson(token, lesson, os.path.join(dest, "lessons"), log)
        except ApiError as e:
            log.fail("урок %s не выгружен: %s" % (lesson["id"], e))
            continue
        log("   урок %d/%d «%s»: медиа %d (файлов %d), вопросов %d, тестов %d"
            % (i, len(lessons), lesson.get("name"), got["медиа"], got["файлы"],
               got["вопросы"], got["тесты"]), count="уроки")
        for k, v in got.items():
            if v:
                log.counts[k] = log.counts.get(k, 0) + v


def _backup_lesson(token, lesson, lessons_dir, log):
    lid = lesson["id"]
    base = os.path.join(lessons_dir, "%s_%s" % (lid, slug(lesson.get("name"))))
    data = {"lesson": lesson, "media": {}, "tests": {}}
    got = {"медиа": 0, "файлы": 0, "вопросы": 0, "тесты": 0}
    for mt in MEDIA_TYPES:
        items = _req("/media/%s/get_order" % mt, token, {"lesson_id": lid}) or []
        data["media"][mt] = items
        got["медиа"] += len(items)
        if mt not in FILE_MEDIA:
            continue
        for it in items:
            path = os.path.join(base, "files", "%s_%s.%s" % (mt, it["id"], it.get("extension") or "bin"))
            if os.path.exists(path):
                continue
            # у файлов авторизация идёт параметром jwt, заголовок тут отдаёт 422
            try:
                blob = _req("/media/%s/%s/get" % (mt, it["id"]), None, {"jwt": token}, binary=True)
            except ApiError as e:
                log.fail("файл «%s» (%s %s): %s" % (it.get("name"), mt, it["id"], e))
                continue
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(blob)
            got["файлы"] += 1
    questions, test_ids = {}, set(lesson.get("extended_tests_ids") or [])
    for tt in TEST_TYPES:
        listing = _req("/test/get_available_tests_teacher", token,
                       {"test_type": tt, "lesson_id": lid})
        data["tests"][tt] = listing
        for t in listing if isinstance(listing, list) else []:
            if isinstance(t, dict) and isinstance(t.get("id"), int):
                test_ids.add(t["id"])
        avail = _req("/question/get_available_questions", token,
                     {"lesson_id": lid, "test_type": tt}) or {}
        for q in avail.get("questions") or []:
            questions[q["id"]] = q  # один вопрос виден и в exam, и в exercise
    data["questions"] = [questions[k] for k in sorted(questions)]
    got["вопросы"], got["тесты"] = len(questions), len(test_ids)
    data["tests"]["details"] = {
        str(tid): {"metadata": _req("/test/%s/get_metadata" % tid, token),
                   "options": _req("/test/%s/get_options" % tid, token)}
        for tid in sorted(test_ids)}

    for qid in sorted(questions):  # картинка к вопросу, если она есть
        path = os.path.join(base, "files", "question_%s.png" % qid)
        if os.path.exists(path):
            continue
        try:
            blob = _req("/question/%s/image/get" % qid, None, {"jwt": token}, binary=True, retries=1)
        except ApiError:
            continue
        if blob:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(blob)

    _save(os.path.join(base, "lesson.json"), data)
    return got


# ---------- восстановление ----------

def restore(user, password, folder, say=print):
    """Создаёт на сайте копии предметов из папки бэкапа. Старые id не сохраняются."""
    token = login(user, password)
    courses = sorted(d for d in os.listdir(folder) if d.startswith("course_"))
    if not courses:
        raise ApiError("в папке нет выгруженных предметов (course_*)")
    log = Log(say, os.path.join(folder, "restore.log"))
    log("Восстанавливаю из %s" % os.path.abspath(folder))
    for n, name in enumerate(courses, 1):
        log("[%d/%d] %s" % (n, len(courses), name))
        try:
            _restore_course(token, os.path.join(folder, name), log)
        except (ApiError, OSError, ValueError) as e:
            log.fail("%s не восстановлен: %s" % (name, e))
    log.total("Готово")


def _create(token, kind, name, parent_id=None):
    res = _req("/structure/%s/create" % kind, token, method="POST",
               body={"name": name, "parentId": parent_id})
    new_id = res.get("id") if isinstance(res, dict) else res  # API отдаёт id строкой
    if not (isinstance(new_id, int) or (isinstance(new_id, str) and new_id.isdigit())):
        raise ApiError("создание %s %r не вернуло id: %r" % (kind, name, res))
    return int(new_id)


def _restore_course(token, course_dir, log):
    with open(os.path.join(course_dir, "course.json"), encoding="utf-8") as f:
        tree = json.load(f)
    files = {}  # id урока в бэкапе -> папка с его файлами
    for entry in sorted(os.listdir(os.path.join(course_dir, "lessons"))) if \
            os.path.isdir(os.path.join(course_dir, "lessons")) else []:
        files[int(entry.split("_", 1)[0])] = os.path.join(course_dir, "lessons", entry)

    extras_path = os.path.join(course_dir, "extras.json")
    extras = {}
    if os.path.exists(extras_path):
        with open(extras_path, encoding="utf-8") as f:
            extras = json.load(f)
    palette = {c.get("id") for c in (_req("/structure/course/colour/get", token) or [])}

    for course in walk(tree, "course"):
        cid = _create(token, "course", course.get("name"))
        log("   предмет «%s» создан, id %s" % (course.get("name"), cid), count="предметы")
        _restore_look(token, "course", course, cid, extras, palette, log)
        teachers = extras.get("teachers", {}).get(str(course["id"]))
        if isinstance(teachers, list) and teachers:
            ids = [t["id"] for t in teachers if isinstance(t, dict) and t.get("id")]
            try:
                _req("/structure/course/%s/teachers/update" % cid, token, method="POST", body=ids)
                log("   преподавателей назначено: %d" % len(ids))
            except ApiError as e:
                log.fail("преподаватели: %s" % e)
        for subject in walk(course.get("contents") or [], "subject"):
            sid = _create(token, "subject", subject.get("name"), cid)
            log("   раздел «%s» создан, id %s" % (subject.get("name"), sid), count="разделы")
            _restore_look(token, "subject", subject, sid, extras, palette, log)
            for lesson in walk(subject.get("contents") or [], "lesson"):
                lid = _create(token, "lesson", lesson.get("name"), sid)
                log("      урок «%s» создан, id %s" % (lesson.get("name"), lid), count="уроки")
                _restore_look(token, "lesson", lesson, lid, extras, palette, log)
                src = files.get(lesson["id"])
                if src:
                    _restore_media(token, src, lid, log)
                    _restore_tests(token, src, lid, log)


def _restore_look(token, kind, node, new_id, extras, palette, log):
    """Описание, заголовок, картинка, цвет, преподаватель и «детали курса» узла.
    Всё это живёт в одной ручке description/update и перезаписывается целиком."""
    desc = extras.get("descriptions", {}).get("%s/%s" % (kind, node["id"])) or {}
    body = {}
    if desc.get("text"):
        body["text"] = desc["text"]
    if node.get("title"):
        body["title"] = node["title"]
    if (node.get("image") or {}).get("id"):
        body["image"] = {"id": node["image"]["id"]}
    if node.get("course_details"):
        body["course_details"] = node["course_details"]
    if (node.get("colour") or {}).get("id") in palette and (node.get("colour") or {}).get("id"):
        body["colour"] = {"id": node["colour"]["id"]}  # цвет не из палитры сервер не принимает
    teacher = node.get("teacher") or {}
    if teacher.get("id"):
        body["teacher"] = {"id": teacher["id"], "name": teacher.get("name") or ""}
    if not body:
        return
    try:
        _req("/structure/description/%s/%s/create" % (kind, new_id), token, method="POST")
    except ApiError:
        pass  # описание уже есть — нормально
    try:
        _req("/structure/description/%s/%s/update" % (kind, new_id), token, method="POST", body=body)
    except ApiError as e:
        if "teacher" not in body:
            return log.fail("оформление %s %s: %s" % (kind, new_id, e))
        body.pop("teacher")  # у преподавателя нет профиля на этом сервере
        try:
            _req("/structure/description/%s/%s/update" % (kind, new_id), token, method="POST", body=body)
            log.fail("преподаватель «%s» не перенесён: %s" % (teacher.get("name"), e))
        except ApiError as e2:
            log.fail("оформление %s %s: %s" % (kind, new_id, e2))


def _restore_media(token, lesson_dir, lesson_id, log):
    with open(os.path.join(lesson_dir, "lesson.json"), encoding="utf-8") as f:
        data = json.load(f)
    for mt, items in (data.get("media") or {}).items():
        for it in items or []:
            try:
                if mt in FILE_MEDIA:
                    path = os.path.join(lesson_dir, "files",
                                        "%s_%s.%s" % (mt, it["id"], it.get("extension") or "bin"))
                    if not os.path.exists(path):
                        continue
                    with open(path, "rb") as f:
                        _req("/media/%s/upload" % mt, token,
                             {"lesson_id": lesson_id, "extension": it.get("extension") or "bin",
                              "name": it.get("name") or ""}, method="POST", body=f.read())
                    log("         файл «%s» залит" % it.get("name"), count="файлы")
                elif it.get("url"):
                    _req("/media/link/upload", token,
                         {"lesson_id": lesson_id, "name": it.get("name") or "", "url": it["url"]},
                         method="POST")
                    log("         ссылка «%s» добавлена" % it.get("name"), count="медиа")
            except ApiError as e:
                log.fail("%s «%s»: %s" % (mt, it.get("name"), e))


def _restore_tests(token, lesson_dir, lesson_id, log):
    """Вопросы урока, затем тесты. /question/create не возвращает id, поэтому
    после заливки читаем вопросы нового урока и сопоставляем по порядку."""
    with open(os.path.join(lesson_dir, "lesson.json"), encoding="utf-8") as f:
        data = json.load(f)
    old = data.get("questions") or []
    if not old:
        return
    made = []
    for q in old:
        body = {k: v for k, v in q.items() if k not in ("id", "creatorId")}
        body["parentLessonId"] = lesson_id
        try:
            _req("/question/create", token, method="POST", body=body)
            made.append(q)
            log("         вопрос «%s…» создан" % str(q.get("text"))[:30], count="вопросы")
        except ApiError as e:
            log.fail("вопрос «%s»: %s" % (str(q.get("text"))[:40], e))

    fresh = (_req("/question/get_available_questions", token,
                  {"lesson_id": lesson_id, "test_type": "exercise"}) or {}).get("questions") or []
    ids = {}  # старый id вопроса -> новый
    for q, new in zip(made, sorted(fresh, key=lambda x: x["id"])):
        ids[q["id"]] = new["id"]
        img = os.path.join(lesson_dir, "files", "question_%s.png" % q["id"])
        if os.path.exists(img):
            with open(img, "rb") as f:
                try:
                    _req("/question/%s/image/upload" % new["id"], token, {"extension": "png"},
                         method="POST", body=f.read())
                except ApiError as e:
                    log.fail("картинка к вопросу %s: %s" % (q["id"], e))

    for tt in TEST_TYPES:
        for t in data.get("tests", {}).get(tt) or []:
            if not isinstance(t, dict):
                continue
            body = {"name": t.get("name") or "Тест",
                    "timeLimit": t.get("timeLimit", 0) or 0,
                    "hideHints": bool(t.get("hideHints")),
                    "hideAnswers": bool(t.get("hideAnswers")),
                    "formationType": t.get("formationType") or "manual",
                    "autoModeTestSize": t.get("autoModeTestSize") or len(ids) or 1,
                    "autoModeAllowedDifficulties": t.get("autoModeAllowedDifficulties") or ["easy", "medium", "hard"],
                    "manualModeQuestionIds": [ids[q] for q in (t.get("manualModeQuestionIds") or []) if q in ids],
                    "isPublic": bool(t.get("isPublic"))}
            try:
                _req("/test/create", token, {"parent_lesson_id": lesson_id, "test_type": tt},
                     method="POST", body=body)
                log("         тест «%s» создан" % body["name"], count="тесты")
            except ApiError as e:
                log.fail("тест «%s»: %s" % (body["name"], e))


def _selftest():
    assert slug('a/b:c*?') == "a_b_c__" and slug("") == "unnamed"
    tree = {"type": "course", "contents": [
        {"type": "subject", "contents": [{"type": "lesson", "id": 1}]},
        {"type": "lesson", "id": 2}]}
    assert [l["id"] for l in walk(tree, "lesson")] == [1, 2]
    assert [c["type"] for c in walk(tree, "course")] == ["course"]
    print("selftest ok")


if __name__ == "__main__":
    _selftest()
