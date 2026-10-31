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
            return json.loads(raw.decode("utf-8")) if raw else None
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

def backup(user, password, course_ids=(), out="egespace_backup", say=print):
    """course_ids пустой — выгружаются все предметы. Возвращает путь к папке бэкапа."""
    token = login(user, password)
    meta = _req("/metadata", token) or {}
    courses = [c for c in meta.get("contents", []) if c.get("type") == "course"]
    if course_ids:
        wanted = set(course_ids)
        courses = [c for c in courses if c["id"] in wanted]
        missing = sorted(wanted - {c["id"] for c in courses})
        if missing:
            say("Нет таких предметов: %s" % ", ".join(map(str, missing)))
    if not courses:
        raise ApiError("нечего выгружать")

    root = os.path.join(out, time.strftime("%Y%m%d-%H%M%S"))
    _save(os.path.join(root, "metadata.json"), meta)
    for n, course in enumerate(courses, 1):
        say("[%d/%d] %s (id %s)" % (n, len(courses), course.get("name"), course["id"]))
        _backup_course(token, course, root, say)
    say("Готово: %s" % os.path.abspath(root))
    return os.path.abspath(root)


def _backup_course(token, course, root, say):
    tree = _req("/content", token, {"course_id": course["id"]})
    dest = os.path.join(root, "course_%s_%s" % (course["id"], slug(course.get("name"))))
    _save(os.path.join(dest, "course.json"), tree)
    lessons = list(walk(tree, "lesson"))
    for i, lesson in enumerate(lessons, 1):
        say("   урок %d/%d: %s" % (i, len(lessons), lesson.get("name")))
        _backup_lesson(token, lesson, os.path.join(dest, "lessons"))


def _backup_lesson(token, lesson, lessons_dir):
    lid = lesson["id"]
    base = os.path.join(lessons_dir, "%s_%s" % (lid, slug(lesson.get("name"))))
    data = {"lesson": lesson, "media": {}, "tests": {}}
    for mt in MEDIA_TYPES:
        items = _req("/media/%s/get_order" % mt, token, {"lesson_id": lid}) or []
        data["media"][mt] = items
        if mt not in FILE_MEDIA:
            continue
        for it in items:
            path = os.path.join(base, "files", "%s_%s.%s" % (mt, it["id"], it.get("extension") or "bin"))
            if os.path.exists(path):
                continue
            # у файлов авторизация идёт параметром jwt, заголовок тут отдаёт 422
            blob = _req("/media/%s/%s/get" % (mt, it["id"]), None, {"jwt": token}, binary=True)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(blob)
    for tt in TEST_TYPES:
        data["tests"][tt] = _req("/test/get_available_tests_teacher", token,
                                 {"test_type": tt, "lesson_id": lid})
    _save(os.path.join(base, "lesson.json"), data)


# ---------- восстановление ----------

def restore(user, password, folder, say=print):
    """Создаёт на сайте копии предметов из папки бэкапа. Старые id не сохраняются."""
    token = login(user, password)
    courses = sorted(d for d in os.listdir(folder) if d.startswith("course_"))
    if not courses:
        raise ApiError("в папке нет выгруженных предметов (course_*)")
    for n, name in enumerate(courses, 1):
        say("[%d/%d] %s" % (n, len(courses), name))
        _restore_course(token, os.path.join(folder, name), say)
    say("Готово.")


def _create(token, kind, name, parent_id=None):
    res = _req("/structure/%s/create" % kind, token, method="POST",
               body={"name": name, "parentId": parent_id})
    new_id = res.get("id") if isinstance(res, dict) else res
    if not isinstance(new_id, int):
        raise ApiError("создание %s %r не вернуло id: %r" % (kind, name, res))
    return new_id


def _restore_course(token, course_dir, say):
    with open(os.path.join(course_dir, "course.json"), encoding="utf-8") as f:
        tree = json.load(f)
    files = {}  # id урока в бэкапе -> папка с его файлами
    for entry in sorted(os.listdir(os.path.join(course_dir, "lessons"))) if \
            os.path.isdir(os.path.join(course_dir, "lessons")) else []:
        files[int(entry.split("_", 1)[0])] = os.path.join(course_dir, "lessons", entry)

    for course in walk(tree, "course"):
        cid = _create(token, "course", course.get("name"))
        say("   предмет «%s» -> id %s" % (course.get("name"), cid))
        for subject in walk(course.get("contents") or [], "subject"):
            sid = _create(token, "subject", subject.get("name"), cid)
            for lesson in walk(subject.get("contents") or [], "lesson"):
                lid = _create(token, "lesson", lesson.get("name"), sid)
                src = files.get(lesson["id"])
                if src:
                    _restore_media(token, src, lid, say)


def _restore_media(token, lesson_dir, lesson_id, say):
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
                elif it.get("url"):
                    _req("/media/link/upload", token,
                         {"lesson_id": lesson_id, "name": it.get("name") or "", "url": it["url"]},
                         method="POST")
            except ApiError as e:
                say("      ! %s «%s»: %s" % (mt, it.get("name"), e))


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
