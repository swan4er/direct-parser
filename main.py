#!/usr/bin/env python3
"""Сбор рекламных объявлений Яндекс.Директа со страницы yandex.ru/search/ads.

Без аргументов в терминале — ручной режим: программа спрашивает запросы и город,
показывает карточку с настройками и временем сбора, любую строку можно поменять.
С аргументами — режим для скриптов и агентов (python main.py --help). Ход работы
пишется в stderr, данные — в файл (по умолчанию в папку results) или в stdout.
"""

import argparse
import codecs
import csv
import difflib
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import textwrap
import time
from pathlib import Path


def _utf8_stdio():
    """Windows пишет перенаправленный вывод в кодировке системы (cp1251, cp866),
    и кириллица в нём ломается. Переключаем stdout и stderr на UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")


def _launched_interactively():
    return len(sys.argv) == 1 and sys.stdin is not None and sys.stdin.isatty()


def _pause():
    """Окно, открытое в Windows двойным щелчком по main.py, закрывается сразу после
    выхода — даём прочитать итог."""
    if sys.platform == "win32" and _launched_interactively():
        try:
            input("\nНажмите Enter, чтобы закрыть окно…")
        except (EOFError, KeyboardInterrupt):
            pass


def _offer_install(missing):
    """Новичку «pip install» ничего не говорит — предлагаем поставить библиотеки сами
    и перезапускаем программу. Возвращает код выхода."""
    requirements = Path(__file__).resolve().parent / "requirements.txt"
    manual = f'{sys.executable} -m pip install -r "{requirements}"'
    print(f"Программе нужны дополнительные библиотеки Python (не хватает «{missing}»). "
          "Их можно установить автоматически, нужен интернет.")
    try:
        answer = input("Установить сейчас? Enter — да, 0 — нет: ").strip()
    except (EOFError, KeyboardInterrupt):
        answer = "0"
    if answer == "0":
        print(f"Установить можно и вручную, командой:\n  {manual}")
        _pause()
        return 1
    if subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(requirements)]).returncode:
        print("Установить не получилось — причина выше. Инструкция по установке — "
              f"в README.md, раздел «Установка». Команда для ручной установки:\n  {manual}")
        _pause()
        return 1
    print("Библиотеки установлены. Запускаю программу…\n")
    child = subprocess.Popen([sys.executable, *sys.argv])
    # Ctrl+C обрабатывает запущенная программа: она сохраняет собранное. Если бы этот
    # процесс тоже прервался, он убил бы её на середине записи.
    signal.signal(signal.SIGINT, signal.SIG_IGN)
    return child.wait()


try:
    import requests
    from bs4 import BeautifulSoup
    from openpyxl import Workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    from openpyxl.styles import Font
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
except ImportError as exc:
    _utf8_stdio()
    if _launched_interactively():
        sys.exit(_offer_install(exc.name))
    print(f"Не установлена библиотека {exc.name}. Выполните: pip install -r requirements.txt",
          file=sys.stderr)
    sys.exit(1)

BASE_URL = "https://yandex.ru/search/ads"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36"
)
REGIONS_DOC = "https://aistudio.yandex.ru/docs/ru/search-api/reference/regions"
RESULTS_DIR = Path(__file__).resolve().parent / "results"

# Коды регионов (параметр lr) из справочника REGIONS_DOC: страны и крупные города.
# Остальные регионы задаются числовым кодом.
REGIONS = {
    "Россия": 225, "Беларусь": 149, "Казахстан": 159, "Украина": 187,
    "Москва": 213, "Москва и Московская область": 1, "Санкт-Петербург": 2,
    "Архангельск": 20, "Астрахань": 37, "Барнаул": 197, "Белгород": 4, "Благовещенск": 77,
    "Брянск": 191, "Великий Новгород": 24, "Владивосток": 75, "Владикавказ": 33,
    "Владимир": 192, "Волгоград": 38, "Вологда": 21, "Воронеж": 193, "Грозный": 1106,
    "Екатеринбург": 54, "Иваново": 5, "Иркутск": 63, "Йошкар-Ола": 41, "Казань": 43,
    "Калининград": 22, "Кемерово": 64, "Кострома": 7, "Краснодар": 35, "Красноярск": 62,
    "Курган": 53, "Курск": 8, "Липецк": 9, "Махачкала": 28, "Мурманск": 23, "Назрань": 1092,
    "Нальчик": 30, "Нижний Новгород": 47, "Новосибирск": 65, "Омск": 66, "Орёл": 10,
    "Оренбург": 48, "Пенза": 49, "Пермь": 50, "Псков": 25, "Ростов-на-Дону": 39,
    "Рязань": 11, "Самара": 51, "Саранск": 42, "Смоленск": 12, "Сочи": 239,
    "Ставрополь": 36, "Сургут": 973, "Тамбов": 13, "Тверь": 14, "Томск": 67, "Тула": 15,
    "Ульяновск": 195, "Уфа": 172, "Хабаровск": 76, "Чебоксары": 45, "Челябинск": 56,
    "Черкесск": 1104, "Ярославль": 16,
}
REGION_ALIASES = {"мск": 213, "москва и область": 1, "спб": 2, "питер": 2, "екб": 54}

# Ключ в данных -> заголовок колонки в Excel/CSV
AD_COLUMNS = {
    "query": "Ключевой запрос",
    "page": "Страница",
    "position": "Лучшая позиция",
    "seen": "Встречалось в прогонах",
    "domain": "Домен",
    "title": "Заголовок",
    "description": "Описание",
    "sitelinks": "Быстрые ссылки",
    "label": "Метка",
    # Переход по этой ссылке засчитывается рекламодателю как клик по объявлению
    "click_url": "Ссылка (счётчик кликов Яндекса)",
}
QUERY_COLUMNS = {
    "query": "Ключевой запрос",
    "status": "Статус",
    "runs_with_ads": "Прогонов с объявлениями",
    "pages": "Страниц загружено",
    "ads": "Объявлений",
    "message": "Пояснение",
}
OUTPUT_FORMATS = (".xlsx", ".csv", ".json")

# Замер 15.10.2026, 10 прогонов подряд по двум запросам: новые рекламодатели приходят
# в основном во 2–6 прогонах, после 6-го — по одному за прогон или ни одного.
# Больше 10 прогонов почти ничего не добавляют, а нагрузка на Яндекс и риск капчи растут.
MAX_RUNS = 10
SECONDS_PER_REQUEST = 0.7  # средний ответ Яндекса в том же замере — 0,64 с

# Коды выхода
EXIT_OK = 0  # собрано хотя бы одно объявление, ошибок нет
EXIT_ERRORS = 1  # капча, нет связи, незнакомая вёрстка или необработанные запросы
EXIT_USAGE = 2  # неверные аргументы (этот же код возвращает argparse)
EXIT_NO_ADS = 3  # ошибок нет, но Яндекс не показал ни одного объявления
EXIT_INTERRUPTED = 130  # остановлено по Ctrl+C (собранное сохранено)

MAX_NETWORK_ERRORS = 3  # столько ошибок связи подряд — и сбор останавливается
TIMEOUT = (10, 30)  # секунд: на соединение и на ответ

# skipped — до запроса не дошли: сбор остановлен раньше
ERROR_STATUSES = {"captcha", "error", "unexpected", "skipped"}
STATUS_MESSAGES = {
    "no_ads": "Яндекс ответил «ничего не нашлось».",
    "captcha": "Яндекс показал капчу.",
    "unexpected": "Незнакомая страница: похоже, Яндекс поменял вёрстку.",
}
STOP_MESSAGES = {
    "captcha": "Яндекс показал капчу — сбор остановлен",
    "network": f"нет связи с Яндексом {MAX_NETWORK_ERRORS} раза подряд — сбор остановлен",
    "interrupted": "сбор остановлен по Ctrl+C",
}


# ---------- Цвет в терминале ----------

STYLES = {"bold": "1", "dim": "2", "red": "31", "green": "32", "yellow": "33", "cyan": "36"}
_ansi = sys.platform != "win32"  # консоль Windows включается в _enable_ansi()


def _enable_ansi():
    """Консоль Windows понимает цветовые коды только после включения режима VT
    (Windows 10 и новее). Не получилось — работаем без цвета."""
    global _ansi
    if sys.platform != "win32":
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        enabled = False
        for handle_id in (-11, -12):  # STD_OUTPUT_HANDLE, STD_ERROR_HANDLE
            handle = kernel32.GetStdHandle(handle_id)
            mode = ctypes.c_uint32()
            if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                # 0x0004 — ENABLE_VIRTUAL_TERMINAL_PROCESSING
                enabled = bool(kernel32.SetConsoleMode(handle, mode.value | 0x0004)) or enabled
        _ansi = enabled
    except Exception:  # цвет — украшение: из-за него программа падать не должна
        _ansi = False


def paint(text, *styles, stream=None):
    """Раскрашивает текст, только если он идёт в настоящий терминал: в файлах
    и выводе, который перехватывает агент, цветовые коды были бы мусором.
    NO_COLOR в окружении выключает цвет (соглашение no-color.org)."""
    stream = stream or sys.stderr
    if (not styles or not _ansi or os.environ.get("NO_COLOR")
            or os.environ.get("TERM") == "dumb" or not stream.isatty()):
        return text
    return f"\033[{';'.join(STYLES[style] for style in styles)}m{text}\033[0m"


def log(message, *styles):
    print(paint(message, *styles, stream=sys.stderr), file=sys.stderr, flush=True)


def _wrap(text):
    """Длинные строки — по словам под ширину окна, иначе терминал рвёт слова («Вых|од»)."""
    if not sys.stdout.isatty():
        return text
    width = max(40, shutil.get_terminal_size().columns - 1)
    lines = []
    for line in text.split("\n"):
        indent = line[:len(line) - len(line.lstrip())]
        lines.append(textwrap.fill(line.strip(), width, initial_indent=indent,
                                   subsequent_indent=indent, break_long_words=False,
                                   break_on_hyphens=False) if line.strip() else line)
    return "\n".join(lines)


def say(message="", *styles):
    print(paint(_wrap(message), *styles, stream=sys.stdout))


# ---------- Регионы и запросы ----------

def _norm(name):
    return " ".join(name.lower().replace("ё", "е").replace("-", " ").split())


_REGION_LOOKUP = {_norm(name): code for name, code in {**REGIONS, **REGION_ALIASES}.items()}
_REGION_NAMES = {code: name for name, code in REGIONS.items()}


def parse_region(value):
    """Код региона из строки: число («36») или название из справочника («Ставрополь»)."""
    value = value.strip()
    if value.isdigit() and int(value) > 0:
        return int(value)
    code = _REGION_LOOKUP.get(_norm(value))
    if code is None:
        raise ValueError(f"Не знаю региона «{value}». Укажите числовой код, все коды: {REGIONS_DOC}")
    return code


def _decode(data):
    """Текст файла запросов: UTF-8 (с BOM и без), UTF-16 с BOM («Юникод» в Excel
    и Блокноте), иначе cp1251 (ANSI в русской Windows, так сохраняет Excel)."""
    if data.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return data.decode("utf-16")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("cp1251", errors="replace")


# Служебные коды клавиш: стрелка без модуля readline даёт «^[[D» прямо в тексте
_KEY_CODES = re.compile(r"\x1b(?:\[[0-?]*[ -/]*[@-~]|O.|.)?")
_CONTROL_CHARS = re.compile(r"[\x00-\x1f\x7f]")


def clean_text(text):
    """Убирает из введённого текста коды клавиш и управляющие символы, лишние пробелы."""
    return " ".join(_CONTROL_CHARS.sub(" ", _KEY_CODES.sub("", text)).split())


def _file_text(path):
    """Текст файла запросов. Excel и другие не текстовые файлы — понятная ошибка
    (ValueError), а не сотни «запросов» из байтов архива."""
    data = Path(path).read_bytes()
    utf16 = data.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE))
    if not utf16 and (data.startswith(b"PK\x03\x04") or b"\x00" in data[:4096]):
        raise ValueError(f"«{Path(path).name}» — не текстовый файл. Нужен .txt с запросами "
                         "по одному на строку (в Excel: «Сохранить как» → «Текст»).")
    return _decode(data)


def read_queries(arg_queries, file_path=None):
    """Собирает запросы из аргументов и файла ("-" — stdin), убирает пустые строки,
    комментарии (#), служебные символы и дубли, сохраняя порядок."""
    lines = list(arg_queries)
    if file_path == "-":
        lines += _decode(sys.stdin.buffer.read()).splitlines()
    elif file_path:
        lines += _file_text(file_path).splitlines()
    cleaned = (clean_text(line) for line in lines)
    return list(dict.fromkeys(q for q in cleaned if q and not q.startswith("#")))


# ---------- Разбор страницы и сбор ----------

def _text(element):
    # Яндекс ставит неразрывные пробелы («Aroma\xa0Box»); для анализа нужны обычные
    return element.get_text(" ", strip=True).replace("\xa0", " ") if element else ""


def parse_page(html, final_url=""):
    """Разбирает страницу выдачи. Возвращает {"status", "ads"}, где status:
    ok — объявления найдены; no_ads — Яндекс ответил «ничего не нашлось»;
    captcha — показана капча; unexpected — ни объявлений, ни пометки
    «ничего не нашлось» (скорее всего, Яндекс поменял вёрстку)."""
    if "showcaptcha" in final_url or "/checkcaptcha" in html:
        return {"status": "captcha", "ads": []}

    soup = BeautifulSoup(html, "html.parser")
    ads = []
    for item in soup.select("#search-result li.serp-item[data-cid]"):
        title = _text(item.select_one(".OrganicTitle-LinkText"))
        link = item.select_one("a.Link[href]")
        if not title and not link:
            continue
        # «randewoo.ru › Парфюм в интернет-магазине…» -> «randewoo.ru»
        path = _text(item.select_one(".Organic-Path .Path-Item"))
        ads.append({
            "domain": path.split("›")[0].strip(),
            "title": title,
            "description": _text(item.select_one(".OrganicTextContentSpan")),
            "sitelinks": [_text(el) for el in item.select(".Sitelinks-Title")],
            "label": _text(item.select_one(".AdvLabel-Text")),
            "click_url": link["href"] if link else "",
        })

    if ads:
        status = "ok"
    elif soup.select_one(".RequestMeta-Level_type_error"):
        status = "no_ads"
    else:
        status = "unexpected"
    return {"status": status, "ads": ads}


def make_session(proxy=None):
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "ru-RU,ru;q=0.9"})
    # Перегрузка Яндекса (429, 5xx) проходит сама — до трёх повторов с паузой. Обрыв связи
    # повторять почти бесполезно — один повтор, иначе каждое обращение ждёт секунды
    retry = Retry(total=3, connect=1, backoff_factor=1,
                  status_forcelist=(429, 500, 502, 503, 504))
    session.mount("https://", HTTPAdapter(max_retries=retry))
    if proxy:
        session.proxies.update({"http": proxy, "https": proxy})
    return session


def _problem(info, status, message):
    """Ошибка важнее «ok», «no_ads» и «skipped»; из нескольких ошибок запоминается первая."""
    if info["status"] not in ERROR_STATUSES - {"skipped"}:
        info.update(status=status, message=message)


def _network_reason(exc):
    """Понятная причина вместо «ProxyError(MaxRetryError(…%D0%BA…))»."""
    if isinstance(exc, requests.HTTPError) and exc.response is not None:
        return f"Яндекс ответил ошибкой {exc.response.status_code}"
    if isinstance(exc, requests.exceptions.RetryError):
        return "Яндекс не отвечает: ошибка на его стороне"
    if isinstance(exc, (requests.ConnectionError, requests.Timeout)):
        return "нет связи с Яндексом"
    return f"ошибка сети ({type(exc).__name__})"


def _safe_name(query):
    return re.sub(r"[^\w-]+", "_", query)[:60]


def _save_unexpected(query, html):
    """Незнакомую страницу сохраняем сами — по ней можно обновить разбор."""
    path = RESULTS_DIR / "debug" / (
        f"unexpected_{time.strftime('%Y-%m-%d_%H-%M-%S')}_{_safe_name(query)}.html")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(html, encoding="utf-8")
    except OSError:
        return None
    return path


def _snapshot(merged, summary, queries, stop=None):
    """Текущий результат: объявления по порядку запросов и позиций, сводка по запросам."""
    order = {query: index for index, query in enumerate(queries)}
    ads = sorted(merged.values(), key=lambda ad: (order[ad["query"]], ad["position"]))
    rows = []
    for query in queries:
        info = dict(summary[query], ads=sum(1 for ad in ads if ad["query"] == query))
        if info["status"] == "no_ads":
            info["message"] = STATUS_MESSAGES["no_ads"]
        elif info["status"] == "skipped":
            info["message"] = (f"Не обработан: {STOP_MESSAGES[stop]}." if stop
                               else "Ещё не обработан.")
        rows.append(info)
    return ads, rows


def collect(session, queries, pages=1, region=225, delay=1.0, runs=1, save_html=None,
            on_progress=None):
    """Обходит запросы и страницы runs раз и объединяет прогоны: Яндекс каждый раз
    показывает немного разный набор объявлений. Одинаковые объявления (тот же домен,
    заголовок и текст) склеиваются: считается, в скольких прогонах они встретились,
    и остаётся лучшая позиция.

    Останавливается на первой капче (дальше Яндекс будет показывать только её),
    после MAX_NETWORK_ERRORS ошибок связи подряд и по Ctrl+C — собранное не теряется.
    on_progress(объявления, сводка) вызывается после каждого запроса: так результат
    можно сохранять по ходу. Возвращает (объявления, сводка, причина остановки или None)."""
    merged = {}
    summary = {query: {"query": query, "status": "skipped", "runs_with_ads": 0, "pages": 0,
                       "ads": 0, "message": ""} for query in queries}
    first_request = True
    network_errors = 0
    stop = None
    try:
        # Сначала все запросы, потом следующий прогон — между повторами больше времени
        for run_number in range(1, runs + 1):
            for number, query in enumerate(queries, 1):
                info = summary[query]
                seen_now = set()
                position = 0
                for page in range(pages):
                    if not first_request and delay:
                        time.sleep(delay)
                    first_request = False
                    prefix = f"[{number}/{len(queries)}] «{query}», стр. {page + 1}"
                    if runs > 1:
                        prefix = f"[прогон {run_number}/{runs}] {prefix}"
                    try:
                        response = session.get(
                            BASE_URL, params={"text": query, "p": page, "lr": region}, timeout=TIMEOUT
                        )
                        response.raise_for_status()
                    except requests.RequestException as exc:
                        reason = _network_reason(exc)
                        _problem(info, "error", f"{reason} ({type(exc).__name__}).")
                        log(f"{prefix}: {reason}.", "red")
                        network_errors += 1
                        if network_errors >= MAX_NETWORK_ERRORS:
                            stop = "network"
                        break
                    network_errors = 0
                    if info["status"] == "skipped":
                        info["status"] = "no_ads"

                    if save_html:
                        run_suffix = f"_r{run_number}" if runs > 1 else ""
                        path = Path(save_html) / (
                            f"{number:03d}_{_safe_name(query)}_p{page + 1}{run_suffix}.html")
                        path.parent.mkdir(parents=True, exist_ok=True)
                        path.write_text(response.text, encoding="utf-8")

                    parsed = parse_page(response.text, response.url)
                    info["pages"] = max(info["pages"], page + 1)

                    if not parsed["ads"]:
                        # На первой странице пустой ответ — результат запроса; дальше —
                        # конец списка объявлений. Капча — ошибка на любой странице.
                        status = parsed["status"]
                        if status == "captcha":
                            _problem(info, "captcha", STATUS_MESSAGES["captcha"])
                            log(f"{prefix}: {STATUS_MESSAGES['captcha']}", "red")
                            stop = "captcha"
                        elif status == "unexpected" and page == 0:
                            saved = _save_unexpected(query, response.text)
                            message = STATUS_MESSAGES["unexpected"]
                            if saved:
                                message += f" Страница сохранена: {saved}"
                            _problem(info, "unexpected", message)
                            log(f"{prefix}: {message}", "red")
                        elif page == 0:
                            log(f"{prefix}: {STATUS_MESSAGES['no_ads']}", "yellow")
                        else:
                            log(f"{prefix}: объявления закончились.", "dim")
                        break

                    for ad in parsed["ads"]:
                        position += 1
                        key = (query, ad["domain"], ad["title"], ad["description"])
                        known = merged.get(key)
                        if known is None:
                            known = merged[key] = {"query": query, "page": page + 1,
                                                   "position": position, "seen": 0, **ad}
                        elif position < known["position"]:
                            known.update(page=page + 1, position=position)
                        if key not in seen_now:
                            seen_now.add(key)
                            known["seen"] += 1
                    log(f"{prefix}: объявлений — {len(parsed['ads'])}", "green")

                if seen_now:
                    info["runs_with_ads"] += 1
                    if info["status"] == "no_ads":
                        info["status"] = "ok"
                if on_progress:
                    on_progress(*_snapshot(merged, summary, queries))
                if stop:
                    break
            if stop:
                break
    except KeyboardInterrupt:
        stop = "interrupted"
        log("\nОстановлено по Ctrl+C — сохраняю собранное.", "yellow")
    return (*_snapshot(merged, summary, queries, stop), stop)


# ---------- Сохранение ----------

def _cell(value):
    """Список (быстрые ссылки) в одну ячейку таблицы."""
    return " | ".join(value) if isinstance(value, list) else value


def _xlsx_cell(value):
    # Excel не принимает управляющие символы: openpyxl на них падает при записи
    value = _cell(value)
    return ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value


def _write_json(target, ads, summary, region, runs):
    json.dump({"region": region, "runs": runs, "queries": summary, "ads": ads}, target,
              ensure_ascii=False, indent=2)
    target.write("\n")


def _write_csv(path, ads):
    # utf-8-sig — чтобы Excel открыл кириллицу без «кракозябр»
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(AD_COLUMNS.values())
        for ad in ads:
            writer.writerow(_cell(ad[key]) for key in AD_COLUMNS)


def _write_xlsx(path, ads, summary):
    workbook = Workbook()
    sheets = (
        (workbook.active, "Объявления", AD_COLUMNS, ads,
         (30, 10, 10, 12, 25, 60, 80, 60, 10, 40)),
        (workbook.create_sheet(), "Запросы", QUERY_COLUMNS, summary, (30, 12, 14, 12, 12, 90)),
    )
    for sheet, title, columns, rows, widths in sheets:
        sheet.title = title
        sheet.append(list(columns.values()))
        for row in rows:
            sheet.append([_xlsx_cell(row[key]) for key in columns])
        for cell, width in zip(sheet[1], widths):
            cell.font = Font(bold=True)
            sheet.column_dimensions[cell.column_letter].width = width
        sheet.freeze_panes = "A2"
    workbook.save(path)


def write_output(output, ads, summary, region, runs=1):
    if output == "-":
        _write_json(sys.stdout, ads, summary, region, runs)
        return
    suffix = Path(output).suffix.lower()
    if suffix == ".json":
        with open(output, "w", encoding="utf-8") as f:
            _write_json(f, ads, summary, region, runs)
    elif suffix == ".csv":
        _write_csv(output, ads)
    else:
        _write_xlsx(output, ads, summary)


def default_output():
    return str(RESULTS_DIR / time.strftime("ads_%Y-%m-%d_%H-%M-%S.xlsx"))


def _prepare_output(output):
    """Проверяет до начала сбора, что результат получится сохранить, и создаёт папку
    results. Иначе файл, открытый в Excel (Windows не даёт в него писать), обнаружится
    только в конце, и все запросы к Яндексу придётся повторять.
    Возвращает текст проблемы или None."""
    if output == "-":
        return None
    path = Path(output)
    if path.suffix.lower() not in OUTPUT_FORMATS:
        return f"Поддерживаются {', '.join(OUTPUT_FORMATS)} или '-'."
    try:
        if path.parent.resolve() == RESULTS_DIR:
            RESULTS_DIR.mkdir(exist_ok=True)
        if path.exists():
            with open(path, "r+b"):
                pass
        elif not path.parent.is_dir():
            raise FileNotFoundError(f"нет папки {path.parent}")
    except OSError as exc:
        return f"Не получается записать {output} ({exc}). Если файл открыт в Excel — закройте его."
    return None


def _duration(seconds):
    return f"{round(seconds)} с" if seconds < 90 else f"{round(seconds / 60)} мин"


def estimate(query_count, pages, runs, delay=1.0):
    """(сколько максимум обращений к Яндексу, сколько примерно секунд)."""
    count = query_count * pages * runs
    return count, count * (delay + SECONDS_PER_REQUEST)


class _Saver:
    """Сохраняет результат по ходу сбора и в конце. Если файл записать не удалось
    (открыт в Excel, нет места, что угодно ещё), кладёт в results запасной JSON:
    собранное не должно пропасть из-за сохранения."""

    def __init__(self, output, region, runs):
        self.output, self.region, self.runs = output, region, runs
        self.warned = self.backup_used = False

    def checkpoint(self, ads, summary):
        try:
            write_output(self.output, ads, summary, self.region, self.runs)
        except Exception as exc:  # сохранение по ходу не должно ронять сбор
            if not self.warned:
                log(f"Не получается сохранить {self.output} по ходу сбора ({exc}). "
                    "Попробую ещё раз в конце.", "yellow")
                self.warned = True

    def final(self, ads, summary):
        """Возвращает, куда сохранено: путь, 'stdout' или None."""
        try:
            write_output(self.output, ads, summary, self.region, self.runs)
            return "stdout" if self.output == "-" else str(Path(self.output).resolve())
        except Exception as exc:  # любая ошибка записи — не повод терять собранное
            log(f"Не удалось сохранить {self.output}: {exc}", "red")
        backup = RESULTS_DIR / time.strftime("backup_%Y-%m-%d_%H-%M-%S.json")
        try:
            RESULTS_DIR.mkdir(exist_ok=True)
            with open(backup, "w", encoding="utf-8") as f:
                _write_json(f, ads, summary, self.region, self.runs)
        except OSError as exc:
            log(f"Запасной файл тоже не записался: {exc}", "red")
            return None
        log(f"Собранное сохранено в запасной файл: {backup}", "yellow")
        self.backup_used = True
        return str(backup)


# Что делать дальше: (совет для режима с параметрами, совет человеку)
HINTS = {
    "captcha": ("Подождите и повторите; помогут --delay побольше или другой IP (--proxy).",
                "Подождите 15–30 минут и запустите снова, лучше с меньшим числом прогонов."),
    "network": ("Проверьте сеть, VPN или --proxy.",
                "Проверьте интернет или VPN и запустите снова."),
    "unexpected": ("Незнакомые страницы сохранены в results/debug — по ним надо обновить "
                   "parse_page().",
                   "Незнакомые страницы сохранены в папке results/debug — отправьте их "
                   "разработчику."),
    "no_ads": ("Объявления сильно зависят от региона — попробуйте другой (-r).",
               "Объявления сильно зависят от города — попробуйте другой."),
}


def _report(query_count, ads, summary, stop, saved_to, human, backup=False):
    """Итог зависит от исхода: зелёный — всё собрано, жёлтый — частично или пусто,
    красный — не получилось. Возвращает код выхода."""
    def hint(kind):
        log(HINTS[kind][1 if human else 0], "yellow")

    where = f" → {saved_to}" if saved_to else " (сохранить не удалось)"
    found = f"объявлений — {len(ads)}"
    problems = [info for info in summary if info["status"] in ERROR_STATUSES]

    if stop == "interrupted":
        log(f"Остановлено по Ctrl+C: {found}{where}", "bold", "yellow")
        return EXIT_INTERRUPTED
    if stop:
        log(f"Собрано не всё: {STOP_MESSAGES[stop]}. {found.capitalize()}{where}", "bold", "red")
        hint(stop)
        return EXIT_ERRORS
    if problems:
        head, style = ("Собрано частично", "yellow") if ads else ("Не получилось", "red")
        log(f"{head}: {found}, запросов с проблемами — {len(problems)} из {query_count}{where}",
            "bold", style)
        if any(info["status"] == "unexpected" for info in problems):
            hint("unexpected")
        return EXIT_ERRORS
    if not ads:
        log(f"Объявлений нет: Яндекс по всем запросам ответил «ничего не нашлось»{where}",
            "bold", "yellow")
        hint("no_ads")
        return EXIT_NO_ADS
    if not saved_to or backup:
        place = "сохранено в запасной файл" if saved_to else "не сохранено"
        log(f"Собрано, но {place}: {found}{where}", "bold", "yellow")
        return EXIT_ERRORS
    log(f"Готово: {found}, запросов — {query_count}{where}", "bold", "green")
    return EXIT_OK


def run(queries, region=225, pages=1, runs=1, output=None, delay=1.0, proxy=None,
        save_html=None, human=False):
    """Собирает объявления, сохраняет результат (по ходу и в конце), пишет итог.
    human — человеку советы без флагов командной строки. Возвращает код выхода."""
    output = output or default_output()
    count, seconds = estimate(len(queries), pages, runs, delay)
    log(f"Обращений к Яндексу: до {count}, это около {_duration(seconds)}.", "dim")
    saver = _Saver(output, region, runs)
    ads, summary, stop = collect(make_session(proxy), queries, pages=pages, region=region,
                                 delay=delay, runs=runs, save_html=save_html,
                                 on_progress=None if output == "-" else saver.checkpoint)
    saved_to = saver.final(ads, summary)
    return _report(len(queries), ads, summary, stop, saved_to, human, saver.backup_used)


# ---------- Ручной режим (python main.py без аргументов) ----------

# Значения по умолчанию в карточке перед стартом
INTERACTIVE_RUNS = 1
INTERACTIVE_PAGES = 1
MAX_PAGES = 10  # дальше 2–3 страниц реклама почти никогда не идёт
DEFAULT_REGION = 225  # Enter в вопросе о городе — вся Россия

# Ввод, похожий на путь к файлу, а не на поисковый запрос
_FILE_SUFFIXES = (".txt", ".csv", ".tsv", ".xlsx", ".xlsm", ".xls", ".docx")


def _clean_path(value):
    """Путь, перетащенный в окно терминала: Windows берёт его в кавычки,
    macOS экранирует пробелы обратной косой чертой."""
    value = value.strip().strip("\"'")
    if sys.platform != "win32":
        value = re.sub(r"\\(.)", r"\1", value)
    return value


def _looks_like_path(text):
    text = text.strip().strip("\"'")
    return (text.startswith(("/", "~", "./", "../", ".\\", "..\\"))
            or re.match(r"[A-Za-z]:[\\/]", text) is not None
            or text.lower().endswith(_FILE_SUFFIXES))


def _find_file(text):
    """Файл из ввода; если по этому пути его нет — ищем такое же имя рядом с программой."""
    path = Path(_clean_path(text)).expanduser()
    if path.exists():
        return path
    nearby = RESULTS_DIR.parent / path.name
    return nearby if nearby.exists() else None


def _first_number(value):
    match = re.search(r"\d+", value)
    return int(match.group()) if match else None


def _ask_until(ask, prompt, convert):
    """Переспрашивает, пока convert не примет ответ (ValueError — показать и спросить снова)."""
    while True:
        try:
            return convert(clean_text(ask(prompt)))
        except ValueError as exc:
            say(f"  {exc}", "yellow")


def _confirm(ask, question):
    """Да/нет цифрой: «н» в русской раскладке стоит на клавише «y», буквы путают."""
    while True:
        answer = clean_text(ask(f"{question} Enter — да, 0 — нет: "))
        if answer in ("", "0"):
            return answer == ""
        say("  Нажмите Enter, чтобы согласиться, или 0, чтобы отказаться.", "yellow")


def _load_queries_file(path):
    """Запросы из файла или None, если файл не подошёл (причина уже показана)."""
    if path.is_dir():
        say(f"  «{path.name}» — это папка, а нужен файл со списком запросов.", "yellow")
        return None
    try:
        queries = read_queries([], str(path))
    except (OSError, UnicodeError, ValueError) as exc:
        say(f"  Не получилось прочитать файл: {exc}", "yellow")
        return None
    if not queries:
        say(f"  В файле «{path.name}» нет запросов.", "yellow")
        return None
    say(f"  Из файла «{path.name}» взято запросов: {len(queries)}.", "green")
    return queries


def _ask_queries(ask):
    say("Введите поисковые запросы по одному на строку или перетащите в это окно файл "
        "со списком (.txt). Когда закончите, нажмите Enter на пустой строке.", "dim")
    queries = []
    while True:
        # Вставка из буфера через readline может прийти сразу несколькими строками
        lines = [clean_text(line) for line in ask(f"{len(queries) + 1}> ").splitlines()]
        lines = [line for line in lines if line]
        if not lines:
            if not queries:
                say("  Нужен хотя бы один запрос.", "yellow")
                continue
            say(f"  Запросов: {len(queries)} — {'; '.join(queries[:5])}"
                f"{'; …' if len(queries) > 5 else ''}", "green")
            more = ask("  Enter — дальше, или допишите ещё: ")
            lines = [line for line in (clean_text(x) for x in more.splitlines()) if line]
            if not lines:
                break
        for line in lines:
            if _looks_like_path(line):
                path = _find_file(line)
                if path is None:
                    say(f"  Файл не найден: {_clean_path(line)}", "yellow")
                    continue
                queries += _load_queries_file(path) or []
            else:
                queries.append(line)
        queries = list(dict.fromkeys(queries))

    if any("," in query for query in queries) and _confirm(
            ask, "  Похоже, в некоторых строках несколько запросов через запятую. Разделить их?"):
        queries = list(dict.fromkeys(
            part.strip() for query in queries for part in query.split(",") if part.strip()))
        say(f"  Теперь запросов: {len(queries)}.", "green")
    return queries


def _region_label(code):
    name = _REGION_NAMES.get(code)
    return f"{name} ({code})" if name else f"код {code}"


def _similar_regions(value):
    """Похожие города: «Ставропольский край» → Ставрополь, «Краснадар» → Краснодар."""
    norm = _norm(value)
    first = norm.split()[0] if norm else ""
    by_norm = {_norm(name): name for name in REGIONS}
    found = [name for key, name in by_norm.items()
             if len(first) >= 4 and (first.startswith(key) or key.startswith(first))]
    found += [by_norm[key] for key in difflib.get_close_matches(norm, list(by_norm), n=3,
                                                                  cutoff=0.75)]
    return list(dict.fromkeys(found))[:3]


def _show_regions():
    say(" · ".join(f"{name} {code}" for name, code in sorted(REGIONS.items())), "dim")
    say(f"Остальные регионы — по коду из справочника Яндекса: {REGIONS_DOC}", "dim")


def _ask_region(ask, current=None):
    say("Реклама сильно зависит от города. Введите город или код региона Яндекса; "
        "? — список городов.", "dim")
    default = current or DEFAULT_REGION
    while True:
        value = clean_text(ask(f"Город [{_region_label(default)}]: "))
        if value == "?":
            _show_regions()
            continue
        if not value:
            return default
        number = _first_number(value)
        if number:  # «36», «36 Ставрополь»
            if number in _REGION_NAMES or _confirm(
                    ask, f"  Кода {number} нет среди известных городов. Всё равно взять его?"):
                return number
            continue
        try:
            return parse_region(value)
        except ValueError:
            pass
        similar = _similar_regions(value)
        if len(similar) == 1 and _confirm(
                ask, f"  Такого в списке нет. Может быть, {_region_label(REGIONS[similar[0]])}?"):
            return REGIONS[similar[0]]
        if len(similar) > 1:
            say("  Такого в списке нет. Похожие: "
                + ", ".join(_region_label(REGIONS[name]) for name in similar) + ".", "yellow")
        elif not similar:
            say("  Такого города в списке нет. ? — список городов; остальные регионы — "
                f"по коду из справочника: {REGIONS_DOC}", "yellow")


def _parse_count(value, default, maximum):
    """Число из ответа: «3», «3 прогона», «3–5» (берётся первое)."""
    if not value:
        return default
    number = _first_number(value)
    if number is None:
        raise ValueError("Введите число, например 3.")
    if not 1 <= number <= maximum:
        raise ValueError(f"Нужно число от 1 до {maximum}.")
    return number


def _ask_completeness(ask, runs, pages):
    say("Прогоны — повторные сборы: Яндекс каждый раз показывает немного разный набор "
        "объявлений, и несколько прогонов дают более полный список. Обычно хватает трёх.", "dim")
    runs = _ask_until(ask, f"Прогонов, от 1 до {MAX_RUNS} [{runs}]: ",
                      lambda v: _parse_count(v, runs, MAX_RUNS))
    say("Страницы — сколько страниц объявлений брать по каждому запросу, на странице до 10. "
        "Обычно хватает 1–2: когда реклама кончится, программа остановится сама.", "dim")
    pages = _ask_until(ask, f"Страниц, от 1 до {MAX_PAGES} [{pages}]: ",
                       lambda v: _parse_count(v, pages, MAX_PAGES))
    return runs, pages


def _free_name(path):
    """«отчёт.xlsx» занят — «отчёт (2).xlsx»: прошлый результат не затираем."""
    number = 2
    candidate = path
    while candidate.exists():
        candidate = path.with_name(f"{path.stem} ({number}){path.suffix}")
        number += 1
    return candidate


def _parse_output(value, default):
    """Пусто — файл по умолчанию; имя — в папку results; папка — внутрь неё; путь — как есть."""
    value = _clean_path(value)
    if not value:
        path = Path(default)
    elif Path(value).expanduser().is_dir():
        path = Path(value).expanduser() / Path(default).name
    elif any(sep and sep in value for sep in (os.sep, os.altsep)):
        path = Path(value).expanduser()
    else:
        path = RESULTS_DIR / value
    if not path.suffix:
        path = path.with_name(path.name + ".xlsx")
    if path.suffix.lower() not in OUTPUT_FORMATS:
        raise ValueError("Поддерживаются .xlsx (Excel), .csv и .json. "
                         "Не знаете, что выбрать, — оставьте Excel.")
    if path.exists():
        free = _free_name(path)
        say(f"  «{path.name}» уже есть — сохраню как «{free.name}», прошлый файл не трону.",
            "yellow")
        path = free
    problem = _prepare_output(str(path))
    if problem:
        raise ValueError(problem)
    return str(path)


def _shown(path):
    """Путь покороче: внутри папки программы — относительный."""
    try:
        return str(Path(path).resolve().relative_to(RESULTS_DIR.parent))
    except ValueError:
        return str(path)


def _ask_output(ask, current):
    say("Просто имя файла — сохраню в папку results. Полный путь или папка — сохраню туда.",
        "dim")
    return _ask_until(ask, f"Файл [{_shown(current)}]: ", lambda v: _parse_output(v, current))


def _open_file(path):
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa — есть только на Windows
        else:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.run([opener, path], check=True, stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError):
        say(f"  Открыть не получилось. Файл лежит здесь: {path}", "yellow")


def _card(queries, region, runs, pages, output):
    count, seconds = estimate(len(queries), pages, runs)
    preview = ", ".join(queries[:3]) + (", …" if len(queries) > 3 else "")
    say("\nПроверьте перед стартом:", "bold")
    say(f"  1. Запросы: {len(queries)} ({preview})")
    say(f"  2. Город: {_region_label(region)}")
    say(f"  3. Полнота: прогонов — {runs}, страниц на запрос — {pages}")
    say(f"  4. Файл: {_shown(output)}")
    say(f"  Сбор займёт около {_duration(seconds)}, обращений к Яндексу — до {count}.", "dim")


def interactive(ask=input):
    """Ручной режим: два вопроса (запросы и город), потом карточка — проверить и при
    необходимости поменять любой пункт, не начиная заново."""
    try:
        import readline  # noqa: F401 — стрелки и правка строки при вводе (macOS, Linux)
    except ImportError:
        pass  # на Windows строку правит сама консоль
    say("Сбор рекламы Яндекс.Директа", "bold", "cyan")
    say("Программа соберёт объявления конкурентов по вашим запросам и сохранит таблицу Excel "
        "в папку results. Выход — Ctrl+C.\n", "dim")
    try:
        say("Шаг 1 из 2. Запросы", "bold")
        queries = _ask_queries(ask)
        say("\nШаг 2 из 2. Город", "bold")
        region = _ask_region(ask)
        runs, pages = INTERACTIVE_RUNS, INTERACTIVE_PAGES
        output = _parse_output("", default_output())
        while True:
            _card(queries, region, runs, pages, output)
            choice = clean_text(ask("Enter — начать, 1–4 — изменить пункт, 0 — выйти: "))
            if not choice:
                break
            if choice == "0":
                say("Отменено.", "yellow")
                return EXIT_USAGE
            if choice == "1":
                queries = _ask_queries(ask)
            elif choice == "2":
                region = _ask_region(ask, region)
            elif choice == "3":
                runs, pages = _ask_completeness(ask, runs, pages)
            elif choice == "4":
                output = _ask_output(ask, output)
            else:
                say("  Нажмите Enter, чтобы начать, или цифру от 1 до 4.", "yellow")
    except EOFError:
        say("\nОтменено.", "yellow")
        return EXIT_USAGE
    say()
    code = run(queries, region=region, pages=pages, runs=runs, output=output, human=True)
    if Path(output).exists():
        try:
            if _confirm(ask, "\nОткрыть таблицу?"):
                _open_file(output)
        except EOFError:
            pass
    return code


# ---------- Режим с аргументами (для скриптов и агентов) ----------

def _region_arg(value):
    try:
        return parse_region(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from None


def build_parser():
    parser = argparse.ArgumentParser(
        description="Собирает объявления Яндекс.Директа по поисковым запросам "
                    "(страница yandex.ru/search/ads). Без аргументов в терминале — "
                    "ручной режим с вопросами.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""примеры:
  python main.py                              ручной режим
  python main.py "купить духи" "парфюм магазин"
  python main.py -f queries.txt -p 2 -r Москва --runs 3 -o ads.xlsx
  python main.py -f queries.txt -r 36 -o -    JSON в stdout (удобно агенту)
  cat queries.txt | python main.py -f - -o ads.csv

файл запросов: один запрос на строку, пустые строки и строки с # пропускаются;
кодировка UTF-8, UTF-16 или cp1251 (ANSI) — определяется сама.

регион (-r) — код Яндекса (уходит в параметр lr) или название крупного города:
225 — Россия, 213 — Москва, 1 — Москва и область, 2 — Санкт-Петербург,
54 — Екатеринбург, 36 — Ставрополь. Объявления сильно зависят от региона.
Все коды: {REGIONS_DOC}

прогоны (--runs): Яндекс каждый раз показывает немного разный набор объявлений.
Прогоны объединяются: одинаковые объявления склеиваются, в колонке «Встречалось
в прогонах» видно, как часто объявление показывалось. Обычно хватает 3–5.

статусы запросов: ok — объявления есть; no_ads — Яндекс ответил «ничего не нашлось»;
captcha — капча; unexpected — незнакомая страница (вёрстка поменялась?);
error — ошибка сети.

коды выхода:
  0 — собрано хотя бы одно объявление
  1 — капча, ошибка сети или незнакомая вёрстка хотя бы по одному запросу
      (собранное всё равно сохраняется)
  2 — неверные аргументы или результат некуда сохранить
  3 — ошибок нет, но объявлений ноль

цвет в терминале выключается переменной окружения NO_COLOR=1.""",
    )
    parser.add_argument("queries", nargs="*", metavar="QUERY", help="поисковые запросы")
    parser.add_argument("-f", "--file", help="файл с запросами, по одному на строку ('-' — stdin)")
    parser.add_argument("-p", "--pages", type=int, default=1,
                        help="сколько страниц объявлений брать на запрос (по умолчанию 1)")
    parser.add_argument("-r", "--region", type=_region_arg, default=225,
                        help="код региона lr или название города (по умолчанию 225 — Россия)")
    parser.add_argument("--runs", type=int, default=1,
                        help=f"сколько прогонов сделать и объединить, от 1 до {MAX_RUNS} "
                             "(по умолчанию 1)")
    parser.add_argument("-o", "--output",
                        help="куда сохранить: .xlsx, .csv или .json; '-' — JSON в stdout "
                             "(по умолчанию results/ads_<дата>_<время>.xlsx рядом с программой)")
    parser.add_argument("--delay", type=float, default=1.0,
                        help="пауза между запросами к Яндексу, сек (по умолчанию 1)")
    parser.add_argument("--proxy", help="прокси, например http://user:pass@host:port "
                                        "(также работают переменные HTTPS_PROXY/HTTP_PROXY)")
    parser.add_argument("--save-html", metavar="DIR",
                        help="сохранять полученные страницы в папку — для отладки селекторов")
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        queries = read_queries(args.queries, args.file)
    except (OSError, UnicodeError, ValueError) as exc:
        parser.error(f"не удалось прочитать файл запросов: {exc}")
    if not queries:
        parser.error("не задано ни одного запроса: передайте их аргументами или через -f FILE "
                     "(без аргументов в терминале запускается ручной режим)")
    if args.pages < 1:
        parser.error("--pages должно быть не меньше 1")
    if not 1 <= args.runs <= MAX_RUNS:
        parser.error(f"--runs должно быть от 1 до {MAX_RUNS}")
    output = args.output or default_output()
    problem = _prepare_output(output)
    if problem:
        parser.error(f"--output: {problem}")

    return run(queries, region=args.region, pages=args.pages, runs=args.runs, output=output,
               delay=args.delay, proxy=args.proxy, save_html=args.save_html)


if __name__ == "__main__":
    _utf8_stdio()
    _enable_ansi()
    if _launched_interactively():
        try:
            code = interactive()
        except KeyboardInterrupt:
            say("\nОстановлено.", "yellow")
            code = 130
        _pause()
        sys.exit(code)
    sys.exit(main())
