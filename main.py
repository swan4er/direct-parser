#!/usr/bin/env python3
"""Сбор рекламных объявлений Яндекс.Директа со страницы yandex.ru/search/ads.

Без аргументов в терминале — пошаговый режим: программа сама спрашивает запросы,
регион, число страниц и прогонов, имя файла. С аргументами — режим для скриптов
и агентов (python main.py --help). Ход работы пишется в stderr, данные — в файл
(по умолчанию в папку results) или в stdout.
"""

import argparse
import codecs
import csv
import json
import os
import re
import sys
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


try:
    import requests
    from bs4 import BeautifulSoup
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
except ImportError as exc:
    _utf8_stdio()
    print(f"Не установлена библиотека {exc.name}. Выполните: pip install -r requirements.txt",
          file=sys.stderr)
    _pause()
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
EXIT_ERRORS = 1  # капча, сетевая ошибка или незнакомая вёрстка хотя бы по одному запросу
EXIT_USAGE = 2  # неверные аргументы (этот же код возвращает argparse)
EXIT_NO_ADS = 3  # ошибок нет, но Яндекс не показал ни одного объявления

ERROR_STATUSES = {"captcha", "error", "unexpected"}
STATUS_MESSAGES = {
    "no_ads": "Яндекс ответил «ничего не нашлось».",
    "captcha": "Яндекс показал капчу: увеличьте --delay или смените IP (--proxy).",
    "unexpected": "Ни объявлений, ни пометки «ничего не нашлось» — вероятно, Яндекс поменял "
                  "вёрстку. Сохраните страницу через --save-html и проверьте селекторы.",
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


def say(message="", *styles):
    print(paint(message, *styles, stream=sys.stdout))


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


def read_queries(arg_queries, file_path=None):
    """Собирает запросы из аргументов и файла ("-" — stdin), убирает пустые строки,
    комментарии (#) и дубли, сохраняя порядок."""
    lines = list(arg_queries)
    if file_path == "-":
        lines += _decode(sys.stdin.buffer.read()).splitlines()
    elif file_path:
        lines += _decode(Path(file_path).read_bytes()).splitlines()
    cleaned = (line.strip() for line in lines)
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
    retry = Retry(total=3, backoff_factor=1, status_forcelist=(429, 500, 502, 503, 504))
    session.mount("https://", HTTPAdapter(max_retries=retry))
    if proxy:
        session.proxies.update({"http": proxy, "https": proxy})
    return session


def _problem(info, status, message):
    """Ошибка важнее «ok» и «no_ads»; из нескольких ошибок запоминается первая."""
    if info["status"] not in ERROR_STATUSES:
        info.update(status=status, message=message)


def collect(session, queries, pages=1, region=225, delay=1.0, runs=1, save_html=None):
    """Обходит запросы и страницы runs раз и объединяет прогоны: Яндекс каждый раз
    показывает немного разный набор объявлений. Одинаковые объявления (тот же домен,
    заголовок и текст) склеиваются: считается, в скольких прогонах они встретились,
    и остаётся лучшая позиция. Возвращает (объявления, сводка по запросам)."""
    merged = {}
    summary = {query: {"query": query, "status": "no_ads", "runs_with_ads": 0, "pages": 0,
                       "ads": 0, "message": ""} for query in queries}
    first_request = True
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
                        BASE_URL, params={"text": query, "p": page, "lr": region}, timeout=30
                    )
                    response.raise_for_status()
                except requests.RequestException as exc:
                    _problem(info, "error", str(exc))
                    log(f"{prefix}: ошибка сети — {exc}", "red")
                    break

                if save_html:
                    safe_name = re.sub(r"[^\w-]+", "_", query)[:60]
                    run_suffix = f"_r{run_number}" if runs > 1 else ""
                    path = Path(save_html) / f"{number:03d}_{safe_name}_p{page + 1}{run_suffix}.html"
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_text(response.text, encoding="utf-8")

                parsed = parse_page(response.text, response.url)
                info["pages"] = max(info["pages"], page + 1)

                if not parsed["ads"]:
                    # На первой странице пустой ответ — результат запроса; дальше — конец
                    # списка объявлений. Капча — ошибка на любой странице.
                    status = parsed["status"]
                    if status == "captcha" or (page == 0 and status == "unexpected"):
                        _problem(info, status, STATUS_MESSAGES[status])
                        log(f"{prefix}: {STATUS_MESSAGES[status]}", "red")
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

    order = {query: index for index, query in enumerate(queries)}
    ads = sorted(merged.values(), key=lambda ad: (order[ad["query"]], ad["position"]))
    for info in summary.values():
        info["ads"] = sum(1 for ad in ads if ad["query"] == info["query"])
        if info["status"] == "no_ads":
            info["message"] = STATUS_MESSAGES["no_ads"]
    return ads, list(summary.values())


# ---------- Сохранение ----------

def _cell(value):
    """Список (быстрые ссылки) в одну ячейку таблицы."""
    return " | ".join(value) if isinstance(value, list) else value


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
            sheet.append([_cell(row[key]) for key in columns])
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


def run(queries, region=225, pages=1, runs=1, output=None, delay=1.0, proxy=None,
        save_html=None):
    """Собирает объявления, сохраняет результат, пишет итог. Возвращает код выхода."""
    output = output or default_output()
    requests_count = len(queries) * pages * runs
    log(f"Запросов к Яндексу: до {requests_count}, это около "
        f"{_duration(requests_count * (delay + SECONDS_PER_REQUEST))}.", "dim")
    ads, summary = collect(make_session(proxy), queries, pages=pages, region=region,
                           delay=delay, runs=runs, save_html=save_html)
    try:
        write_output(output, ads, summary, region, runs)
    except OSError as exc:
        log(f"Не удалось сохранить результат: {exc}", "red")
        return EXIT_ERRORS

    where = "stdout" if output == "-" else Path(output).resolve()
    log(f"Готово: запросов {len(queries)}, объявлений {len(ads)} → {where}", "bold", "green")

    if any(info["status"] in ERROR_STATUSES for info in summary):
        return EXIT_ERRORS
    if not ads:
        log("Объявлений нет. Они сильно зависят от региона — попробуйте другой регион.", "yellow")
        return EXIT_NO_ADS
    return EXIT_OK


# ---------- Пошаговый режим (python main.py без аргументов) ----------

def _clean_path(value):
    """Путь, перетащенный в окно терминала: Windows берёт его в кавычки,
    macOS экранирует пробелы обратной косой чертой."""
    value = value.strip().strip("\"'")
    if sys.platform != "win32":
        value = re.sub(r"\\(.)", r"\1", value)
    return value


def _prompt(label, default):
    return f"{paint(label, 'bold', stream=sys.stdout)} {paint(f'[{default}]', 'dim', stream=sys.stdout)}: "


def _ask_until(ask, prompt, convert):
    """Переспрашивает, пока convert не примет ответ (ValueError — показать и спросить снова)."""
    while True:
        try:
            return convert(ask(prompt).strip())
        except ValueError as exc:
            say(f"  {exc}", "yellow")


def _parse_number(value, default, maximum=None):
    if not value:
        return default
    if not value.isdigit() or int(value) < 1 or (maximum and int(value) > maximum):
        raise ValueError(f"Нужно целое число от 1{f' до {maximum}' if maximum else ''}.")
    return int(value)


def _parse_output(value, default):
    """Пустой ответ — файл по умолчанию; просто имя — в папку results; путь — как указан."""
    value = _clean_path(value)
    if not value:
        path = Path(default)
    elif any(sep and sep in value for sep in (os.sep, os.altsep)):
        path = Path(value)
    else:
        path = RESULTS_DIR / value
    if not path.suffix:
        path = path.with_name(path.name + ".xlsx")
    problem = _prepare_output(str(path))
    if problem:
        raise ValueError(problem)
    return str(path)


def _ask_queries(ask):
    say("Запросы — по одному на строку, в конце пустая строка.", "bold")
    say("Если запросов много, укажите путь к текстовому файлу с ними "
        "(можно перетащить файл в это окно).", "dim")
    while True:
        first = ask("").strip()
        if not first:
            say("  Нужен хотя бы один запрос.", "yellow")
            continue
        path = Path(_clean_path(first))
        if path.is_file():
            try:
                queries = read_queries([], str(path))
            except OSError as exc:
                say(f"  Не удалось прочитать файл: {exc}", "yellow")
                continue
            if not queries:
                say("  В файле нет запросов.", "yellow")
                continue
            say(f"  Из файла взято запросов: {len(queries)}.", "green")
            return queries
        lines = [first]
        while line := ask("").strip():
            lines.append(line)
        queries = read_queries(lines)
        say(f"  Запросов: {len(queries)}.", "green")
        return queries


def interactive(ask=input):
    """Пошаговый режим для человека: вопросы вместо флагов."""
    say("Сбор объявлений Яндекс.Директа", "bold", "cyan")
    say("В квадратных скобках — значение по умолчанию, для него просто нажмите Enter. "
        "Выход — Ctrl+C.\n", "dim")
    try:
        queries = _ask_queries(ask)

        say()
        say("Регион — код или название города. Популярные: 225 Россия, 213 Москва, "
            "1 Москва и область, 2 Санкт-Петербург, 54 Екатеринбург, 43 Казань, "
            "35 Краснодар, 36 Ставрополь.", "dim")
        say(f"Все коды регионов: {REGIONS_DOC}", "dim")
        region = _ask_until(ask, _prompt("Регион", "225 — Россия"),
                            lambda v: parse_region(v or "225"))
        name = _REGION_NAMES.get(region)
        if name:
            say(f"  Регион: {name} ({region}).", "green")
        else:
            say(f"  Регион: код {region} — его нет среди популярных, сверьтесь со справочником.",
                "yellow")

        say()
        pages = _ask_until(ask, _prompt("Сколько страниц объявлений на запрос (на странице до 10)", 1),
                           lambda v: _parse_number(v, 1))

        say()
        say("Прогоны — повторные сборы: Яндекс каждый раз показывает немного разный набор "
            f"объявлений, и несколько прогонов дают более полный список. Обычно хватает 3–5, "
            f"максимум {MAX_RUNS}.", "dim")
        runs = _ask_until(ask, _prompt("Сколько прогонов", 1),
                          lambda v: _parse_number(v, 1, MAX_RUNS))

        say()
        default = default_output()
        shown = Path(default).relative_to(RESULTS_DIR.parent)
        say("Просто имя файла — сохраню в папку results, полный путь — куда укажете.", "dim")
        output = _ask_until(ask, _prompt("Куда сохранить (.xlsx, .csv или .json)", shown),
                            lambda v: _parse_output(v, default))
    except EOFError:
        say("\nОтменено.", "yellow")
        return EXIT_USAGE
    say()
    return run(queries, region=region, pages=pages, runs=runs, output=output)


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
                    "пошаговый режим с вопросами.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""примеры:
  python main.py                              пошаговый режим
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
    except (OSError, UnicodeError) as exc:
        parser.error(f"не удалось прочитать файл запросов: {exc}")
    if not queries:
        parser.error("не задано ни одного запроса: передайте их аргументами или через -f FILE "
                     "(без аргументов в терминале запускается пошаговый режим)")
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
