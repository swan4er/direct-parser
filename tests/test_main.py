"""Тесты без сети: образцы в fixtures/ — ужатые настоящие страницы yandex.ru/search/ads
(14.10.2026): ads_page.html — первые два объявления по «купить духи»,
no_ads_page.html — ответ «ничего не нашлось»."""

import csv
import io
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
import requests
from openpyxl import load_workbook

import main

ROOT = Path(__file__).parent.parent
FIXTURES = Path(__file__).parent / "fixtures"
ADS_PAGE = (FIXTURES / "ads_page.html").read_text(encoding="utf-8")
NO_ADS_PAGE = (FIXTURES / "no_ads_page.html").read_text(encoding="utf-8")
CAPTCHA_PAGE = '<form action="/checkcaptcha?key=x"></form>'


def ads_html(*domains):
    """Страница с объявлениями заданных доменов — в той же разметке, что ads_page.html."""
    items = "".join(
        f'<li class="serp-item serp-item_card" data-cid="{i}">'
        f'<a class="Link OrganicTitle-Link" href="https://yabs.yandex.ru/count/{domain}">'
        f'<h2 class="OrganicTitle-LinkText">Объявление {domain}</h2></a>'
        f'<div class="Path Organic-Path"><a class="Link Path-Item" href="#"><b>{domain}</b>'
        f'<span class="Path-Separator">›</span>страница</a></div>'
        f'<span class="OrganicTextContentSpan">Текст {domain}</span></li>'
        for i, domain in enumerate(domains))
    return f'<ul id="search-result">{items}</ul>'


class FakeResponse:
    def __init__(self, text, url="https://yandex.ru/search/ads"):
        self.text, self.url = text, url

    def raise_for_status(self):
        pass


class FakeSession:
    """Отдаёт страницы по номеру p; всё, что за пределами pages, — «ничего не нашлось»."""

    def __init__(self, pages, url="https://yandex.ru/search/ads"):
        self.pages, self.url, self.calls = pages, url, []

    def get(self, url, params, timeout):
        self.calls.append(params)
        page = params["p"]
        return FakeResponse(self.pages[page] if page < len(self.pages) else NO_ADS_PAGE, self.url)


class QueueSession(FakeSession):
    """Отдаёт страницы по очереди вызовов (удобно для прогонов); потом — «ничего не нашлось».
    Исключение в очереди не отдаётся, а выбрасывается: так изображаются обрыв связи и Ctrl+C."""

    def __init__(self, *texts):
        super().__init__([])
        self.texts = list(texts)

    def get(self, url, params, timeout):
        self.calls.append(params)
        item = self.texts.pop(0) if self.texts else NO_ADS_PAGE
        if isinstance(item, BaseException):
            raise item
        return FakeResponse(item, self.url)


class FakeTTY(io.StringIO):
    def isatty(self):
        return True


@pytest.fixture(autouse=True)
def isolated_results(monkeypatch, tmp_path):
    """Результаты «по умолчанию» — во временную папку, а не в results/ проекта."""
    monkeypatch.setattr(main, "RESULTS_DIR", tmp_path.resolve() / "results")


@pytest.fixture
def no_sleep(monkeypatch):
    monkeypatch.setattr(main.time, "sleep", lambda seconds: None)


def use_session(monkeypatch, session):
    monkeypatch.setattr(main, "make_session", lambda proxy=None: session)
    return session


def run_cli(monkeypatch, session, *argv):
    use_session(monkeypatch, session)
    return main.main([*argv, "--delay", "0"])


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def answers(*values):
    """Подменяет input(): печатает приглашение, как настоящий input(), и отдаёт ответы
    по очереди, а когда они кончились — EOF."""
    queue = iter(values)

    def ask(prompt=""):
        print(prompt, end="")
        try:
            return next(queue)
        except StopIteration:
            raise EOFError from None
    return ask


# ---------- разбор страницы ----------

def test_parse_real_ads_page():
    result = main.parse_page(ADS_PAGE)
    assert result["status"] == "ok"
    assert [ad["domain"] for ad in result["ads"]] == ["randewoo.ru", "impexoils.ru"]
    first = result["ads"][0]
    assert first["title"] == "Парфюм в интернет-магазине Randewoo — скидки до 55%"
    assert first["description"].startswith("Раскройте свой потенциал")
    # неразрывные пробелы Яндекса заменены обычными
    assert first["sitelinks"] == ["Aroma Box", "Бестселлеры", "Конструктор Aroma Box", "SALE"]
    assert first["label"] == "Промо"
    assert first["click_url"].startswith("https://yabs.yandex.ru/count/")


def test_parse_real_no_ads_page():
    assert main.parse_page(NO_ADS_PAGE) == {"status": "no_ads", "ads": []}


def test_parse_captcha():
    assert main.parse_page("<html></html>", "https://yandex.ru/showcaptcha?x")["status"] == "captcha"
    assert main.parse_page(CAPTCHA_PAGE)["status"] == "captcha"


def test_parse_unknown_markup():
    assert main.parse_page('<ul id="search-result"></ul>')["status"] == "unexpected"


def test_parse_ad_with_missing_parts_does_not_crash():
    html = ('<ul id="search-result"><li class="serp-item" data-cid="0">'
            '<a class="Link" href="https://yabs.yandex.ru/count/x">'
            '<h2 class="OrganicTitle-LinkText">Только заголовок</h2></a></li></ul>')
    ad = main.parse_page(html)["ads"][0]
    assert ad == {"domain": "", "title": "Только заголовок", "description": "",
                  "sitelinks": [], "label": "", "click_url": "https://yabs.yandex.ru/count/x"}


def test_synthetic_page_parses_like_real_one():
    assert [ad["domain"] for ad in main.parse_page(ads_html("a.ru", "b.ru"))["ads"]] == [
        "a.ru", "b.ru"]


# ---------- регионы и запросы ----------

@pytest.mark.parametrize("value, code", [
    ("36", 36), ("Ставрополь", 36), ("  ростов на дону ", 39), ("Ростов-на-Дону", 39),
    ("Орел", 10), ("спб", 2), ("Москва", 213), ("москва и область", 1), ("99999", 99999),
])
def test_parse_region(value, code):
    assert main.parse_region(value) == code


@pytest.mark.parametrize("value", ["Атлантида", "0", ""])
def test_parse_region_rejects_unknown(value):
    with pytest.raises(ValueError, match="aistudio.yandex.ru"):
        main.parse_region(value)


def test_read_queries_from_args_and_file(tmp_path):
    path = tmp_path / "queries.txt"
    # BOM — так сохраняет Блокнот Windows
    path.write_text("﻿купить духи\n\n# комментарий\n  парфюм магазин  \nкупить духи\n",
                    encoding="utf-8")
    assert main.read_queries(["купить парфюм"], str(path)) == [
        "купить парфюм", "купить духи", "парфюм магазин"]


@pytest.mark.parametrize("encoding", ["cp1251", "utf-16"])
def test_read_queries_windows_encodings(tmp_path, encoding):
    # cp1251 — «Текст» из Excel в русской Windows, utf-16 — «Текст Юникод»
    path = tmp_path / "queries.txt"
    path.write_bytes("купить духи\r\nпарфюм магазин\r\n".encode(encoding))
    assert main.read_queries([], str(path)) == ["купить духи", "парфюм магазин"]


def make_excel(path, first_column, second_column=()):
    from openpyxl import Workbook
    workbook = Workbook()
    for row, value in enumerate(first_column, 1):
        workbook.active.cell(row=row, column=1, value=value)
    for row, value in enumerate(second_column, 1):
        workbook.active.cell(row=row, column=2, value=value)
    workbook.create_sheet("Другой лист").append(["не запрос"])
    workbook.save(path)
    return path


def test_read_queries_from_excel_first_column(tmp_path):
    path = make_excel(tmp_path / "запросы.xlsx",
                      ["Ключевые слова", "купить духи", None, "парфюм магазин", 2026],
                      ["Частота", "100", "", "50"])
    # заголовок пропущен, пустая клетка тоже, второй лист и вторая колонка не читаются
    assert main.read_queries([], str(path)) == ["купить духи", "парфюм магазин", "2026"]


def test_read_queries_from_csv_with_semicolons(tmp_path):
    path = tmp_path / "запросы.csv"
    # так сохраняет CSV Excel в русской Windows: «;» и cp1251
    path.write_bytes("Запрос;Частота\r\nкупить духи, недорого;100\r\nпарфюм магазин;50\r\n"
                     .encode("cp1251"))
    assert main.read_queries([], str(path)) == ["купить духи, недорого", "парфюм магазин"]


@pytest.mark.parametrize("data, message", [
    (b"\xd0\xcf\x11\xe0" + b"\x00" * 100, "старый формат Excel"),
    (b"PK\x03\x04" + b"\x00" * 100, "не таблица Excel и не текст"),
    (b"\x89PNG\r\n\x1a\n\x00\x00", "не текстовый файл"),
])
def test_read_queries_rejects_other_files(tmp_path, data, message):
    path = tmp_path / "файл"
    path.write_bytes(data)
    with pytest.raises(ValueError, match=message):
        main.read_queries([], str(path))


def test_cli_reads_queries_from_excel(monkeypatch, tmp_path):
    session = FakeSession([ADS_PAGE])
    excel = make_excel(tmp_path / "q.xlsx", ["Запрос", "окна", "двери"])
    run_cli(monkeypatch, session, "-f", str(excel), "-o", str(tmp_path / "out.json"))
    assert [call["text"] for call in session.calls] == ["окна", "двери"]


def test_read_queries_from_stdin(monkeypatch):
    stdin = io.TextIOWrapper(io.BytesIO("купить духи\nпарфюм магазин\n".encode("utf-8")))
    monkeypatch.setattr(sys, "stdin", stdin)
    assert main.read_queries([], "-") == ["купить духи", "парфюм магазин"]


# ---------- режим с аргументами ----------

def test_cli_json_paginates_until_ads_end(monkeypatch, tmp_path):
    session = FakeSession([ads_html("a.ru", "b.ru"), ads_html("c.ru", "d.ru")])
    out = tmp_path / "out.json"
    assert run_cli(monkeypatch, session, "купить духи", "-p", "5", "-r", "213",
                   "-o", str(out)) == main.EXIT_OK

    # третья страница — «ничего не нашлось», дальше не ходим
    assert [call["p"] for call in session.calls] == [0, 1, 2]
    assert session.calls[0] == {"text": "купить духи", "p": 0, "lr": 213}
    data = read_json(out)
    assert (data["region"], data["runs"]) == (213, 1)
    assert data["queries"] == [{"query": "купить духи", "status": "ok", "runs_with_ads": 1,
                                "pages": 3, "ads": 4, "message": ""}]
    assert [(ad["domain"], ad["page"], ad["position"], ad["seen"]) for ad in data["ads"]] == [
        ("a.ru", 1, 1, 1), ("b.ru", 1, 2, 1), ("c.ru", 2, 3, 1), ("d.ru", 2, 4, 1)]


def test_cli_region_by_name(monkeypatch, tmp_path):
    session = FakeSession([ADS_PAGE])
    run_cli(monkeypatch, session, "создание сайтов", "-r", "Ставрополь",
            "-o", str(tmp_path / "out.json"))
    assert session.calls[0]["lr"] == 36


def test_cli_json_to_stdout(monkeypatch, capsys):
    assert run_cli(monkeypatch, FakeSession([ADS_PAGE]), "купить духи", "-o", "-") == main.EXIT_OK
    data = json.loads(capsys.readouterr().out)
    assert len(data["ads"]) == 2
    assert data["ads"][0]["sitelinks"][0] == "Aroma Box"


def test_cli_xlsx_has_ads_and_queries_sheets(monkeypatch, tmp_path):
    out = tmp_path / "ads.xlsx"
    assert run_cli(monkeypatch, FakeSession([ADS_PAGE]), "купить духи", "-o", str(out)) == 0
    workbook = load_workbook(out)
    assert workbook.sheetnames == ["Объявления", "Запросы"]
    rows = list(workbook["Объявления"].values)
    assert rows[0] == tuple(main.AD_COLUMNS.values())
    assert rows[1][:6] == ("купить духи", 1, 1, 1, "randewoo.ru",
                           "Парфюм в интернет-магазине Randewoo — скидки до 55%")
    assert rows[1][7:9] == ("Aroma Box | Бестселлеры | Конструктор Aroma Box | SALE", "Промо")
    assert list(workbook["Запросы"].values)[1][:5] == ("купить духи", "ok", 1, 1, 2)


def test_cli_csv(monkeypatch, tmp_path):
    out = tmp_path / "ads.csv"
    assert run_cli(monkeypatch, FakeSession([ADS_PAGE]), "купить духи", "-o", str(out)) == 0
    with open(out, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    assert rows[0] == list(main.AD_COLUMNS.values())
    assert [row[4] for row in rows[1:]] == ["randewoo.ru", "impexoils.ru"]


def test_cli_default_output_goes_to_results(monkeypatch):
    assert run_cli(monkeypatch, FakeSession([ADS_PAGE]), "купить духи") == main.EXIT_OK
    assert len(list(main.RESULTS_DIR.glob("ads_*.xlsx"))) == 1


def test_cli_no_ads_exit_code(monkeypatch, tmp_path):
    out = tmp_path / "out.json"
    assert run_cli(monkeypatch, FakeSession([]), "редкий запрос", "-o", str(out)) == main.EXIT_NO_ADS
    query = read_json(out)["queries"][0]
    assert (query["status"], query["message"]) == ("no_ads", main.STATUS_MESSAGES["no_ads"])


def test_cli_captcha_exit_code(monkeypatch, tmp_path):
    session = FakeSession([ADS_PAGE], url="https://yandex.ru/showcaptcha?retpath=abc")
    out = tmp_path / "out.json"
    assert run_cli(monkeypatch, session, "купить духи", "-o", str(out)) == main.EXIT_ERRORS
    assert read_json(out)["queries"][0]["status"] == "captcha"


def test_cli_logs_request_estimate(monkeypatch, tmp_path, capsys):
    run_cli(monkeypatch, FakeSession([]), "а", "б", "-p", "3", "--runs", "2",
            "-o", str(tmp_path / "out.json"))
    assert "Обращений к Яндексу: до 12" in capsys.readouterr().err


@pytest.mark.parametrize("argv", [
    [], ["купить", "-o", "out.txt"], ["купить", "-p", "0"], ["купить", "-r", "Атлантида"],
    ["купить", "--runs", "0"], ["купить", "--runs", str(main.MAX_RUNS + 1)],
])
def test_cli_bad_arguments(argv):
    with pytest.raises(SystemExit) as exc:
        main.main(argv)
    assert exc.value.code == main.EXIT_USAGE


def test_cli_checks_output_before_requests(monkeypatch, tmp_path):
    # Так в Windows ведёт себя файл, открытый в Excel: писать в него нельзя
    locked = tmp_path / "locked.xlsx"
    locked.write_bytes(b"")
    locked.chmod(0o444)
    readonly = tmp_path / "только-чтение"
    readonly.mkdir()
    readonly.chmod(0o555)
    session = use_session(monkeypatch, FakeSession([ADS_PAGE]))
    for output in (str(locked), str(tmp_path / "нет-папки" / "out.xlsx"),
                   str(readonly / "out.xlsx"), "имя\x00с-нулём.xlsx"):
        with pytest.raises(SystemExit) as exc:
            main.main(["купить духи", "-o", output])
        assert exc.value.code == main.EXIT_USAGE
    assert session.calls == []
    assert not any(readonly.iterdir())  # пробный файл не остаётся


# ---------- прогоны ----------

def test_runs_are_merged(monkeypatch, tmp_path):
    session = QueueSession(ads_html("a.ru", "b.ru"), ads_html("b.ru", "c.ru"), NO_ADS_PAGE)
    out = tmp_path / "out.json"
    assert run_cli(monkeypatch, session, "окна", "--runs", "3", "-o", str(out)) == main.EXIT_OK

    data = read_json(out)
    assert data["runs"] == 3
    # b.ru встретилось дважды и во втором прогоне было выше — берётся лучшая позиция
    assert [(ad["domain"], ad["position"], ad["seen"]) for ad in data["ads"]] == [
        ("a.ru", 1, 1), ("b.ru", 1, 2), ("c.ru", 2, 1)]
    assert data["queries"][0] == {"query": "окна", "status": "ok", "runs_with_ads": 2,
                                  "pages": 1, "ads": 3, "message": ""}


def test_runs_go_through_all_queries_before_repeating(monkeypatch, tmp_path):
    session = FakeSession([ADS_PAGE])
    run_cli(monkeypatch, session, "первый", "второй", "--runs", "2",
            "-o", str(tmp_path / "out.json"))
    assert [call["text"] for call in session.calls] == ["первый", "второй", "первый", "второй"]


def test_same_ad_on_two_pages_counts_once_per_run(monkeypatch, tmp_path):
    session = FakeSession([ads_html("a.ru", "b.ru"), ads_html("a.ru", "c.ru")])
    out = tmp_path / "out.json"
    run_cli(monkeypatch, session, "окна", "-p", "2", "-o", str(out))
    assert [(ad["domain"], ad["position"], ad["seen"]) for ad in read_json(out)["ads"]] == [
        ("a.ru", 1, 1), ("b.ru", 2, 1), ("c.ru", 4, 1)]


def test_captcha_in_later_run_keeps_collected_ads(monkeypatch, tmp_path):
    session = QueueSession(ads_html("a.ru"), CAPTCHA_PAGE)
    out = tmp_path / "out.json"
    assert run_cli(monkeypatch, session, "окна", "--runs", "2", "-o", str(out)) == main.EXIT_ERRORS
    data = read_json(out)
    assert (data["queries"][0]["status"], data["queries"][0]["runs_with_ads"]) == ("captcha", 1)
    assert [ad["domain"] for ad in data["ads"]] == ["a.ru"]


# ---------- надёжность: остановки, сохранение, мусор во вводе ----------

def test_key_codes_and_control_chars_are_removed_from_queries():
    # стрелки без readline приходят как «^[[D» прямо в тексте
    assert main.read_queries(["купить дхи\x1b[D\x1b[Dу", "парфюм\tмагазин\x07"]) == [
        "купить дхиу", "парфюм магазин"]


def test_xlsx_survives_control_characters(tmp_path):
    out = tmp_path / "ads.xlsx"
    ad = {"query": "купить\x1bдухи", "page": 1, "position": 1, "seen": 1, "domain": "a.ru",
          "title": "Заголовок\x08", "description": "", "sitelinks": [], "label": "",
          "click_url": ""}
    main.write_output(str(out), [ad], [], 225)
    assert list(load_workbook(out)["Объявления"].values)[1][0] == "купитьдухи"


def test_captcha_stops_collection_and_marks_the_rest(monkeypatch, tmp_path, capsys):
    session = QueueSession(ads_html("a.ru"), CAPTCHA_PAGE)
    out = tmp_path / "out.json"
    code = run_cli(monkeypatch, session, "первый", "второй", "третий", "-o", str(out))
    assert code == main.EXIT_ERRORS
    assert len(session.calls) == 2  # после капчи к Яндексу больше не ходим
    queries = read_json(out)["queries"]
    assert [q["status"] for q in queries] == ["ok", "captcha", "skipped"]
    assert "капчу" in queries[2]["message"]
    err = capsys.readouterr().err
    assert "Собрано не всё" in err and "--proxy" in err


def test_network_errors_stop_after_three_in_a_row(monkeypatch, tmp_path, capsys):
    down = requests.ConnectionError("ProxyError(MaxRetryError(%D0%BA…))")
    session = QueueSession(down, down, down, down)
    out = tmp_path / "out.json"
    assert run_cli(monkeypatch, session, *"абвгд", "-o", str(out)) == main.EXIT_ERRORS
    assert len(session.calls) == main.MAX_NETWORK_ERRORS
    queries = read_json(out)["queries"]
    assert [q["status"] for q in queries] == ["error"] * 3 + ["skipped"] * 2
    assert queries[0]["message"] == "нет связи с Яндексом (ConnectionError)."
    err = capsys.readouterr().err
    assert "%D0%BA" not in err  # вместо служебного текста — понятная причина


def test_ctrl_c_keeps_what_was_collected(monkeypatch, tmp_path):
    session = QueueSession(ads_html("a.ru"), KeyboardInterrupt())
    out = tmp_path / "out.json"
    assert run_cli(monkeypatch, session, "первый", "второй", "-o", str(out)) == main.EXIT_INTERRUPTED
    data = read_json(out)
    assert [ad["domain"] for ad in data["ads"]] == ["a.ru"]
    assert data["queries"][1]["message"] == "Не обработан: сбор остановлен по Ctrl+C."


def test_result_is_saved_after_each_query(monkeypatch, tmp_path):
    # Падение посреди сбора (или закрытое окно) не уносит уже собранное
    session = QueueSession(ads_html("a.ru"), RuntimeError("окно закрыли"))
    out = tmp_path / "ads.xlsx"
    with pytest.raises(RuntimeError):
        run_cli(monkeypatch, session, "первый", "второй", "-o", str(out))
    assert [row[4] for row in list(load_workbook(out)["Объявления"].values)[1:]] == ["a.ru"]


def test_failed_write_falls_back_to_backup_json(monkeypatch, tmp_path, capsys):
    real_write = main.write_output

    def broken(output, *args, **kwargs):
        if output.endswith(".xlsx"):
            raise RuntimeError("что-то сломалось при записи")
        return real_write(output, *args, **kwargs)

    monkeypatch.setattr(main, "write_output", broken)
    code = run_cli(monkeypatch, FakeSession([ADS_PAGE]), "купить духи", "-o", str(tmp_path / "a.xlsx"))
    assert code == main.EXIT_ERRORS
    backups = list(main.RESULTS_DIR.glob("backup_*.json"))
    assert len(backups) == 1 and len(read_json(backups[0])["ads"]) == 2
    err = capsys.readouterr().err
    assert "запасной файл" in err and "Готово" not in err


def test_unexpected_page_is_saved_for_the_developer(monkeypatch, tmp_path):
    out = tmp_path / "out.json"
    session = FakeSession(['<ul id="search-result"></ul>'])
    assert run_cli(monkeypatch, session, "окна", "-o", str(out)) == main.EXIT_ERRORS
    saved = list((main.RESULTS_DIR / "debug").glob("unexpected_*_окна.html"))
    assert len(saved) == 1
    assert str(saved[0]) in read_json(out)["queries"][0]["message"]


@pytest.mark.parametrize("pages, expected", [
    ([ADS_PAGE], "Готово"),
    ([], "Объявлений нет"),
    (['<ul id="search-result"></ul>'], "Не получилось"),
])
def test_final_message_matches_the_outcome(monkeypatch, tmp_path, capsys, pages, expected):
    run_cli(monkeypatch, FakeSession(pages), "окна", "-o", str(tmp_path / "out.json"))
    err = capsys.readouterr().err
    assert expected in err
    assert ("Готово" in err) == (expected == "Готово")


def test_partial_result_is_not_called_done(monkeypatch, tmp_path, capsys):
    session = QueueSession(ads_html("a.ru"), '<ul id="search-result"></ul>')
    run_cli(monkeypatch, session, "первый", "второй", "-o", str(tmp_path / "out.json"))
    assert "Собрано частично" in capsys.readouterr().err


def test_human_hints_have_no_command_line_flags(monkeypatch, tmp_path, capsys, no_sleep):
    use_session(monkeypatch, QueueSession(CAPTCHA_PAGE))
    main.run(["окна"], output=str(tmp_path / "out.json"), human=True)
    err = capsys.readouterr().err
    assert "Подождите 15–30 минут" in err and "--" not in err


# ---------- цвет ----------

def test_paint_colors_only_real_terminals(monkeypatch):
    monkeypatch.delenv("NO_COLOR", raising=False)
    monkeypatch.setenv("TERM", "xterm-256color")
    monkeypatch.setattr(main, "_ansi", True)
    assert main.paint("ок", "bold", "green", stream=FakeTTY()) == "\033[1;32mок\033[0m"
    assert main.paint("ок", "green", stream=io.StringIO()) == "ок"  # файл или агент
    monkeypatch.setenv("NO_COLOR", "1")
    assert main.paint("ок", "green", stream=FakeTTY()) == "ок"


def test_windows_console_without_color_support_falls_back_to_plain_text(monkeypatch):
    # На Linux ctypes.windll нет — так же, как если включить цвет в консоли не удалось
    monkeypatch.setattr(main.sys, "platform", "win32")
    monkeypatch.setattr(main, "_ansi", True)
    main._enable_ansi()
    assert main._ansi is False


# ---------- запуск отдельным процессом: кодировки и зависимости ----------

def run_script(*args, python_flags=()):
    # cp1252 — как в Windows, когда вывод программы перехватывает агент или файл
    env = {**os.environ, "PYTHONIOENCODING": "cp1252"}
    return subprocess.run([sys.executable, *python_flags, str(ROOT / "main.py"), *args],
                          stdin=subprocess.DEVNULL, capture_output=True, env=env, timeout=60)


def test_redirected_output_is_utf8_whatever_the_system_encoding():
    result = run_script()  # без запросов и без терминала: ошибка, а не пошаговый режим
    assert result.returncode == main.EXIT_USAGE
    stderr = result.stderr.decode("utf-8")
    assert "не задано ни одного запроса" in stderr
    assert "\033[" not in stderr  # и без цветовых кодов


def test_missing_dependency_is_explained():
    result = run_script("купить духи", python_flags=("-S",))  # -S: без site-packages
    assert result.returncode == 1
    assert "pip install -r requirements.txt" in result.stderr.decode("utf-8")


# ---------- ручной режим: два вопроса и карточка ----------

def asked(session):
    """Какие запросы ушли в Яндекс — без повторов прогонов, по порядку."""
    return list(dict.fromkeys(call["text"] for call in session.calls))


def start_interactive(monkeypatch, session, *values):
    use_session(monkeypatch, session)
    monkeypatch.setattr(main.time, "sleep", lambda seconds: None)
    opened = []
    monkeypatch.setattr(main, "_open_file", opened.append)
    return main.interactive(answers(*values)), opened


def test_interactive_two_questions_and_start(monkeypatch, capsys):
    session = FakeSession([ADS_PAGE])
    # запросы, пустая строка, Enter «дальше», город, Enter на карточке
    code, opened = start_interactive(monkeypatch, session, "купить духи", "парфюм магазин",
                                     "", "", "Ставрополь", "")
    assert code == main.EXIT_OK
    # по умолчанию в ручном режиме — 3 прогона
    assert [(c["text"], c["p"], c["lr"]) for c in session.calls] == [
        ("купить духи", 0, 36), ("парфюм магазин", 0, 36)] * main.INTERACTIVE_RUNS
    printed = capsys.readouterr().out
    assert "Шаг 1 из 2" in printed and "Шаг 2 из 2" in printed
    assert "1. Запросы: 2 (купить духи, парфюм магазин)" in printed
    assert "2. Город: Ставрополь (36)" in printed
    assert "3. Полнота: прогонов — 3, страниц на запрос — 1" in printed
    assert "обращений к Яндексу — до 6" in printed
    assert len(list(main.RESULTS_DIR.glob("ads_*.xlsx"))) == 1
    assert opened == []  # на «Открыть таблицу?» ответа не было (EOF) — не открываем


def test_interactive_card_items_can_be_changed(monkeypatch):
    session = QueueSession(ads_html("a.ru"), ads_html("b.ru"))
    code, opened = start_interactive(
        monkeypatch, session, "окна", "", "", "36",
        "3", "2 прогона", "", "4", "итог", "", "")
    assert code == main.EXIT_OK
    assert len(session.calls) == 2  # 2 прогона × 1 страница
    output = main.RESULTS_DIR / "итог.xlsx"
    assert [row[4] for row in list(load_workbook(output)["Объявления"].values)[1:]] == [
        "a.ru", "b.ru"]
    assert opened == [str(output)]  # Enter на «Открыть таблицу?»


def test_interactive_reasks_with_helpful_messages(monkeypatch, capsys):
    session = FakeSession([ADS_PAGE])
    code, _ = start_interactive(
        monkeypatch, session,
        "", "купить духи", "", "",              # пустой ввод, потом запрос
        "Атлантида", "Ставропольский край", "",  # незнакомый город, похожий — «да»
        "7", "3", "11", "3-5", "x", "",          # неверный пункт, прогоны, страницы
        "")
    assert code == main.EXIT_OK
    printed = capsys.readouterr().out
    assert "Нужен хотя бы один запрос" in printed
    assert "Такого города в списке нет" in printed
    assert "Может быть, Ставрополь (36)?" in printed
    assert "цифру от 1 до 4" in printed
    assert f"Нужно число от 1 до {main.MAX_RUNS}." in printed
    assert "Введите число, например 3." in printed
    # «3-5» — берётся первое число: 3 прогона
    assert [c["lr"] for c in session.calls] == [36, 36, 36]


@pytest.mark.parametrize("typed, code", [
    ("36 Ставрополь", 36), ("мск", 213), ("краснадар", 35), ("Россия", 225),
])
def test_interactive_region_forms(monkeypatch, typed, code):
    session = FakeSession([ADS_PAGE])
    extra = ("",) if typed == "краснадар" else ()  # «Может быть, Краснодар?» — да
    start_interactive(monkeypatch, session, "окна", "", "", typed, *extra, "")
    assert session.calls[0]["lr"] == code


def test_interactive_city_is_required(monkeypatch, capsys):
    session = FakeSession([ADS_PAGE])
    start_interactive(monkeypatch, session, "окна", "", "", "", "Россия", "")
    assert "Укажите город" in capsys.readouterr().out
    assert session.calls[0]["lr"] == 225


def test_interactive_region_list_and_unknown_code(monkeypatch, capsys):
    session = FakeSession([ADS_PAGE])
    start_interactive(monkeypatch, session, "окна", "", "", "?", "999999", "0", "Сочи", "")
    printed = capsys.readouterr().out
    assert "Ставрополь 36" in printed and "aistudio.yandex.ru" in printed
    assert "Кода 999999 нет среди известных городов" in printed
    assert session.calls[0]["lr"] == 239


def test_interactive_paste_with_blank_line_keeps_all_queries(monkeypatch):
    session = FakeSession([ADS_PAGE])
    start_interactive(monkeypatch, session, "а", "б", "", "в", "", "", "36", "")
    assert asked(session) == ["а", "б", "в"]


def test_interactive_multiline_paste_in_one_input(monkeypatch):
    # readline со «вставкой целиком» отдаёт сразу несколько строк
    session = FakeSession([ADS_PAGE])
    start_interactive(monkeypatch, session, "а\nб\n\nв", "", "", "36", "")
    assert asked(session) == ["а", "б", "в"]


def test_interactive_comma_separated_queries_can_be_split(monkeypatch):
    session = FakeSession([ADS_PAGE])
    start_interactive(monkeypatch, session, "купить духи, парфюм магазин, туалетная вода",
                      "", "", "", "36", "")
    assert asked(session) == ["купить духи", "парфюм магазин", "туалетная вода"]


def test_interactive_queries_from_dragged_file(monkeypatch, tmp_path):
    session = FakeSession([ADS_PAGE])
    queries = tmp_path / "мои запросы.txt"
    queries.write_text("купить духи\nпарфюм магазин\n", encoding="utf-8")
    # Windows при перетаскивании файла в окно берёт путь в кавычки
    start_interactive(monkeypatch, session, f'"{queries}"', "", "", "213", "")
    assert asked(session) == ["купить духи", "парфюм магазин"]


def test_interactive_file_typos_and_broken_files_are_not_queries(monkeypatch, tmp_path,
                                                                    capsys):
    session = FakeSession([ADS_PAGE])
    broken = tmp_path / "отчёт.docx"
    broken.write_bytes(b"PK\x03\x04" + b"\x00" * 100)  # zip-архив, но не таблица
    start_interactive(monkeypatch, session, "/нет/такого/queris.txt", str(broken), "окна",
                      "", "", "36", "")
    printed = capsys.readouterr().out
    assert "Файл не найден: /нет/такого/queris.txt" in printed
    assert "не таблица Excel и не текст" in printed
    assert asked(session) == ["окна"]


def test_interactive_queries_from_dragged_excel(monkeypatch, tmp_path, capsys):
    session = FakeSession([ADS_PAGE])
    excel = make_excel(tmp_path / "запросы.xlsx", ["Ключевые слова", "купить духи",
                                                    "парфюм магазин"])
    start_interactive(monkeypatch, session, str(excel), "", "", "36", "")
    assert "Из файла «запросы.xlsx» взято запросов: 2" in capsys.readouterr().out
    assert asked(session) == ["купить духи", "парфюм магазин"]


def test_interactive_finds_file_next_to_the_program(monkeypatch, tmp_path):
    session = FakeSession([ADS_PAGE])
    (main.RESULTS_DIR.parent / "queries.txt").write_text("окна\n", encoding="utf-8")
    elsewhere = tmp_path / "другая папка"
    elsewhere.mkdir()
    monkeypatch.chdir(elsewhere)  # программа запущена не из своей папки
    start_interactive(monkeypatch, session, "queries.txt", "", "", "36", "")
    assert asked(session) == ["окна"]


def test_interactive_output_folder_existing_file_and_wrong_format(monkeypatch, tmp_path,
                                                                    capsys):
    session = FakeSession([ADS_PAGE])
    main.RESULTS_DIR.mkdir()
    old = main.RESULTS_DIR / "отчёт.xlsx"
    old.write_bytes("старый отчёт".encode())
    start_interactive(monkeypatch, session, "окна", "", "", "36",
                      "4", "отчёт.txt", "отчёт", "")
    printed = capsys.readouterr().out
    assert "Не знаете, что выбрать, — оставьте Excel." in printed and "'-'" not in printed
    assert "сохраню как «отчёт (2).xlsx»" in printed
    assert old.read_bytes() == "старый отчёт".encode()
    assert (main.RESULTS_DIR / "отчёт (2).xlsx").exists()

    folder = tmp_path / "Мои отчёты"
    folder.mkdir()
    start_interactive(monkeypatch, FakeSession([ADS_PAGE]), "окна", "", "", "36",
                      "4", str(folder), "")
    assert len(list(folder.glob("ads_*.xlsx"))) == 1  # папка — сохраняем внутрь


def test_interactive_cancel_and_eof(monkeypatch, capsys):
    session = FakeSession([ADS_PAGE])
    code, _ = start_interactive(monkeypatch, session, "окна", "", "", "36", "0")
    assert code == main.EXIT_USAGE and session.calls == []
    assert main.interactive(answers()) == main.EXIT_USAGE
    assert "Отменено" in capsys.readouterr().out


def test_clean_path_unescapes_macos_drag_and_drop():
    if sys.platform == "win32":
        pytest.skip("на Windows обратная косая черта — разделитель пути")
    assert main._clean_path(r"/Users/me/My\ Files/q.txt ") == "/Users/me/My Files/q.txt"


def test_long_lines_wrap_by_words(monkeypatch):
    monkeypatch.setattr(main.sys, "stdout", FakeTTY())
    monkeypatch.setattr(main.shutil, "get_terminal_size", lambda: os.terminal_size((41, 24)))
    text = "  " + "Реклама сильно зависит от города Ставрополь и Ростов-на-Дону " * 2
    lines = main._wrap(text).split("\n")
    assert all(len(line) <= 40 for line in lines)
    assert all(line.startswith("  ") for line in lines)  # отступ сохраняется
    assert "Ростов-на-Дону" in main._wrap(text)  # слова и дефисы не рвутся


def test_offer_install_declined_and_accepted(monkeypatch, capsys):
    monkeypatch.setattr("builtins.input", lambda prompt="": "0")
    assert main._offer_install("bs4") == 1
    assert "pip install -r" in capsys.readouterr().out

    calls = []

    class FakeChild:
        def __init__(self, command):
            calls.append(command)

        def wait(self):
            return 0

    def fake_run(command, **kwargs):
        calls.append(command)
        return subprocess.CompletedProcess(command, 0)

    monkeypatch.setattr("builtins.input", lambda prompt="": "")
    monkeypatch.setattr(main.subprocess, "run", fake_run)
    monkeypatch.setattr(main.subprocess, "Popen", FakeChild)
    monkeypatch.setattr(main.signal, "signal", lambda *args: None)
    assert main._offer_install("bs4") == 0
    assert calls[0][1:4] == ["-m", "pip", "install"]
    assert calls[1] == [sys.executable, *sys.argv]  # после установки — перезапуск
