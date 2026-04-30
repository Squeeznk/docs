"""Build the Automacon WCS project plan as an XLSX workbook with a Gantt-style chart.

The script generates `automacon-wcs-project-plan.xlsx` in the same folder.

The workbook contains four sheets:
    1. "План-график"   — main task list with computed durations.
    2. "Вехи"           — list of milestones (M1..M10).
    3. "Обязательства"  — customer obligations with deadlines.
    4. "Диаграмма"      — embedded native bar chart rendered as a Gantt:
                          start dates as the invisible "spacer" series and
                          durations as the visible bar.

The data source is the master plan in `automacon-wcs-project-plan.md`.
Re-run `python3 build_project_plan_xlsx.py` whenever the plan changes.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

OUT_PATH = Path(__file__).resolve().parent / "automacon-wcs-project-plan.xlsx"

# ---------------------------------------------------------------------------
# Master data
# ---------------------------------------------------------------------------

# Format: (level, code, title, start, end, owner, milestone_code_or_None)
# level: "stage" or "task". Stages have empty start/end (computed from tasks).

STAGES_TASKS: list[tuple[str, str, str, dt.date | None, dt.date | None, str, str | None]] = [
    # Этап 1
    ("stage", "Этап 1",  "Предпроектная стадия",                                                            None, None, "Андрей Соколов",       "М1"),
    ("task",  "1.1",    "Получение ответов от подразделения «Охрана труда»",                                dt.date(2026, 4, 30), dt.date(2026, 5, 12), "Андрей Соколов",  None),
    ("task",  "1.2",    "Получение ответов от подразделения «ИТ»",                                          dt.date(2026, 4, 30), dt.date(2026, 5, 12), "Антон Руль",      None),
    ("task",  "1.3",    "Получение ответов от подразделения «Информационная безопасность»",                 dt.date(2026, 4, 30), dt.date(2026, 5, 15), "Антон Руль",      None),
    ("task",  "1.4",    "Получение ответов от подразделения «Производственная логистика и WMS Axelot»",     dt.date(2026, 4, 30), dt.date(2026, 5, 12), "Андрей Соколов",  None),
    ("task",  "1.5",    "Получение ответов от подразделения «Производственные участки»",                    dt.date(2026, 4, 30), dt.date(2026, 5, 15), "Андрей Соколов",  None),
    ("task",  "1.6",    "Финализация сценариев работы парка и технического задания",                        dt.date(2026, 5, 15), dt.date(2026, 5, 22), "Андрей Соколов",  None),
    ("task",  "1.7",    "Подготовка и передача документа «Требования к среде эксплуатации»",                dt.date(2026, 5, 4),  dt.date(2026, 5, 12), "Антон Руль",      None),
    ("task",  "1.8",    "Проработка и согласование тележки под автономный транспортный робот",              dt.date(2026, 4, 23), dt.date(2026, 5, 15), "Андрей Соколов",  None),
    ("task",  "1.9",    "Подготовка и согласование устава проекта",                                         dt.date(2026, 4, 21), dt.date(2026, 5, 12), "Даниил Солнцев",  None),
    ("task",  "1.10",   "Проработка альтернативной логистики",                                              dt.date(2026, 4, 28), dt.date(2026, 5, 15), "Андрей Соколов",  None),
    ("task",  "1.11",   "Финализация план-графика проекта",                                                 dt.date(2026, 5, 18), dt.date(2026, 5, 22), "Андрей Соколов",  None),

    # Этап 2
    ("stage", "Этап 2",  "Производство оборудования",                                                       None, None, "Алексей Матюшкин",  "М2"),
    ("task",  "2.1",    "Оплата производственного заказа",                                                  dt.date(2026, 5, 4),  dt.date(2026, 5, 6),  "Даниил Солнцев",  None),
    ("task",  "2.2",    "Производственный цикл (5 роботов и 3 зарядные станции)",                           dt.date(2026, 5, 6),  dt.date(2026, 6, 12), "Алексей Матюшкин", None),
    ("task",  "2.3",    "Заводские приёмо-сдаточные испытания",                                             dt.date(2026, 6, 8),  dt.date(2026, 6, 15), "Алексей Матюшкин", None),
    ("task",  "2.4",    "Подготовка отгрузочной документации",                                              dt.date(2026, 6, 12), dt.date(2026, 6, 15), "Алексей Матюшкин", None),

    # Этап 3
    ("stage", "Этап 3",  "Логистика и таможенное оформление",                                               None, None, "Андрей Соколов",     "М3"),
    ("task",  "3.1",    "Подготовка отправки и оформление на авиатранспорт",                                dt.date(2026, 6, 15), dt.date(2026, 6, 18), "Андрей Соколов",  None),
    ("task",  "3.2",    "Авиаперевозка",                                                                    dt.date(2026, 6, 18), dt.date(2026, 6, 23), "Андрей Соколов",  None),
    ("task",  "3.3",    "Таможенное оформление",                                                            dt.date(2026, 6, 23), dt.date(2026, 6, 30), "Андрей Соколов",  None),
    ("task",  "3.4",    "Доставка на площадку заказчика",                                                   dt.date(2026, 6, 30), dt.date(2026, 7, 2),  "Андрей Соколов",  None),

    # Этап 4
    ("stage", "Этап 4",  "Подготовка площадки заказчика",                                                   None, None, "Заказчик",           "М4"),
    ("task",  "4.1",    "Электротехнические работы под зарядные станции",                                   dt.date(2026, 5, 11), dt.date(2026, 6, 19), "Заказчик: гл. энергетик", None),
    ("task",  "4.2",    "Подготовка сетевой инфраструктуры (ВМ, VLAN, межсетевое экранирование)",           dt.date(2026, 5, 11), dt.date(2026, 6, 19), "Заказчик: ИТ-инфраструктура", None),
    ("task",  "4.3",    "Беспроводная сеть в зонах работы (Wi-Fi, радиоплан)",                              dt.date(2026, 5, 11), dt.date(2026, 6, 19), "Заказчик: сетевые инженеры", None),
    ("task",  "4.4",    "Согласование маршрутов, точек REST, зон зарядки",                                  dt.date(2026, 5, 18), dt.date(2026, 6, 12), "Заказчик: технология/ОТ/диспетчер", None),
    ("task",  "4.5",    "Выдача доступов команде Automacon",                                                dt.date(2026, 5, 11), dt.date(2026, 5, 22), "Заказчик: ИТ",   None),
    ("task",  "4.6",    "Подготовка зон работы (чистота, разметка)",                                        dt.date(2026, 5, 25), dt.date(2026, 6, 26), "Заказчик: эксплуатация участка", None),
    ("task",  "4.7",    "Согласование ИБ-регламентов, технологические учётные записи",                      dt.date(2026, 5, 18), dt.date(2026, 6, 19), "Заказчик: ИБ",   None),

    # Этап 5
    ("stage", "Этап 5",  "Развёртывание серверной части WCS",                                               None, None, "Антон Руль",         "М5"),
    ("task",  "5.1",    "Получение виртуальных машин и сетевых доступов",                                   dt.date(2026, 5, 22), dt.date(2026, 5, 29), "Антон Руль",     None),
    ("task",  "5.2",    "Установка платформы в тестовом контуре, базовая настройка",                        dt.date(2026, 5, 29), dt.date(2026, 6, 5),  "Антон Руль",     None),
    ("task",  "5.3",    "Интеграция с корпоративной аутентификацией, ролевая модель",                       dt.date(2026, 6, 5),  dt.date(2026, 6, 12), "Антон Руль",     None),
    ("task",  "5.4",    "Интеграция с WMS Axelot — реализация коннектора",                                  dt.date(2026, 6, 5),  dt.date(2026, 6, 26), "Антон Руль",     None),
    ("task",  "5.5",    "Развёртывание контура мониторинга",                                                dt.date(2026, 6, 12), dt.date(2026, 6, 26), "Антон Руль",     None),
    ("task",  "5.6",    "Установка платформы в промышленном контуре, регрессионная проверка",               dt.date(2026, 6, 26), dt.date(2026, 7, 3),  "Антон Руль",     None),
    ("task",  "5.7",    "Подготовка к приёмке роботов: справочники, шаблоны устройств",                     dt.date(2026, 6, 26), dt.date(2026, 7, 3),  "Антон Руль",     None),

    # Этап 6
    ("stage", "Этап 6",  "Приёмка, калибровка и подключение оборудования",                                  None, None, "Андрей Соколов",     "М6"),
    ("task",  "6.1",    "Распаковка, входной контроль, оформление приёмочных актов",                        dt.date(2026, 7, 2),  dt.date(2026, 7, 6),  "Андрей Соколов", None),
    ("task",  "6.2",    "Монтаж и пусконаладка зарядных станций",                                           dt.date(2026, 7, 6),  dt.date(2026, 7, 13), "Алексей Матюшкин", None),
    ("task",  "6.3",    "Первичная настройка и калибровка роботов",                                         dt.date(2026, 7, 6),  dt.date(2026, 7, 24), "Алексей Матюшкин", None),
    ("task",  "6.4",    "Регистрация роботов в WCS, технологические учётные записи",                        dt.date(2026, 7, 13), dt.date(2026, 7, 24), "Антон Руль",     None),
    ("task",  "6.5",    "Подключение периферии к WCS",                                                      dt.date(2026, 7, 13), dt.date(2026, 7, 24), "Антон Руль",     None),
    ("task",  "6.6",    "Проверки безопасности (аварийный останов, режимы)",                                dt.date(2026, 7, 20), dt.date(2026, 7, 24), "Алексей Матюшкин", None),

    # Этап 7
    ("stage", "Этап 7",  "Картография",                                                                     None, None, "Команда пусконаладки", "М7"),
    ("task",  "7.1",    "Лидарное сканирование зон работы",                                                 dt.date(2026, 7, 24), dt.date(2026, 7, 27), "Команда пусконаладки", None),
    ("task",  "7.2",    "Подготовка карты в формате LIF",                                                   dt.date(2026, 7, 27), dt.date(2026, 7, 30), "Команда пусконаладки", None),
    ("task",  "7.3",    "Валидация графа, согласование с технологией и ОТ",                                 dt.date(2026, 7, 30), dt.date(2026, 7, 31), "Андрей Соколов", None),
    ("task",  "7.4",    "Загрузка карты в WCS, контрольный прогон роботов",                                 dt.date(2026, 7, 30), dt.date(2026, 7, 31), "Антон Руль",     None),

    # Этап 8
    ("stage", "Этап 8",  "Отладка сценариев и обучение персонала",                                          None, None, "Андрей Соколов",     "М8"),
    ("task",  "8.1",    "Поочередная отладка сценариев работы парка",                                       dt.date(2026, 7, 31), dt.date(2026, 8, 21), "Андрей Соколов", None),
    ("task",  "8.2",    "Тестирование точек погрузки и разгрузки, точность позиционирования",               dt.date(2026, 8, 10), dt.date(2026, 8, 24), "Андрей Соколов", None),
    ("task",  "8.3",    "Тестирование сценариев исключений",                                                dt.date(2026, 8, 17), dt.date(2026, 8, 28), "Антон Руль",     None),
    ("task",  "8.4",    "Обучение операторов и диспетчеров",                                                dt.date(2026, 8, 17), dt.date(2026, 8, 28), "Команда внедрения", None),
    ("task",  "8.5",    "Инструктажи под подпись (охрана труда)",                                           dt.date(2026, 8, 24), dt.date(2026, 8, 28), "Совместно с заказчиком", None),
    ("task",  "8.6",    "Передача эксплуатационной документации",                                           dt.date(2026, 8, 24), dt.date(2026, 8, 28), "Команда внедрения", None),

    # Этап 9
    ("stage", "Этап 9",  "Опытная эксплуатация",                                                            None, None, "Андрей Соколов",     "М9"),
    ("task",  "9.1",    "Запуск опытной эксплуатации, дежурство команды",                                   dt.date(2026, 8, 31), dt.date(2026, 9, 11), "Андрей Соколов", None),
    ("task",  "9.2",    "Журналирование инцидентов, ежедневный разбор",                                     dt.date(2026, 8, 31), dt.date(2026, 9, 13), "Андрей Соколов", None),
    ("task",  "9.3",    "Закрытие замечаний — корректировка сценариев и настроек",                          dt.date(2026, 9, 3),  dt.date(2026, 9, 13), "Антон Руль",     None),
    ("task",  "9.4",    "Подведение итогов опытной эксплуатации с заказчиком",                              dt.date(2026, 9, 11), dt.date(2026, 9, 13), "Андрей Соколов", None),

    # Этап 10
    ("stage", "Этап 10", "ПСИ и передача в промышленную эксплуатацию",                                      None, None, "Даниил Солнцев",     "М10"),
    ("task",  "10.1",   "Подготовка программы и методики ПСИ",                                              dt.date(2026, 9, 14), dt.date(2026, 9, 16), "Даниил Солнцев", None),
    ("task",  "10.2",   "Проведение ПСИ по согласованной программе",                                        dt.date(2026, 9, 16), dt.date(2026, 9, 23), "Даниил Солнцев", None),
    ("task",  "10.3",   "Подготовка и подписание акта приёмки",                                             dt.date(2026, 9, 23), dt.date(2026, 9, 25), "Даниил Солнцев", None),
    ("task",  "10.4",   "Передача на сопровождение, активация SLA",                                         dt.date(2026, 9, 25), dt.date(2026, 9, 27), "Команда сопровождения", None),
]

MILESTONES: list[tuple[str, dt.date, str]] = [
    ("М1",  dt.date(2026, 5, 22), "Предпроектная стадия закрыта (ТЗ, сценарии, тележка, требования к среде, устав)."),
    ("М2",  dt.date(2026, 6, 15), "Оборудование (5 роботов, 3 ЗУ) готово к отгрузке."),
    ("М3",  dt.date(2026, 7, 2),  "Оборудование на площадке заказчика."),
    ("М4",  dt.date(2026, 6, 30), "Площадка заказчика готова к приёмке оборудования."),
    ("М5",  dt.date(2026, 7, 3),  "Серверная часть WCS установлена и интегрирована."),
    ("М6",  dt.date(2026, 7, 24), "Парк подключён к WCS, прошёл проверку безопасности."),
    ("М7",  dt.date(2026, 7, 31), "Карта в формате LIF утверждена и загружена."),
    ("М8",  dt.date(2026, 8, 28), "Сценарии отлажены, персонал обучен."),
    ("М9",  dt.date(2026, 9, 13), "Опытная эксплуатация завершена."),
    ("М10", dt.date(2026, 9, 27), "Промышленная эксплуатация и сопровождение."),
]

OBLIGATIONS: list[tuple[str, dt.date, str]] = [
    ("Ответы по опросным листам всех пяти подразделений",                          dt.date(2026, 5, 15), "М1"),
    ("Утверждение технического задания и сценариев работы парка",                  dt.date(2026, 5, 22), "М1"),
    ("Подписание устава проекта",                                                  dt.date(2026, 5, 22), "М1"),
    ("Утверждение конструкции тележки",                                            dt.date(2026, 5, 15), "М1"),
    ("Выдача доступов команде Automacon",                                          dt.date(2026, 5, 22), "М4"),
    ("Выделение виртуальных машин под платформу",                                  dt.date(2026, 5, 29), "М5"),
    ("Готовность электропитания зарядной зоны",                                    dt.date(2026, 6, 19), "М4"),
    ("Готовность беспроводной сети в зонах работы",                                dt.date(2026, 6, 19), "М4"),
    ("Согласование маршрутов, REST-точек, зон зарядки (технология/ОТ/диспетчер)",  dt.date(2026, 6, 12), "М4"),
    ("Согласование ИБ-регламентов и выпуск технологических учётных записей",       dt.date(2026, 6, 19), "М4 / М5"),
    ("Подготовка зон работы (чистота, разметка)",                                  dt.date(2026, 6, 26), "М4"),
    ("Назначение ответственных за инструктажи и обучение",                         dt.date(2026, 8, 17), "М8"),
    ("Назначение приёмочной комиссии",                                             dt.date(2026, 9, 14), "М10"),
]

# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------

THIN = Side(style="thin", color="C8D6CC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

HEADER_FILL = PatternFill("solid", fgColor="0D9373")
STAGE_FILL = PatternFill("solid", fgColor="D6F0E1")
ALT_FILL = PatternFill("solid", fgColor="F4FBF6")
MILESTONE_FILL = PatternFill("solid", fgColor="FFF7D6")

HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
STAGE_FONT = Font(bold=True, color="0A3E2D")
TITLE_FONT = Font(bold=True, size=14, color="0A3E2D")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def autosize(ws: Worksheet, widths: dict[str, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def apply_border(ws: Worksheet, cell_range: str) -> None:
    for row in ws[cell_range]:
        for cell in row:
            cell.border = BORDER


# ---------------------------------------------------------------------------
# Sheet 1: Plan
# ---------------------------------------------------------------------------


def build_plan_sheet(ws: Worksheet) -> dict:
    """Fill the main plan sheet, return a dict with row references for the chart."""
    ws.title = "План-график"
    ws.freeze_panes = "A4"

    ws["A1"] = "Automacon WCS — план-график проекта"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:G1")

    ws["A2"] = "Целевая дата перехода в промышленную эксплуатацию: 27.09.2026"
    ws["A2"].font = Font(italic=True, color="666666")
    ws.merge_cells("A2:G2")

    headers = ["№", "Задача", "Дата начала", "Дата окончания", "Длит., дн.", "Ответственный", "Веха"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[3].height = 28

    # Compute stage spans from member tasks
    stage_indices: dict[int, tuple[dt.date | None, dt.date | None]] = {}
    current_stage_idx: int | None = None

    row_for_chart: list[tuple[int, str, dt.date, dt.date, bool]] = []
    # (excel_row, label, start, end, is_stage)

    excel_row = 4
    alt = False
    for entry in STAGES_TASKS:
        level, code, title, start, end, owner, milestone = entry
        if level == "stage":
            stage_indices[excel_row] = (None, None)
            current_stage_idx = excel_row
            ws.cell(row=excel_row, column=1, value=code).font = STAGE_FONT
            ws.cell(row=excel_row, column=2, value=title).font = STAGE_FONT
            ws.cell(row=excel_row, column=6, value=owner).font = STAGE_FONT
            ws.cell(row=excel_row, column=7, value=milestone or "").font = STAGE_FONT
            for col in range(1, 8):
                ws.cell(row=excel_row, column=col).fill = STAGE_FILL
                ws.cell(row=excel_row, column=col).border = BORDER
                if col not in (3, 4, 5):
                    ws.cell(row=excel_row, column=col).alignment = LEFT
            ws.row_dimensions[excel_row].height = 22
        else:
            duration = (end - start).days if start and end else 0
            ws.cell(row=excel_row, column=1, value=code)
            ws.cell(row=excel_row, column=2, value=title)
            c_start = ws.cell(row=excel_row, column=3, value=start)
            c_end = ws.cell(row=excel_row, column=4, value=end)
            c_dur = ws.cell(row=excel_row, column=5, value=duration)
            ws.cell(row=excel_row, column=6, value=owner)
            ws.cell(row=excel_row, column=7, value="")

            c_start.number_format = "DD.MM.YYYY"
            c_end.number_format = "DD.MM.YYYY"

            for col in range(1, 8):
                cell = ws.cell(row=excel_row, column=col)
                cell.border = BORDER
                if alt:
                    cell.fill = ALT_FILL
                cell.alignment = LEFT if col in (1, 2, 6) else CENTER
            ws.row_dimensions[excel_row].height = 18

            # Track stage span
            if current_stage_idx is not None:
                cur_start, cur_end = stage_indices[current_stage_idx]
                stage_indices[current_stage_idx] = (
                    start if cur_start is None or start < cur_start else cur_start,
                    end if cur_end is None or end > cur_end else cur_end,
                )

            row_for_chart.append((excel_row, code, title, start, end, False))
            alt = not alt

        excel_row += 1

    # Now write computed start/end/duration for stages
    for r, (s, e) in stage_indices.items():
        if s and e:
            cs = ws.cell(row=r, column=3, value=s)
            ce = ws.cell(row=r, column=4, value=e)
            cd = ws.cell(row=r, column=5, value=(e - s).days)
            cs.number_format = "DD.MM.YYYY"
            ce.number_format = "DD.MM.YYYY"
            cs.font = STAGE_FONT
            ce.font = STAGE_FONT
            cd.font = STAGE_FONT
            cs.alignment = CENTER
            ce.alignment = CENTER
            cd.alignment = CENTER

    autosize(
        ws,
        {
            "A": 9,
            "B": 70,
            "C": 14,
            "D": 14,
            "E": 12,
            "F": 26,
            "G": 8,
        },
    )

    return {
        "tasks": row_for_chart,
        "last_row": excel_row - 1,
    }


# ---------------------------------------------------------------------------
# Sheet 2: Milestones
# ---------------------------------------------------------------------------


def build_milestones_sheet(ws: Worksheet) -> None:
    ws.title = "Вехи"
    ws.freeze_panes = "A4"

    ws["A1"] = "Сводный список вех проекта"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:C1")
    ws["A2"] = "Контрольные точки М1–М10 — условия перехода между этапами."
    ws["A2"].font = Font(italic=True, color="666666")
    ws.merge_cells("A2:C2")

    headers = ["Веха", "Дата", "Содержание"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[3].height = 28

    for idx, (code, date, text) in enumerate(MILESTONES, start=4):
        ws.cell(row=idx, column=1, value=code).alignment = CENTER
        c = ws.cell(row=idx, column=2, value=date)
        c.number_format = "DD.MM.YYYY"
        c.alignment = CENTER
        ws.cell(row=idx, column=3, value=text).alignment = LEFT
        for col in range(1, 4):
            ws.cell(row=idx, column=col).fill = MILESTONE_FILL
            ws.cell(row=idx, column=col).border = BORDER
        ws.row_dimensions[idx].height = 22

    autosize(ws, {"A": 8, "B": 14, "C": 90})


# ---------------------------------------------------------------------------
# Sheet 3: Obligations
# ---------------------------------------------------------------------------


def build_obligations_sheet(ws: Worksheet) -> None:
    ws.title = "Обязательства"
    ws.freeze_panes = "A4"

    ws["A1"] = "Обязательства заказчика"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:C1")
    ws["A2"] = "Действия и материалы, которые предоставляет заказчик к контрольным точкам."
    ws["A2"].font = Font(italic=True, color="666666")
    ws.merge_cells("A2:C2")

    headers = ["Требование", "Срок", "Контрольная точка"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[3].height = 28

    for idx, (req, deadline, m) in enumerate(OBLIGATIONS, start=4):
        ws.cell(row=idx, column=1, value=req).alignment = LEFT
        c = ws.cell(row=idx, column=2, value=deadline)
        c.number_format = "DD.MM.YYYY"
        c.alignment = CENTER
        ws.cell(row=idx, column=3, value=m).alignment = CENTER
        for col in range(1, 4):
            ws.cell(row=idx, column=col).border = BORDER
            if idx % 2 == 0:
                ws.cell(row=idx, column=col).fill = ALT_FILL
        ws.row_dimensions[idx].height = 20

    autosize(ws, {"A": 70, "B": 14, "C": 18})


# ---------------------------------------------------------------------------
# Sheet 4: Gantt chart
# ---------------------------------------------------------------------------


def build_gantt_sheet(ws: Worksheet, plan_meta: dict) -> None:
    ws.title = "Диаграмма"

    ws["A1"] = "Диаграмма Ганта по задачам проекта"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:E1")
    ws["A2"] = "Источник данных — лист «План-график». Видимая полоса = длительность задачи."
    ws["A2"].font = Font(italic=True, color="666666")
    ws.merge_cells("A2:E2")

    headers = ["№", "Задача", "Старт (смещение, дни)", "Длительность (дни)", "Дата начала"]
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=i, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[3].height = 28

    project_start = min(t[3] for t in plan_meta["tasks"])

    last_row = 3
    for i, (excel_row, code, title, start, end, _) in enumerate(plan_meta["tasks"], start=4):
        offset = (start - project_start).days
        duration = max((end - start).days, 1)
        ws.cell(row=i, column=1, value=code).alignment = CENTER
        ws.cell(row=i, column=2, value=title).alignment = LEFT
        ws.cell(row=i, column=3, value=offset).alignment = CENTER
        ws.cell(row=i, column=4, value=duration).alignment = CENTER
        c = ws.cell(row=i, column=5, value=start)
        c.number_format = "DD.MM.YYYY"
        c.alignment = CENTER
        for col in range(1, 6):
            ws.cell(row=i, column=col).border = BORDER
            if i % 2 == 0:
                ws.cell(row=i, column=col).fill = ALT_FILL
        last_row = i

    autosize(ws, {"A": 8, "B": 70, "C": 22, "D": 22, "E": 16})

    # Build a stacked bar chart with two series: invisible offset + visible duration.
    chart = BarChart()
    chart.type = "bar"
    chart.style = 11
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Automacon WCS — диаграмма Ганта"
    chart.y_axis.title = None
    chart.x_axis.title = "Дни от старта проекта"
    chart.legend = None

    data_offset = Reference(ws, min_col=3, min_row=3, max_row=last_row, max_col=3)
    data_duration = Reference(ws, min_col=4, min_row=3, max_row=last_row, max_col=4)
    cats = Reference(ws, min_col=2, min_row=4, max_row=last_row)
    chart.add_data(data_offset, titles_from_data=True)
    chart.add_data(data_duration, titles_from_data=True)
    chart.set_categories(cats)

    # Hide the offset series — make the bar transparent and lineless.
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.fill import ColorChoice
    from openpyxl.drawing.line import LineProperties

    offset_series = chart.series[0]
    offset_series.graphicalProperties = GraphicalProperties(solidFill="FFFFFF")
    offset_series.graphicalProperties.line = LineProperties(noFill=True)
    offset_series.graphicalProperties.solidFill = "FFFFFF"

    duration_series = chart.series[1]
    duration_series.graphicalProperties = GraphicalProperties(solidFill="07C983")
    duration_series.graphicalProperties.line = LineProperties(solidFill="0D9373")

    # Reverse the categories so the first task is at the top (Excel default reverses bars).
    chart.y_axis.scaling.orientation = "maxMin"

    # Generous chart size. ~ width 30 cm × height (count*0.5+4) cm.
    n = len(plan_meta["tasks"])
    chart.width = 32
    chart.height = max(18, 0.55 * n + 6)

    ws.add_chart(chart, "G3")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main() -> None:
    wb = Workbook()
    plan_ws = wb.active
    plan_meta = build_plan_sheet(plan_ws)

    milestones_ws = wb.create_sheet("Вехи")
    build_milestones_sheet(milestones_ws)

    oblig_ws = wb.create_sheet("Обязательства")
    build_obligations_sheet(oblig_ws)

    gantt_ws = wb.create_sheet("Диаграмма")
    build_gantt_sheet(gantt_ws, plan_meta)

    # Reorder: Plan, Diagram, Milestones, Obligations
    wb._sheets = [plan_ws, gantt_ws, milestones_ws, oblig_ws]

    wb.save(OUT_PATH)
    print(f"Saved: {OUT_PATH} ({OUT_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
