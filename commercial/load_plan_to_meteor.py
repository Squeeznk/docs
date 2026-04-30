"""Load Automacon WCS project plan into METEOR (OpenProject-compatible) tracker.

Run:
    METEOR_KEY=... python3 load_plan_to_meteor.py [--dry-run] [--rollback-from N]

Defaults:
    BASE          = https://avg.u-meteor.ru/op
    PROJECT_ID    = 326

The script creates a hierarchy of work packages:
  10 Epics (one per project stage) → child Tasks under each epic.

It is idempotent in the sense that it never reuses existing work packages —
on each run it creates a fresh batch. Use --rollback-from to delete previously
created work packages by ID range if needed.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from dataclasses import dataclass
from datetime import date
from typing import Any

import urllib.request
import urllib.error

DEFAULT_BASE = "https://avg.u-meteor.ru/op"
DEFAULT_PROJECT_ID = 326

# OpenProject-compatible type ids (taken from /api/v3/projects/326/types):
TYPE_EPIC = 5            # «Эпик»
TYPE_TASK = 1            # «Задача (dev)»
TYPE_IMPLEMENTATION = 41 # «Внедрение» — best fit for stages of an implementation
TYPE_BUG = 7             # «Ошибка»

PRIORITY_NORMAL = 8


@dataclass
class TaskSpec:
    subject: str
    start: date | None
    end: date | None
    description: str = ""


@dataclass
class EpicSpec:
    code: str
    subject: str
    start: date
    end: date
    description: str
    milestone: str
    children: list[TaskSpec]


# --------------------------------------------------------------------------- #
# Plan data — single source of truth, mirrors the XLSX/Markdown plan.        #
# --------------------------------------------------------------------------- #

PLAN: list[EpicSpec] = [
    EpicSpec(
        code="Этап 1",
        subject="Этап 1. Предпроектная стадия",
        start=date(2026, 4, 21),
        end=date(2026, 5, 22),
        milestone="М1",
        description=(
            "Цель: получить ответы по опросным листам от всех пяти подразделений "
            "заказчика, финализировать ТЗ и сценарии, согласовать тележку, утвердить "
            "устав и финальный план-график.\n\n"
            "Веха М1 (22.05.2026): предпроектная стадия закрыта."
        ),
        children=[
            TaskSpec("Получение ответов от подразделения «Охрана труда»", date(2026,4,30), date(2026,5,12)),
            TaskSpec("Получение ответов от подразделения «ИТ»", date(2026,4,30), date(2026,5,12)),
            TaskSpec("Получение ответов от подразделения «Информационная безопасность»", date(2026,4,30), date(2026,5,15)),
            TaskSpec("Получение ответов от подразделения «Производственная логистика и WMS Axelot»", date(2026,4,30), date(2026,5,12)),
            TaskSpec("Получение ответов от подразделения «Производственные участки»", date(2026,4,30), date(2026,5,15)),
            TaskSpec("Финализация сценариев работы парка и технического задания", date(2026,5,15), date(2026,5,22)),
            TaskSpec("Подготовка и передача документа «Требования к среде эксплуатации»", date(2026,5,4), date(2026,5,12)),
            TaskSpec("Проработка и согласование тележки под автономный транспортный робот", date(2026,4,23), date(2026,5,15)),
            TaskSpec("Подготовка и согласование устава проекта", date(2026,4,21), date(2026,5,12)),
            TaskSpec("Проработка альтернативной логистики", date(2026,4,28), date(2026,5,15)),
            TaskSpec("Финализация план-графика проекта", date(2026,5,18), date(2026,5,22)),
        ],
    ),
    EpicSpec(
        code="Этап 2",
        subject="Этап 2. Производство оборудования",
        start=date(2026, 5, 4),
        end=date(2026, 6, 15),
        milestone="М2",
        description=(
            "Цель: обеспечить готовность парка из 5 роботов и 3 зарядных станций "
            "к отгрузке с производственной площадки Automacon в Китае.\n\n"
            "Веха М2 (15.06.2026): оборудование готово к отгрузке."
        ),
        children=[
            TaskSpec("Оплата производственного заказа", date(2026,5,4), date(2026,5,6)),
            TaskSpec("Производственный цикл (5 роботов и 3 зарядные станции)", date(2026,5,6), date(2026,6,12)),
            TaskSpec("Заводские приёмо-сдаточные испытания", date(2026,6,8), date(2026,6,15)),
            TaskSpec("Подготовка отгрузочной документации", date(2026,6,12), date(2026,6,15)),
        ],
    ),
    EpicSpec(
        code="Этап 3",
        subject="Этап 3. Логистика и таможенное оформление",
        start=date(2026, 6, 15),
        end=date(2026, 7, 2),
        milestone="М3",
        description=(
            "Цель: доставить оборудование с производственной площадки на объект "
            "заказчика. Авиатранспорт. Резерв на таможне до 5 рабочих дней.\n\n"
            "Веха М3 (02.07.2026): оборудование на площадке заказчика."
        ),
        children=[
            TaskSpec("Подготовка отправки и оформление на авиатранспорт", date(2026,6,15), date(2026,6,18)),
            TaskSpec("Авиаперевозка", date(2026,6,18), date(2026,6,23)),
            TaskSpec("Таможенное оформление", date(2026,6,23), date(2026,6,30)),
            TaskSpec("Доставка на площадку заказчика", date(2026,6,30), date(2026,7,2)),
        ],
    ),
    EpicSpec(
        code="Этап 4",
        subject="Этап 4. Подготовка площадки заказчика",
        start=date(2026, 5, 11),
        end=date(2026, 6, 30),
        milestone="М4",
        description=(
            "Цель: обеспечить готовность площадки заказчика к приёмке оборудования "
            "и пусконаладочным работам. Контролируется руководителем проекта со стороны "
            "Automacon в части соответствия требованиям к среде эксплуатации.\n\n"
            "Веха М4 (30.06.2026): площадка готова к приёмке оборудования."
        ),
        children=[
            TaskSpec("Электротехнические работы под зарядные станции (заказчик)", date(2026,5,11), date(2026,6,19)),
            TaskSpec("Подготовка сетевой инфраструктуры: ВМ, VLAN, межсетевое экранирование (заказчик)", date(2026,5,11), date(2026,6,19)),
            TaskSpec("Беспроводная сеть в зонах работы парка (заказчик): Wi-Fi, радиоплан", date(2026,5,11), date(2026,6,19)),
            TaskSpec("Согласование маршрутов, точек REST, зон зарядки (технология/ОТ/диспетчер)", date(2026,5,18), date(2026,6,12)),
            TaskSpec("Выдача доступов команде Automacon к серверам и средствам разработки", date(2026,5,11), date(2026,5,22)),
            TaskSpec("Подготовка зон работы (чистота, разметка, удаление посторонних предметов)", date(2026,5,25), date(2026,6,26)),
            TaskSpec("Согласование ИБ-регламентов и выпуск технологических учётных записей", date(2026,5,18), date(2026,6,19)),
        ],
    ),
    EpicSpec(
        code="Этап 5",
        subject="Этап 5. Развёртывание серверной части WCS",
        start=date(2026, 5, 22),
        end=date(2026, 7, 3),
        milestone="М5",
        description=(
            "Цель: установить и настроить платформу Automacon WCS в инфраструктуре "
            "заказчика, выполнить интеграцию с WMS Axelot и с корпоративной системой "
            "управления учётными записями.\n\n"
            "Веха М5 (03.07.2026): серверная часть WCS установлена и интегрирована."
        ),
        children=[
            TaskSpec("Получение виртуальных машин и сетевых доступов", date(2026,5,22), date(2026,5,29)),
            TaskSpec("Установка платформы в тестовом контуре, базовая настройка", date(2026,5,29), date(2026,6,5)),
            TaskSpec("Интеграция с корпоративной аутентификацией и настройка ролевой модели", date(2026,6,5), date(2026,6,12)),
            TaskSpec("Интеграция с WMS Axelot — реализация коннектора, отладка контракта", date(2026,6,5), date(2026,6,26)),
            TaskSpec("Развёртывание контура мониторинга (метрики, журналы, дашборды)", date(2026,6,12), date(2026,6,26)),
            TaskSpec("Установка платформы в промышленном контуре, регрессионная проверка", date(2026,6,26), date(2026,7,3)),
            TaskSpec("Подготовка к приёмке роботов: справочники, шаблоны устройств", date(2026,6,26), date(2026,7,3)),
        ],
    ),
    EpicSpec(
        code="Этап 6",
        subject="Этап 6. Приёмка, калибровка и подключение оборудования",
        start=date(2026, 7, 2),
        end=date(2026, 7, 24),
        milestone="М6",
        description=(
            "Цель: ввести в эксплуатацию парк роботов и зарядные станции, подключить "
            "их к платформе, выполнить калибровку и проверку безопасности.\n\n"
            "Веха М6 (24.07.2026): парк подключён к WCS, прошёл проверку безопасности."
        ),
        children=[
            TaskSpec("Распаковка, входной контроль, оформление приёмочных актов", date(2026,7,2), date(2026,7,6)),
            TaskSpec("Монтаж и пусконаладка зарядных станций", date(2026,7,6), date(2026,7,13)),
            TaskSpec("Первичная настройка и калибровка роботов", date(2026,7,6), date(2026,7,24)),
            TaskSpec("Регистрация роботов в WCS, технологические учётные записи", date(2026,7,13), date(2026,7,24)),
            TaskSpec("Подключение периферии к WCS (зарядные станции, ворота, конвейеры)", date(2026,7,13), date(2026,7,24)),
            TaskSpec("Проверки безопасности (аварийный останов, режимы остановки)", date(2026,7,20), date(2026,7,24)),
        ],
    ),
    EpicSpec(
        code="Этап 7",
        subject="Этап 7. Картография",
        start=date(2026, 7, 24),
        end=date(2026, 7, 31),
        milestone="М7",
        description=(
            "Цель: подготовить и утвердить лидарную карту цехов и логическую карту "
            "дорог в формате LIF, загрузить карту в платформу.\n\n"
            "Веха М7 (31.07.2026): карта в формате LIF утверждена и загружена."
        ),
        children=[
            TaskSpec("Лидарное сканирование зон работы", date(2026,7,24), date(2026,7,27)),
            TaskSpec("Подготовка карты в формате LIF (узлы, рёбра, точки, станции)", date(2026,7,27), date(2026,7,30)),
            TaskSpec("Валидация графа карты, согласование с технологией и охраной труда", date(2026,7,30), date(2026,7,31)),
            TaskSpec("Загрузка карты в WCS, контрольный прогон по основным маршрутам", date(2026,7,30), date(2026,7,31)),
        ],
    ),
    EpicSpec(
        code="Этап 8",
        subject="Этап 8. Отладка сценариев и обучение персонала",
        start=date(2026, 7, 31),
        end=date(2026, 8, 28),
        milestone="М8",
        description=(
            "Цель: последовательно ввести в эксплуатацию все сценарии работы парка, "
            "отладить взаимодействие с WMS Axelot и периферией, обучить персонал.\n\n"
            "Веха М8 (28.08.2026): сценарии отлажены, персонал обучен."
        ),
        children=[
            TaskSpec("Поочередная отладка сценариев работы парка", date(2026,7,31), date(2026,8,21)),
            TaskSpec("Тестирование точек погрузки и разгрузки, точность позиционирования", date(2026,8,10), date(2026,8,24)),
            TaskSpec("Тестирование сценариев исключений", date(2026,8,17), date(2026,8,28)),
            TaskSpec("Обучение операторов и диспетчеров заказчика", date(2026,8,17), date(2026,8,28)),
            TaskSpec("Инструктажи под подпись персонала участков (охрана труда)", date(2026,8,24), date(2026,8,28)),
            TaskSpec("Передача эксплуатационной документации", date(2026,8,24), date(2026,8,28)),
        ],
    ),
    EpicSpec(
        code="Этап 9",
        subject="Этап 9. Опытная эксплуатация",
        start=date(2026, 8, 31),
        end=date(2026, 9, 13),
        milestone="М9",
        description=(
            "Цель: проверить работу комплекса на реальном производственном цикле "
            "в течение двух недель, собрать и закрыть замечания.\n\n"
            "Веха М9 (13.09.2026): опытная эксплуатация завершена."
        ),
        children=[
            TaskSpec("Запуск опытной эксплуатации, дежурство команды Automacon на площадке", date(2026,8,31), date(2026,9,11)),
            TaskSpec("Журналирование инцидентов и замечаний, ежедневный разбор", date(2026,8,31), date(2026,9,13)),
            TaskSpec("Закрытие замечаний — корректировка сценариев, настроек, документации", date(2026,9,3), date(2026,9,13)),
            TaskSpec("Подведение итогов опытной эксплуатации совместно с заказчиком", date(2026,9,11), date(2026,9,13)),
        ],
    ),
    EpicSpec(
        code="Этап 10",
        subject="Этап 10. ПСИ и передача в промышленную эксплуатацию",
        start=date(2026, 9, 14),
        end=date(2026, 9, 27),
        milestone="М10",
        description=(
            "Цель: провести формальные испытания, подписать акт приёмки, передать "
            "комплекс на промышленную эксплуатацию и сопровождение.\n\n"
            "Веха М10 (27.09.2026): промышленная эксплуатация и сопровождение."
        ),
        children=[
            TaskSpec("Подготовка программы и методики приёмо-сдаточных испытаний", date(2026,9,14), date(2026,9,16)),
            TaskSpec("Проведение приёмо-сдаточных испытаний по согласованной программе", date(2026,9,16), date(2026,9,23)),
            TaskSpec("Подготовка и подписание акта приёмки", date(2026,9,23), date(2026,9,25)),
            TaskSpec("Передача на регламентное сопровождение, активация SLA", date(2026,9,25), date(2026,9,27)),
        ],
    ),
]


# --------------------------------------------------------------------------- #
# HTTP helpers
# --------------------------------------------------------------------------- #


def http_request(
    method: str,
    url: str,
    api_key: str,
    payload: dict | None = None,
    timeout: int = 20,
) -> dict[str, Any]:
    body = None
    headers = {
        "Accept": "application/json",
    }
    if payload is not None:
        body = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"

    # OpenProject-compatible Basic auth: "apikey:<token>"
    import base64

    auth = base64.b64encode(f"apikey:{api_key}".encode("ascii")).decode("ascii")
    headers["Authorization"] = f"Basic {auth}"

    req = urllib.request.Request(url=url, method=method, data=body, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            content = resp.read().decode("utf-8")
            return {"status": resp.status, "body": json.loads(content) if content else None}
    except urllib.error.HTTPError as e:
        try:
            err_body = json.loads(e.read().decode("utf-8"))
        except Exception:
            err_body = None
        return {"status": e.code, "body": err_body, "error": str(e)}


# --------------------------------------------------------------------------- #
# Work-package payload helpers
# --------------------------------------------------------------------------- #


def epic_payload(epic: EpicSpec) -> dict:
    return {
        "subject": f"[{epic.code}] {epic.subject[len(epic.code) + 2:]}" if epic.subject.startswith(epic.code + ".") else epic.subject,
        "description": {"format": "markdown", "raw": epic.description},
        "startDate": epic.start.isoformat(),
        "dueDate": epic.end.isoformat(),
        "_links": {
            "type": {"href": f"/api/v3/types/{TYPE_EPIC}"},
            "priority": {"href": f"/api/v3/priorities/{PRIORITY_NORMAL}"},
        },
    }


def task_payload(task: TaskSpec, parent_id: int) -> dict:
    body: dict[str, Any] = {
        "subject": task.subject,
        "_links": {
            "type": {"href": f"/api/v3/types/{TYPE_TASK}"},
            "priority": {"href": f"/api/v3/priorities/{PRIORITY_NORMAL}"},
            "parent": {"href": f"/api/v3/work_packages/{parent_id}"},
        },
    }
    if task.start:
        body["startDate"] = task.start.isoformat()
    if task.end:
        body["dueDate"] = task.end.isoformat()
    if task.description:
        body["description"] = {"format": "markdown", "raw": task.description}
    return body


# --------------------------------------------------------------------------- #
# Driver
# --------------------------------------------------------------------------- #


def run(
    base_url: str,
    project_id: int,
    api_key: str,
    *,
    dry_run: bool,
    pause_seconds: float = 0.15,
) -> int:
    print(f"Base URL:    {base_url}")
    print(f"Project ID:  {project_id}")
    print(f"Dry-run:     {dry_run}")

    # Verify access
    me = http_request("GET", f"{base_url}/api/v3/users/me", api_key)
    if me["status"] != 200:
        print("Authentication failed:", me)
        return 2
    print(f"Authenticated as: {me['body'].get('name')} (id={me['body'].get('id')})")

    proj = http_request("GET", f"{base_url}/api/v3/projects/{project_id}", api_key)
    if proj["status"] != 200:
        print("Project lookup failed:", proj)
        return 2
    print(f"Project: {proj['body'].get('name')} (identifier={proj['body'].get('identifier')})")

    create_endpoint = f"{base_url}/api/v3/projects/{project_id}/work_packages"
    update_endpoint_template = f"{base_url}/api/v3/work_packages/{{id}}"

    created_summary: list[tuple[str, int, str]] = []
    failed: list[tuple[str, dict]] = []

    for epic in PLAN:
        if dry_run:
            print(f"\n[DRY-RUN] Would create epic: {epic.subject}")
            for t in epic.children:
                print(f"           ↳ Task: {t.subject}")
            continue

        # 1. Create epic.
        result = http_request("POST", create_endpoint, api_key, payload=epic_payload(epic))
        if result["status"] not in (200, 201):
            print(f"FAIL epic «{epic.subject}»: {result}")
            failed.append((epic.subject, result))
            continue
        epic_id = result["body"]["id"]
        epic_subject = result["body"]["subject"]
        print(f"✓ Epic #{epic_id}: {epic_subject}")
        created_summary.append((epic.code, epic_id, epic_subject))

        # 2. Create children.
        for child in epic.children:
            child_result = http_request(
                "POST", create_endpoint, api_key, payload=task_payload(child, epic_id)
            )
            if child_result["status"] not in (200, 201):
                print(f"  ✗ child «{child.subject}»: {child_result}")
                failed.append((child.subject, child_result))
                continue
            child_id = child_result["body"]["id"]
            print(f"  ↳ #{child_id} {child.subject}")
            created_summary.append(("  ", child_id, child.subject))
            time.sleep(pause_seconds)
        time.sleep(pause_seconds)

    print()
    print("=" * 70)
    print(f"Created total:   {len(created_summary)}")
    print(f"Failed total:    {len(failed)}")
    if failed:
        print("Errors:")
        for subj, err in failed:
            print(f"  - {subj}: status={err.get('status')}")
    return 0 if not failed else 3


def main() -> int:
    parser = argparse.ArgumentParser(description="Load Automacon WCS plan into METEOR.")
    parser.add_argument("--base", default=DEFAULT_BASE)
    parser.add_argument("--project", type=int, default=DEFAULT_PROJECT_ID)
    parser.add_argument("--key", default=os.environ.get("METEOR_KEY"))
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    if not args.key:
        print("ERROR: METEOR API key is not provided. Pass --key or set METEOR_KEY.", file=sys.stderr)
        return 2
    return run(args.base, args.project, args.key, dry_run=args.dry_run)


if __name__ == "__main__":
    sys.exit(main())
