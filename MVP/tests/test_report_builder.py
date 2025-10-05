"""Unit-тесты для модуля report_builder."""

from datetime import date
from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[2]))

import pytest

from MVP.report_builder import (
    DateRange,
    ReportConfig,
    auto_map_columns,
    build_report,
    build_total_header,
    FieldDefinition,
)


@pytest.fixture
def sample_config() -> ReportConfig:
    violations = [
        {
            "ID": "A1",
            "Статус": "На контроле инспектора ОАТИ",
            "Нарушение": "Неубранный снег",
            "Результат": "Нарушение выявлено",
            "Тип": "ОДХ",
            "Объект": "Объект А",
            "Дата": date(2024, 4, 10),
            "Округ": "Центральный административный округ",
        },
        {
            "ID": "B1",
            "Статус": "Снят с контроля",
            "Нарушение": "Неубранный снег",
            "Результат": "Нарушение выявлено",
            "Тип": "ОДХ",
            "Объект": "Объект Б",
            "Дата": date(2024, 4, 11),
            "Округ": "ЦАО",
        },
        {
            "ID": "A0",
            "Статус": "На контроле инспектора ОАТИ",
            "Нарушение": "Неубранный снег",
            "Результат": "Нарушение выявлено",
            "Тип": "ОДХ",
            "Объект": "Объект А",
            "Дата": date(2024, 3, 5),
            "Округ": "ЦАО",
        },
    ]
    objects = [
        {
            "Тип объекта": "ОДХ",
            "Наименование": "Объект А",
            "Округ": "Центральный административный округ",
        },
        {
            "Тип объекта": "ОДХ",
            "Наименование": "Объект Б",
            "Округ": "ЦАО",
        },
    ]
    violation_mapping = {
        "id": "ID",
        "status": "Статус",
        "violationName": "Нарушение",
        "inspectionResult": "Результат",
        "objectType": "Тип",
        "objectName": "Объект",
        "inspectionDate": "Дата",
        "district": "Округ",
    }
    object_mapping = {
        "objectType": "Тип объекта",
        "objectName": "Наименование",
        "district": "Округ",
    }
    return ReportConfig(
        violation_mapping=violation_mapping,
        object_mapping=object_mapping,
        current_period=DateRange(date(2024, 4, 1), date(2024, 4, 30)),
        previous_period=DateRange(date(2024, 3, 1), date(2024, 3, 31)),
    ), violations, objects


def test_build_report_basic(sample_config):
    config, violations, objects = sample_config
    report = build_report(violations, objects, config)
    assert len(report.rows) == 1
    row = report.rows[0]
    assert row.total_objects == 2
    assert row.inspected_objects == 2
    assert round(row.inspected_percent, 1) == 100.0
    assert round(row.violation_percent, 1) == 100.0
    assert row.total_violations == 3
    assert row.current_violations == 2
    assert row.previous_control == 1
    assert row.resolved == 1
    assert row.on_control == 1
    assert report.total_row.total_objects == 2
    assert report.total_row.total_violations == 3


def test_auto_map_columns():
    columns = ["ID", "Статус нарушения", "Дата обследования", "Произвольный"]
    mapping = auto_map_columns(columns, [])
    assert mapping == {}

    mapping = auto_map_columns(
        columns,
        [FieldDefinition(key="status", label="Статус", candidates=["статус нарушения"])],
    )
    assert mapping["status"] == "Статус нарушения"


def test_build_total_header() -> None:
    assert build_total_header("all", []) == "Всего ОДХ"
    assert build_total_header("custom", ["ОДХ"]) == "Всего ОДХ"
    assert (
        build_total_header("custom", ["МФЦ", "ОДХ"]) == "Всего (МФЦ, ОДХ)"
    )
