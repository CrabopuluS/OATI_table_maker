"""Core report-building logic for the desktop MVP.

This module mirrors the calculation rules from the web version of the
OATI table builder. It is deliberately framework-agnostic so it can be
unit-tested without the UI layer.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

CONTROL_STATUSES = {"на устранении", "на контроле инспектора оати"}
RESOLVED_STATUS = "снят с контроля"
INSPECTION_RESULT_VIOLATION = "нарушение выявлено"

MOSCOW_DISTRICT_ABBREVIATIONS = {
    "центральный административный округ": "ЦАО",
    "северный административный округ": "САО",
    "северо-восточный административный округ": "СВАО",
    "восточный административный округ": "ВАО",
    "юго-восточный административный округ": "ЮВАО",
    "южный административный округ": "ЮАО",
    "юго-западный административный округ": "ЮЗАО",
    "западный административный округ": "ЗАО",
    "северо-западный административный округ": "СЗАО",
    "зеленоградский административный округ": "ЗелАО",
    "новомосковский административный округ": "НАО",
    "троицкий административный округ": "ТАО",
    "троицкий и новомосковский административный округ": "ТиНАО",
}

MOSCOW_DISTRICT_ABBREVIATION_KEYS = {
    value.strip().lower() for value in MOSCOW_DISTRICT_ABBREVIATIONS.values()
}
DISTRICT_FALLBACK_KEY = "без округа"


@dataclass(frozen=True)
class FieldDefinition:
    """Description of a logical field that should be mapped to a column."""

    key: str
    label: str
    candidates: Sequence[str]
    optional: bool = False


@dataclass(frozen=True)
class DateRange:
    """Inclusive date range."""

    start: date
    end: date

    def contains(self, value: date) -> bool:
        return self.start <= value <= self.end


@dataclass
class ReportRow:
    label: str
    total_objects: int
    inspected_objects: int
    inspected_percent: float
    violation_percent: float
    total_violations: int
    current_violations: int
    previous_control: int
    resolved: int
    on_control: int


@dataclass
class ReportResult:
    rows: List[ReportRow]
    total_row: ReportRow


@dataclass
class ReportConfig:
    violation_mapping: Dict[str, str]
    object_mapping: Dict[str, str]
    current_period: DateRange
    previous_period: DateRange
    type_mode: str = "all"
    selected_types: Optional[Sequence[str]] = None
    violation_mode: str = "all"
    selected_violations: Optional[Sequence[str]] = None
    data_source_mode: str = "all"


VIOLATION_FIELD_DEFINITIONS: Sequence[FieldDefinition] = (
    FieldDefinition("id", "Идентификатор нарушения", ("идентификатор", "id", "uid")),
    FieldDefinition("status", "Статус нарушения", ("статус нарушения", "статус")),
    FieldDefinition(
        "violationName",
        "Наименование нарушения",
        ("наименование нарушения", "нарушение"),
    ),
    FieldDefinition(
        "inspectionResult",
        "Результат обследования",
        ("результат обследования", "результат осмотра", "результат проверки"),
        optional=True,
    ),
    FieldDefinition(
        "objectType",
        "Тип объекта",
        ("тип объекта", "тип объекта контроля"),
    ),
    FieldDefinition(
        "objectName",
        "Наименование объекта",
        ("наименование объекта", "наименование объекта контроля"),
    ),
    FieldDefinition(
        "inspectionDate",
        "Дата обследования",
        ("дата обследования", "дата осмотра", "дата контроля"),
    ),
    FieldDefinition(
        "district",
        "Округ",
        ("округ", "административный округ", "округ объекта"),
    ),
    FieldDefinition(
        "dataSource",
        "Источник данных",
        ("источник данных",),
        optional=True,
    ),
)


OBJECT_FIELD_DEFINITIONS: Sequence[FieldDefinition] = (
    FieldDefinition(
        "objectType",
        "Вид объекта",
        ("вид объекта", "тип объекта", "тип объекта контроля"),
    ),
    FieldDefinition(
        "objectName",
        "Наименование объекта",
        ("наименование объекта", "наименование объекта контроля"),
    ),
    FieldDefinition(
        "district",
        "Округ",
        ("округ", "административный округ", "округ объекта"),
    ),
)


def normalize_key(value: Optional[str]) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def get_value_as_string(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return str(value).strip()


def parse_date_value(value: object) -> Optional[date]:
    if value in {None, ""}:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, (int, float)):
        # Excel stores dates as days since 1899-12-30.
        excel_start = date(1899, 12, 30)
        try:
            return excel_start + timedelta(days=float(value))
        except Exception:  # pylint: disable=broad-except
            return None
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            return datetime.strptime(text, "%d.%m.%Y").date()
        except ValueError:
            pass
        try:
            parsed = datetime.fromisoformat(text)
            return parsed.date()
        except ValueError:
            pass
        # Try flexible parsing for strings like "2024-03-15T00:00:00"
        try:
            parsed = datetime.strptime(text, "%Y-%m-%d %H:%M:%S")
            return parsed.date()
        except ValueError:
            pass
    return None


# Needed for parse_date_value when Excel serial numbers appear.
from datetime import timedelta  # noqa: E402  (import after function definitions)


def compute_percent(part: int, total: int) -> float:
    if not total:
        return 0.0
    return (part / total) * 100.0


def auto_map_columns(columns: Sequence[str], definitions: Sequence[FieldDefinition]) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    normalized_columns = {normalize_key(column): column for column in columns if column}
    for definition in definitions:
        matched = ""
        for candidate in definition.candidates:
            normalized_candidate = normalize_key(candidate)
            if normalized_candidate in normalized_columns:
                matched = normalized_columns[normalized_candidate]
                break
        mapping[definition.key] = matched
    return mapping


def ensure_required_columns(mapping: Dict[str, str], definitions: Sequence[FieldDefinition]) -> None:
    missing = [
        definition.label
        for definition in definitions
        if not definition.optional and not mapping.get(definition.key)
    ]
    if missing:
        raise ValueError(
            "Не все обязательные поля сопоставлены: " + ", ".join(missing)
        )


def build_district_lookup(
    object_records: Iterable[Dict[str, object]],
    object_district_column: Optional[str],
    violation_records: Iterable[Dict[str, object]],
    violation_district_column: Optional[str],
) -> Dict[str, str]:
    lookup: Dict[str, str] = {}

    def register_value(raw_value: object) -> None:
        label = get_value_as_string(raw_value)
        normalized = normalize_district_key(label)
        key = normalized or DISTRICT_FALLBACK_KEY
        if not key:
            return
        display_label = get_district_display_label(label) if normalized else "Без округа"
        if key not in lookup:
            lookup[key] = display_label
            return
        if should_replace_district_label(lookup[key], display_label):
            lookup[key] = display_label

    def register_from_records(
        records: Iterable[Dict[str, object]], column: Optional[str]
    ) -> None:
        if not column:
            return
        for record in records:
            register_value(record.get(column, ""))

    register_from_records(object_records, object_district_column)
    register_from_records(violation_records, violation_district_column)
    return lookup


def normalize_district_key(value: str) -> str:
    normalized = normalize_key(value)
    if not normalized:
        return ""
    if normalized in MOSCOW_DISTRICT_ABBREVIATION_KEYS:
        return normalized
    return MOSCOW_DISTRICT_ABBREVIATIONS.get(normalized, normalized)


def get_district_display_label(value: str) -> str:
    normalized = normalize_key(value)
    if normalized in MOSCOW_DISTRICT_ABBREVIATION_KEYS:
        return value.strip()
    if normalized in MOSCOW_DISTRICT_ABBREVIATIONS:
        return MOSCOW_DISTRICT_ABBREVIATIONS[normalized]
    return value.strip() or "Без округа"


def should_replace_district_label(current: str, candidate: str) -> bool:
    if not candidate or candidate == current:
        return False
    if not current:
        return True
    current_key = normalize_district_key(current)
    candidate_key = normalize_district_key(candidate)
    current_is_abbr = current_key in MOSCOW_DISTRICT_ABBREVIATION_KEYS
    candidate_is_abbr = candidate_key in MOSCOW_DISTRICT_ABBREVIATION_KEYS
    if candidate_is_abbr and not current_is_abbr:
        return True
    if not candidate_is_abbr and current_is_abbr:
        return False
    return len(candidate) < len(current)


def categorize_data_source(value: object) -> str:
    text = get_value_as_string(value)
    normalized = normalize_key(text)
    if not normalized:
        return ""
    has_oati = "оати" in normalized
    has_cafap = "цафап" in normalized
    if has_oati and has_cafap:
        return "both"
    if has_oati:
        return "oati"
    if has_cafap:
        return "cafap"
    return ""


class RecordFilter:
    """Precomputed filters for types, violations and data sources."""

    def __init__(self, config: ReportConfig) -> None:
        self.type_mode = config.type_mode
        self.violation_mode = config.violation_mode
        self.data_source_mode = config.data_source_mode

        self.allowed_types = {
            normalize_key(value)
            for value in (config.selected_types or [])
            if normalize_key(value)
        }
        self.allowed_violations = {
            normalize_key(value)
            for value in (config.selected_violations or [])
            if normalize_key(value)
        }

    def allow_type(self, value: str) -> bool:
        if self.type_mode != "custom" or not self.allowed_types:
            return True
        normalized = normalize_key(value)
        if not normalized:
            return False
        return normalized in self.allowed_types

    def allow_violation(self, value: str) -> bool:
        if self.violation_mode != "custom" or not self.allowed_violations:
            return True
        normalized = normalize_key(value)
        if not normalized:
            return False
        return normalized in self.allowed_violations

    def allow_data_source(self, record: Dict[str, object], column: Optional[str]) -> bool:
        if not column:
            return True
        category = categorize_data_source(record.get(column, ""))
        if not category:
            return self.data_source_mode == "all"
        if self.data_source_mode == "oati":
            return category == "oati"
        if self.data_source_mode == "cafap":
            return category == "cafap"
        return True


@dataclass
class TotalsAccumulator:
    total_objects: int = 0
    inspected_objects: int = 0
    objects_with_detected_violations: int = 0
    total_violations: set = None  # type: ignore[assignment]
    current_violations: set = None  # type: ignore[assignment]
    previous_control: set = None  # type: ignore[assignment]
    resolved: set = None  # type: ignore[assignment]
    on_control: set = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        self.total_violations = set()
        self.current_violations = set()
        self.previous_control = set()
        self.resolved = set()
        self.on_control = set()


def build_report(
    violations: Sequence[Dict[str, object]],
    objects: Sequence[Dict[str, object]],
    config: ReportConfig,
) -> ReportResult:
    ensure_required_columns(config.violation_mapping, VIOLATION_FIELD_DEFINITIONS)
    ensure_required_columns(config.object_mapping, OBJECT_FIELD_DEFINITIONS)

    filters = RecordFilter(config)
    violation_mapping = config.violation_mapping
    object_mapping = config.object_mapping

    district_data: Dict[str, Dict[str, object]] = {}
    district_lookup = build_district_lookup(
        objects,
        object_mapping.get("district"),
        violations,
        violation_mapping.get("district"),
    )

    def ensure_entry(raw_label: str) -> Dict[str, object]:
        normalized = normalize_district_key(raw_label)
        key = normalized or DISTRICT_FALLBACK_KEY
        display_label = (
            district_lookup.get(key)
            or (get_district_display_label(raw_label) if normalized else "Без округа")
        )
        aggregated_key = normalize_key(display_label) or DISTRICT_FALLBACK_KEY
        if aggregated_key not in district_data:
            district_data[aggregated_key] = {
                "label": display_label,
                "total_objects": set(),
                "inspected_objects": set(),
                "objects_with_violations": set(),
                "current_violation_ids": set(),
                "previous_control_ids": set(),
                "resolved_ids": set(),
                "control_ids": set(),
            }
        else:
            entry = district_data[aggregated_key]
            if should_replace_district_label(entry["label"], display_label):
                entry["label"] = display_label
        return district_data[aggregated_key]

    allowed_violation_objects = set()
    if config.violation_mode == "custom" and config.selected_violations:
        violation_name_column = violation_mapping.get("violationName")
        violation_object_column = violation_mapping.get("objectName")
        if violation_name_column and violation_object_column:
            for record in violations:
                if not filters.allow_data_source(record, violation_mapping.get("dataSource")):
                    continue
                violation_name = get_value_as_string(record.get(violation_name_column, ""))
                if not violation_name or not filters.allow_violation(violation_name):
                    continue
                related_object = get_value_as_string(record.get(violation_object_column, ""))
                if related_object:
                    allowed_violation_objects.add(normalize_key(related_object))

    for record in objects:
        object_type = get_value_as_string(record.get(object_mapping.get("objectType", ""), ""))
        if object_type and not filters.allow_type(object_type):
            continue
        if not object_type and config.type_mode == "custom":
            continue
        district_label = get_value_as_string(record.get(object_mapping.get("district", ""), ""))
        object_name = get_value_as_string(record.get(object_mapping.get("objectName", ""), ""))
        if not object_name:
            continue
        if config.violation_mode == "custom" and allowed_violation_objects:
            if normalize_key(object_name) not in allowed_violation_objects:
                continue
        entry = ensure_entry(district_label)
        entry["total_objects"].add(object_name)

    inspection_result_column = violation_mapping.get("inspectionResult")

    for record in violations:
        if not filters.allow_data_source(record, violation_mapping.get("dataSource")):
            continue
        type_value = get_value_as_string(record.get(violation_mapping.get("objectType", ""), ""))
        if type_value and not filters.allow_type(type_value):
            continue
        if not type_value and config.type_mode == "custom":
            continue
        district_label = get_value_as_string(record.get(violation_mapping.get("district", ""), ""))
        object_name = get_value_as_string(record.get(violation_mapping.get("objectName", ""), ""))
        violation_id = get_value_as_string(record.get(violation_mapping.get("id", ""), ""))
        status_value = get_value_as_string(record.get(violation_mapping.get("status", ""), ""))
        violation_name = get_value_as_string(record.get(violation_mapping.get("violationName", ""), ""))
        if violation_name and not filters.allow_violation(violation_name):
            continue
        if not violation_name and config.violation_mode == "custom":
            continue
        normalized_status = normalize_key(status_value)
        inspection_date = parse_date_value(record.get(violation_mapping.get("inspectionDate", ""), ""))
        if not inspection_date:
            continue
        entry = ensure_entry(district_label)
        inspection_result_value = (
            get_value_as_string(record.get(inspection_result_column, ""))
            if inspection_result_column
            else ""
        )
        normalized_inspection_result = normalize_key(inspection_result_value)

        if config.current_period.contains(inspection_date):
            if object_name:
                entry["inspected_objects"].add(object_name)
            if (
                inspection_result_column
                and object_name
                and normalized_inspection_result == INSPECTION_RESULT_VIOLATION
            ):
                entry["objects_with_violations"].add(object_name)
            if violation_id:
                entry["current_violation_ids"].add(violation_id)
            if normalized_status == normalize_key(RESOLVED_STATUS) and violation_id:
                entry["resolved_ids"].add(violation_id)
            if violation_id and normalized_status in CONTROL_STATUSES:
                entry["control_ids"].add(violation_id)

        if config.previous_period.contains(inspection_date):
            if normalized_status in CONTROL_STATUSES and violation_id:
                entry["previous_control_ids"].add(violation_id)

    totals = TotalsAccumulator()
    rows: List[ReportRow] = []
    sorted_entries = sorted(
        district_data.values(), key=lambda entry: entry["label"].lower()
    )
    for entry in sorted_entries:
        total_objects_count = len(entry["total_objects"])
        inspected_count = len(entry["inspected_objects"])
        violation_objects_count = len(entry["objects_with_violations"])
        current_violations_count = len(entry["current_violation_ids"])
        previous_control_count = len(entry["previous_control_ids"])
        resolved_count = len(entry["resolved_ids"])
        on_control_count = len(entry["control_ids"])
        total_violations_count = len(
            entry["current_violation_ids"].union(entry["previous_control_ids"])
        )

        row = ReportRow(
            label=entry["label"],
            total_objects=total_objects_count,
            inspected_objects=inspected_count,
            inspected_percent=compute_percent(inspected_count, total_objects_count),
            violation_percent=compute_percent(violation_objects_count, inspected_count),
            total_violations=total_violations_count,
            current_violations=current_violations_count,
            previous_control=previous_control_count,
            resolved=resolved_count,
            on_control=on_control_count,
        )
        rows.append(row)

        totals.total_objects += total_objects_count
        totals.inspected_objects += inspected_count
        totals.objects_with_detected_violations += violation_objects_count
        totals.current_violations.update(entry["current_violation_ids"])
        totals.previous_control.update(entry["previous_control_ids"])
        totals.resolved.update(entry["resolved_ids"])
        totals.on_control.update(entry["control_ids"])
        totals.total_violations.update(entry["current_violation_ids"])
        totals.total_violations.update(entry["previous_control_ids"])

    total_row = ReportRow(
        label="ИТОГО",
        total_objects=totals.total_objects,
        inspected_objects=totals.inspected_objects,
        inspected_percent=compute_percent(
            totals.inspected_objects, totals.total_objects
        ),
        violation_percent=compute_percent(
            totals.objects_with_detected_violations, totals.inspected_objects
        ),
        total_violations=len(totals.total_violations),
        current_violations=len(totals.current_violations),
        previous_control=len(totals.previous_control),
        resolved=len(totals.resolved),
        on_control=len(totals.on_control),
    )

    return ReportResult(rows=rows, total_row=total_row)


def extract_unique_dates(
    records: Sequence[Dict[str, object]], column: Optional[str]
) -> List[date]:
    if not column:
        return []
    seen = set()
    dates: List[date] = []
    for record in records:
        parsed = parse_date_value(record.get(column, ""))
        if not parsed:
            continue
        iso = parsed.isoformat()
        if iso in seen:
            continue
        seen.add(iso)
        dates.append(parsed)
    dates.sort()
    return dates


def collect_unique_values(
    records: Sequence[Dict[str, object]], column: Optional[str]
) -> List[str]:
    if not column:
        return []
    seen: Dict[str, str] = {}
    for record in records:
        value = get_value_as_string(record.get(column, ""))
        key = normalize_key(value)
        if not key or key in seen:
            continue
        seen[key] = value
    return list(seen.values())


def format_integer(value: int) -> str:
    return f"{value:,}".replace(",", " ")


def format_percent(value: float) -> str:
    return f"{value:.1f}%" if value else "0%"


def describe_data_source(mode: str) -> str:
    if mode == "oati":
        return "накопленные только ОАТИ"
    if mode == "cafap":
        return "накопленные только ЦАФАП"
    return "выявленные ОАТИ и ЦАФАП"


def build_total_header(type_mode: str, selected_types: Sequence[str] | None) -> str:
    if type_mode != "custom" or not selected_types:
        return "Всего ОДХ"
    if len(selected_types) == 1:
        return f"Всего {selected_types[0]}"
    return "Всего (" + ", ".join(selected_types) + ")"


__all__ = [
    "FieldDefinition",
    "DateRange",
    "ReportConfig",
    "ReportResult",
    "ReportRow",
    "VIOLATION_FIELD_DEFINITIONS",
    "OBJECT_FIELD_DEFINITIONS",
    "auto_map_columns",
    "build_report",
    "extract_unique_dates",
    "collect_unique_values",
    "format_integer",
    "format_percent",
    "describe_data_source",
    "build_total_header",
]
