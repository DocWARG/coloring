from __future__ import annotations

from io import BytesIO
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


TITLE = "Лист испытаний"
MIXING_LABEL = "Смешение отвердителя на 100 частей основы"
MAIN_SECTION = "Основной состав"
HARDENER_SECTION = "Отвердитель"
EXPERIMENTS_SECTION = "Эксперименты"
EXPERIMENT_FIELDS = ("Атрибут", "Материал", "Метод эксперимента", "Стандарт", "Норма")


def _variant_value(values: Any, variant_id: int) -> Any:
    if not isinstance(values, dict):
        return 0
    return values.get(str(variant_id), values.get(variant_id, 0))


def _append_variant_values(sheet: Any, label: str, variants: list[dict[str, Any]], values: Any) -> None:
    sheet.append([label, *[_variant_value(values, variant["id"]) for variant in variants]])


def _style_sheet(sheet: Any, variant_count: int) -> None:
    last_column = max(6, variant_count + 5)

    sheet.sheet_view.showGridLines = True
    sheet.column_dimensions["A"].width = 42
    for column in range(2, last_column + 1):
        sheet.column_dimensions[get_column_letter(column)].width = 22
    sheet.freeze_panes = "B10"

    for row in sheet.iter_rows():
        first_value = row[0].value if row else None
        if first_value == TITLE:
            for cell in row:
                cell.font = Font(bold=True, color="245C43", size=15)
        elif first_value in {MAIN_SECTION, HARDENER_SECTION, EXPERIMENTS_SECTION, "Варианты"}:
            for cell in row:
                cell.font = Font(bold=True, color="245C43")
        elif first_value in {"Вариант", "Компонент", "Компонент отвердителя", *EXPERIMENT_FIELDS}:
            for cell in row:
                cell.font = Font(bold=True, color="4C594C")
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0


def workbook_bytes(data: dict[str, Any]) -> bytes:
    variants = data.get("variants") or [{"id": 1, "name": "Вариант № 1"}]
    variant_names = [variant.get("name") or f"Вариант № {index + 1}" for index, variant in enumerate(variants)]
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Экспериментальные данные"

    sheet.append([TITLE])
    product = data.get("product") or {}
    conditions = data.get("testConditions") or {}
    thickness = conditions.get("dryFilmThicknessMicrometers") or {}
    curing = conditions.get("curingBeforeTest") or {}
    sheet.append(["Название продукта", product.get("name", 0)])
    sheet.append(["Назначение продукта", product.get("purpose", 0)])
    sheet.append(["Подготовка поверхности", conditions.get("surfacePreparation", 0)])
    sheet.append(["Толщина сухой пленки, мкм", thickness.get("min", 0), thickness.get("max", 0)])
    sheet.append(["Отверждение до испытания", "Режим", curing.get("mode", 0), "Срок", curing.get("duration", 0)])
    sheet.append([])

    sheet.append(["Варианты"])
    sheet.append(["Вариант", *variant_names])
    sheet.append([])

    _write_composition(
        sheet,
        MAIN_SECTION,
        "Компонент",
        data.get("mainComposition") or {},
        variants,
        "Итог по основному составу",
        experiment_values={
            variant["id"]: variant.get("mainCompositionExperiment", 0)
            for variant in variants
        },
    )
    sheet.append([])

    hardener = data.get("hardener") or {}
    _write_composition(
        sheet,
        HARDENER_SECTION,
        "Компонент отвердителя",
        hardener,
        variants,
        "Итог по отвердителю",
        mixing_values=hardener.get("mixingRatioPer100PartsBaseByVariant", {}),
    )
    sheet.append([])

    sheet.append([EXPERIMENTS_SECTION])
    sheet.append([*EXPERIMENT_FIELDS, *[f"Показатели: {name}" for name in variant_names]])
    for experiment in data.get("experiments", []):
        indicators = experiment.get("indicatorsByVariant", {})
        sheet.append([
            experiment.get("attribute", 0),
            experiment.get("material", 0),
            experiment.get("method", 0),
            experiment.get("standard", 0),
            experiment.get("norm", 0),
            *[_variant_value(indicators, variant["id"]) for variant in variants],
        ])
    sheet.append([
        "Статус варианта", "", "", "", "",
        *[variant.get("status", 0) for variant in variants],
    ])
    sheet.append([
        "Вывод", "", "", "", "",
        *[variant.get("conclusion", 0) for variant in variants],
    ])

    _style_sheet(sheet, len(variants))
    stream = BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def _write_composition(
    sheet: Any,
    section: str,
    component_label: str,
    composition: dict[str, Any],
    variants: list[dict[str, Any]],
    total_label: str,
    mixing_values: Any = None,
    experiment_values: Any = None,
) -> None:
    sheet.append([section])
    sheet.append([component_label, *[variant.get("name", "") for variant in variants]])
    rows = composition.get("components") or [{}]
    first_data_row = sheet.max_row + 1
    for component in rows:
        values = component.get("massPercentByVariant", {})
        sheet.append([
            component.get("name", ""),
            *[_variant_value(values, variant["id"]) for variant in variants],
        ])
    last_data_row = sheet.max_row

    if experiment_values is not None:
        _append_variant_values(sheet, "Суть эксперимента", variants, experiment_values)
    if mixing_values is not None:
        _append_variant_values(sheet, MIXING_LABEL, variants, mixing_values)
    total_row = [total_label]
    for column in range(2, len(variants) + 2):
        letter = get_column_letter(column)
        total_row.append(f"=SUM({letter}{first_data_row}:{letter}{last_data_row})")
    sheet.append(total_row)


def template_bytes() -> bytes:
    data = {
        "product": {"name": "", "purpose": ""},
        "testConditions": {
            "surfacePreparation": "",
            "dryFilmThicknessMicrometers": {"min": "", "max": ""},
            "curingBeforeTest": {"mode": "", "duration": ""},
        },
        "variants": [{
            "id": 1,
            "name": "Вариант № 1",
            "mainCompositionExperiment": "",
            "status": "",
            "conclusion": "",
        }],
        "mainComposition": {"components": [{}]},
        "hardener": {"components": [{}], "mixingRatioPer100PartsBaseByVariant": {}},
        "experiments": [{
            "attribute": "",
            "material": "",
            "method": "",
            "standard": "",
            "norm": "",
            "indicatorsByVariant": {},
        }],
    }
    return workbook_bytes(data)


def _find_row(sheet: Any, label: str, start_row: int = 1) -> int:
    for row_index in range(start_row, sheet.max_row + 1):
        if sheet.cell(row_index, 1).value == label:
            return row_index
    raise ValueError(f"Не найдена строка «{label}»")


def _read_variants(sheet: Any, row_index: int) -> list[dict[str, Any]]:
    names = [cell.value for cell in sheet[row_index][1:] if cell.value not in (None, "")]
    if not names:
        raise ValueError("В таблице не указаны варианты")
    return [{"id": index + 1, "name": str(name)} for index, name in enumerate(names)]


def _read_variant_row(sheet: Any, label: str, variants: list[dict[str, Any]], start_column: int = 2) -> dict[int, Any]:
    row_index = _find_row(sheet, label)
    return {
        variant["id"]: sheet.cell(row_index, start_column + index).value
        for index, variant in enumerate(variants)
    }


def _read_composition(
    sheet: Any,
    section_row: int,
    variants: list[dict[str, Any]],
    component_label: str,
    stop_labels: set[str],
) -> list[dict[str, Any]]:
    header_row = section_row + 1
    if sheet.cell(header_row, 1).value != component_label:
        raise ValueError(f"Не найдена колонка «{component_label}»")
    components = []
    for row_index in range(header_row + 1, sheet.max_row + 1):
        name = sheet.cell(row_index, 1).value
        if name in stop_labels:
            break
        if name in (None, ""):
            continue
        components.append({
            "id": len(components) + 1,
            "name": name,
            "massPercentByVariant": {
                str(variant["id"]): sheet.cell(row_index, index + 2).value or 0
                for index, variant in enumerate(variants)
            },
        })
    return components


def data_from_workbook(content: bytes) -> dict[str, Any]:
    try:
        workbook = load_workbook(BytesIO(content), data_only=True, read_only=True)
    except Exception as error:
        raise ValueError("Не удалось прочитать Excel-файл") from error
    if len(workbook.worksheets) != 1:
        raise ValueError("Шаблон должен содержать один лист")
    sheet = workbook.active
    if sheet.cell(1, 1).value != TITLE:
        raise ValueError("Неизвестный формат листа: ожидается «Лист испытаний»")

    product = {
        "name": sheet.cell(_find_row(sheet, "Название продукта"), 2).value or 0,
        "purpose": sheet.cell(_find_row(sheet, "Назначение продукта"), 2).value or 0,
    }
    surface_preparation = sheet.cell(_find_row(sheet, "Подготовка поверхности"), 2).value or 0
    thickness_row = _find_row(sheet, "Толщина сухой пленки, мкм")
    curing_row = _find_row(sheet, "Отверждение до испытания")
    variants = _read_variants(sheet, _find_row(sheet, "Варианты") + 1)

    for index, variant in enumerate(variants):
        variant["mainCompositionExperiment"] = _read_variant_row(sheet, "Суть эксперимента", variants)[variant["id"]] or 0
        variant["status"] = _read_variant_row(sheet, "Статус варианта", variants)[variant["id"]] or 0
        variant["conclusion"] = _read_variant_row(sheet, "Вывод", variants)[variant["id"]] or 0

    main_section_row = _find_row(sheet, MAIN_SECTION)
    hardener_section_row = _find_row(sheet, HARDENER_SECTION)
    experiments_section_row = _find_row(sheet, EXPERIMENTS_SECTION)
    main_components = _read_composition(
        sheet,
        main_section_row,
        variants,
        "Компонент",
        {"Суть эксперимента", "Итог по основному составу", HARDENER_SECTION},
    )
    hardener_components = _read_composition(
        sheet,
        hardener_section_row,
        variants,
        "Компонент отвердителя",
        {MIXING_LABEL, "Итог по отвердителю", EXPERIMENTS_SECTION},
    )
    mixing_values = _read_variant_row(sheet, MIXING_LABEL, variants)

    experiment_header_row = experiments_section_row + 1
    if tuple(sheet.cell(experiment_header_row, index + 1).value for index in range(len(EXPERIMENT_FIELDS))) != EXPERIMENT_FIELDS:
        raise ValueError("Заголовки таблицы экспериментов не совпадают с шаблоном")
    status_row = _find_row(sheet, "Статус варианта", start_row=experiments_section_row + 1)
    conclusion_row = _find_row(sheet, "Вывод", start_row=experiments_section_row + 1)
    experiments = []
    for row_index in range(experiment_header_row + 1, status_row):
        attribute = sheet.cell(row_index, 1).value
        if attribute in (None, ""):
            continue
        experiments.append({
            "id": len(experiments) + 1,
            "attribute": attribute,
            "material": sheet.cell(row_index, 2).value or 0,
            "method": sheet.cell(row_index, 3).value or 0,
            "standard": sheet.cell(row_index, 4).value or 0,
            "norm": sheet.cell(row_index, 5).value or 0,
            "indicatorsByVariant": {
                str(variant["id"]): sheet.cell(row_index, index + 6).value or 0
                for index, variant in enumerate(variants)
            },
        })

    for index, variant in enumerate(variants):
        variant["status"] = sheet.cell(status_row, index + 6).value or variant["status"]
        variant["conclusion"] = sheet.cell(conclusion_row, index + 6).value or variant["conclusion"]

    return {
        "schemaVersion": 1,
        "product": product,
        "testConditions": {
            "surfacePreparation": surface_preparation,
            "dryFilmThicknessMicrometers": {
                "min": sheet.cell(thickness_row, 2).value or 0,
                "max": sheet.cell(thickness_row, 3).value or 0,
            },
            "curingBeforeTest": {
                "mode": sheet.cell(curing_row, 3).value or 0,
                "duration": sheet.cell(curing_row, 5).value or 0,
            },
        },
        "variants": variants,
        "mainComposition": {"components": main_components},
        "hardener": {
            "components": hardener_components,
            "mixingRatioPer100PartsBaseByVariant": {
                str(variant["id"]): mixing_values[variant["id"]] or 0 for variant in variants
            },
        },
        "experiments": experiments,
    }