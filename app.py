# app.py
# Python >= 3.14
#
# Приложение Streamlit для сравнения двух спецификаций Excel.
# Формат результата: XLSX или JSON.

import io
import json
import re
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


DEFAULT_HEADER_ROW = 2
PLACEHOLDER_RE = re.compile(r"^\d+[.)]?$")


@dataclass(slots=True)
class FieldChange:
    """
    Изменение одного поля/столбца в строке.
    """
    field: str
    old: str
    new: str


@dataclass(slots=True)
class RowRecord:
    """
    Нормализованная строка спецификации.
    """
    excel_row: int
    section: str
    is_section: bool
    key: str
    values: dict[str, str]


@dataclass(slots=True)
class ParsedSpec:
    """
    Разобранный файл спецификации.
    """
    file_name: str
    title: str
    columns: list[str]
    records: list[RowRecord]


@dataclass(slots=True)
class RowChange:
    """
    Результат сравнения одной строки.
    """
    key: str
    section: str
    old_row_num: int | None = None
    new_row_num: int | None = None
    old_row: dict[str, str] | None = None
    new_row: dict[str, str] | None = None
    changed_fields: list[FieldChange] = field(default_factory=list)


@dataclass(slots=True)
class ComparisonResult:
    """
    Полный результат сравнения двух спецификаций.
    """
    meta: dict[str, Any]
    summary: dict[str, int]
    columns: list[str]
    field_map: dict[str, str | None]
    added: list[RowChange]
    deleted: list[RowChange]
    modified: list[RowChange]


def normalize_text(value: Any) -> str:
    """
    Нормализует значение ячейки к строке:
    - None/NaN -> пустая строка;
    - числа 2.0 -> 2;
    - лишние пробелы убираются;
    - одиночные тире/прочерки считаются пустым значением.
    """
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, (int, float)):
        try:
            f = float(value)
            if f.is_integer():
                return str(int(f))
        except Exception:
            pass
        return str(value)

    s = str(value).replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()

    if s in {"-", "—", "–", "--", "---"}:
        return ""

    return s


def clean_header(value: Any, idx: int) -> str:
    """
    Нормализует заголовок столбца.
    """
    result = normalize_text(value)
    return result if result else f"Столбец {idx + 1}"


def make_unique_columns(columns: list[str]) -> list[str]:
    """
    Делает имена столбцов уникальными, если есть дубли.
    """
    seen: dict[str, int] = {}
    result: list[str] = []

    for col in columns:
        if col not in seen:
            seen[col] = 1
            result.append(col)
        else:
            seen[col] += 1
            result.append(f"{col} ({seen[col]})")

    return result


def find_column(columns: list[str], keywords: tuple[str, ...]) -> str | None:
    """
    Ищет столбец по ключевым словам в имени столбца.
    """
    for keyword in keywords:
        for col in columns:
            if keyword in col.lower():
                return col
    return None


def build_field_map(columns: list[str]) -> dict[str, str | None]:
    """
    Определяет служебные столбцы спецификации по их названиям.
    """
    return {
        "poz": find_column(columns, ("поз. по спецификации", "поз")),
        "name": find_column(columns, ("наименование",)),
        "type": find_column(columns, ("тип,", "тип")),
        "code": find_column(columns, ("код оборудования", "код")),
        "factory": find_column(columns, ("завод",)),
        "unit": find_column(columns, ("ед. изм", "ед")),
        "qty": find_column(columns, ("кол-во", "количество")),
        "mass": find_column(columns, ("масса",)),
        "note": find_column(columns, ("примечание",)),
    }


def read_title(file_bytes: bytes) -> str:
    """
    Читает название спецификации из ячейки A1.
    """
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    title = ws.cell(row=1, column=1).value if ws is not None else ""
    wb.close()
    return normalize_text(title)


def parse_spec(
    file_bytes: bytes,
    file_name: str,
    header_row: int,
    ignore_placeholders: bool = True,
) -> ParsedSpec:
    """
    Разбирает файл спецификации.

    Логика:
    - заголовок таблицы берётся из строки header_row;
    - данные идут ниже;
    - строка считается разделом, если заполнен только столбец
      "Наименование и техническая характеристика";
    - для строк создаётся ключ сравнения.
    """
    title = read_title(file_bytes)

    df = pd.read_excel(
        io.BytesIO(file_bytes),
        header=header_row - 1,
        dtype=object,
        engine="openpyxl",
    )

    columns = make_unique_columns(
        [clean_header(col, i) for i, col in enumerate(df.columns)]
    )
    df.columns = columns

    field_map = build_field_map(columns)

    name_col = field_map["name"]
    poz_col = field_map["poz"]
    qty_col = field_map["qty"]
    mass_col = field_map["mass"]
    note_col = field_map["note"]

    volatile_cols = {col for col in (qty_col, mass_col, note_col) if col}

    records: list[RowRecord] = []

    current_section = ""
    section_counts: dict[str, int] = {}
    key_counts: dict[str, int] = {}

    for offset, (_, row) in enumerate(df.iterrows()):
        excel_row = header_row + offset + 1

        values = {col: normalize_text(row.get(col)) for col in columns}
        non_empty = [col for col, val in values.items() if val]

        # Полностью пустая строка.
        if not non_empty:
            continue

        # Служебные пустые строки вида "4." или "5.".
        if (
            ignore_placeholders
            and name_col
            and len(non_empty) == 1
            and non_empty[0] == name_col
            and PLACEHOLDER_RE.fullmatch(values[name_col])
        ):
            continue

        # Если заполнен только столбец "Наименование...", считаем строку разделом.
        is_section = bool(name_col) and len(non_empty) == 1 and non_empty[0] == name_col

        if is_section:
            section_name = values[name_col]
            count = section_counts.get(section_name, 0) + 1
            section_counts[section_name] = count

            current_section = (
                section_name if count == 1 else f"{section_name} ({count})"
            )
            key = f"section::{current_section}"
        else:
            poz = values.get(poz_col, "") if poz_col else ""

            # Если есть позиция, используем её как основной ключ.
            if poz:
                base_key = f"poz::{current_section}::{poz}"
            else:
                # Иначе используем все поля, кроме обычно изменяемых количественных/служебных.
                identity_parts: list[str] = []

                for col in columns:
                    if col in volatile_cols:
                        continue

                    val = values.get(col, "")
                    if val:
                        identity_parts.append(f"{col}={val}")

                if identity_parts:
                    base_key = f"row::{current_section}::" + "|".join(identity_parts)
                else:
                    base_key = f"empty::{current_section}::{offset}"

            count = key_counts.get(base_key, 0) + 1
            key_counts[base_key] = count
            key = base_key if count == 1 else f"{base_key}#{count}"

        records.append(
            RowRecord(
                excel_row=excel_row,
                section=current_section,
                is_section=is_section,
                key=key,
                values=values,
            )
        )

    return ParsedSpec(
        file_name=file_name,
        title=title,
        columns=columns,
        records=records,
    )


def compare_specs(
    old: ParsedSpec,
    new: ParsedSpec,
    header_row: int,
) -> ComparisonResult:
    """
    Сравнивает две спецификации по ключам строк.
    """
    columns = list(dict.fromkeys(old.columns + new.columns))
    field_map = build_field_map(columns)

    old_map = {record.key: record for record in old.records}
    new_map = {record.key: record for record in new.records}

    old_keys = [record.key for record in old.records]
    new_keys = [record.key for record in new.records]

    added_keys = [key for key in new_keys if key not in old_map]
    deleted_keys = [key for key in old_keys if key not in new_map]
    common_keys = [key for key in new_keys if key in old_map]

    added: list[RowChange] = []
    deleted: list[RowChange] = []
    modified: list[RowChange] = []
    unchanged = 0

    for key in added_keys:
        rec = new_map[key]
        added.append(
            RowChange(
                key=key,
                section=rec.section,
                new_row_num=rec.excel_row,
                new_row=rec.values,
            )
        )

    for key in deleted_keys:
        rec = old_map[key]
        deleted.append(
            RowChange(
                key=key,
                section=rec.section,
                old_row_num=rec.excel_row,
                old_row=rec.values,
            )
        )

    for key in common_keys:
        old_rec = old_map[key]
        new_rec = new_map[key]

        changed_fields: list[FieldChange] = []

        for col in columns:
            old_val = old_rec.values.get(col, "")
            new_val = new_rec.values.get(col, "")

            if old_val != new_val:
                changed_fields.append(
                    FieldChange(
                        field=col,
                        old=old_val,
                        new=new_val,
                    )
                )

        if changed_fields:
            modified.append(
                RowChange(
                    key=key,
                    section=new_rec.section,
                    old_row_num=old_rec.excel_row,
                    new_row_num=new_rec.excel_row,
                    old_row=old_rec.values,
                    new_row=new_rec.values,
                    changed_fields=changed_fields,
                )
            )
        else:
            unchanged += 1

    summary = {
        "added": len(added),
        "deleted": len(deleted),
        "modified": len(modified),
        "unchanged": unchanged,
        "total_old": len(old.records),
        "total_new": len(new.records),
    }

    meta = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "header_row": header_row,
        "old_file": old.file_name,
        "new_file": new.file_name,
        "old_title": old.title,
        "new_title": new.title,
        "headers_equal": old.columns == new.columns,
    }

    return ComparisonResult(
        meta=meta,
        summary=summary,
        columns=columns,
        field_map=field_map,
        added=added,
        deleted=deleted,
        modified=modified,
    )


def changed_fields_to_df(result: ComparisonResult) -> pd.DataFrame:
    """
    Возвращает DataFrame с изменениями по каждому изменённому полю.
    """
    headers = [
        "Статус",
        "Раздел",
        "Позиция",
        "Наименование",
        "Изменённое поле",
        "Значение в r0",
        "Значение в r1",
        "Строка r0",
        "Строка r1",
    ]

    rows: list[dict[str, Any]] = []

    poz_col = result.field_map.get("poz")
    name_col = result.field_map.get("name")

    for item in result.modified:
        source = item.new_row or item.old_row or {}

        for fc in item.changed_fields:
            rows.append(
                {
                    "Статус": "Изменено",
                    "Раздел": item.section,
                    "Позиция": source.get(poz_col, "") if poz_col else "",
                    "Наименование": source.get(name_col, "") if name_col else "",
                    "Изменённое поле": fc.field,
                    "Значение в r0": fc.old,
                    "Значение в r1": fc.new,
                    "Строка r0": item.old_row_num,
                    "Строка r1": item.new_row_num,
                }
            )

    return pd.DataFrame(rows, columns=headers)


def items_to_df(
    items: list[RowChange],
    columns: list[str],
    mode: str,
) -> pd.DataFrame:
    """
    Возвращает DataFrame для добавленных или удалённых строк.
    """
    headers = ["Раздел", "Excel row", *columns]
    rows: list[dict[str, Any]] = []

    for item in items:
        values = item.new_row if mode == "added" else item.old_row
        row_num = item.new_row_num if mode == "added" else item.old_row_num

        row: dict[str, Any] = {
            "Раздел": item.section,
            "Excel row": row_num,
        }

        values = values or {}

        for col in columns:
            row[col] = values.get(col, "")

        rows.append(row)

    return pd.DataFrame(rows, columns=headers)


def to_cell(value: Any) -> Any:
    """
    Подготавливает значение для записи в Excel.
    """
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass

    return value


def df_to_rows(df: pd.DataFrame) -> list[list[Any]]:
    """
    Преобразует DataFrame в список строк для записи в Excel.
    """
    if df.empty:
        return []

    return [[to_cell(v) for v in row] for row in df.astype(object).values.tolist()]


def export_xlsx(result: ComparisonResult) -> bytes:
    """
    Формирует итоговый XLSX-файл сравнения.
    """
    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)

    added_fill = PatternFill("solid", fgColor="C6EFCE")
    deleted_fill = PatternFill("solid", fgColor="FFC7CE")
    modified_fill = PatternFill("solid", fgColor="FFEB9C")

    def write_table(
        ws,
        headers: list[str],
        rows: list[list[Any]],
        fill: PatternFill | None = None,
    ) -> None:
        ws.append(headers)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(vertical="top", wrap_text=True)

        if rows:
            for row in rows:
                ws.append(row)

                if fill:
                    for cell in ws[ws.max_row]:
                        cell.fill = fill
        else:
            ws.append(["Нет данных"])

        # Примерная автоподборка ширины столбцов.
        for col_idx, header in enumerate(headers, start=1):
            max_len = len(str(header))
            max_row = min(ws.max_row, 100)

            for row in ws.iter_rows(
                min_row=2,
                max_row=max_row,
                min_col=col_idx,
                max_col=col_idx,
            ):
                for cell in row:
                    if cell.value is not None:
                        max_len = max(
                            max_len,
                            min(len(str(cell.value)), 80),
                        )

            ws.column_dimensions[get_column_letter(col_idx)].width = min(
                max_len + 2,
                80,
            )

        ws.freeze_panes = "A2"

    # Лист "Итог"
    ws_summary = wb.active
    ws_summary.title = "Итог"

    summary_rows = [
        ["Файл до корректировки", result.meta.get("old_file", "")],
        ["Файл после корректировки", result.meta.get("new_file", "")],
        ["Название в r0", result.meta.get("old_title", "")],
        ["Название в r1", result.meta.get("new_title", "")],
        ["Строка заголовков", result.meta.get("header_row", "")],
        ["Дата сравнения (UTC)", result.meta.get("generated_at", "")],
        ["Состав столбцов одинаковый", result.meta.get("headers_equal", "")],
        ["Добавлено строк", result.summary.get("added", 0)],
        ["Удалено строк", result.summary.get("deleted", 0)],
        ["Изменено строк", result.summary.get("modified", 0)],
        ["Без изменений", result.summary.get("unchanged", 0)],
        ["Всего строк в r0", result.summary.get("total_old", 0)],
        ["Всего строк в r1", result.summary.get("total_new", 0)],
    ]

    write_table(ws_summary, ["Параметр", "Значение"], summary_rows)

    # Лист "Изменения"
    changes_df = changed_fields_to_df(result)
    write_table(
        wb.create_sheet("Изменения"),
        list(changes_df.columns),
        df_to_rows(changes_df),
        modified_fill,
    )

    # Лист "Добавленные"
    added_df = items_to_df(result.added, result.columns, "added")
    write_table(
        wb.create_sheet("Добавленные"),
        list(added_df.columns),
        df_to_rows(added_df),
        added_fill,
    )

    # Лист "Удалённые"
    deleted_df = items_to_df(result.deleted, result.columns, "deleted")
    write_table(
        wb.create_sheet("Удалённые"),
        list(deleted_df.columns),
        df_to_rows(deleted_df),
        deleted_fill,
    )

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def export_json(result: ComparisonResult) -> bytes:
    """
    Формирует JSON-результат сравнения.
    """
    payload = asdict(result)
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def main() -> None:
    st.set_page_config(
        page_title="Сравнение спецификаций",
        layout="wide",
    )

    st.title("Сравнение спецификаций")
    st.caption(
        "Загрузите два файла Excel: r0 — до корректировки и r1 — после корректировки."
    )

    with st.sidebar:
        st.subheader("Настройки")

        header_row = st.number_input(
            "Строка заголовков таблицы",
            min_value=1,
            max_value=100,
            value=DEFAULT_HEADER_ROW,
            help=(
                "Если название спецификации в A1, а заголовки таблицы во второй строке, "
                "оставьте 2. Если заголовки ниже, укажите номер строки вручную."
            ),
        )

        ignore_placeholders = st.checkbox(
            "Игнорировать служебные строки вида '4.'",
            value=True,
            help=(
                "Если в таблице есть пустые служебные строки с текстом '4.' или '5.', "
                "они не будут участвовать в сравнении."
            ),
        )

        output_format = st.selectbox(
            "Формат результата",
            ("xlsx", "json"),
            format_func=lambda x: x.upper(),
        )

    file_r0 = st.file_uploader("Файл r0 (до корректировки)", type=["xlsx"])
    file_r1 = st.file_uploader("Файл r1 (после корректировки)", type=["xlsx"])

    if st.button("Сравнить", type="primary"):
        if not file_r0 or not file_r1:
            st.warning("Нужно загрузить оба файла: r0 и r1.")
        else:
            try:
                old_spec = parse_spec(
                    file_r0.getvalue(),
                    file_r0.name,
                    int(header_row),
                    ignore_placeholders,
                )
                new_spec = parse_spec(
                    file_r1.getvalue(),
                    file_r1.name,
                    int(header_row),
                    ignore_placeholders,
                )

                result = compare_specs(old_spec, new_spec, int(header_row))
                st.session_state.result = result

            except Exception as exc:
                st.error(f"Ошибка при сравнении: {exc}")
                st.exception(exc)

    result = st.session_state.get("result")

    if result is None:
        st.info("Загрузите файлы и нажмите «Сравнить».")
        return

    if not result.meta.get("headers_equal", True):
        st.warning("Состав столбцов в файлах различается.")

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Добавлено", result.summary["added"])
    c2.metric("Удалено", result.summary["deleted"])
    c3.metric("Изменено", result.summary["modified"])
    c4.metric("Без изменений", result.summary["unchanged"])

    changes_df = changed_fields_to_df(result)
    added_df = items_to_df(result.added, result.columns, "added")
    deleted_df = items_to_df(result.deleted, result.columns, "deleted")

    tabs = st.tabs(
        [
            "Итог",
            "Изменения",
            "Добавленные",
            "Удалённые",
            "JSON",
        ]
    )

    with tabs[0]:
        st.subheader("Метаданные")
        st.json(result.meta)

        st.subheader("Сводка")
        st.json(result.summary)

    with tabs[1]:
        st.subheader("Изменённые поля")
        st.dataframe(changes_df, use_container_width=True)

    with tabs[2]:
        st.subheader("Добавленные строки")
        st.dataframe(added_df, use_container_width=True)

    with tabs[3]:
        st.subheader("Удалённые строки")
        st.dataframe(deleted_df, use_container_width=True)

    with tabs[4]:
        st.subheader("Полный JSON-результат")
        st.json(asdict(result))

    st.subheader("Скачать результат")

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    if output_format == "xlsx":
        st.download_button(
            "Скачать результат в XLSX",
            data=export_xlsx(result),
            file_name=f"spec_comparison_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_xlsx",
        )
    else:
        st.download_button(
            "Скачать результат в JSON",
            data=export_json(result),
            file_name=f"spec_comparison_{timestamp}.json",
            mime="application/json",
            key="download_json",
        )


if __name__ == "__main__":
    main()