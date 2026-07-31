from __future__ import annotations

from copy import copy
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side

from config import ROOT_DIR, baseline_paths, ensure_generated_directories


DAY_SHEET_PREFIX = "Day "
MACHINE_COUNT = 10
WINDOW_COUNT = 4
ROWS_PER_MACHINE = 12
FIRST_SCHEDULE_ROW = 6
FIRST_WINDOW_COLUMN = 2  # B

REQUIRED_COLUMNS = {
    "patient",
    "fraction",
    "machine",
    "day",
    "window",
    "duration_minutes",
    "priority",
    "is_preferred_machine",
}


def _template_candidates(instance_id: str) -> list[Path]:
    filename = f"{instance_id} base.xlsx"
    return [
        ROOT_DIR / "templates" / filename,
        ROOT_DIR / "data" / "template" / filename,
        ROOT_DIR / filename,
    ]


def find_template_workbook(instance_id: str, template_path: str | Path | None = None) -> Path:
    if template_path is not None:
        path = Path(template_path)
        if not path.is_absolute():
            path = ROOT_DIR / path
        if not path.exists():
            raise FileNotFoundError(f"Template workbook not found: {path}")
        return path

    for candidate in _template_candidates(instance_id):
        if candidate.exists():
            return candidate

    searched = "\n".join(f"  - {path}" for path in _template_candidates(instance_id))
    raise FileNotFoundError(
        "Could not find the baseline Excel template. Place it in one of:\n"
        f"{searched}"
    )


def _read_schedule_csv(csv_path: Path) -> pd.DataFrame:
    if not csv_path.exists():
        raise FileNotFoundError(f"Baseline schedule CSV not found: {csv_path}")

    schedule = pd.read_csv(csv_path)
    missing = REQUIRED_COLUMNS.difference(schedule.columns)
    if missing:
        raise ValueError(
            "Baseline schedule CSV is missing required columns: "
            + ", ".join(sorted(missing))
        )

    numeric_columns = [
        "patient",
        "fraction",
        "machine",
        "day",
        "window",
        "duration_minutes",
        "priority",
    ]
    for column in numeric_columns:
        schedule[column] = pd.to_numeric(schedule[column], errors="raise").astype(int)

    preferred = schedule["is_preferred_machine"]
    if preferred.dtype == bool:
        schedule["is_preferred_machine"] = preferred
    else:
        schedule["is_preferred_machine"] = (
            preferred.astype(str)
            .str.strip()
            .str.lower()
            .map({"true": True, "false": False, "1": True, "0": False})
        )

    if schedule["is_preferred_machine"].isna().any():
        raise ValueError("Column 'is_preferred_machine' contains unrecognised values.")
    if not schedule["machine"].between(1, MACHINE_COUNT).all():
        raise ValueError(f"Machine values must be between 1 and {MACHINE_COUNT}.")
    if not schedule["window"].between(1, WINDOW_COUNT).all():
        raise ValueError(f"Window values must be between 1 and {WINDOW_COUNT}.")
    if not schedule["priority"].isin([1, 2, 3]).all():
        raise ValueError("Priority values must be 1, 2, or 3.")

    return schedule.sort_values(
        ["day", "machine", "window", "patient", "fraction"],
        kind="stable",
    ).reset_index(drop=True)


def _apply_style(style, target) -> None:
    """Apply an immutable copied StyleArray to a target cell."""
    target._style = copy(style)


def _find_style_cells(workbook) -> dict[str, Any]:
    """Find reusable appointment, empty, and neutral slot styles."""
    styles: dict[str, Any] = {}
    neutral_by_offset: dict[int, Any] = {}

    fill_to_key = {
        "00C6E0B4": "priority_1",
        "00FCE4D6": "priority_2",
        "00F4CCCC": "priority_3",
        "00E7E6E6": "empty",
    }

    schedule_last_row = FIRST_SCHEDULE_ROW + MACHINE_COUNT * ROWS_PER_MACHINE - 1

    for sheet_name in workbook.sheetnames:
        if not sheet_name.startswith(DAY_SHEET_PREFIX):
            continue
        worksheet = workbook[sheet_name]

        for row in range(FIRST_SCHEDULE_ROW, schedule_last_row + 1):
            offset = (row - FIRST_SCHEDULE_ROW) % ROWS_PER_MACHINE
            for column in range(FIRST_WINDOW_COLUMN, FIRST_WINDOW_COLUMN + WINDOW_COUNT):
                cell = worksheet.cell(row, column)
                rgb = cell.fill.fgColor.rgb

                key = fill_to_key.get(rgb)
                if key and key not in styles:
                    styles[key] = copy(cell._style)

                # Neutral unused schedule cells use the very light blue fill.
                if rgb == "00F8FAFC" and offset not in neutral_by_offset:
                    neutral_by_offset[offset] = copy(cell._style)

        if len(styles) == 4 and len(neutral_by_offset) == ROWS_PER_MACHINE:
            break

    missing = {"priority_1", "priority_2", "priority_3", "empty"}.difference(styles)
    if missing:
        raise ValueError(
            "The template does not contain all required appointment styles: "
            + ", ".join(sorted(missing))
        )

    if not neutral_by_offset:
        raise ValueError("The template does not contain a neutral schedule-cell style.")

    # All unused slot rows share the same neutral style in this template.
    # Offset 0 is normally occupied by an appointment or an EMPTY merge, so it
    # may not exist as a neutral cell in the populated template. Reuse the first
    # neutral style for any missing offset.
    fallback_neutral = next(iter(neutral_by_offset.values()))
    for offset in range(ROWS_PER_MACHINE):
        neutral_by_offset.setdefault(offset, fallback_neutral)

    styles["neutral_by_offset"] = neutral_by_offset
    return styles


def _appointment_text(row: pd.Series) -> str:
    patient = int(row["patient"])
    fraction = int(row["fraction"])
    duration = int(row["duration_minutes"])
    suffix = "" if bool(row["is_preferred_machine"]) else "  NP"
    return f"P{patient:02d}  |  F{fraction}{suffix}\n{duration} min"


def _apply_nonpreferred_border(cell) -> None:
    blue = Side(style="medium", color="4472C4")
    cell.border = Border(left=blue, right=blue, top=blue, bottom=blue)


def _clear_day_body(worksheet, styles: dict[str, Any]) -> None:
    """Reset the complete schedule body to the template's neutral slot style."""
    schedule_last_row = FIRST_SCHEDULE_ROW + MACHINE_COUNT * ROWS_PER_MACHINE - 1

    for merged_range in list(worksheet.merged_cells.ranges):
        if (
            merged_range.min_col >= FIRST_WINDOW_COLUMN
            and merged_range.max_col <= FIRST_WINDOW_COLUMN + WINDOW_COUNT - 1
            and merged_range.min_row >= FIRST_SCHEDULE_ROW
            and merged_range.max_row <= schedule_last_row
        ):
            worksheet.unmerge_cells(str(merged_range))

    neutral_by_offset = styles["neutral_by_offset"]
    for row in range(FIRST_SCHEDULE_ROW, schedule_last_row + 1):
        offset = (row - FIRST_SCHEDULE_ROW) % ROWS_PER_MACHINE
        for column in range(FIRST_WINDOW_COLUMN, FIRST_WINDOW_COLUMN + WINDOW_COUNT):
            cell = worksheet.cell(row=row, column=column)
            cell.value = None
            _apply_style(neutral_by_offset[offset], cell)


def _populate_day_sheet(worksheet, day_schedule: pd.DataFrame, styles: dict[str, Any]) -> None:
    _clear_day_body(worksheet, styles)

    if not day_schedule.empty:
        day_number = int(day_schedule["day"].iloc[0])
    else:
        day_number = int(worksheet.title.replace(DAY_SHEET_PREFIX, ""))
    worksheet["A1"] = f"DAY {day_number} — MACHINE APPOINTMENT BOARD"

    # Appointment rows are 35 pt whenever at least one window uses that row.
    for machine in range(1, MACHINE_COUNT + 1):
        block_start = FIRST_SCHEDULE_ROW + (machine - 1) * ROWS_PER_MACHINE
        machine_schedule = day_schedule[day_schedule["machine"] == machine]
        max_appointments = 0
        for window in range(1, WINDOW_COUNT + 1):
            count = int((machine_schedule["window"] == window).sum())
            max_appointments = max(max_appointments, count)
        for offset in range(max_appointments):
            worksheet.row_dimensions[block_start + offset].height = 35

    for machine in range(1, MACHINE_COUNT + 1):
        block_start = FIRST_SCHEDULE_ROW + (machine - 1) * ROWS_PER_MACHINE
        block_end = block_start + ROWS_PER_MACHINE - 1

        for window in range(1, WINDOW_COUNT + 1):
            column = FIRST_WINDOW_COLUMN + window - 1
            appointments = day_schedule[
                (day_schedule["machine"] == machine)
                & (day_schedule["window"] == window)
            ]

            if appointments.empty:
                worksheet.merge_cells(
                    start_row=block_start,
                    start_column=column,
                    end_row=block_end,
                    end_column=column,
                )
                cell = worksheet.cell(block_start, column)
                _apply_style(styles["empty"], cell)
                cell.value = "— EMPTY —"
                continue

            if len(appointments) > ROWS_PER_MACHINE:
                raise ValueError(
                    f"Day {day_number}, machine M{machine}, window {window} has "
                    f"{len(appointments)} appointments; the template supports at most "
                    f"{ROWS_PER_MACHINE}."
                )

            for offset, (_, appointment) in enumerate(appointments.iterrows()):
                row_number = block_start + offset
                cell = worksheet.cell(row_number, column)
                priority = int(appointment["priority"])
                _apply_style(styles[f"priority_{priority}"], cell)
                cell.value = _appointment_text(appointment)

                if not bool(appointment["is_preferred_machine"]):
                    _apply_nonpreferred_border(cell)


def build_baseline_workbook(
    instance_id: str,
    *,
    force: bool = False,
    template_path: str | Path | None = None,
) -> Path:
    """Build the day-by-day baseline workbook from baseline_schedule.csv."""
    ensure_generated_directories(instance_id)

    baseline = baseline_paths(instance_id)
    source_csv = Path(baseline["schedule_csv"])
    output_excel = Path(baseline["schedule_excel"])
    template = find_template_workbook(instance_id, template_path)

    if output_excel.exists() and not force:
        newest_input = max(source_csv.stat().st_mtime, template.stat().st_mtime)
        if output_excel.stat().st_mtime >= newest_input:
            return output_excel

    schedule = _read_schedule_csv(source_csv)
    workbook = load_workbook(template)
    styles = _find_style_cells(workbook)

    available_days = sorted(schedule["day"].unique().tolist())
    for sheet_name in workbook.sheetnames:
        if not sheet_name.startswith(DAY_SHEET_PREFIX):
            continue
        try:
            day = int(sheet_name.replace(DAY_SHEET_PREFIX, ""))
        except ValueError:
            continue
        day_schedule = schedule[schedule["day"] == day]
        _populate_day_sheet(workbook[sheet_name], day_schedule, styles)

    missing_sheets = [
        day for day in available_days if f"Day {day:02d}" not in workbook.sheetnames
    ]
    if missing_sheets:
        raise ValueError(
            "Template is missing required day sheets: "
            + ", ".join(f"Day {day:02d}" for day in missing_sheets)
        )

    output_excel.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_excel)
    return output_excel


def load_baseline_workbook(
    instance_id: str,
    *,
    force_rebuild: bool = False,
    template_path: str | Path | None = None,
):
    workbook_path = build_baseline_workbook(
        instance_id,
        force=force_rebuild,
        template_path=template_path,
    )
    return load_workbook(workbook_path, data_only=False)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("instance_id")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    path = build_baseline_workbook(args.instance_id, force=args.force)
    print(f"Baseline workbook created: {path}")