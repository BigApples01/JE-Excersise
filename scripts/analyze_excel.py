import argparse
import json
from pathlib import Path
from typing import Dict, List, Any

import pandas as pd

DATE_PARSE_THRESHOLD = 0.6


def detect_date_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
        parsed = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
        non_null = series.notna().sum()
        if non_null == 0:
            return pd.Series(dtype="datetime64[ns]")
        parsed_non_null = parsed.notna().sum()
        if parsed_non_null / non_null >= DATE_PARSE_THRESHOLD:
            return parsed
    return pd.Series(dtype="datetime64[ns]")


def summarize_sheet(name: str, df: pd.DataFrame) -> Dict[str, Any]:
    sheet_summary: Dict[str, Any] = {
        "sheet": name,
        "row_count": int(df.shape[0]),
        "column_count": int(df.shape[1]),
        "columns": [],
        "date_ranges": {},
    }

    date_ranges: Dict[str, Dict[str, Any]] = {}
    for column in df.columns:
        series = df[column]
        non_null = int(series.notna().sum())
        nulls = int(series.isna().sum())
        unique_count = int(series.nunique(dropna=True))
        col_info: Dict[str, Any] = {
            "column": str(column),
            "dtype": str(series.dtype),
            "non_null_count": non_null,
            "null_count": nulls,
            "unique_count": unique_count,
        }

        if pd.api.types.is_numeric_dtype(series):
            col_info.update(
                {
                    "min": float(series.min()),
                    "max": float(series.max()),
                    "mean": float(series.mean()),
                    "median": float(series.median()),
                    "std": float(series.std()),
                }
            )
        else:
            date_series = detect_date_series(series)
            if not date_series.empty:
                date_min = date_series.min()
                date_max = date_series.max()
                col_info.update(
                    {
                        "min": date_min.isoformat() if pd.notna(date_min) else None,
                        "max": date_max.isoformat() if pd.notna(date_max) else None,
                    }
                )
                date_ranges[str(column)] = {
                    "min": col_info["min"],
                    "max": col_info["max"],
                }

        sheet_summary["columns"].append(col_info)

    sheet_summary["date_ranges"] = date_ranges
    if date_ranges:
        all_mins = [v["min"] for v in date_ranges.values() if v["min"]]
        all_maxs = [v["max"] for v in date_ranges.values() if v["max"]]
        if all_mins and all_maxs:
            sheet_summary["overall_date_range"] = {
                "min": min(all_mins),
                "max": max(all_maxs),
            }

    return sheet_summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze Excel file and emit summary statistics.")
    parser.add_argument(
        "--input",
        default="je_samples_exercise.xlsx",
        help="Path to the Excel file.",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs",
        help="Directory to write output files.",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    workbook = pd.ExcelFile(input_path)

    summaries: List[Dict[str, Any]] = []
    for sheet_name in workbook.sheet_names:
        df = workbook.parse(sheet_name)
        summaries.append(summarize_sheet(sheet_name, df))

    summary_payload = {
        "file": str(input_path.name),
        "sheet_count": len(workbook.sheet_names),
        "sheets": summaries,
    }

    summary_json_path = output_dir / "summary.json"
    summary_json_path.write_text(json.dumps(summary_payload, indent=2))

    sheet_overview_rows = [
        {
            "sheet": sheet["sheet"],
            "row_count": sheet["row_count"],
            "column_count": sheet["column_count"],
            "overall_date_min": sheet.get("overall_date_range", {}).get("min"),
            "overall_date_max": sheet.get("overall_date_range", {}).get("max"),
        }
        for sheet in summaries
    ]
    pd.DataFrame(sheet_overview_rows).to_csv(output_dir / "summary.csv", index=False)

    column_rows = []
    for sheet in summaries:
        for column in sheet["columns"]:
            row = {"sheet": sheet["sheet"], **column}
            column_rows.append(row)

    pd.DataFrame(column_rows).to_csv(output_dir / "column_stats.csv", index=False)
    (output_dir / "column_stats.json").write_text(json.dumps(column_rows, indent=2))


if __name__ == "__main__":
    main()
