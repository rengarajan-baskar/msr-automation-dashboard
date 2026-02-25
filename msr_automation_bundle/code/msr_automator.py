#!/usr/bin/env python3
import argparse
import os
import sys
import re
import json
from typing import Dict, List, Optional
import pandas as pd
import yaml


# ----------------- CONFIG LOADER -----------------
def load_config(path: str) -> dict:
    """Load YAML configuration safely."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Configuration file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


# ----------------- DATA HELPERS -----------------
def normalize_columns(df: pd.DataFrame, aliases: Dict[str, List[str]]) -> pd.DataFrame:
    """Normalize column names based on config aliases."""
    lc_map = {c.lower().strip(): c for c in df.columns}
    mapped = {}
    for logical, candidates in aliases.items():
        for cand in candidates:
            if cand.lower().strip() in lc_map:
                mapped[logical] = lc_map[cand.lower().strip()]
                break
    return df.rename(columns={v: k for k, v in mapped.items()})


def infer_root_cause(text: str, rules: Dict[str, List[str]]) -> Optional[str]:
    """Infer root cause by scanning description keywords."""
    if not isinstance(text, str) or not text.strip():
        return None
    text_lower = text.lower()
    for cause, kw_list in rules.items():
        for kw in kw_list:
            if kw in text_lower:
                return cause
    return None


def coalesce_root_cause(row, rc_col: Optional[str], rules: Dict[str, List[str]]):
    """Use provided root cause if available, else infer from text."""
    if rc_col and pd.notna(row.get(rc_col)):
        val = str(row.get(rc_col)).strip()
        if val:
            return val
    return infer_root_cause(row.get("short_description"), rules) or "Unspecified"


def safe_get_col(df, logical_name: str):
    """Return column name if it exists."""
    return logical_name if logical_name in df.columns else None


def build_pivot(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """Build pivot table for a given column."""
    if col not in df.columns:
        return pd.DataFrame(columns=[col, "Count"])
    pv = (
        df.groupby(col, dropna=False)
        .size()
        .reset_index(name="Count")
        .sort_values("Count", ascending=False)
    )
    return pv


# ----------------- MAIN FUNCTION -----------------
def main():
    ap = argparse.ArgumentParser(description="Automate MSR reporting from Excel tracker.")
    ap.add_argument("--input", default="../data/msr_sample.xlsx", help="Path to input Excel (.xlsx)")
    ap.add_argument("--outdir", default="../out", help="Directory to write outputs")
    ap.add_argument("--config", default="config.yaml", help="Path to config YAML")
    ap.add_argument("--charts", action="store_true", help="Add charts to output workbook")
    args = ap.parse_args()

    # --- Auto-detect config.yaml if not found ---
    config_path = args.config
    if not os.path.exists(config_path):
        script_dir = os.path.dirname(__file__)
        config_path = os.path.join(script_dir, "config.yaml")
        if not os.path.exists(config_path):
            raise FileNotFoundError("Could not find config.yaml. Place it in the 'code' folder.")

    cfg = load_config(config_path)

    # --- Load configuration details ---
    aliases = cfg.get("column_aliases", {})
    rc_rules = cfg.get("root_cause_rules", {})
    out_cfg = cfg.get("output", {})
    add_charts = args.charts or bool(out_cfg.get("add_charts", False))

    os.makedirs(args.outdir, exist_ok=True)

    # --- Auto-detect Excel file if not found ---
    if not os.path.exists(args.input):
        script_dir = os.path.dirname(__file__)
        fallback_path = os.path.join(script_dir, "../data/msr_sample.xlsx")
        fallback_path = os.path.abspath(fallback_path)
        if os.path.exists(fallback_path):
            args.input = fallback_path
        else:
            raise FileNotFoundError(f"Input Excel not found at: {args.input} or {fallback_path}")

    # --- Read Excel safely ---
    df = pd.read_excel(args.input, sheet_name=0)

    # --- Normalize column names ---
    df = normalize_columns(df, aliases)

    # --- Determine root causes ---
    rc_col = safe_get_col(df, "root_cause")
    df["RootCauseFinal"] = df.apply(lambda r: coalesce_root_cause(r, rc_col, rc_rules), axis=1)

    # --- Clean type column ---
    if "type" in df.columns:
        df["type"] = df["type"].astype(str).str.strip().str.title()

    # --- Build pivot tables ---
    wanted = [
        "type", "priority", "state", "category", "subcategory",
        "RootCauseFinal", "assignment_group", "customer", "sla_breached"
    ]
    pivots = {col: build_pivot(df, col) for col in wanted}

    # --- Save output workbook ---
     # Always put the output inside the "out" folder next to this script
    script_dir = os.path.dirname(__file__)
    out_folder = os.path.abspath(os.path.join(script_dir, "../out"))
    os.makedirs(out_folder, exist_ok=True)
    out_file = os.path.join(out_folder, out_cfg.get("filename", "MSR_Summary.xlsx"))

    os.makedirs(os.path.dirname(out_file), exist_ok=True)

    with pd.ExcelWriter(out_file, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        total = len(df)
        overview = pd.DataFrame({"Metric": ["Total Tickets"], "Value": [total]})
        overview.to_excel(writer, index=False, sheet_name="Overview")

        sheet_order = out_cfg.get("sheets", [])
        name_map = {
            "Type": "type",
            "Priority": "priority",
            "State": "state",
            "Category": "category",
            "Subcategory": "subcategory",
            "RootCause": "RootCauseFinal",
            "AssignmentGroup": "assignment_group",
            "Customer": "customer",
            "SLA": "sla_breached",
        }

        for sheet in sheet_order:
            key = name_map.get(sheet, sheet)
            df_p = pivots.get(key, pd.DataFrame())
            df_p.to_excel(writer, index=False, sheet_name=sheet)

        # Details tab
        cols_keep = [
            c for c in [
                "ticket_id", "type", "priority", "state", "category",
                "subcategory", "short_description", "RootCauseFinal",
                "assignment_group", "assignee", "opened_at", "resolved_at",
                "sla_breached", "customer"
            ]
            if c in df.columns or c == "RootCauseFinal"
        ]
        df[cols_keep].to_excel(writer, index=False, sheet_name="Details")

        # Adjust formatting
        for ws_name, ws in writer.sheets.items():
            ws.set_column(0, 6, 20)
            ws.set_column(6, 6, 60)

        # Optional charts
        if add_charts:
            wb = writer.book
            for sheet in ["Type", "Priority", "State", "RootCause"]:
                if sheet in writer.sheets and not pivots.get(name_map.get(sheet, sheet), pd.DataFrame()).empty:
                    ws = writer.sheets[sheet]
                    chart = wb.add_chart({"type": "column"})
                    max_row = len(pivots[name_map.get(sheet, sheet)]) + 1
                    chart.add_series({
                        "name": "Count",
                        "categories": f"='{sheet}'!$A$2:$A${max_row}",
                        "values": f"='{sheet}'!$B$2:$B${max_row}",
                    })
                    chart.set_title({"name": f"{sheet} Distribution"})
                    ws.insert_chart("E2", chart)

    print(json.dumps({"summary_path": out_file}, indent=2))
    print(f"\n✅ MSR Summary generated successfully at: {out_file}")

    # --- (Optional) Auto-open Excel file on Windows ---
    try:
        if sys.platform.startswith("win"):
            os.startfile(out_file)
    except Exception:
        pass


# ----------------- ENTRY POINT -----------------
if __name__ == "__main__":
    main()
