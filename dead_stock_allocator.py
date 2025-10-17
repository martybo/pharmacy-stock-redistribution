#!/usr/bin/env python3
"""
Dead Stock Allocator  (v13.2)
-----------------------------
Fixes:
- "Reason" is now set AFTER allocation from a definitive Store->AM map:
  • If destination == f"{source} NHD" and we allocated >0: "NHD priority"
  • Else "Same AM" if AM(source) == AM(dest), otherwise "Cross AM"
- Optional override map: store_am_map.csv (columns: Store,AM) beside the script.

Also includes (from v13.1):
- Sender workbooks (“<Store> - Dead_Stock_Transfers.xlsx”)
    → “Flat Transfers” sorted by Destination, then Trade Description.
    → Blank “Sent QTY” column right of “Qty”.
- Receiver workbook (Dead_Stock_Receiver_Schedule.xlsx)
    → Each per-destination tab sorted by Trade Description only.
    → Blank “Received QTY” right of “Qty”.
"""

import os, re, glob, sys, time, csv
from typing import Dict, List, Tuple, Callable, Set
import pandas as pd
import numpy as np

# ---------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------
FILE_GLOB = "* - Dead Stock & Transfers By Medicine.xlsx"
SHEET_FALLBACKS = ["Sheet 1", "Sheet1", 0]
USAGE_EPS = 1e-6
WEIGHT_EXPONENT = 0.9
ASSUME_BASELINE_STOCK_AT_RECEIVER = 1
OUTPUT_DIR_NAME = "Dead_Stock_Transfer_Output"
MASTER_SUMMARY_NAME = "Dead_Stock_Summary.xlsx"
RECEIVER_WORKBOOK_NAME = "Dead_Stock_Receiver_Schedule.xlsx"
RETRY_ATTEMPTS = 3
RETRY_SLEEP_SEC = 0.5
# ---------------------------------------------------------------------


# ---------- Helpers to read the inbound matrix ----------
def open_matrix(path: str):
    last_err = None
    for sn in SHEET_FALLBACKS:
        try:
            df = pd.read_excel(path, sheet_name=sn, header=1)
            return df, str(sn)
        except Exception as e:
            last_err = e
    raise RuntimeError(
        f"Failed to open sheet in '{os.path.basename(path)}'; tried {SHEET_FALLBACKS}. Last error: {last_err}"
    )


def read_store_matrix(path: str):
    """
    Returns:
      df: tidy 3-row-per-product layout (usage_4m, st_packs, st_value)
      store_to_am_heur: heuristic AM map from column header (may be incomplete)
      store_columns: list[(store_name, original_col_name)]
    """
    df, chosen = open_matrix(path)
    if df.shape[0] < 3 or df.shape[1] < 5:
        raise RuntimeError(
            f"Unexpected shape in '{os.path.basename(path)}' (sheet {chosen}): {df.shape}"
        )

    # row 0 carries store names
    store_names_row = df.iloc[0, :]
    store_to_am_heur: Dict[str, str] = {}
    store_columns: List[Tuple[str, str]] = []
    for j, col in enumerate(df.columns):
        if j < 3:
            continue
        store_name = (
            str(store_names_row.iloc[j]).strip()
            if pd.notna(store_names_row.iloc[j])
            else ""
        )
        if not store_name:
            continue
        # Heuristic AM string from column header (e.g., "RL.Ballycolman")
        am_guess = ""
        col_s = str(col)
        if "." in col_s:
            am_guess = col_s.split(".", 1)[0].strip()
        elif " - " in col_s:
            am_guess = col_s.split(" - ", 1)[0].strip()
        elif ":" in col_s:
            am_guess = col_s.split(":", 1)[0].strip()
        store_to_am_heur[store_name] = am_guess
        store_columns.append((store_name, col))

    if not store_columns:
        raise RuntimeError(
            f"No store columns detected in '{os.path.basename(path)}' (sheet {chosen})."
        )

    # Rename first 3 columns
    df = df.rename(
        columns={
            df.columns[0]: "virtual_product",
            df.columns[1]: "trade_description",
            df.columns[2]: "metric",
        }
    ).copy()

    # Filter to expected metrics, forward-fill product & description
    valid = {
        "Packs Disp Last 4 Months": "usage_4m",
        "ST Packs Count": "st_packs",
        "Stocktake Value": "st_value",
    }
    df = df[df["metric"].isin(valid.keys())].copy()
    df["metric"] = df["metric"].map(valid)
    df["virtual_product"] = df["virtual_product"].ffill()
    df["trade_description"] = df["trade_description"].ffill()
    # remove totals/grand totals lines
    df = df.loc[
        ~df["trade_description"].astype(str).str.contains(r"\btotal\b", case=False, na=False)
    ].copy()
    return df, store_to_am_heur, store_columns


def infer_source_store(df, store_columns, filename_hint):
    m = re.match(r"(.+?)\s*-\s*Dead Stock", os.path.basename(filename_hint), flags=re.I)
    if m:
        name = m.group(1).strip()
        for s, _c in store_columns:
            if s.lower() == name.lower():
                return s
    st = df[df["metric"] == "st_packs"]
    counts = {s: st[c].notna().sum() for (s, c) in store_columns}
    return max(counts, key=counts.get) if counts else store_columns[0][0]


def build_product_blocks(df, store_columns):
    usage = df[df["metric"] == "usage_4m"].reset_index(drop=True)
    stpk  = df[df["metric"] == "st_packs"].reset_index(drop=True)
    stvl  = df[df["metric"] == "st_value"].reset_index(drop=True)
    n = min(len(usage), len(stpk), len(stvl))
    usage, stpk, stvl = usage.iloc[:n], stpk.iloc[:n], stvl.iloc[:n]
    merged = pd.DataFrame(
        {"virtual_product": usage["virtual_product"], "trade_description": usage["trade_description"]}
    )
    for (s, col) in store_columns:
        merged[f"usage::{s}"] = pd.to_numeric(usage.get(col), errors="coerce")
        merged[f"st::{s}"]    = pd.to_numeric(stpk.get(col), errors="coerce")
        merged[f"stval::{s}"] = pd.to_numeric(stvl.get(col), errors="coerce")
    return merged


# ---------- Allocation ----------
def allocate_for_product(row, source_store: str, store_to_am: Dict[str, str], stores: List[str]):
    avail = row.get(f"st::{source_store}", np.nan)
    if pd.isna(avail) or avail <= 0:
        return []
    avail = int(round(avail))

    candidates = []
    for s in stores:
        if s == source_store:
            continue
        usage = row.get(f"usage::{s}", np.nan)
        if pd.isna(usage) or usage <= USAGE_EPS:
            continue
        has_st = not pd.isna(row.get(f"st::{s}", np.nan))
        baseline = ASSUME_BASELINE_STOCK_AT_RECEIVER if not has_st else 0
        capacity = max(0, int(round(usage)) - baseline)
        if capacity > 0:
            candidates.append((s, float(usage), capacity))

    if not candidates:
        return []

    allocations = []
    remaining = avail

    # 1) NHD priority
    nhd = f"{source_store} NHD"
    for i, (s, usage, cap) in enumerate(list(candidates)):
        if s.lower() == nhd.lower() and cap > 0:
            qty = min(remaining, cap)
            if qty > 0:
                allocations.append((s, qty, "NHD priority"))
                remaining -= qty
            # reduce/remove this candidate
            candidates[i] = (s, usage, cap - qty)
            candidates = [c for c in candidates if c[2] > 0]
            break

    if remaining <= 0 or not candidates:
        return allocations

    # Split remaining across same-AM first, then others (weight ~ usage^exp)
    def spread(cands, qty):
        if qty <= 0 or not cands:
            return []
        usages = np.array([c[1] for c in cands], float)
        caps   = np.array([c[2] for c in cands], int)
        weights = np.power(np.maximum(usages, 0.0), WEIGHT_EXPONENT)
        if weights.sum() <= 0:
            weights = np.ones_like(weights)
        desired = weights / weights.sum() * qty
        base = np.floor(desired).astype(int)
        base = np.minimum(base, caps)
        alloc = base.copy()
        rem = qty - alloc.sum()
        remainders = desired - base
        for idx in np.argsort(-remainders):
            if rem <= 0:
                break
            if alloc[idx] < caps[idx]:
                alloc[idx] += 1
                rem -= 1
        return [(cands[i][0], int(alloc[i])) for i in range(len(cands)) if alloc[i] > 0]

    # NOTE: AM fairness and priority are handled later when we label reason.
    # Here we only spread by need (usage & capacity).
    for s, q in spread(candidates, remaining):
        allocations.append((s, q, ""))  # reason will be set later
        remaining -= q

    return allocations


# ---------- Utilities ----------
def make_output_folder(base):
    out = os.path.join(base, OUTPUT_DIR_NAME)
    os.makedirs(out, exist_ok=True)
    return out


def safe_write_excel(path, fn):
    for i in range(RETRY_ATTEMPTS):
        try:
            fn(path)
            print("[OK] Wrote", path)
            return
        except Exception as e:
            print(f"[WARN] Write attempt {i+1} failed for {path}: {e}")
            time.sleep(RETRY_SLEEP_SEC)
    print("[ERROR] Could not write", path)


def load_store_am_override(script_dir: str) -> Dict[str, str]:
    """
    Optional CSV beside the script: store_am_map.csv with columns: Store,AM
    Returns lowercase-key dict for robust lookup.
    """
    out: Dict[str, str] = {}
    csv_path = os.path.join(script_dir, "store_am_map.csv")
    if not os.path.exists(csv_path):
        return out
    try:
        with open(csv_path, newline="", encoding="utf-8-sig") as f:
            rdr = csv.DictReader(f)
            for row in rdr:
                store = (row.get("Store") or "").strip()
                am = (row.get("AM") or "").strip()
                if store:
                    out[store.lower()] = am
    except Exception as e:
        print(f"[WARN] Could not read store_am_map.csv: {e}")
    return out


# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------
def main():
    base = sys.argv[1] if len(sys.argv) > 1 else "."
    script_dir = os.path.dirname(os.path.abspath(__file__))

    files = sorted(glob.glob(os.path.join(base, FILE_GLOB)))
    if not files:
        print("No input files found.")
        return 1

    outdir = make_output_folder(base)

    # Optional override AM map
    am_override = load_store_am_override(script_dir)

    routes = []      # (src, dest, vp, desc, qty, note_tmp, est_val)  note_tmp may be blank / provisional
    per_source = {}  # src -> list of rows for sender file
    unalloc = []

    # Heuristic map aggregated across files (fallback if no override)
    global_heur_am: Dict[str, str] = {}

    for path in files:
        df, store_to_am_heur, cols = read_store_matrix(path)
        # merge heuristics
        for k, v in store_to_am_heur.items():
            if k not in global_heur_am and v:
                global_heur_am[k] = v

        src = infer_source_store(df, cols, path)
        merged = build_product_blocks(df, cols)
        stores = [s for s, _ in cols]

        for _, row in merged.iterrows():
            stp = row.get(f"st::{src}", np.nan)
            stv = row.get(f"stval::{src}", np.nan)
            if pd.isna(stp) or stp <= 0:
                continue
            # dead stock only (no recent usage at source)
            if not (pd.isna(row.get(f"usage::{src}", np.nan)) or row.get(f"usage::{src}", np.nan) <= USAGE_EPS):
                continue

            unit = stv / stp if (pd.notna(stv) and stv > 0) else np.nan
            allocs = allocate_for_product(row, src, store_to_am_heur, stores)

            total = 0
            for dest, qty, note_tmp in allocs:
                est = qty * unit if pd.notna(unit) else np.nan
                routes.append((src, dest, row["virtual_product"], row["trade_description"], int(qty), note_tmp, est))
                per_source.setdefault(src, []).append({
                    "Virtual Product": row["virtual_product"],
                    "Trade Description": row["trade_description"],
                    "ST Packs (source)": int(round(stp)),
                    "Destination": dest,
                    "Qty": int(qty),
                    "Est Value (£)": est,
                    # Reason will be recalculated later
                })
                total += qty

            un = int(round(stp)) - total
            if un > 0:
                unalloc.append({
                    "Source Store": src,
                    "Virtual Product": row["virtual_product"],
                    "Trade Description": row["trade_description"],
                    "ST Packs (source)": int(round(stp)),
                    "Unallocated": un,
                    "Est Unit (£)": unit if pd.notna(unit) else np.nan,
                    "Est Unallocated (£)": unit * un if pd.notna(unit) else np.nan,
                })

    # ----- Build definitive AM map (override first, then heuristic) -----
    def am_of(store: str) -> str:
        if not store:
            return ""
        s = store.strip().lower()
        if s in am_override:
            return am_override[s]
        return global_heur_am.get(store, "")

    # ----- Recompute Reason consistently -----
    # Preserve "NHD priority" if that was explicitly set; otherwise recompute by AM equality.
    def recompute_reason(src: str, dest: str, prev_note: str, qty: int) -> str:
        if prev_note and prev_note.lower() == "nhd priority":
            return "NHD priority"
        if dest.strip().lower() == f"{src.strip().lower()} nhd" and qty > 0:
            return "NHD priority"
        a_src = am_of(src)
        a_dst = am_of(dest)
        if a_src and a_dst and (a_src == a_dst):
            return "Same AM"
        return "Cross AM"

    routes_df = pd.DataFrame(routes, columns=[
        "Source Store","Destination Store","Virtual Product","Trade Description","Qty","_Note_TMP","Est Value (£)"
    ])
    if not routes_df.empty:
        routes_df["Reason"] = routes_df.apply(
            lambda r: recompute_reason(r["Source Store"], r["Destination Store"], r["_Note_TMP"], r["Qty"]), axis=1
        )
        routes_df.drop(columns=["_Note_TMP"], inplace=True)

    # ----------- Summaries -----------
    unalloc_df = pd.DataFrame(unalloc)
    route_tot = (routes_df.groupby(["Source Store","Destination Store"], as_index=False)
                 .agg(Qty=("Qty","sum"), **{"Est Value (£)":("Est Value (£)","sum")}))

    receipts = routes_df.rename(columns={
        "Destination Store":"Destination",
        "Source Store":"From"
    })[["Destination","From","Virtual Product","Trade Description","Qty","Est Value (£)","Reason"]]

    bydest = (receipts.groupby("Destination", as_index=False)
              .agg(Qty=("Qty","sum"), **{"Est Value (£)":("Est Value (£)","sum")}))

    un_by_src = (unalloc_df.groupby("Source Store", as_index=False)
                 .agg(**{"Unallocated Qty":("Unallocated","sum"),
                         "Est Unallocated (£)":("Est Unallocated (£)","sum")})) if not unalloc_df.empty else pd.DataFrame(columns=["Source Store","Unallocated Qty","Est Unallocated (£)"])

    # ---------------- Master Summary ----------------
    mpath = os.path.join(outdir, MASTER_SUMMARY_NAME)
    def write_master(p):
        with pd.ExcelWriter(p, engine="xlsxwriter") as xw:
            route_tot.to_excel(xw, "Summary", index=False)
            routes_df.to_excel(xw, "Routes (Detail)", index=False)
            receipts.to_excel(xw, "All Receipts (Flat)", index=False)
            bydest.to_excel(xw, "Summary (By Destination)", index=False)
            unalloc_df.to_excel(xw, "Unallocated (Detail)", index=False)
            un_by_src.to_excel(xw, "Unallocated (By Source)", index=False)
    safe_write_excel(mpath, write_master)

    # ---------------- Sender Workbooks ----------------
    # Sort rule: Destination, then Trade Description
    for src, rows in per_source.items():
        df_src = pd.DataFrame(rows)
        if not df_src.empty:
            # attach correct Reason into per_source rows too
            df_src["Reason"] = df_src.apply(
                lambda r: recompute_reason(src, r["Destination"], "", r["Qty"]), axis=1
            )
        path = os.path.join(outdir, f"{src} - Dead_Stock_Transfers.xlsx")

        def write_src(p):
            with pd.ExcelWriter(p, engine="xlsxwriter") as xw:
                if df_src.empty:
                    pd.DataFrame([{"Note":"No allocations"}]).to_excel(xw, "Flat Transfers", index=False)
                    return
                flat = df_src.drop(columns=["Est Value (£)"], errors="ignore").copy()
                flat.insert(0, "From", src)
                sort_cols = [c for c in ["Destination","Trade Description"] if c in flat.columns]
                if sort_cols:
                    flat = flat.sort_values(sort_cols, kind="mergesort", ignore_index=True)
                if "Qty" in flat.columns:
                    i = list(flat.columns).index("Qty") + 1
                    flat.insert(i, "Sent QTY", "")
                flat.to_excel(xw, "Flat Transfers", index=False)
        safe_write_excel(path, write_src)

    # ---------------- Receiver Workbook ----------------
    # Sort rule: Trade Description only
    rpath = os.path.join(outdir, RECEIVER_WORKBOOK_NAME)
    def write_rx(p):
        with pd.ExcelWriter(p, engine="xlsxwriter") as xw:
            if receipts.empty:
                pd.DataFrame([{"Note":"No receipts"}]).to_excel(xw, "No Receipts", index=False)
                return
            for dest, sub in receipts.groupby("Destination"):
                sheet_name = re.sub(r"[:\\/?*\[\]]", "_", str(dest))[:31]
                if "Trade Description" in sub.columns:
                    sub = sub.sort_values(["Trade Description"], kind="mergesort", ignore_index=True)
                sub2 = sub.drop(columns=["Est Value (£)"], errors="ignore").copy()
                if "Qty" in sub2.columns:
                    i = list(sub2.columns).index("Qty") + 1
                    sub2.insert(i, "Received QTY", "")
                sub2.to_excel(xw, sheet_name=sheet_name, index=False)
    safe_write_excel(rpath, write_rx)

    print("[DONE] Master:", mpath)
    print("Receiver workbook:", rpath)
    print("Per-store files:", outdir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
