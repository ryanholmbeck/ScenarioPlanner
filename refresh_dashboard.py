"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         GERON S&OP DASHBOARD — REFRESH SCRIPT  (Enhanced Workbook)         ║
║         Reads the Enhanced workbook and rebuilds the HTML dashboard         ║
║                                                                             ║
║  HOW TO RUN:                                                                ║
║    1. Make sure Excel has the workbook OPEN and fully calculated            ║
║       (Ctrl+Alt+F9 to force recalc, then Ctrl+S to save)                   ║
║    2. Open Command Prompt in this folder                                    ║
║    3. Type:  python refresh_dashboard.py                                    ║
║    4. Press Enter — done. Open the HTML file in your browser.               ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import os, sys, json, re, math, shutil
from datetime import datetime, date
from pathlib import Path

# ── AUTO-INSTALL DEPENDENCIES ──────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    print("📦 Installing openpyxl...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--quiet"])
    import openpyxl

# ══════════════════════════════════════════════════════════════════════════════
#  PATHS — Edit these if you move the files
# ══════════════════════════════════════════════════════════════════════════════

WORKBOOK_PATH = r"C:\Users\rholmbeck\OneDrive - Geron Corporation\Documents\Current Files\Demand_Inventory_Supply Planning_Expiry_Workbook.xlsm"
HTML_IN       = Path(r"C:\Users\rholmbeck\OneDrive - Geron Corporation\Documents\Current Files\Geron_SOP_Executive_Dashboard.html")
HTML_OUT      = HTML_IN  # overwrite in place

# Scenario Planner — set to None to skip, or point to the file path
SCENARIO_PLANNER_PATH = Path(r"C:\Users\rholmbeck\OneDrive - Geron Corporation\Documents\Current Files\Geron_SOP_Scenario_Planner.html")

# ══════════════════════════════════════════════════════════════════════════════
#  WORKBOOK STRUCTURE MAP  (Enhanced workbook — verified Mar 2026)
#
#  Controls        B5=model_start  B6=dash_start  B7=lookahead  B15=spotlight
#                  Row 18: headers, Rows 19-28: product parameters (cols A-H)
#                  Col A=Product  D=StartInv  E=BatchSize  F=LeadTime
#                  G=TargetMOH  H=CriticalMOH
#
#  Planning Engine Row 3: headers, Rows 4+: data
#                  A=Month  B=Product  C=Category  D=Unit
#                  E=BI  F=Demand(Agg)  G=PlannedProd  H=EI
#                  I=AvgFutureDemand  J=MOH  K=ProdRequired  L=RiskStatus
#                  R=PrimarySupplier  S=ReleaseMonth  T=ReleaseQty
#
#  Lot Inputs      Row 4: headers, Rows 5+: data
#                  Col 1=Product  2=LotID  3=Category  4=ReceiptDate
#                  5=ExpiryDate  6=QtyAvailable  7=QtyReserved  8=NetAvailable
#
#  Cost Inputs     Row 5: headers (Table 1)  Rows 6-18: unit costs
#                  Col 1=Product  4=Price/Unit  5=BatchCost
#                  Row 21: budget headers  Rows 22+: budget by CDMO
#
#  Dependencies    Row 3: headers  Rows 4+: parent/child/factor pairs
#                  Col 1=Parent  2=Child  4=Factor
#
#  Expiry Risk     Row 3: headers  Rows 4+: per-product risk summary
#  Summary         Col 1=Product  3=NetLotInv  7=DemandNext6m  9=RiskStatus
# ══════════════════════════════════════════════════════════════════════════════

# PRODUCT_META is now built dynamically from Controls + Supplier Registry by
# build_product_meta(wb).  This stub is only used as a last-resort fallback
# when the workbook cannot be opened (e.g. during unit tests).
_PRODUCT_META_FALLBACK = {}


def build_product_meta(wb):
    """
    Build the PRODUCT_META dict dynamically from the workbook.

    Priority order per field:
      1. Controls sheet  rows 19-34 (cols A-H)
      2. Supplier Registry rows 5+ (cols A = product, B = supplier,
         C = category, D = unit, E = batch_size, H = overstock_moh (opt))
      3. Sensible defaults

    Returns a dict keyed by product name.
    """
    meta = {}

    # ── Pass 1: Supplier Registry (broader product list) ─────────────────────
    ws_sr = get_sheet(wb, "Supplier Registry")
    if ws_sr:
        for ri_vals in ws_sr.iter_rows(min_row=5, values_only=True):
            row = list(ri_vals)
            if not row[0] or not row[1]: continue
            prod     = safe_str(row[0])
            supplier = safe_str(row[1])
            cat      = safe_str(row[2]) if len(row) > 2 and row[2] else ""
            unit     = safe_str(row[3]) if len(row) > 3 and row[3] else "units"
            if not prod or prod in ("Product", "product"): continue
            if prod not in meta:
                meta[prod] = {
                    "category":     cat,
                    "unit":         unit,
                    "target_moh":   12,
                    "critical_moh": 6,
                    "overstock_moh": 18,
                    "supplier":     supplier,
                    "batch_size":   0,
                    "lead_time":    0,
                }
            else:
                # Update supplier if blank; don't overwrite category/unit once set
                if not meta[prod].get("supplier"):
                    meta[prod]["supplier"] = supplier
                if not meta[prod].get("category"):
                    meta[prod]["category"] = cat

    # ── Pass 2: Controls sheet (rows 19-34) — authoritative for thresholds ───
    ws_ctrl = get_sheet(wb, "Controls")
    if ws_ctrl:
        for ri in range(19, 35):
            row = list(ws_ctrl.iter_rows(min_row=ri, max_row=ri, values_only=True))[0]
            if not row[0]: continue
            prod = safe_str(row[0])
            if not prod or prod in ("Product", "product"): continue
            target_moh   = safe_float(row[6]) if len(row) > 6 and row[6] else 0
            critical_moh = safe_float(row[7]) if len(row) > 7 and row[7] else 0
            batch_size   = safe_float(row[4]) if len(row) > 4 and row[4] else 0
            lead_time    = safe_float(row[5]) if len(row) > 5 and row[5] else 0
            if prod not in meta:
                meta[prod] = {
                    "category":     "",
                    "unit":         "units",
                    "target_moh":   target_moh or 12,
                    "critical_moh": critical_moh or 6,
                    "overstock_moh": 18,
                    "supplier":     "",
                    "batch_size":   batch_size,
                    "lead_time":    lead_time,
                }
            else:
                if target_moh:   meta[prod]["target_moh"]   = target_moh
                if critical_moh: meta[prod]["critical_moh"] = critical_moh
                if batch_size:   meta[prod]["batch_size"]   = batch_size
                if lead_time:    meta[prod]["lead_time"]    = lead_time

    # ── Fill unit/category defaults from Planning Engine header rows if blank ─
    ws_pe = get_sheet(wb, "Planning Engine")
    if ws_pe:
        seen = {}
        for ri_vals in ws_pe.iter_rows(min_row=4, values_only=True):
            row = list(ri_vals)
            if not row[1]: continue
            prod = safe_str(row[1])
            if prod in seen: continue
            cat  = safe_str(row[2]) if len(row) > 2 and row[2] else ""
            unit = safe_str(row[3]) if len(row) > 3 and row[3] else ""
            if prod in meta:
                if cat  and not meta[prod].get("category"): meta[prod]["category"] = cat
                if unit and not meta[prod].get("unit"):     meta[prod]["unit"] = unit
            seen[prod] = True

    return meta


# Module-level cache: populated in main() after workbook is opened.
PRODUCT_META = _PRODUCT_META_FALLBACK

# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def safe_float(v, default=0.0):
    if v is None or v == "" or v == "-": return default
    try: return float(v)
    except: return default

def safe_str(v):
    return "" if v is None else str(v).strip()

def to_iso(v):
    if v is None: return None
    if hasattr(v, "strftime"): return v.strftime("%Y-%m-%d")
    s = str(v)[:10]
    try: datetime.strptime(s, "%Y-%m-%d"); return s
    except: return None

def to_iso_01(v):
    iso = to_iso(v)
    if iso: return iso[:7] + "-01"
    return None

def fmt_label(v):
    if isinstance(v, (datetime, date)): return v.strftime("%b %Y")
    iso = to_iso(v)
    if iso:
        try: return datetime.strptime(iso, "%Y-%m-%d").strftime("%b %Y")
        except: pass
    return safe_str(v)

def open_wb(path):
    """Open workbook safely — works even if Excel has it open."""
    try:
        tmp = str(path) + ".~refresh_tmp.xlsx"
        shutil.copy2(str(path), tmp)
        wb = openpyxl.load_workbook(tmp, data_only=True)
        return wb, tmp
    except Exception as e:
        print(f"  ⚠  copy method failed ({e}), trying direct open...")
        try:
            wb = openpyxl.load_workbook(str(path), data_only=True)
            return wb, None
        except Exception as e2:
            return None, str(e2)

def get_sheet(wb, name):
    for s in wb.sheetnames:
        if s.strip().lower() == name.strip().lower(): return wb[s]
    for s in wb.sheetnames:
        if name.lower() in s.lower():
            print(f"  ⚠  Using '{s}' for '{name}'")
            return wb[s]
    return None

# ══════════════════════════════════════════════════════════════════════════════
#  READ CONTROLS
# ══════════════════════════════════════════════════════════════════════════════

def read_controls(wb):
    ws = get_sheet(wb, "Controls")
    if not ws:
        return {"model_start": None, "dash_start": None, "lookahead": 6,
                "spotlight": "188mg LP", "products": {}}

    model_start = ws["B5"].value
    dash_start  = ws["B6"].value
    lookahead   = int(safe_float(ws["B7"].value, 6))
    spotlight   = safe_str(ws["B15"].value) or "188mg LP"

    # Product parameters — rows 19-28, cols A=prod, D=start_inv, E=batch,
    # F=lead_time, G=target_moh, H=critical_moh
    products = {}
    for ri in range(19, 35):
        row = list(ws.iter_rows(min_row=ri, max_row=ri, values_only=True))[0]
        if not row[0]: continue
        prod = safe_str(row[0])
        if not prod or prod in ("Product", "product"): continue
        meta = PRODUCT_META.get(prod, {}).copy()
        meta["start_inv"]    = safe_float(row[3])  # col D (index 3)
        meta["batch_size"]   = safe_float(row[4])  # col E
        meta["lead_time"]    = safe_float(row[5])  # col F
        meta["target_moh"]   = safe_float(row[6]) or meta["target_moh"]   # col G
        meta["critical_moh"] = safe_float(row[7]) or meta["critical_moh"] # col H
        products[prod] = meta

    if isinstance(model_start, datetime): model_start = model_start.date()
    if isinstance(dash_start,  datetime): dash_start  = dash_start.date()

    return {
        "model_start": model_start,
        "dash_start":  dash_start,
        "lookahead":   lookahead,
        "spotlight":   spotlight,
        "products":    products,
    }

# ══════════════════════════════════════════════════════════════════════════════
#  READ PLANNING ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def read_planning_engine(wb, controls):
    ws = get_sheet(wb, "Planning Engine")
    if not ws:
        print("  ❌ Planning Engine sheet not found")
        return []

    # Use today as the filter floor
    today_first = datetime.now().replace(day=1).date()
    filter_dt   = today_first

    # Column positions (0-indexed):
    # A=0=Month  B=1=Product  C=2=Category  D=3=Unit
    # E=4=BI  F=5=Demand  G=6=PlannedProd  H=7=EI
    # I=8=AvgFutureDemand  J=9=MOH  K=10=ProdRequired  L=11=RiskStatus
    # R=17=PrimarySupplier  S=18=ReleaseMonth  T=19=ReleaseQty
    COL = dict(month=0, product=1, category=2, unit=3, bi=4, demand=5,
               prod=6, ei=7, avg_future=8, moh=9, prod_req=10, risk=11,
               supplier=17, rel_month=18, rel_qty=19)

    rows = []
    for ri_vals in ws.iter_rows(min_row=4, values_only=True):
        row = list(ri_vals)
        if not row[COL["month"]] or not row[COL["product"]]: continue
        prod = safe_str(row[COL["product"]])
        if not prod or prod == "Product": continue

        # Parse month
        m_val = row[COL["month"]]
        iso   = to_iso_01(m_val)
        if not iso: continue
        try:
            m_dt = datetime.strptime(iso, "%Y-%m-%d").date()
        except:
            continue
        if m_dt < filter_dt: continue

        meta = PRODUCT_META.get(prod, {})
        moh  = safe_float(row[COL["moh"]])
        risk = safe_str(row[COL["risk"]]) or _calc_risk(
            moh,
            meta.get("target_moh", 6),
            meta.get("critical_moh", 3),
            meta.get("overstock_moh", 18)
        )

        rows.append({
            "month":    iso,
            "ml":       fmt_label(m_val),
            "product":  prod,
            "category": safe_str(row[COL["category"]]) or meta.get("category", ""),
            "unit":     safe_str(row[COL["unit"]]) or meta.get("unit", "units"),
            "bi":       round(safe_float(row[COL["bi"]]), 2),
            "demand":   round(safe_float(row[COL["demand"]]), 2),
            "prod":     round(safe_float(row[COL["prod"]]), 2),
            "ei":       round(safe_float(row[COL["ei"]]), 2),
            "moh":      round(moh, 2),
            "risk":     risk,
            "rm":       to_iso_01(row[COL["rel_month"]]) or iso,
            "rq":       round(safe_float(row[COL["rel_qty"]]), 2),
            "supplier": safe_str(row[COL["supplier"]]) or meta.get("supplier", ""),
        })

    print(f"  ✅ Planning Engine: {len(rows)} rows ({len(set(r['product'] for r in rows))} products)")
    return rows

def _calc_risk(moh, target=6, critical=3, overstock=18):
    if moh == 0: return "No Demand"
    if moh < 0:  return "High"
    if moh < critical: return "High"
    if moh < target:   return "Medium"
    if moh > overstock: return "Overstock"
    return "Low"

# ══════════════════════════════════════════════════════════════════════════════
#  BUILD RAW METADATA
# ══════════════════════════════════════════════════════════════════════════════

def build_raw_meta(prod_data, controls):
    """Build per-product metadata for RAW JS constant."""
    raw = {}
    seen = {}
    for r in prod_data:
        p = r["product"]
        if p not in seen:
            seen[p] = r.get("bi", 0)
        if p in raw: continue
        meta = PRODUCT_META.get(p, {})
        ctrl_meta = controls.get("products", {}).get(p, {})
        raw[p] = {
            "category":     r.get("category") or meta.get("category", ""),
            "unit":         r.get("unit")     or meta.get("unit", "units"),
            "target_moh":   ctrl_meta.get("target_moh")   or meta.get("target_moh", 6),
            "critical_moh": ctrl_meta.get("critical_moh") or meta.get("critical_moh", 3),
            "overstock_moh":meta.get("overstock_moh", 18),
            "supplier":     r.get("supplier") or meta.get("supplier", ""),
            "batch_size":   ctrl_meta.get("batch_size")   or meta.get("batch_size", 0),
            "lead_time":    ctrl_meta.get("lead_time")    or meta.get("lead_time", 0),
        }
    return raw

# ══════════════════════════════════════════════════════════════════════════════
#  READ COST INPUTS
# ══════════════════════════════════════════════════════════════════════════════

def read_cost_inputs(wb):
    ws = get_sheet(wb, "Cost Inputs")
    if not ws: return {"unit_costs": {}, "budget": {}, "actuals": {}}

    unit_costs = {}
    budget     = {}
    actuals    = {}
    months_hdr = []

    mode = None
    for ri_vals in ws.iter_rows(values_only=True):
        row = list(ri_vals)
        if not any(v for v in row): continue
        first = safe_str(row[0]).upper()

        if "TABLE 1" in first and "UNIT COST" in first:
            mode = "unit_costs"; continue
        if "TABLE 2" in first and "BUDGET" in first:
            mode = "budget"; months_hdr = []; continue
        if "TABLE 3" in first and "ACTUAL" in first:
            mode = "actuals"; months_hdr = []; continue
        if "TABLE 4" in first or "TABLE 5" in first:
            mode = None; continue

        if mode == "unit_costs":
            if safe_str(row[0]).lower() == "product": continue
            prod = safe_str(row[0])
            if prod and row[3] is not None:
                price_per_unit = round(safe_float(row[3]), 4)   # col D
                batch_cost_raw = safe_float(row[4]) if len(row) > 4 and row[4] is not None else 0.0
                # If batch_cost is absent, estimate from unit_cost × batch_size
                if batch_cost_raw == 0.0 and price_per_unit > 0:
                    bsize = PRODUCT_META.get(prod, {}).get("batch_size", 0)
                    batch_cost_raw = round(price_per_unit * bsize, 2) if bsize else 0.0
                unit_costs[prod] = {
                    "price_per_unit": price_per_unit,
                    "batch_cost":     round(batch_cost_raw, 2),
                }

        elif mode in ("budget", "actuals"):
            target = budget if mode == "budget" else actuals
            # Detect month header row (has date-like values in cols 2+)
            if not months_hdr and row[0] and (isinstance(row[1], (datetime, date)) or
               (isinstance(row[1], str) and re.match(r'[A-Za-z]{3}-\d{2}', str(row[1])))):
                months_hdr = [fmt_label(v) if v else "" for v in row[1:]]
                continue
            elif not months_hdr and row[0] and row[1] and safe_str(row[1]).lower().startswith("jan"):
                months_hdr = [safe_str(v) for v in row[1:]]
                continue
            cdmo = safe_str(row[0])
            if cdmo and months_hdr:
                target[cdmo] = {}
                for i, mh in enumerate(months_hdr):
                    if mh and i+1 < len(row):
                        target[cdmo][mh] = round(safe_float(row[i+1]), 2)

    print(f"  ✅ Cost Inputs: {len(unit_costs)} products, {len(budget)} budget CDMOs, {len(actuals)} actual CDMOs")
    return {"unit_costs": unit_costs, "budget": budget, "actuals": actuals}

# ══════════════════════════════════════════════════════════════════════════════
#  READ SUPPLIER COSTS  (from Supplier Registry cols X / Y / Z)
#  Col X (idx 23) = Inv Value / UOM   — book value per unit or kg
#  Col Y (idx 24) = Prod Cost / Batch — what you pay CMO per batch (mutually exclusive with Z)
#  Col Z (idx 25) = Prod Cost / UOM   — what you pay CMO per unit/kg (mutually exclusive with Y)
#
#  Only primary active supplier rows are used for costs.
#  Returns cost_data dict compatible with COST_DATA JS constant:
#    { "products": { <prod>: { "inv_value_uom": float, "prod_cost_batch": float,
#                              "prod_cost_uom": float, "cost_by_batch": bool,
#                              "unit": str, "supplier": str } },
#      "budget": { ... },   <- preserved from Cost Inputs for backward compat
#      "actuals": { ... } }
# ══════════════════════════════════════════════════════════════════════════════

def read_supplier_costs(wb):
    """
    Read production and inventory costs from Supplier Registry cols X/Y/Z.
    Also reads budget/actuals from Cost Inputs for backward compatibility.
    """
    ws = get_sheet(wb, "Supplier Registry")
    products = {}

    if ws:
        for ri_vals in ws.iter_rows(min_row=5, values_only=True):
            row = list(ri_vals)
            if not row[0] or not row[1]: continue

            prod      = safe_str(row[0]).strip()
            is_primary = safe_float(row[12]) if len(row) > 12 and row[12] is not None else 0
            is_active  = safe_str(row[13]).strip().lower() if len(row) > 13 and row[13] else ""
            unit       = safe_str(row[3]).strip() if len(row) > 3 and row[3] else "units"
            supplier   = safe_str(row[1]).strip() if row[1] else ""

            # Only primary active rows carry cost data
            if is_primary != 1 or is_active != "active":
                continue

            inv_val    = safe_float(row[23]) if len(row) > 23 and row[23] is not None else 0.0
            prod_batch = safe_float(row[24]) if len(row) > 24 and row[24] is not None else 0.0
            prod_uom   = safe_float(row[25]) if len(row) > 25 and row[25] is not None else 0.0

            # Skip rows with no cost data at all
            if inv_val == 0 and prod_batch == 0 and prod_uom == 0:
                continue

            # If product already registered, don't overwrite (first primary wins)
            if prod in products:
                continue

            products[prod] = {
                "inv_value_uom":   round(inv_val,    4),
                "prod_cost_batch": round(prod_batch, 2),
                "prod_cost_uom":   round(prod_uom,   4),
                "cost_by_batch":   prod_batch > 0,   # True = batch-costed, False = UOM-costed
                "unit":            unit,
                "supplier":        supplier,
            }

    # Also pull budget/actuals from Cost Inputs for exec dashboard backward compat
    ci = read_cost_inputs(wb)

    n = len(products)
    print(f"  ✅ Supplier costs: {n} products with cost data")
    for p, v in sorted(products.items()):
        cost_str = f"${v['prod_cost_batch']:,.0f}/batch" if v['cost_by_batch'] else f"${v['prod_cost_uom']:,.2f}/uom"
        print(f"     {p}: inv=${v['inv_value_uom']:,.2f}/{v['unit']}  prod={cost_str}")

    return {
        "products": products,
        "budget":   ci.get("budget",  {}),
        "actuals":  ci.get("actuals", {}),
        # Backward-compat unit_costs for exec dashboard (price_per_unit / batch_cost)
        "unit_costs": {
            p: {
                "price_per_unit": v["inv_value_uom"],
                "batch_cost":     v["prod_cost_batch"] if v["cost_by_batch"] else 0,
            }
            for p, v in products.items()
        },
    }


# ══════════════════════════════════════════════════════════════════════════════
#  READ DEPENDENCIES
# ══════════════════════════════════════════════════════════════════════════════

def read_dependencies(wb):
    """Build deps dict: child_product -> factor (first/primary row).
    Backward-compatible flat shape used by older HTML code.
    """
    ws = get_sheet(wb, "Dependencies")
    if not ws: return {}
    deps = {}
    for ri_vals in ws.iter_rows(min_row=4, values_only=True):
        row = list(ri_vals)
        if not row[0] or not row[1]: continue
        child  = safe_str(row[1])
        factor = safe_float(row[3])  # col D
        if child and child not in deps and factor:
            deps[child] = factor
    print(f"  ✅ Dependencies: {len(deps)} child products mapped")
    return deps


def read_dependencies_full(wb):
    """Rich dependency map keyed by (parent, child) -> list of supplier-specific
    factors with active-date windows. Columns:
      A=Parent  B=Child  C=Supplier  D=Factor  E=Basis  F=ActStart  G=ActEnd
    """
    ws = get_sheet(wb, "Dependencies")
    if not ws: return {}
    deps_full = {}
    rows_read = 0
    for ri_vals in ws.iter_rows(min_row=4, values_only=True):
        row = list(ri_vals)
        if not row or len(row) < 4: continue
        if not row[0] or not row[1]: continue
        parent  = safe_str(row[0])
        child   = safe_str(row[1])
        supplier = safe_str(row[2]) if len(row) > 2 else ""
        factor  = safe_float(row[3])
        basis   = safe_str(row[4]) if len(row) > 4 else ""
        act_from = to_iso_01(row[5]) if len(row) > 5 else None
        act_to   = to_iso_01(row[6]) if len(row) > 6 else None
        if not parent or not child or not factor: continue
        key = f"{parent}||{child}"
        deps_full.setdefault(key, []).append({
            "parent":    parent,
            "child":     child,
            "supplier":  supplier,
            "factor":    factor,
            "basis":     basis,
            "active_from": act_from,
            "active_through": act_to,
        })
        rows_read += 1
    print(f"  ✅ Dependencies (full): {rows_read} rows across {len(deps_full)} parent/child pairs")
    return deps_full


def read_supplier_registry(wb):
    """Read Supplier Registry tab → keyed by product:
      [{supplier, batch_size, lead_time, target_moh, critical_moh, overstock_moh,
        active_from, active_through, manual_toggle, primary, binding_mo, category, unit}, ...]
    Columns (1-indexed from row 4):
      A=1 Product, B=2 Supplier, C=3 Category, D=4 Unit, E=5 BatchSize,
      F=6 LeadTime(mo), G=7 TargetMOH, H=8 CriticalMOH, I=9 OverstockMOH,
      J=10 ActiveStart, K=11 ActiveEnd, L=12 ManualToggle, M=13 Primary,
      N=14 IsActive(formula), O=15 Notes, P=16 BindingPeriod(mo)
    """
    ws = get_sheet(wb, "Supplier Registry")
    if not ws:
        print("  ⚠  Supplier Registry tab not found — SUPPLIER_REGISTRY will be empty")
        return {}
    registry = {}
    rows_read = 0
    for ri_vals in ws.iter_rows(min_row=5, values_only=True):
        row = list(ri_vals)
        if not row or len(row) < 13: continue
        product = safe_str(row[0])
        supplier = safe_str(row[1])
        if not product or not supplier: continue
        # Skip header/instruction rows
        if product.lower().startswith("product") or supplier.lower().startswith("supplier"):
            continue
        entry = {
            "supplier":       supplier,
            "category":       safe_str(row[2]),
            "unit":           safe_str(row[3]) or "units",
            "batch_size":     safe_float(row[4], 1.0),
            "lead_time":      safe_float(row[5], 3.0),
            "target_moh":     safe_float(row[6], 9.0),
            "critical_moh":   safe_float(row[7], 6.0),
            "overstock_moh":  safe_float(row[8], 18.0),
            "active_from":    to_iso_01(row[9]),
            "active_through": to_iso_01(row[10]),
            "manual_toggle":  int(safe_float(row[11], 0)),
            "primary":        int(safe_float(row[12], 0)),
            "binding_mo":     int(safe_float(row[15], 0)) if len(row) > 15 else 0,
            "notes":          safe_str(row[14]) if len(row) > 14 else "",
        }
        registry.setdefault(product, []).append(entry)
        rows_read += 1
    n_active = sum(1 for prod_rows in registry.values()
                    for r in prod_rows if r["manual_toggle"] == 1)
    print(f"  ✅ Supplier Registry: {rows_read} rows across {len(registry)} products "
          f"({n_active} toggled active)")
    return registry

# ══════════════════════════════════════════════════════════════════════════════
#  READ LOT DATA (FEFO)
# ══════════════════════════════════════════════════════════════════════════════

def read_fefo_data(wb, prod_data):
    """
    Read Lot Inputs and compute runout / risk for each lot.

    Column map (0-indexed):
      0  Product        1  Lot ID         4  Expiry Date
      5  Qty Available  6  Qty Reserved   7  Net Available
      9  Sort Key       10 Notes

    Sort Key (col J, 0-indexed = 9):
      Controls consumption order within each product block.
      Lower Sort Key = consumed first.
      Falls back to Expiry Date if blank (pure FEFO).
      To defer a lot (e.g. deviation hold), set its Sort Key to a later
      date — the lot drops to the back of the queue and the next lot
      is consumed first.
    """
    ws = get_sheet(wb, "Lot Inputs")
    if not ws:
        print("  ⚠  Lot Inputs sheet not found")
        return []

    from collections import defaultdict
    demand_by_product = defaultdict(dict)
    for r in prod_data:
        demand_by_product[r["product"]][r["month"]] = r.get("demand", 0)

    # Read all lots
    # Cols (0-indexed): 0=Product 1=LotID 4=Expiry 5=QtyAvail 6=QtyReserved
    #                   7=NetAvail 9=SortKey
    lots_by_product = defaultdict(list)
    for ri_vals in ws.iter_rows(min_row=5, values_only=True):
        row = list(ri_vals)
        if not row[0]: continue
        product  = safe_str(row[0])
        lot_id   = safe_str(row[1]) if len(row) > 1 and row[1] else ""
        expiry   = to_iso(row[4]) if len(row) > 4 else None

        # Net available: prefer col 7 (formula cache), fallback to col5 - col6
        if len(row) > 7 and row[7] is not None:
            net_avail = int(safe_float(row[7]))
        elif len(row) > 5 and row[5] is not None:
            net_avail = int(safe_float(row[5])) - int(safe_float(row[6] if len(row) > 6 else None))
        else:
            net_avail = 0

        if not expiry or not lot_id or net_avail <= 0: continue

        # Sort Key (col J, 0-indexed = 9): consumption order; fall back to expiry
        sort_key = expiry
        if len(row) > 9 and row[9] is not None:
            sk = to_iso(row[9])
            if sk:
                sort_key = sk

        # Status (col L, 0-indexed = 11): Active / On Hold
        lot_status = "Active"
        if len(row) > 11 and row[11] is not None:
            s = safe_str(row[11]).strip()
            if s:
                lot_status = s

        lots_by_product[product].append({
            "lot_id":   lot_id,
            "expiry":   expiry,
            "avail":    net_avail,
            "sort_key": sort_key,
            "lot_status": lot_status,
        })

    # Sort each product's lots by Sort Key ascending
    # Lots with a later Sort Key (e.g. deferred deviation hold) sort to the back
    for p in lots_by_product:
        lots_by_product[p].sort(key=lambda l: l["sort_key"])

    # Compute runout for each lot in Sort Key order
    result = []
    seq_counter = {}
    for product, lots in lots_by_product.items():
        d_map  = demand_by_product.get(product, {})
        months = sorted(d_map.keys())
        cum_demand = 0.0
        demand_at  = {}
        for m in months:
            cum_demand += d_map[m]
            demand_at[m] = cum_demand

        # Build cumulative chain from Active lots only.
        # On Hold lots are skipped in the demand chain — they sit at their
        # expiry position in the Gantt but demand does not burn through them.
        cum_at_lot = 0.0
        for lot in lots:
            on_hold = lot.get("lot_status", "Active").lower() in ("on hold", "hold")

            if on_hold:
                # Not consumed — runout = expiry (worst case), full qty at risk
                runout_month = lot["expiry"][:7] + "-01"
                is_risk  = True
                risk_qty = lot["avail"]
                cum_for_display = cum_at_lot  # don't advance chain
            else:
                cum_at_lot += lot["avail"]
                runout_month = None
                for m in months:
                    if demand_at[m] >= cum_at_lot:
                        runout_month = m
                        break
                if runout_month is None:
                    runout_month = lot["expiry"][:7] + "-01"
                try:
                    exp_dt   = datetime.strptime(lot["expiry"][:10], "%Y-%m-%d")
                    run_dt   = datetime.strptime(runout_month[:10], "%Y-%m-%d")
                    is_risk  = run_dt > exp_dt
                    risk_qty = lot["avail"] if is_risk else 0
                except:
                    is_risk = False; risk_qty = 0
                cum_for_display = int(cum_at_lot)

            seq_counter[product] = seq_counter.get(product, 0) + 1
            result.append({
                "seq":      seq_counter[product],
                "lot":      lot["lot_id"],
                "expiry":   lot["expiry"],
                "runout":   runout_month,
                "avail":    lot["avail"],
                "cumAvail": cum_for_display,
                "risk":     risk_qty,
                "status":   "On Hold" if on_hold else ("At Risk" if is_risk else "OK"),
                "product":  product,
            })

    at_risk = sum(1 for l in result if l["status"] == "At Risk")
    print(f"  ✅ FEFO lots: {len(result)} lots ({at_risk} at risk) across {len(lots_by_product)} products")
    return result

def build_fefo_demand(prod_data, product="47mg LP"):
    """Build cumulative demand array for FEFO chart baseline."""
    rows = sorted([r for r in prod_data if r["product"] == product], key=lambda r: r["month"])
    result = []
    cum = 0.0
    for r in rows:
        cum += r.get("demand", 0)
        result.append({"m": r["month"], "d": round(r.get("demand", 0), 2), "c": round(cum, 2)})
    return result

# ══════════════════════════════════════════════════════════════════════════════
#  READ CANCELLATION POLICY
# ══════════════════════════════════════════════════════════════════════════════

def read_cancellation_policy(wb):
    ws = get_sheet(wb, "Cancellation Policy")
    if not ws:
        print("  ⚠  Cancellation Policy sheet not found")
        return {"tiers": [], "exposure": []}

    # TABLE A — tiers (rows 5-13, cols A-N)
    tiers = []
    for ri in range(5, 14):
        row = list(ws.iter_rows(min_row=ri, max_row=ri, values_only=True))[0]
        if not row[0]: continue
        tiers.append({
            "supplier":      safe_str(row[0]),
            "product":       safe_str(row[1]),
            "binding_mo":    int(safe_float(row[2])),
            "notice_mo":     int(safe_float(row[3])),
            "q1_pct":        safe_float(row[4]),
            "q2_pct":        safe_float(row[5]),
            "q3_pct":        safe_float(row[6]),
            "q4_pct":        safe_float(row[7]),
            "q5_pct":        safe_float(row[8]),
            "q6_pct":        safe_float(row[9]),
            "notes":         safe_str(row[10]),
            "active":        safe_str(row[13]) if len(row) > 13 else "",
        })

    # TABLE B — exposure by quarter (rows 18-26)
    exposure = []
    for ri in range(18, 27):
        row = list(ws.iter_rows(min_row=ri, max_row=ri, values_only=True))[0]
        if not row[0]: continue
        exposure.append({
            "supplier":  safe_str(row[0]),
            "product":   safe_str(row[1]),
            "q1_prod":   round(safe_float(row[2]),  2),
            "q1_pct":    safe_float(row[3]),
            "q1_exp":    round(safe_float(row[4]),  2),
            "q2_prod":   round(safe_float(row[5]),  2),
            "q2_pct":    safe_float(row[6]),
            "q2_exp":    round(safe_float(row[7]),  2),
            "q3_prod":   round(safe_float(row[8]),  2),
            "q3_pct":    safe_float(row[9]),
            "q3_exp":    round(safe_float(row[10]), 2),
            "q4_prod":   round(safe_float(row[11]), 2),
            "q4_pct":    safe_float(row[12]),
            "q4_exp":    round(safe_float(row[13]), 2),
            "q5_prod":   round(safe_float(row[14]), 2),
            "q5_pct":    safe_float(row[15]),
            "q5_exp":    round(safe_float(row[16]), 2),
            "q6_prod":   round(safe_float(row[17]), 2),
            "q6_pct":    safe_float(row[18]),
            "q6_exp":    round(safe_float(row[19]), 2),
            "total_exp": round(safe_float(row[20]), 2),
        })

    total = sum(r["total_exp"] for r in exposure)
    print(f"  ✅ Cancellation policy: {len(tiers)} suppliers, total 6Q exposure ${total:,.0f}")
    return {"tiers": tiers, "exposure": exposure}

# ══════════════════════════════════════════════════════════════════════════════
#  READ PRODUCTION INPUTS (per-supplier monthly schedule + binding periods)
# ══════════════════════════════════════════════════════════════════════════════

def parse_mon_yy(s):
    """Parse 'Mar-26' style headers into ISO month string '2026-03-01'."""
    if not s: return None
    s = str(s).strip()
    for fmt in ('%b-%y', '%b %Y', '%B-%y', '%B %Y'):
        try:
            d = datetime.strptime(s, fmt)
            return d.strftime('%Y-%m-01')
        except: pass
    return to_iso_01(s)

def read_prod_inputs(wb, controls):
    """
    Read Production Inputs tab — per-supplier monthly batch schedule.
    Active status and batch size come from Supplier Registry (cached numeric values).
    Month headers are 'Mon-YY' strings parsed with parse_mon_yy().
    Binding period from Supplier Registry col P (index 15, cached numeric).
    Returns list of dicts:
      { product, supplier, unit, batch_size, binding_mo, binding_through,
        months: [{month: iso, qty: float}] }
    """
    ws_pi = get_sheet(wb, "Production Inputs")
    ws_sr = get_sheet(wb, "Supplier Registry")
    if not ws_pi:
        print("  ⚠  Production Inputs sheet not found")
        return []

    # ── Build Supplier Registry lookup ───────────────────────────────────────
    # Key: (product, supplier_key)  Value: {batch_size, binding_mo, active, primary}
    sr_lookup = {}  # (prod, sup_key) -> dict
    if ws_sr:
        for ri_vals in ws_sr.iter_rows(min_row=5, values_only=True):
            row = list(ri_vals)
            if not row[0] or not row[1]: continue
            prod     = safe_str(row[0])
            sup      = safe_str(row[1])
            sup_key  = sup[:12].lower().strip()
            bsize    = safe_float(row[4])   # col E = batch size (cached)
            toggle   = safe_float(row[11])  # col L = manual toggle 1/0 (cached)
            primary  = safe_float(row[12])  # col M = primary flag (cached)
            binding_mo = int(safe_float(row[15])) if len(row) > 15 and row[15] else 9  # col P
            sr_lookup[(prod, sup_key)] = {
                'batch_size': bsize,
                'binding_mo': binding_mo,
                'active': toggle,
                'primary': primary,
            }

    # ── Parse month headers from row 3 (col K onward = index 10+) ────────────
    header_row = list(ws_pi.iter_rows(min_row=3, max_row=3, values_only=True))[0]
    month_headers = []
    for i, v in enumerate(header_row[10:], start=10):
        if not v: continue
        s = str(v).strip()
        if s.startswith('FY'): break  # stop at annual summary columns
        iso = parse_mon_yy(s)
        if iso:
            month_headers.append((i, iso))

    # Limit to 24 months from model start
    model_start = controls.get('model_start')
    if model_start:
        ms_iso = model_start.strftime('%Y-%m-01') if hasattr(model_start, 'strftime') else str(model_start)[:7] + '-01'
        month_headers = [(i, iso) for i, iso in month_headers if iso >= ms_iso][:24]

    if not month_headers:
        print("  ⚠  No valid month headers found in Production Inputs")
        return []

    # ── Read supplier rows ────────────────────────────────────────────────────
    # All products are included — active flag from Supplier Registry controls visibility.
    result = []
    for ri_vals in ws_pi.iter_rows(min_row=4, values_only=True):
        row = list(ri_vals)
        prod = safe_str(row[0]) if row[0] else ''
        sup  = safe_str(row[1]) if len(row) > 1 and row[1] else ''

        if not prod or prod.startswith('▸'): continue
        if not sup or sup == 'TOTAL ACTIVE':  continue

        # Look up SR metadata
        sup_key = sup[:12].lower().strip()
        sr      = sr_lookup.get((prod, sup_key), {})

        # Active: use SR toggle (1 = active). Default to include if not found.
        if sr.get('active', 1) == 0:
            continue

        batch_size = sr.get('batch_size', 0) or safe_float(row[6] if len(row) > 6 else None)
        binding_mo = sr.get('binding_mo', 9)

        # Binding through date
        binding_through = None
        if model_start:
            ms = model_start if isinstance(model_start, date) else model_start.date() if hasattr(model_start, 'date') else None
            if ms:
                import calendar
                total_m = ms.month - 1 + binding_mo
                bt_yr   = ms.year + total_m // 12
                bt_mo   = total_m % 12 + 1
                binding_through = f'{bt_yr:04d}-{bt_mo:02d}-01'

        # Monthly production values
        months = []
        for col_idx, iso in month_headers:
            v = row[col_idx] if col_idx < len(row) else None
            qty = safe_float(v) if isinstance(v, (int, float)) else 0
            months.append({'month': iso, 'qty': round(qty, 2)})

        # Skip rows with zero production across all months
        if not any(m['qty'] > 0 for m in months):
            continue

        unit = safe_str(row[3]) if len(row) > 3 and row[3] else 'units'

        result.append({
            'product':         prod,
            'supplier':        sup,
            'unit':            unit,
            'batch_size':      batch_size,
            'binding_mo':      binding_mo,
            'binding_through': binding_through,
            'months':          months,
        })

    print(f"  ✅ Production Inputs: {len(result)} active supplier rows across {len({r['product'] for r in result})} products")
    for r in result:
        total = sum(m['qty'] for m in r['months'])
        print(f"     {r['product']:15} {r['supplier'][:25]:25} total={total:,.0f} {r['unit']}  binding={r['binding_mo']}mo")
    return result




def read_demand_streams(wb):
    """
    Read Demand Inputs tab and aggregate monthly demand by Indication × Region
    per drug product.  Returns a dict keyed by stream label, where each value
    is a list of monthly demand totals in chronological order.

    Stream labels follow the pattern:
        "{Region} {Indication} ({Drug Product})"
    e.g. "US LR-MDS (188mg LP)", "EU MF (47mg LP)"

    Active? column (col E / index 4) holds either a numeric 1/0 cached from
    the SKU Master formula, or the formula string itself.  We treat any non-zero
    number as active and skip inactive rows.  AGGREGATE rows (SKU Code = AGGREGATE)
    and section-header rows (Drug Product starts with ▸) are skipped.
    """
    ws = get_sheet(wb, "Demand Inputs")
    if not ws:
        print("  ⚠  Demand Inputs sheet not found — stream data skipped")
        return {}

    # Row 3 is the header; months start at col F (index 5)
    header_row = list(ws.iter_rows(min_row=3, max_row=3, values_only=True))[0]
    month_headers = []
    for i, v in enumerate(header_row[5:], start=5):
        iso = to_iso_01(v)
        if iso:
            month_headers.append((i, iso))

    if not month_headers:
        print("  ⚠  No month headers found in Demand Inputs — stream data skipped")
        return {}

    # Accumulate demand per stream per month index
    # streams[label][month_index] += demand
    from collections import defaultdict
    stream_totals = defaultdict(lambda: defaultdict(float))

    for ri_vals in ws.iter_rows(min_row=4, values_only=True):
        row = list(ri_vals)
        prod = safe_str(row[0]) if row[0] else ""
        sku  = safe_str(row[1]) if len(row) > 1 and row[1] else ""

        # Skip section headers, blank rows, and aggregate rows
        if not prod or prod.startswith("▸"):
            continue
        if sku in ("", "AGGREGATE", "SKU Code"):
            continue
        if prod in ("Drug Product",):
            continue

        # Only LP products contribute to the demand chart
        if prod not in ("188mg LP", "47mg LP"):
            continue

        indication = safe_str(row[2]) if len(row) > 2 and row[2] else ""
        region     = safe_str(row[3]) if len(row) > 3 and row[3] else ""

        # Active check — col E (index 4) is either a number (1/0) or formula string
        active_val = row[4] if len(row) > 4 else None
        if isinstance(active_val, (int, float)):
            if active_val == 0:
                continue   # inactive SKU — contributes 0, skip
        elif isinstance(active_val, str):
            if active_val.strip() == "0":
                continue
        elif active_val is None:
            continue

        # Include all indications and regions present in the data

        label = f"{region} {indication} ({prod})"

        # Sum demand across all month columns
        for col_idx, iso in month_headers:
            v = row[col_idx] if col_idx < len(row) else None
            stream_totals[label][iso] += safe_float(v)

    # Convert to sorted lists of {"m": iso, "d": value} matching PROD_DATA shape
    # All streams should cover the same month range
    all_months = sorted({iso for totals in stream_totals.values() for iso in totals})
    result = {}
    for label, totals in sorted(stream_totals.items()):
        result[label] = [
            {"m": iso, "d": round(totals.get(iso, 0.0), 2)}
            for iso in all_months
        ]

    stream_count = len(result)
    month_count  = len(all_months)
    print(f"  ✅ Demand streams: {stream_count} streams × {month_count} months")
    for label, rows in sorted(result.items()):
        total = sum(r["d"] for r in rows)
        print(f"     {label}: total={total:,.0f}")
    return result



def inject_html(html_path, out_path, prod_data, raw_meta, cost_data,
                controls, deps_data, fefo_data, cancel_data, demand_streams, prod_inputs,
                supplier_registry, deps_full, refresh_time):

    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()

    def replace_const(html, name, value_json, is_array=False):
        bracket = r'\[.*?\]' if is_array else r'\{.*?\}'
        pattern = rf'(const {re.escape(name)}\s*=\s*){bracket}(;)'
        # Use lambda so value_json is never treated as a regex replacement string
        # (avoids errors on \u, \n, \g etc. that appear in JSON)
        found = []
        def _repl(m):
            found.append(1)
            return m.group(1) + value_json + m.group(2)
        new_html = re.sub(pattern, _repl, html, flags=re.DOTALL)
        if not found:
            print(f"  ⚠  Could not find 'const {name}' in HTML — skipping")
        return new_html

    # Controls inventory = first-month BI per product
    ctrl_inv = {}
    seen = set()
    for r in prod_data:
        if r["product"] not in seen:
            ctrl_inv[r["product"]] = r.get("bi", 0)
            seen.add(r["product"])

    fefo_demand = build_fefo_demand(prod_data, "47mg LP")

    html = replace_const(html, "RAW",                json.dumps(raw_meta,  separators=(",",":")))
    html = replace_const(html, "PROD_DATA",           json.dumps(prod_data, separators=(",",":")), is_array=True)
    html = replace_const(html, "COST_DATA",           json.dumps(cost_data, separators=(",",":")))
    html = replace_const(html, "CONTROLS_INVENTORY",  json.dumps(ctrl_inv,  separators=(",",":")))
    html = replace_const(html, "DEPENDENCIES",        json.dumps(deps_data, separators=(",",":")))
    html = replace_const(html, "FEFO_LOTS",           json.dumps(fefo_data, separators=(",",":")), is_array=True)
    html = replace_const(html, "FEFO_DEMAND",         json.dumps(fefo_demand, separators=(",",":")), is_array=True)
    html = replace_const(html, "CANCEL_DATA",         json.dumps(cancel_data, separators=(",",":")))

    # ── PROD_INPUTS — inject if const exists; add before DEMAND_BY_STREAM if not ─
    prod_inputs_json = json.dumps(prod_inputs, separators=(",",":"))
    if "const PROD_INPUTS" in html:
        html = replace_const(html, "PROD_INPUTS", prod_inputs_json, is_array=True)
    else:
        html = html.replace(
            "const DEMAND_BY_STREAM =",
            f"const PROD_INPUTS = {prod_inputs_json};\nconst DEMAND_BY_STREAM ="
        )

    # ── DEMAND_BY_STREAM — inject if const exists; add it if not ─────────────
    streams_json = json.dumps(demand_streams, separators=(",",":"))
    if "const DEMAND_BY_STREAM" in html:
        html = replace_const(html, "DEMAND_BY_STREAM", streams_json)
    else:
        # Insert after CANCEL_DATA line so the constant is always present
        html = html.replace(
            "const CANCEL_DATA =",
            f"const DEMAND_BY_STREAM = {streams_json};\nconst CANCEL_DATA ="
        )

    # ── SUPPLIER_REGISTRY ────────────────────────────────────────────────────
    sr_json = json.dumps(supplier_registry, separators=(",",":"))
    if "const SUPPLIER_REGISTRY" in html:
        html = replace_const(html, "SUPPLIER_REGISTRY", sr_json)
    else:
        html = html.replace(
            "const CANCEL_DATA =",
            f"const SUPPLIER_REGISTRY = {sr_json};\nconst CANCEL_DATA ="
        )

    # ── DEPENDENCIES_FULL ────────────────────────────────────────────────────
    deps_full_json = json.dumps(deps_full, separators=(",",":"))
    if "const DEPENDENCIES_FULL" in html:
        html = replace_const(html, "DEPENDENCIES_FULL", deps_full_json)
    else:
        html = html.replace(
            "const CANCEL_DATA =",
            f"const DEPENDENCIES_FULL = {deps_full_json};\nconst CANCEL_DATA ="
        )
    # Footer date
    refresh_str = refresh_time.strftime("%B %d, %Y at %I:%M %p")
    html = re.sub(r'(<span id="footer-date">).*?(</span>)',
                  f'\\1Last refreshed: {refresh_str}\\2', html)

    # Period badge
    if controls.get("dash_start"):
        ds = controls["dash_start"]
        if isinstance(ds, date):
            period = ds.strftime("%b %Y")
        else:
            period = fmt_label(ds)
    else:
        period = refresh_time.strftime("%b %Y")
    html = re.sub(r'(<span class="header-date" id="period-badge">)[^<]*(</span>)',
                  f'\\1· Period: {period}\\2', html)

    # Lookahead default
    lookahead = controls.get("lookahead", 6)
    html = re.sub(r'let lookahead\s*=\s*\d+;', f'let lookahead = {lookahead};', html)
    html = re.sub(r'value="\d+" step="3"', f'value="{lookahead}" step="3"', html)

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

# ══════════════════════════════════════════════════════════════════════════════
#  SCENARIO PLANNER INJECT
#  Writes the same JS data constants into the Scenario Planner HTML so it is
#  fully self-contained and usable by anyone who opens the file.
# ══════════════════════════════════════════════════════════════════════════════

def inject_scenario_planner(planner_path, prod_data, raw_meta, cost_data,
                             deps_data, fefo_data, demand_streams, prod_inputs,
                             supplier_registry, deps_full, refresh_time):
    """Inject live workbook data constants into the Scenario Planner HTML."""

    if planner_path is None:
        return False
    planner_path = Path(planner_path)
    if not planner_path.exists():
        print(f"\n⚠  Scenario Planner not found at:\n   {planner_path}")
        print("   Skipping Scenario Planner update.")
        return False

    with open(planner_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Build the constants block — identical shape to what inject_html() writes
    prod_json    = json.dumps(prod_data,         separators=(",",":"))
    raw_json     = json.dumps(raw_meta,          separators=(",",":"))
    cost_json    = json.dumps(cost_data,         separators=(",",":"))
    deps_json    = json.dumps(deps_data,         separators=(",",":"))
    deps_full_js = json.dumps(deps_full,         separators=(",",":"))
    fefo_json    = json.dumps(fefo_data,         separators=(",",":"))
    streams_json = json.dumps(demand_streams,    separators=(",",":"))
    inputs_json  = json.dumps(prod_inputs,       separators=(",",":"))
    sr_json      = json.dumps(supplier_registry, separators=(",",":"))
    ts           = refresh_time.strftime("%B %d, %Y at %I:%M %p")

    # The Scenario Planner uses window.X lookups so we just need the consts
    # declared before the guard lines run. We inject a dedicated data block
    # right after the opening <script> tag.
    data_block = f"""
// ══════════════════════════════════════════════════════════════════════════════
//  INJECTED DATA — written by refresh_dashboard.py on {ts}
//  Do not edit manually — re-run refresh_dashboard.py to update.
// ══════════════════════════════════════════════════════════════════════════════
var PROD_DATA          = {prod_json};
var RAW                = {raw_json};
var COST_DATA          = {cost_json};
var DEPENDENCIES       = {deps_json};
var DEPENDENCIES_FULL  = {deps_full_js};
var FEFO_LOTS          = {fefo_json};
var FEFO_DEMAND        = [];
var DEMAND_BY_STREAM   = {streams_json};
var PROD_INPUTS        = {inputs_json};
var SUPPLIER_REGISTRY  = {sr_json};
// ─────────────────────────────────────────────────────────────────────────────
"""

    # Remove any previously injected data block if present
    html = re.sub(
        r"// ══+\n//  INJECTED DATA.*?// ─────+\n",
        "",
        html,
        flags=re.DOTALL
    )

    # Insert fresh block right after the opening <script> tag in the JS section
    # (the one that contains the guard lines — identified by the comment above them)
    marker = "// ══════════════════════════════════════════════════════════════════════════════\n//  DATA — injected"
    if marker in html:
        html = html.replace(marker, data_block + marker)
    else:
        # Fallback: insert after first <script> tag in body
        html = html.replace("<script>", "<script>" + data_block, 1)

    # Update the refresh timestamp badge if present
    html = re.sub(
        r'(<span[^>]*id="planner-refresh-date"[^>]*>).*?(</span>)',
        f'\\1Last refreshed: {ts}\\2',
        html
    )

    with open(planner_path, "w", encoding="utf-8") as f:
        f.write(html)

    return True


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print("=" * 62)
    print("  GERON S&OP DASHBOARD REFRESH  (Enhanced Workbook)")
    print(f"  {datetime.now().strftime('%A, %B %d, %Y  %I:%M %p')}")
    print("=" * 62)

    wb_path = Path(WORKBOOK_PATH)
    if not wb_path.exists():
        print(f"\n❌ Workbook not found:\n   {WORKBOOK_PATH}")
        print("\nCheck that:")
        print("  1. OneDrive is synced (look for the cloud icon in the taskbar)")
        print("  2. The file name is exactly: Demand_Inventory_Supply Planning_Expiry_Workbook_Enhanced.xlsx")
        input("\nPress Enter to exit...")
        sys.exit(1)

    mod = datetime.fromtimestamp(wb_path.stat().st_mtime)
    print(f"\n📂 Workbook: {wb_path.name}")
    print(f"   Last saved: {mod.strftime('%b %d, %Y at %I:%M %p')}")

    if not HTML_IN.exists():
        print(f"\n❌ Dashboard HTML not found:\n   {HTML_IN}")
        print("   Make sure Geron_SOP_Executive_Dashboard.html is in the same folder.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # Check workbook recency — warn if not saved in last 10 min
    age_min = (datetime.now() - mod).total_seconds() / 60
    if age_min > 30:
        print(f"\n⚠  Workbook was last saved {int(age_min)} min ago.")
        print("   For live formula values: open Excel, press Ctrl+Alt+F9, then Ctrl+S.")
        ans = input("   Continue anyway? (y/N): ").strip().lower()
        if ans not in ("y", "yes"):
            print("   Cancelled.")
            sys.exit(0)

    print(f"\n📖 Opening workbook...")
    wb, tmp = open_wb(wb_path)
    if wb is None:
        print(f"\n❌ Could not open workbook: {tmp}")
        input("\nPress Enter to exit...")
        sys.exit(1)

    print(f"   Sheets: {', '.join(wb.sheetnames[:8])}...")

    print(f"\n⚙️  Reading Controls...")
    global PRODUCT_META
    PRODUCT_META = build_product_meta(wb)
    print(f"   Products discovered: {', '.join(PRODUCT_META.keys())}")
    controls = read_controls(wb)
    print(f"  Model start: {controls['model_start']}  |  Spotlight: {controls['spotlight']}")

    print(f"\n📊 Reading Planning Engine...")
    prod_data = read_planning_engine(wb, controls)

    print(f"\n💰 Reading Supplier Costs (Supplier Registry X/Y/Z) + Cost Inputs...")
    cost_data = read_supplier_costs(wb)

    print(f"\n🔗 Reading Dependencies...")
    deps_data = read_dependencies(wb)
    deps_full = read_dependencies_full(wb)

    print(f"\n🏭 Reading Supplier Registry...")
    supplier_registry = read_supplier_registry(wb)

    print(f"\n📦 Reading Lot Inputs (FEFO)...")
    fefo_data = read_fefo_data(wb, prod_data)

    print(f"\n❌ Reading Cancellation Policy...")
    cancel_data = read_cancellation_policy(wb)

    print(f"\n📈 Reading Demand Streams (Indication × Region)...")
    demand_streams = read_demand_streams(wb)

    print(f"\n📅 Reading Production Inputs (per-supplier schedule)...")
    prod_inputs = read_prod_inputs(wb, controls)

    wb.close()
    if tmp and os.path.exists(tmp):
        try: os.unlink(tmp)
        except: pass

    if not prod_data:
        print("\n⚠  No planning data read — check that the workbook has been saved after recalculation.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # ── SANITY CHECK: PE Demand column vs DEMAND_BY_STREAM ────────────────────
    # If Planning Engine col F has zero demand for any LP product BUT the
    # streams have demand for that product, warn loudly. This catches the
    # exact regression we fixed (PE demand column wasn't recalculated before
    # workbook was saved).
    lp_products = [p for p in {r["product"] for r in prod_data}
                   if p.endswith(" LP") or " LP - " in p]
    for lp_prod in sorted(lp_products):
        pe_total = sum(r.get("demand", 0) for r in prod_data if r["product"] == lp_prod)
        stream_total = 0
        for key, entries in demand_streams.items():
            if f"({lp_prod})" in key:
                stream_total += sum(e.get("d", 0) for e in entries)
        if pe_total == 0 and stream_total > 0:
            print(f"\n  ⚠  WARNING: PE demand for '{lp_prod}' is ZERO but streams sum to "
                  f"{stream_total:,.0f}. The Planning Engine col F may not have been "
                  f"recalculated before saving — open the workbook, press F9 to recalc, "
                  f"save, then re-run refresh.")
        elif pe_total > 0 and stream_total > 0:
            ratio = pe_total / stream_total
            if ratio < 0.5 or ratio > 2.0:
                print(f"\n  ⚠  WARNING: PE demand for '{lp_prod}' ({pe_total:,.0f}) and "
                      f"stream sum ({stream_total:,.0f}) differ by more than 2x. "
                      f"Verify the demand aggregation in Excel is correct.")

    raw_meta = build_raw_meta(prod_data, controls)
    print(f"\n📦 Products: {', '.join(raw_meta.keys())}")

    refresh_time = datetime.now()
    print(f"\n✏️  Writing dashboard: {HTML_OUT.name}...")
    inject_html(HTML_IN, HTML_OUT, prod_data, raw_meta, cost_data,
                controls, deps_data, fefo_data, cancel_data, demand_streams, prod_inputs,
                supplier_registry, deps_full, refresh_time)

    # ── Write Scenario Planner ─────────────────────────────────────────────────
    planner_ok = False
    if SCENARIO_PLANNER_PATH is not None:
        print(f"\n✏️  Writing Scenario Planner: {Path(SCENARIO_PLANNER_PATH).name}...")
        planner_ok = inject_scenario_planner(
            SCENARIO_PLANNER_PATH, prod_data, raw_meta, cost_data,
            deps_data, fefo_data, demand_streams, prod_inputs,
            supplier_registry, deps_full, refresh_time
        )
        if planner_ok:
            print(f"   ✅ Scenario Planner updated.")

    print("\n" + "=" * 62)
    print("  ✅ REFRESH COMPLETE")
    print("=" * 62)
    print(f"  Products:   {len(raw_meta)}")
    print(f"  Data rows:  {len(prod_data)}")
    print(f"  FEFO lots:  {len(fefo_data)}")
    print(f"  Streams:    {len(demand_streams)} demand streams (indication × region)")
    print(f"  Prod rows:  {len(prod_inputs)} active supplier-product rows")
    print(f"  Refreshed:  {refresh_time.strftime('%B %d, %Y  %I:%M %p')}")
    print(f"  Dashboard:  {HTML_OUT}")
    if planner_ok:
        print(f"  Planner:    {SCENARIO_PLANNER_PATH}")
    print("=" * 62)

    print("\n  Open in browser?")
    print("  [1] Executive Dashboard")
    if planner_ok:
        print("  [2] Scenario Planner")
    print("  [Enter] Skip")
    ans = input("\n  Choice: ").strip().lower()
    import webbrowser
    if ans in ("", "1", "y", "yes"):
        webbrowser.open(HTML_OUT.as_uri())
        print("  🌐 Opening Dashboard...")
    elif ans == "2" and planner_ok:
        webbrowser.open(Path(SCENARIO_PLANNER_PATH).as_uri())
        print("  🌐 Opening Scenario Planner...")
    print("\n  Done! Close this window.\n")

if __name__ == "__main__":
    main()
