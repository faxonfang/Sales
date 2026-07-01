#!/usr/bin/env python3
"""從總表產生 momo 摩天商城的改價批次檔。"""
import argparse
import math
from datetime import date, timedelta, datetime

import csv
from collections import defaultdict

import openpyxl
import xlrd
from openpyxl.utils import column_index_from_string
from xlutils.copy import copy as xls_copy

TAX_RATE = 1.05
MARGIN_THRESHOLD = 0.40

MOMO_HEADERS = [
    "商品編號", "商品名稱", "生效日期", "進價", "進價(未稅)",
    "售價", "市價", "毛利率(%)", "備註", "負責MD聯絡資訊", "負責企劃聯絡資訊",
]

MO_COLS = {name: column_index_from_string(letter) for name, letter in {
    "商品編號": "C", "商品名稱": "D", "銷售狀況": "H",
    "商品原廠編號": "J", "售價含稅": "AB", "市價": "AC",
}.items()}


def truncate(value):
    """無條件捨去到整數（不四捨五入）。"""
    return math.floor(value)


def month_bounds(year, month):
    first = date(year, month, 1)
    next_first = date(year + month // 12, month % 12 + 1, 1)
    return first, next_first - timedelta(days=1)


def load_master_sheet(ws):
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    col = {name: i + 1 for i, name in enumerate(headers) if name}
    required = ["品號", "ERP成本", "XP價", "活動價", "促銷起日", "促銷訖日"]
    missing = [name for name in required if name not in col]
    if missing:
        raise ValueError(f"總表缺少必要欄位: {missing}")

    master = {}
    for r in range(2, ws.max_row + 1):
        sku = ws.cell(row=r, column=col["品號"]).value
        if not sku:
            continue
        master[sku] = {
            "cost": ws.cell(row=r, column=col["ERP成本"]).value,
            "base_price": ws.cell(row=r, column=col["XP價"]).value,
            "promo_price": ws.cell(row=r, column=col["活動價"]).value,
            "promo_start": ws.cell(row=r, column=col["促銷起日"]).value,
            "promo_end": ws.cell(row=r, column=col["促銷訖日"]).value,
        }
    return master


def as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def resolve_selling_price(item, target_first, target_last):
    start = as_date(item["promo_start"])
    end = as_date(item["promo_end"])
    if start and end and item["promo_price"] is not None:
        if start <= target_last and end >= target_first:
            return item["promo_price"]
    return item["base_price"]


def build_rows(master, mo_sheet, target_year, target_month):
    """momo 批次檔以「商品編號」（母碼）為單位，一個母碼底下可能有多個規格
    （不同單品編號/原廠貨號），但只能提交一筆售價，所以先依商品編號分組，
    每組取所有規格算出來售價中最低的一筆代表整個商品編號。
    """
    target_first, target_last = month_bounds(target_year, target_month)
    effective_date = f"{target_first.year}/{target_first.month}/{target_first.day}"

    groups = defaultdict(list)
    for r in range(2, mo_sheet.max_row + 1):
        if mo_sheet.cell(row=r, column=MO_COLS["銷售狀況"]).value != "進行":
            continue
        parent_id = mo_sheet.cell(row=r, column=MO_COLS["商品編號"]).value
        groups[parent_id].append(r)

    stats = {
        "active_parents": len(groups), "unmatched_parents": 0,
        "price_conflict_parents": 0, "no_cost_or_price": 0,
        "unchanged": 0, "output": 0,
    }
    rows = []
    missing_variants = []

    for parent_id, variant_rows in groups.items():
        candidates = []
        for r in variant_rows:
            sku = mo_sheet.cell(row=r, column=MO_COLS["商品原廠編號"]).value
            item = master.get(sku)
            if not item:
                missing_variants.append([
                    sku, str(parent_id),
                    str(mo_sheet.cell(row=r, column=MO_COLS["商品名稱"]).value),
                ])
                continue
            sell_price = resolve_selling_price(item, target_first, target_last)
            if sell_price is None or item["cost"] is None:
                continue
            candidates.append({"row": r, "item": item, "price": sell_price})

        if not candidates:
            stats["unmatched_parents"] += 1
            continue
        if len({c["price"] for c in candidates}) > 1:
            stats["price_conflict_parents"] += 1

        chosen = min(candidates, key=lambda c: c["price"])
        sell_price = chosen["price"]
        item = chosen["item"]
        row = chosen["row"]

        current_price = mo_sheet.cell(row=row, column=MO_COLS["售價含稅"]).value
        try:
            current_price = float(current_price)
        except (TypeError, ValueError):
            current_price = None

        new_price_int = truncate(sell_price)
        if current_price is not None and new_price_int == truncate(current_price):
            stats["unchanged"] += 1
            continue

        sell_price_untaxed = sell_price / TAX_RATE
        margin = (sell_price_untaxed - item["cost"]) / sell_price_untaxed
        cost_ratio = 0.95 if margin > MARGIN_THRESHOLD else 0.97
        purchase_price = sell_price * cost_ratio
        purchase_price_untaxed = purchase_price / TAX_RATE

        rows.append([
            str(mo_sheet.cell(row=row, column=MO_COLS["商品編號"]).value),
            str(mo_sheet.cell(row=row, column=MO_COLS["商品名稱"]).value),
            effective_date,
            str(truncate(purchase_price)),
            str(truncate(purchase_price_untaxed)),
            str(new_price_int),
            str(mo_sheet.cell(row=row, column=MO_COLS["市價"]).value),
            "", "", "", "",
        ])
        stats["output"] += 1
    return rows, missing_variants, stats


def write_momo_xls(template_path, rows, out_path):
    """把資料寫入 momo 原始範本（保留範本格式與說明分頁），輸出舊版 .xls。"""
    template_book = xlrd.open_workbook(template_path, formatting_info=True)
    out_book = xls_copy(template_book)
    ws = out_book.get_sheet(0)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row):
            ws.write(r, c, value)
    out_book.save(out_path)


def write_missing_variants_report(missing_variants, out_path):
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["總表品號(原廠貨號)", "momo商品編號", "商品名稱"])
        writer.writerows(missing_variants)


def main():
    parser = argparse.ArgumentParser(description="產生 momo 改價批次檔")
    parser.add_argument("--source", required=True, help="總表檔案路徑（需含「總表」與「MO(K)」工作表）")
    parser.add_argument("--template", required=True, help="momo 官方改價批次範本檔路徑（.xls）")
    parser.add_argument("--month", required=True, help="要處理的月份，格式 YYYY-MM，例如 2026-08")
    parser.add_argument("--output", help="輸出批次檔路徑，預設 momo_價格異動_YYYYMM.xls")
    parser.add_argument("--missing-output", help="規格對照不到總表品號的清單，預設 momo_對照不到_YYYYMM.csv")
    args = parser.parse_args()

    year, month = (int(x) for x in args.month.split("-"))

    wb = openpyxl.load_workbook(args.source, data_only=True)
    master = load_master_sheet(wb["總表"])
    rows, missing_variants, stats = build_rows(master, wb["MO(K)"], year, month)

    out_path = args.output or f"momo_價格異動_{year}{month:02d}.xls"
    write_momo_xls(args.template, rows, out_path)

    missing_path = args.missing_output or f"momo_對照不到_{year}{month:02d}.csv"
    write_missing_variants_report(missing_variants, missing_path)

    print(f"momo 目前上架中商品編號（母碼）: {stats['active_parents']}")
    print(f"  完全對照不到總表品號（整個商品編號跳過）: {stats['unmatched_parents']}")
    print(f"  同一商品編號內規格售價不一致（已自動取最低價）: {stats['price_conflict_parents']}")
    print(f"  缺成本或售價無法計算: {stats['no_cost_or_price']}")
    print(f"  售價未變動（略過）: {stats['unchanged']}")
    print(f"  需要改價（寫入批次檔）: {stats['output']}")
    print(f"部分規格對照不到總表品號（其餘規格仍正常處理）: {len(missing_variants)} -> {missing_path}")
    print(f"輸出批次檔: {out_path}")


if __name__ == "__main__":
    main()
