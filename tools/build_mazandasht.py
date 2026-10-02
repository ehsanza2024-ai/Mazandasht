#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mazandasht Excel System Generator
=================================
Generates the two core workbooks of the Mazandasht accounting system:

  1) Mazandasht_INFO.xlsx         -> the "mother" (definitions) workbook
  2) Mazandasht_Transactions.xlsx -> the bank-transactions workbook (Entry + Bank1..Bank10 + Summary)

Everything (sheet/file names & formulas) is Excel-2013 compatible:
only classic functions are used (IF, IFERROR, VLOOKUP, INDEX, MATCH, LOOKUP,
COUNTIF(S), SUMIF(S), LEFT/RIGHT/MID/FIND/LEN/TRIM/VALUE, ...). No dynamic
arrays, no IFS/TEXTJOIN/XLOOKUP.

The Transactions workbook pulls all lists from the INFO workbook through a
"Sync" sheet using real external-workbook links (ECMA-376 externalLink part,
formulas stored as [1]Sheet!Ref).

Usage:
    python3 tools/build_mazandasht.py --all                     # full deliverables at repo root
    python3 tools/build_mazandasht.py --qa                      # small internal-linked build (for tests)
    python3 tools/build_mazandasht.py --qa --external           # small externally-linked pair (for tests)
"""

import argparse
import os
import re
import sys
import zipfile
from dataclasses import dataclass, field

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.comments import Comment

# ============================================================================
# 1. CONFIG
# ============================================================================

FONT_NAME = "Tahoma"          # universally available, renders Persian well
INFO_FILE = "Mazandasht_INFO.xlsx"
TXN_FILE = "Mazandasht_Transactions.xlsx"

BRAND_NAME = "مازندشت"
BRAND_DESC = "سورتینگ و بسته‌بندی مرکبات"
BRAND_ADDR = "باغدشت، بخش مرکزی، شهرستان قائم‌شهر، استان مازندران، ایران"
BRAND_EN = "MAZANDASHT · CITRUS SORTING & PACKING"

# palette (citrus: deep green + orange)
C = dict(
    green_d="1B5E20", green="2E7D32", green_l="E8F5E9", green_xl="F1F8F2",
    orange_d="E65100", orange="EF6C00", orange_l="FFF3E0",
    input="FFF9C4", auto="EEF3F1", zebra="F3F8F3",
    border="B0BEC5", border_d="78909C",
    txt="263238", gray="607D8B", hdr_en="546E7A",
    red="C62828", red_l="FFCDD2", ok="2E7D32", ok_l="C8E6C9",
    amber="E65100", amber_l="FFE0B2", white="FFFFFF",
    tab_green="2E7D32", tab_orange="EF6C00", tab_gray="90A4AE",
    tab_blue="1565C0", tab_teal="00695C", tab_olive="558B2F",
)

# widths of shared text helpers
OK_MARK = "√"          # green tick
BAD_MARK = "×"         # red cross
WARN_MARK = "؟"        # amber question (cannot compute)
DUP_MARK = "تکرار"
NONE_MARK = "—"


@dataclass
class Cap:
    """Row capacities of the system."""
    entry_rows: int = 2000      # Entry data rows  (rows 9..2008)
    bank_rows: int = 500        # rows in each Bank sheet (rows 17..516)
    seller_rows: int = 2000     # rows in each seller statement sheet
    banks: int = 10
    customers: int = 200
    sellers: int = 100
    suppliers: int = 100
    accounts: int = 50
    staff: int = 50
    party_types: int = 20
    txn_types: int = 14
    products_buy: int = 100
    products_sell: int = 100
    baskets: int = 30
    warehouses: int = 15


QA_CAP = Cap(entry_rows=40, bank_rows=20, seller_rows=12, banks=10, customers=10, sellers=8,
             suppliers=8, accounts=6, staff=6, party_types=10, txn_types=5,
             products_buy=8, products_sell=8, baskets=5, warehouses=4)

# ---- fixed layout coordinates --------------------------------------------
E_DATA0 = 9                       # first Entry data row
B_DATA0 = 17                      # first Bank-sheet data row
S_DATA0 = 7                       # first data row in INFO registers
SY_DATA0 = 6                      # first data row of Sync mirrors


def e_data1(cap): return E_DATA0 + cap.entry_rows - 1
def b_data1(cap): return B_DATA0 + cap.bank_rows - 1


# ============================================================================
# 2. SEED DATA (realistic samples the user can edit / extend)
# ============================================================================

SEED_BANKS = [
    ("BANK01", "ملت — جاری شرکت", 500000000, "بله", "حساب جاری اصلی شرکت (نمونه — ویرایش کنید)"),
    ("BANK02", "صادرات — جاری شرکت", 120000000, "بله", ""),
    ("BANK03", "حساب ۳ — تعریف نشده", 0, "بله", "عنوان و مانده اولیه را در همین ردیف تعریف کنید"),
    ("BANK04", "حساب ۴ — تعریف نشده", 0, "بله", ""),
    ("BANK05", "حساب ۵ — تعریف نشده", 0, "بله", ""),
    ("BANK06", "حساب ۶ — تعریف نشده", 0, "بله", ""),
    ("BANK07", "حساب ۷ — تعریف نشده", 0, "بله", ""),
    ("BANK08", "حساب ۸ — تعریف نشده", 0, "بله", ""),
    ("BANK09", "حساب ۹ — تعریف نشده", 0, "بله", ""),
    ("BANK10", "حساب ۱۰ — تعریف نشده", 0, "بله", ""),
]

SEED_CUSTOMERS = [
    ("C001", "مشتری عمده", "علی محمدی (نمونه)", "011-33330001", "09110000001", "بابل — میدان مرکبات", "نمونه — ویرایش کنید"),
    ("C002", "مشتری صادراتی", "شرکت پارس میوه (نمونه)", "011-33330002", "09110000002", "ساری — شهرک صنعتی", "نمونه"),
    ("C003", "مشتری خرده‌فروش", "فروشگاه بهار (نمونه)", "011-33330003", "09110000003", "قائم‌شهر — خیابان طالب آملی", "نمونه"),
]

SEED_SUPPLIERS = [
    ("S001", "باغ‌دار", "حسن رستمی (نمونه)", "011-44440001", "09110000004", "باغدشت", "خرید از باغ"),
    ("S002", "حجره‌دار میدان بار", "کریم نوری (نمونه)", "011-44440002", "09110000005", "قائم‌شهر — میدان بار", "خرید از میدان بار"),
    ("S003", "عامل خرید", "محسن کاظمی (نمونه)", "011-44440003", "09110000006", "جویبار", "نمونه"),
]

SEED_ACCOUNTS = [
    ("X001", "حساب اداری", "کارمزد و هزینه بانکی", "کارمزد، Thomson انتقال و similares"),
    ("X002", "حساب اداری", "مالیات و عوارض", ""),
    ("X003", "حساب داخلی", "صندوق / نقدی", "انتقال بین صندوق و بانک"),
]

SEED_STAFF = [
    ("E001", "(نام مدیرعامل)", "مدیرعامل", "09110000010", ""),
    ("E002", "(نام حسابدار)", "حسابدار", "09110000011", ""),
    ("E003", "(نام فروشنده ۱)", "فروشنده", "09110000012", ""),
    ("E004", "(نام انباردار)", "انباردار", "09110000013", ""),
]

SEED_SELLERS = [
    ("SL001", "حاج رضا محمدی (نمونه)", "حجره‌دار میدان بار",
     "021-55550001", "09120000001", "تهران — میدان بار میوه و تره‌بار", "نمونه — ویرایش کنید"),
    ("SL002", "شرکت گلسرای شمال (نمونه)", "مشتری صادراتی",
     "011-55550002", "09120000002", "ساری — شهرک صنعتی", "نمونه"),
    ("SL003", "عباس نیک‌پی (نمونه — واسطه)", "واسطه",
     "011-55550003", "09120000003", "قائم‌شهر", "نمونه"),
    ("SL004", "مشتری نقدی — پاساژ مرکزی (نمونه)", "مشتری نقدی",
     "011-55550004", "09120000004", "قائم‌شهر — پاساژ مرکزی", "نمونه"),
]

SEED_PARTY_TYPES = [
    ("PT01", "مشتری عمده", "مشتری"),
    ("PT02", "مشتری صادراتی", "مشتری"),
    ("PT03", "مشتری خرده‌فروش", "مشتری"),
    ("PT04", "باغ‌دار", "تأمین‌کننده"),
    ("PT05", "حجره‌دار میدان بار", "تأمین‌کننده"),
    ("PT06", "عامل خرید", "تأمین‌کننده"),
    ("PT07", "حساب اداری/داخلی", "سایر"),
    ("PT08", "واسطه", "مشتری"),
    ("PT09", "مشتری نقدی", "مشتری"),
    ("PT10", "سایر", "سایر"),
]

SEED_TXN_TYPES = [
    ("T01", "واریز", "افزایش مانده"),
    ("T02", "برداشت", "کاهش مانده"),
    ("T03", "انتقال بین بانکی", "خنثی در مجموع"),
    ("T04", "اصلاحیه / تعدیل", "طبق ماهیت"),
    ("T05", "سایر", "طبق ماهیت"),
]

SEED_PRODUCTS_BUY = [
    ("PB01", "پرتقال (عمومی)", "پرتقال", "کیلوگرم", "خرید از باغ/میدان بار"),
    ("PB02", "نارنگی", "نارنگی", "کیلوگرم", ""),
    ("PB03", "لیمو ترش", "لیمو", "کیلوگرم", ""),
    ("PB04", "گریپ‌فرویت", "گریپ‌فرویت", "کیلوگرم", ""),
]

SEED_PRODUCTS_SELL = [
    ("PS01", "پرتقال درجه ۱ — بسته‌بندی", "پرتقال", "کیلوگرم", "BK03", "خروج خط سورتینگ"),
    ("PS02", "پرتقال درجه ۲ — عمومی", "پرتقال", "کیلوگرم", "BK02", ""),
    ("PS03", "نارنگی درجه ۱ — بسته‌بندی", "نارنگی", "کیلوگرم", "BK03", ""),
    ("PS04", "نارنگی عمومی", "نارنگی", "کیلوگرم", "BK01", ""),
]

SEED_BASKETS = [
    ("BK01", "سبد ۵ کیلویی", 5, "سبد پلاستیکی"),
    ("BK02", "سبد ۱۰ کیلویی", 10, "سبد پلاستیکی"),
    ("BK03", "جعبه ۱۰ کیلویی", 10, "کارتن"),
]

SEED_WAREHOUSES = [
    ("W01", "انبار اصلی (باغدشت)", "باغدشت", "E004", "انبار محصول"),
    ("W02", "سردخانه ۱", "باغدشت", "E004", ""),
    ("W03", "سایبان سورتینگ", "باغدشت", "E004", "ورودی خط سورتینگ"),
]

SEED_LISTS_UNITS = ["کیلوگرم", "جعبه", "سبد", "تن", "عدد"]
SEED_LISTS_ROLES = ["مدیرعامل", "حسابدار", "فروشنده", "انباردار", "سرپرست خط", "کارگر"]
SEED_LISTS_CATS = ["پرتقال", "نارنگی", "لیمو", "گریپ‌فرویت", "گیلاس", "سایر"]
SEED_LISTS_YESNO = ["بله", "خیر"]

SEED_COMPANY = [
    ("نام کامل شرکت", "شرکت سورتینگ و بسته‌بندی مرکبات مازندشت"),
    ("برند", "مازندشت"),
    ("زمینه فعالیت", "سورتینگ، بسته‌بندی و بازرگانی مرکبات"),
    ("نشانی", BRAND_ADDR),
    ("تلفن ثابت", "011-00000000"),
    ("تلفن همراه", "0911-0000000"),
    ("مدیرعامل", "(نام مدیرعامل)"),
    ("سال مالی", "1405"),
    ("واحد پول", "ریال"),
    ("یادداشت", "این اطلاعات به‌صورت خودکار در سرصفحه فایل‌های دیگر تکرار می‌شود."),
]

# Entry demo rows: (date, time, bank_label, party_code, party_type, txn_type,
#                   deposit, withdraw, manual_balance, note)
SEED_ENTRY = [
    ("1405/07/01", "08:30", "BANK01 – ملت — جاری شرکت", "C001", "مشتری عمده", "واریز",
     250000000, "", 750000000, "واریز بابت خرید مرکبات (ردیف نمونه)"),
    ("1405/07/01", "10:15", "BANK01 – ملت — جاری شرکت", "X001", "حساب اداری/داخلی", "برداشت",
     "", 1200000, 748800000, "کارمزد بانکی (نمونه)"),
    ("1405/07/02", "09:00", "BANK01 – ملت — جاری شرکت", "S002", "حجره‌دار میدان بار", "برداشت",
     "", 80000000, 668800000, "خرید پرتقال از میدان بار (نمونه)"),
    ("1405/07/02", "12:45", "BANK02 – صادرات — جاری شرکت", "C002", "مشتری صادراتی", "واریز",
     90000000, "", 215000000, "واریز مشتری — مغایرت عمدی برای نمایش × (نمونه)"),
    ("1405/07/03", "11:20", "BANK02 – صادرات — جاری شرکت", "X003", "حساب اداری/داخلی", "برداشت",
     "", 15000000, "", "برداشت نقدی به صندوق — بدون کنترل مانده (نمونه)"),
    ("1405/07/01", "08:30", "BANK01 – ملت — جاری شرکت", "C001", "مشتری عمده", "واریز",
     250000000, "", 750000000, "ردیف تکراری — اثر آن را بر ستون وضعیت ببینید (نمونه)"),
    ("1405/07/04", "10:30", "BANK01 – ملت — جاری شرکت", "SL001", "حجره‌دار میدان بار", "واریز",
     300000000, "", "", "واریز فروشنده SL001 (نمونه)"),
    ("1405/07/09", "11:30", "BANK02 – صادرات — جاری شرکت", "SL002", "مشتری صادراتی", "واریز",
     1600000000, "", "", "تسویه فروشنده SL002 (نمونه)"),
]

MONTHS_FA = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
             "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]

# ============================================================================
# 3. STYLE HELPERS
# ============================================================================

def fnt(size=9, bold=False, color=None, italic=False, name=FONT_NAME):
    return Font(name=name, size=size, bold=bold, italic=italic,
                color=color or C["txt"])

def fl(color):
    return PatternFill("solid", fgColor=color)

def al(h="right", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, readingOrder=0)

def sd(style="thin", color=None):
    return Side(style=style, color=color or C["border"])

def bd(l=None, r=None, t=None, b=None, diag=False):
    l = l or sd(); r = r or sd(); t = t or sd(); b = b or sd()
    return Border(left=l, right=r, top=t, bottom=b)

BORDER_ALL = bd()
BORDER_BOX_D = bd(sd("medium", C["border_d"]), sd("medium", C["border_d"]),
                  sd("medium", C["border_d"]), sd("medium", C["border_d"]))


def put(ws, ref, value=None, font=None, fill=None, alignment=None,
        border=None, number_format=None):
    """Set a single cell with styles."""
    cell = ws[ref]
    if value is not None:
        cell.value = value
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if border: cell.border = border
    if number_format: cell.number_format = number_format
    return cell


def style_range(ws, ref, font=None, fill=None, alignment=None, border=None,
                number_format=None):
    """Apply styles to every cell of a (possibly merged) range."""
    from openpyxl.utils.cell import range_boundaries
    c0, r0, c1, r1 = range_boundaries(ref)
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            if font: cell.font = font
            if fill: cell.fill = fill
            if alignment: cell.alignment = alignment
            if border: cell.border = border
            if number_format: cell.number_format = number_format


def merge_put(ws, ref, value=None, font=None, fill=None, alignment=None,
              border=None, number_format=None):
    """Merge a range, write value into its top-left cell and style all cells."""
    ws.merge_cells(ref)
    style_range(ws, ref, font, fill, alignment, border, number_format)
    if value is not None:
        first = ref.split(":")[0]
        ws[first] = value
    return ws[ref.split(":")[0]]


def set_widths(ws, widths):
    """widths: dict like {'A': 9, 'B': 24, ...}"""
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def setup_print(ws, orientation="portrait", title_rows=None, print_area=None,
                header_center=None, footer_left=None):
    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = 9  # A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(left=0.35, right=0.35, top=0.62, bottom=0.55,
                                  header=0.28, footer=0.25)
    ws.print_options.horizontalCentered = True
    if title_rows:
        ws.print_title_rows = title_rows
    if print_area:
        ws.print_area = print_area
    h = ws.oddHeader
    h.left.text = "مازندشت — سامانه حساب‌های بانکی"; h.left.size = 8; h.left.font = FONT_NAME
    h.center.text = header_center or ""; h.center.size = 8; h.center.font = FONT_NAME
    h.right.text = "&A"; h.right.size = 8; h.right.font = FONT_NAME
    f = ws.oddFooter
    f.left.text = footer_left or BRAND_ADDR; f.left.size = 7; f.left.font = FONT_NAME
    f.center.text = "صفحه &P از &N"; f.center.size = 8; f.center.font = FONT_NAME
    f.right.text = "چاپ: &D"; f.right.size = 7; f.right.font = FONT_NAME


def rtl(ws, zoom=90):
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom


def brand_band(ws, ncols, mode="info", comp_col="AJ", print_ncols=None):
    """Rows 1-2 brand band. mode='info' -> static text, 'txn' -> pulls from Sync.
    If print_ncols is given (sales Entry: print zone is narrower than the sheet),
    the band is split so printed pages still show the text."""
    last = get_column_letter(ncols)
    if mode == "txn":
        line1 = (f'=IF(Sync!${comp_col}$6="","مازندشت",Sync!${comp_col}$6)&" — "&'
                 f'IF(Sync!${comp_col}$9="","سورتینگ و بسته‌بندی مرکبات",Sync!${comp_col}$9)')
        line2 = (f'=IF(Sync!${comp_col}$7="","%s",Sync!${comp_col}$7)&"     |     %s"'
                 % (BRAND_ADDR, BRAND_EN))
    else:
        line1 = f"{BRAND_NAME} — {BRAND_DESC}"
        line2 = BRAND_ADDR + ("     |     " + BRAND_EN if ncols >= 8 else "")
    if print_ncols and print_ncols < ncols:
        p_last = get_column_letter(print_ncols)
        r_last = get_column_letter(print_ncols + 1)
        merge_put(ws, f"A1:{p_last}1", line1, fnt(15, True, C["white"]), fl(C["green_d"]),
                  al("center"), None)
        merge_put(ws, f"{r_last}1:{last}1", BRAND_EN + "  |  ناحیه فقط-اکسل",
                  fnt(8, False, "A5D6A7"), fl(C["green_d"]), al("center"), None)
        merge_put(ws, f"A2:{p_last}2", line2, fnt(8, False, "C8E6C9"),
                  fl(C["green"]), al("center"), None)
        merge_put(ws, f"{r_last}2:{last}2", None, fnt(8), fl(C["green"]), al("center"), None)
    else:
        merge_put(ws, f"A1:{last}1", line1, fnt(15, True, C["white"]), fl(C["green_d"]),
                  al("center"), None)
        merge_put(ws, f"A2:{last}2", line2, fnt(8, False, "C8E6C9"),
                  fl(C["green"]), al("center"), None)
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 14


def title_row(ws, ncols, row, text, fill_color=None, txt_color=None, size=12):
    last = get_column_letter(ncols)
    merge_put(ws, f"A{row}:{last}{row}", text, fnt(size, True, txt_color or C["orange_d"]),
              fl(fill_color or C["green_xl"]), al("center"),
              bd(b=sd("medium", C["orange"])))
    ws.row_dimensions[row].height = 22
    return row


def table_headers(ws, row_fa, row_en, headers, col0=1, fa_fill=None):
    """headers: list of (fa, en). Writes Persian row + small English row below."""
    for i, (fa, en) in enumerate(headers):
        col = get_column_letter(col0 + i)
        put(ws, f"{col}{row_fa}", fa, fnt(9.5, True, C["white"]), fl(fa_fill or C["green"]),
            al("center", wrap=True), bd(b=sd("medium", C["orange"])))
        put(ws, f"{col}{row_en}", en, fnt(6.5, False, C["hdr_en"]), fl(C["green_l"]),
            al("center", wrap=True), bd(b=sd("medium", C["border_d"])))
    ws.row_dimensions[row_fa].height = 30
    ws.row_dimensions[row_en].height = 13


def card(ws, row_lbl, row_val, c0, c1, label, value, fmt="#,##0",
         vfont=None, vfill=None):
    """A small KPI card spanning columns c0..c1 (label row + value row)."""
    a, b = get_column_letter(c0), get_column_letter(c1)
    merge_put(ws, f"{a}{row_lbl}:{b}{row_lbl}", label, fnt(7.5, True, C["gray"]),
              fl(C["green_l"]), al("center"), BORDER_ALL)
    merge_put(ws, f"{a}{row_val}:{b}{row_val}", value, vfont or fnt(11, True, C["txt"]),
              fl(vfill or C["white"]), al("center"), BORDER_ALL, fmt)
    return f"{a}{row_val}"


def data_area(ws, r0, r1, c0, c1, col_kinds, fmts=None, zebra=True):
    """
    Style a data area. col_kinds: dict col_index -> 'in' | 'auto' | 'plain'
    fmts: dict col_index -> number format
    """
    fmts = fmts or {}
    for r in range(r0, r1 + 1):
        zeb = fl(C["zebra"]) if (zebra and (r - r0) % 2 == 1) else fl(C["white"])
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            kind = col_kinds.get(c, "plain")
            if kind == "in":
                cell.fill = fl(C["input"])
            elif kind == "auto":
                cell.fill = fl(C["auto"])
            else:
                cell.fill = zeb
            cell.border = BORDER_ALL
            cell.font = fnt(9)
            cell.alignment = al("center")
            if c in fmts:
                cell.number_format = fmts[c]


# ============================================================================
# 4. INFO WORKBOOK
# ============================================================================

INFO_SHEET_ORDER = ["Home", "Company", "Banks", "Customers", "Sellers",
                    "Suppliers", "Accounts", "Staff", "PartyTypes", "TxnTypes",
                    "Products_Buy", "Products_Sell", "Baskets", "Warehouses", "Lists"]

# register specs: sheet -> (headers_fa_en, columns used for seed tuples, widths)
REG_SPECS = {
    "Banks": dict(
        headers=[("کد حساب", "Code"), ("عنوان نمایشی (روزمره)", "Display Title"),
                 ("مانده اولیه (ریال)", "Opening Balance"), ("فعال", "Active"),
                 ("توضیحات", "Notes")],
        widths=[10, 30, 17, 8, 40],
        fmts={3: "#,##0"},
        title="بانک‌ها و حساب‌های شرکت — ۱۰ حساب (کد ثابت، عنوان آزاد)",
        hint="کد حساب‌ها ثابت است (BANK01 تا BANK10) و در همه فایل‌ها همان است؛ عنوان نمایشی را آزادانه عوض کنید — فایل تراکنش‌ها خودکار به‌روز می‌شود. مشخصات رسمی (شعبه، شماره حساب، شبا و…) در سربرگ هر برگهٔ بانکی در فایل تراکنش‌ها به‌صورت دستی وارد می‌شود.",
        seed=SEED_BANKS, ncols=5, tab="orange",
    ),
    "Customers": dict(
        headers=[("کد", "Code"), ("نوع طرف حساب", "Party Type"), ("نام / نام شرکت", "Name"),
                 ("تلفن", "Phone"), ("همراه", "Mobile"), ("نشانی", "Address"),
                 ("توضیحات", "Notes")],
        widths=[9, 20, 26, 15, 15, 32, 26],
        title="دفتر مشتریان",
        hint="کدها از C001 شروع می‌شوند و از قبل درج شده‌اند؛ فقط نام و مشخصات را بنویسید. کد را تغییر ندهید — ردیف‌های ثبت‌شده با همین کد خوانده می‌شوند.",
        seed=SEED_CUSTOMERS, ncols=7, tab="green",
    ),
    "Sellers": dict(
        headers=[("کد", "Code"), ("نام صاحب حساب", "Account Holder"), ("نوع مشتری", "Customer Type"),
                 ("تلفن", "Phone"), ("همراه", "Mobile"), ("نشانی", "Address"),
                 ("توضیحات", "Notes")],
        widths=[10, 30, 20, 15, 15, 32, 26],
        title="دفتر فروشندگان (خریداران-بازفروش‌ها) — منبع کد حساب در فایل فروش",
        hint="کدها از SL001 شروع می‌شوند و از قبل درج شده‌اند. «نوع مشتری» را از کشویی انتخاب کنید — انواع آن در شیت PartyTypes تعریف می‌شود (حجره‌دار، صادراتی، واسطه، نقدی و…) و هر زمان قابل گسترش است.",
        seed=SEED_SELLERS, ncols=7, tab="green",
    ),
    "Suppliers": dict(
        headers=[("کد", "Code"), ("نوع طرف حساب", "Party Type"), ("نام / باغ‌دار / حجره", "Name"),
                 ("تلفن", "Phone"), ("همراه", "Mobile"), ("نشانی", "Address"),
                 ("توضیحات", "Notes")],
        widths=[9, 20, 26, 15, 15, 32, 26],
        title="دفتر تأمین‌کنندگان (حجره‌داران میدان بار، باغ‌داران و عوامل)",
        hint="کدها از S001 شروع می‌شوند.",
        seed=SEED_SUPPLIERS, ncols=7, tab="green",
    ),
    "Accounts": dict(
        headers=[("کد", "Code"), ("نوع", "Type"), ("نام حساب", "Account Name"),
                 ("توضیحات", "Notes")],
        widths=[9, 20, 26, 40],
        title="سایر حساب‌ها و عوامل (اداری، داخلی و متفرقه)",
        hint="کدها از X001 شروع می‌شوند — برای طرف‌هایی که مشتری یا تأمین‌کننده نیستند (کارمزد بانکی، مالیات، صندوق و…).",
        seed=SEED_ACCOUNTS, ncols=4, tab="green",
    ),
    "Staff": dict(
        headers=[("کد", "Code"), ("نام", "Name"), ("سمت", "Role"),
                 ("همراه", "Mobile"), ("توضیحات", "Notes")],
        widths=[9, 24, 18, 16, 34],
        title="کارکنان و فروشندگان",
        hint="کدها از E001 شروع می‌شوند — برای هر فروشنده/کارمند کد جدا بدهید تا فایل‌های آینده (فروش هر فروشنده و…) روی همین کد ساخته شود.",
        seed=SEED_STAFF, ncols=5, tab="green",
    ),
    "PartyTypes": dict(
        headers=[("کد", "Code"), ("عنوان", "Title"), ("گروه", "Group"), ("توضیح", "Note")],
        widths=[9, 24, 16, 36],
        title="انواع طرف حساب (در فایل تراکنش‌ها به‌صورت کشویی انتخاب می‌شود)",
        hint="این لیست منبع کشوی «نوع طرف حساب» در شیت Entry است؛ می‌توانید هر تعداد نوع جدید تعریف کنید.",
        seed=SEED_PARTY_TYPES, ncols=4, tab="gray",
    ),
    "TxnTypes": dict(
        headers=[("کد", "Code"), ("عنوان", "Title"), ("تأثیر بر مانده", "Balance Effect"),
                 ("توضیح", "Note")],
        widths=[9, 20, 18, 36],
        title="انواع تراکنش (قابل توسعه برای ماهیت‌های آینده)",
        hint="ستون نوع تراکنش در Entry فقط اطلاعاتی/تحلیلی است؛ واریز و برداشت از ستون‌های مبلغ خوانده می‌شود. در آینده می‌توانید ماهیت‌های تازه (انتقال، اصلاحیه و…) تعریف کنید.",
        seed=SEED_TXN_TYPES, ncols=4, tab="gray",
    ),
    "Products_Buy": dict(
        headers=[("کد", "Code"), ("نام محصول", "Product"), ("دسته", "Category"),
                 ("واحد", "Unit"), ("توضیحات", "Notes")],
        widths=[9, 24, 14, 12, 34],
        title="محصولات خرید (ورودی کارخانه)",
        hint="کدها از PB01 شروع می‌شوند.",
        seed=SEED_PRODUCTS_BUY, ncols=5, tab="olive",
    ),
    "Products_Sell": dict(
        headers=[("کد", "Code"), ("نام محصول", "Product"), ("دسته (نوع مرکبات)", "Category"),
                 ("واحد", "Unit"), ("نوع سبد", "Basket"), ("توضیحات", "Notes")],
        widths=[9, 24, 15, 11, 22, 30],
        title="محصولات فروش (خروجی سورتینگ و بسته‌بندی)",
        hint="کدها از PS01 شروع می‌شوند. «نوع سبد» را از کشویی انتخاب کنید (از شیت Baskets) — در فایل فروش، با انتخاب هر محصول، نوع سبد و نوع مرکبات آن خودکار نوشته می‌شود.",
        seed=SEED_PRODUCTS_SELL, ncols=6, tab="olive",
    ),
    "Baskets": dict(
        headers=[("کد", "Code"), ("نام", "Name"), ("ظرفیت (کیلوگرم)", "Capacity"),
                 ("توضیحات", "Notes"), ("برچسب کشویی (خودکار)", "Label (auto)")],
        widths=[9, 22, 15, 28, 30],
        title="انواع سبدها و ظروف",
        hint="کدها از BK01 شروع می‌شوند. ستون آخر خودکار است و در کشوی «نوع سبد» محصولات استفاده می‌شود.",
        seed=SEED_BASKETS, ncols=5, tab="olive",
    ),
    "Warehouses": dict(
        headers=[("کد", "Code"), ("نام انبار", "Warehouse"), ("موقعیت", "Location"),
                 ("انباردار", "Keeper"), ("توضیحات", "Notes")],
        widths=[9, 24, 18, 14, 30],
        title="انبارها",
        hint="کدها از W01 شروع می‌شوند.",
        seed=SEED_WAREHOUSES, ncols=5, tab="olive",
    ),
}

CAP_OF = dict(Banks="banks", Customers="customers", Sellers="sellers", Suppliers="suppliers",
              Accounts="accounts", Staff="staff", PartyTypes="party_types",
              TxnTypes="txn_types", Products_Buy="products_buy",
              Products_Sell="products_sell", Baskets="baskets",
              Warehouses="warehouses")

PREFIX_OF = dict(Banks="BANK", Customers="C", Sellers="SL", Suppliers="S", Accounts="X",
                 Staff="E", PartyTypes="PT", TxnTypes="T", Products_Buy="PB",
                 Products_Sell="PS", Baskets="BK", Warehouses="W")

PAD_OF = dict(Banks=2, Customers=3, Sellers=3, Suppliers=3, Accounts=3, Staff=3,
              PartyTypes=2, TxnTypes=2, Products_Buy=2, Products_Sell=2,
              Baskets=2, Warehouses=2)


def code_prefix_rows(sheet, n, cap_obj):
    """Prefilled code list for a register, e.g. BANK01..BANK10 or C001..C200."""
    pref = PREFIX_OF[sheet]
    pad = PAD_OF[sheet]
    return [f"{pref}{str(i+1).zfill(pad)}" for i in range(n)]


def build_register_sheet(wb, name, cap, dv_sources=None):
    """Build one INFO register sheet inside workbook wb (also used for QA internal build)."""
    spec = REG_SPECS[name]
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = C[f"tab_{spec['tab']}"]
    rtl(ws)
    ncols = spec["ncols"]
    brand_band(ws, ncols, mode="info")
    title_row(ws, ncols, 3, spec["title"])
    # hint row
    merge_put(ws, f"A4:{get_column_letter(ncols)}4", spec["hint"],
              fnt(8, False, C["gray"], italic=True), fl(C["white"]), al("right", wrap=True))
    ws.row_dimensions[4].height = 26
    table_headers(ws, 5, 6, spec["headers"])
    n = getattr(cap, CAP_OF[name])
    codes = code_prefix_rows(name, n, cap)
    # data area styling
    kinds = {1: "auto"}
    fmts = {i + 1: f for i, f in (spec.get("fmts") or {}).items()}
    data_area(ws, S_DATA0, S_DATA0 + n - 1, 1, ncols, kinds, fmts)
    # prefill codes + seed rows
    for i, code in enumerate(codes):
        r = S_DATA0 + i
        ws.cell(row=r, column=1, value=code).font = fnt(9, True, C["green_d"])
    # Baskets: auto label column (last col) used by dropdowns elsewhere
    if name == "Baskets":
        lc = get_column_letter(spec["ncols"])
        for i in range(n):
            r = S_DATA0 + i
            ws[f"{lc}{r}"] = (f'=IF($A{r}="","",$A{r}&" – "&IF($B{r}="","(بدون عنوان)",$B{r}))')
            ws[f"{lc}{r}"].font = fnt(8.5, False, C["gray"])
            ws[f"{lc}{r}"].alignment = al("center")
            ws[f"{lc}{r}"].fill = fl(C["auto"])
    for i, row in enumerate(spec["seed"]):
        if i >= n:
            break
        r = S_DATA0 + i
        for j, v in enumerate(row):
            if j == 0:
                continue  # code already set
            cell = ws.cell(row=r, column=j + 1, value=(v if v != "" else None))
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = "#,##0"
                cell.alignment = al("center")
    set_widths(ws, {get_column_letter(i + 1): w for i, w in enumerate(spec["widths"])})
    # freeze + print
    ws.freeze_panes = "A7"
    setup_print(ws, "portrait", title_rows="5:6",
                print_area=f"A1:{get_column_letter(ncols)}{S_DATA0 + n - 1}",
                header_center=spec["title"])
    # data validations
    if dv_sources:
        for col_letter, dv_spec in dv_sources.items():
            dv = DataValidation(type="list", formula1=dv_spec, allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"{col_letter}{S_DATA0}:{col_letter}{S_DATA0 + n - 1}")
    return ws


def build_info(cap, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.title = "Mazandasht INFO — فایل مادر"
    wb.properties.creator = BRAND_NAME
    wb.properties.company = "Mazandasht"
    wb.calculation.fullCalcOnLoad = True

    # ---------- Lists (built first: source of INFO dropdowns) ----------
    ws = wb.create_sheet("Lists")
    ws.sheet_properties.tabColor = C["tab_gray"]
    rtl(ws)
    brand_band(ws, 4, mode="info")
    title_row(ws, 4, 3, "لیست‌های عمومی (منبع کشوهای این فایل)")
    lists_headers = [("واحد شمارش", "Units"), ("سمت‌ها", "Roles"),
                     ("دسته محصولات", "Categories"), ("بله/خیر", "Yes/No")]
    table_headers(ws, 5, 6, lists_headers)
    lists_data = [SEED_LISTS_UNITS, SEED_LISTS_ROLES, SEED_LISTS_CATS, SEED_LISTS_YESNO]
    max_len = max(len(x) for x in lists_data)
    data_area(ws, S_DATA0, S_DATA0 + max_len - 1, 1, 4, {}, zebra=True)
    for ci, values in enumerate(lists_data):
        for ri, v in enumerate(values):
            ws.cell(row=S_DATA0 + ri, column=ci + 1, value=v)
    set_widths(ws, {"A": 16, "B": 16, "C": 16, "D": 10})
    ws.freeze_panes = "A7"
    setup_print(ws, "portrait", title_rows="5:6", header_center="لیست‌های عمومی")

    # ---------- Company ----------
    ws = wb.create_sheet("Company")
    ws.sheet_properties.tabColor = C["tab_teal"]
    rtl(ws)
    brand_band(ws, 8, mode="info")
    title_row(ws, 8, 3, "مشخصات شرکت (به‌صورت خودکار در سرصفحه فایل‌های دیگر درج می‌شود)")
    ws.row_dimensions[4].height = 8
    r = 5
    for i, (label, value) in enumerate(SEED_COMPANY):
        row = r + i
        put(ws, f"B{row}", label, fnt(9.5, True, C["green_d"]), fl(C["green_l"]),
            al("right"), BORDER_ALL)
        merge_put(ws, f"C{row}:G{row}", value, fnt(10), fl(C["input"]), al("right"), BORDER_ALL)
        ws.row_dimensions[row].height = 20
    # logo placeholder
    merge_put(ws, "H5:H14", "جای لوگو\n(Insert ← Picture)", fnt(8, False, C["gray"], italic=True),
              fl(C["white"]), al("center", wrap=True), BORDER_BOX_D)
    set_widths(ws, {"A": 3, "B": 18, "C": 22, "D": 22, "E": 22, "F": 22, "G": 22, "H": 16})
    setup_print(ws, "portrait", header_center="مشخصات شرکت")
    # company value cells: C5..C14 (label B, value C merged C:G) — used by Sync

    # ---------- Registers ----------
    dv_map = {
        "Banks": {"D": "ListYesNo"},
        "Customers": {"B": "ListPT"},
        "Sellers": {"C": "ListPT"},
        "Suppliers": {"B": "ListPT"},
        "Accounts": {"B": "ListPT"},
        "Staff": {"C": "ListRoles"},
        "Products_Buy": {"C": "ListCategories", "D": "ListUnits"},
        "Products_Sell": {"C": "ListCategories", "D": "ListUnits", "E": "ListBaskets"},
        "Warehouses": {"D": "ListRoles"},
    }
    for name in ["Banks", "Customers", "Sellers", "Suppliers", "Accounts", "Staff",
                 "PartyTypes", "TxnTypes", "Products_Buy", "Products_Sell",
                 "Baskets", "Warehouses"]:
        build_register_sheet(wb, name, cap, dv_map.get(name))

    # INFO defined names
    info_names = {
        "ListUnits": f"Lists!$A${S_DATA0}:$A${S_DATA0 + 25}",
        "ListRoles": f"Lists!$B${S_DATA0}:$B${S_DATA0 + 25}",
        "ListCategories": f"Lists!$C${S_DATA0}:$C${S_DATA0 + 25}",
        "ListYesNo": f"Lists!$D${S_DATA0}:$D${S_DATA0 + 1}",
        "ListPT": f"PartyTypes!$B${S_DATA0}:$B${S_DATA0 + cap.party_types - 1}",
        "ListBaskets": (f"OFFSET(Baskets!$E${S_DATA0},0,0,"
                        f"MAX(COUNTIF(Baskets!$E${S_DATA0}:$E${S_DATA0 + cap.baskets - 1},\"?*\"),1),1)"),
    }
    for n, ref in info_names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)

    # ---------- Home (dashboard / guide) ----------
    ws = wb.create_sheet("Home", 0)
    ws.sheet_properties.tabColor = C["green_d"]
    rtl(ws, 100)
    brand_band(ws, 8, mode="info")
    merge_put(ws, "A3:H3", "فایل مادر (INFO) — مرکز تعریف داده‌های سامانه مازندشت",
              fnt(13, True, C["orange_d"]), fl(C["green_xl"]), al("center"),
              bd(b=sd("medium", C["orange"])))
    ws.row_dimensions[3].height = 24

    stats = [
        ("بانک‌ها", f"=COUNTA(Banks!$B${S_DATA0}:$B${S_DATA0 + cap.banks - 1})"),
        ("مشتریان", f"=COUNTA(Customers!$C${S_DATA0}:$C${S_DATA0 + cap.customers - 1})"),
        ("فروشندگان", f"=COUNTA(Sellers!$B${S_DATA0}:$B${S_DATA0 + cap.sellers - 1})"),
        ("تأمین‌کنندگان", f"=COUNTA(Suppliers!$C${S_DATA0}:$C${S_DATA0 + cap.suppliers - 1})"),
        ("سایر حساب‌ها", f"=COUNTA(Accounts!$C${S_DATA0}:$C${S_DATA0 + cap.accounts - 1})"),
        ("کارکنان", f"=COUNTA(Staff!$B${S_DATA0}:$B${S_DATA0 + cap.staff - 1})"),
        ("محصولات خرید", f"=COUNTA(Products_Buy!$B${S_DATA0}:$B${S_DATA0 + cap.products_buy - 1})"),
        ("محصولات فروش", f"=COUNTA(Products_Sell!$B${S_DATA0}:$B${S_DATA0 + cap.products_sell - 1})"),
        ("سبدها", f"=COUNTA(Baskets!$B${S_DATA0}:$B${S_DATA0 + cap.baskets - 1})"),
        ("انبارها", f"=COUNTA(Warehouses!$B${S_DATA0}:$B${S_DATA0 + cap.warehouses - 1})"),
    ]
    r0 = 5
    put(ws, f"A{r0}", "آمار ثبت‌شده‌ها:", fnt(9, True, C["green_d"]), None, al("right"))
    for i, (lbl, formula) in enumerate(stats):
        col = get_column_letter(i + 1)
        put(ws, f"{col}{r0 + 1}", lbl, fnt(7.5, True, C["gray"]), fl(C["green_l"]),
            al("center"), BORDER_ALL)
        put(ws, f"{col}{r0 + 2}", formula, fnt(12, True, C["txt"]), fl(C["white"]),
            al("center"), BORDER_ALL, "0")
    ws.row_dimensions[r0 + 1].height = 14
    ws.row_dimensions[r0 + 2].height = 20

    # sheet index
    r = r0 + 4
    merge_put(ws, f"A{r}:H{r}", "فهرست شیت‌های این فایل", fnt(11, True, C["green_d"]), None, al("right"))
    idx = [
        ("Company", "مشخصات شرکت — منبع سرصفحه همه فایل‌ها"),
        ("Banks", "بانک‌ها — کد، عنوان روزمره، مانده اولیه، فعال بودن"),
        ("Customers", "مشتریان — کد C…"),
        ("Sellers", "فروشندگان (خریداران-بازفروش‌ها) — کد SL… و نوع مشتری (کشویی)"),
        ("Suppliers", "تأمین‌کنندگان (حجره‌داران، باغ‌داران، عوامل) — کد S…"),
        ("Accounts", "سایر حساب‌ها (کارمزد، مالیات، صندوق و…) — کد X…"),
        ("Staff", "کارکنان و فروشندگان — کد E…"),
        ("PartyTypes", "انواع طرف حساب — منبع کشوی Entry"),
        ("TxnTypes", "انواع تراکنش — قابل توسعه برای ماهیت‌های آینده"),
        ("Products_Buy", "محصولات خرید — کد PB…"),
        ("Products_Sell", "محصولات فروش — کد PS…"),
        ("Baskets", "انواع سبدها — کد BK…"),
        ("Warehouses", "انبارها — کد W…"),
        ("Lists", "لیست‌های عمومی (واحد، سمت، دسته، بله/خیر)"),
    ]
    for i, (nm, desc) in enumerate(idx):
        rr = r + 1 + i
        put(ws, f"A{rr}", nm, fnt(9, True, C["blue"] if False else C["green_d"]),
            fl(C["auto"]), al("left"), BORDER_ALL)
        merge_put(ws, f"B{rr}:H{rr}", desc, fnt(9), fl(C["white"] if i % 2 == 0 else C["zebra"]),
                  al("right"), BORDER_ALL)
        ws.row_dimensions[rr].height = 17

    # coding rules
    r = r + len(idx) + 3
    merge_put(ws, f"A{r}:H{r}", "قواعد کدگذاری (در همه فایل‌ها یکسان است)", fnt(11, True, C["green_d"]),
              None, al("right"))
    rules = [
        ("بانک‌ها", "BANK01 … BANK10", "کد ثابت — عنوان آزاد"),
        ("مشتریان", "C001, C002, …", "کد ثابت پس از تعریف"),
        ("فروشندگان", "SL001, SL002, …", "منبع کد حساب فایل فروش"),
        ("تأمین‌کنندگان", "S001, S002, …", ""),
        ("سایر حساب‌ها", "X001, X002, …", ""),
        ("کارکنان/فروشندگان", "E001, E002, …", ""),
        ("انواع طرف حساب", "PT01 …", "حجره‌دار، صادراتی، واسطه، نقدی و… — قابل گسترش"),
        ("انواع تراکنش", "T01 …", ""),
        ("محصولات خرید / فروش", "PB01… / PS01…", "محصول فروش: نوع سبد + نوع مرکبات دارد"),
        ("سبدها / انبارها", "BK01… / W01…", ""),
    ]
    hdr_r = r + 1
    for i, h in enumerate(["موضوع", "قالب کد", "نکته"]):
        c0 = ["A", "C", "F"][i]
        c1 = ["B", "E", "H"][i]
        merge_put(ws, f"{c0}{hdr_r}:{c1}{hdr_r}", h, fnt(9, True, C["white"]),
                  fl(C["green"]), al("center"), bd(b=sd("medium", C["orange"])))
    for i, (a, b, c_) in enumerate(rules):
        rr = hdr_r + 1 + i
        merge_put(ws, f"A{rr}:B{rr}", a, fnt(9, True), fl(C["white"]), al("right"), BORDER_ALL)
        merge_put(ws, f"C{rr}:E{rr}", b, fnt(9, name="Consolas"), fl(C["auto"]), al("center"), BORDER_ALL)
        merge_put(ws, f"F{rr}:H{rr}", c_, fnt(8.5, False, C["gray"]), fl(C["white"]), al("right"), BORDER_ALL)
        ws.row_dimensions[rr].height = 17

    # golden rules
    r = hdr_r + len(rules) + 2
    merge_put(ws, f"A{r}:H{r}", "قواعد طلایی", fnt(11, True, C["orange_d"]), None, al("right"))
    golden = [
        "۱) هرگز کدها را تغییر ندهید؛ عنوان‌ها را آزادانه عوض کنید (سیستم با «کد» کار می‌کند، نه با عنوان).",
        "۲) ردیف تازه = ردیف خالی آماده بعدی؛ کدها از قبل درج شده‌اند. بین ردیف‌ها ردیف خالی نگذارید.",
        "۳) برای حذف، محتوای سلول‌ها را پاک کنید (Delete) — خودِ ردیف را حذف نکنید.",
        "۴) این فایل و فایل تراکنش‌ها همیشه باید در یک پوشه باشند؛ مسیر لینک نسبی است.",
        "۵) بعد از تغییر در این فایل، در فایل تراکنش‌ها: Data ← Edit Links ← Update Values.",
        "۶) تاریخ‌ها همیشه شمسی با قالب 1405/07/10 (یا 14050710) و ساعت با قالب 09:30 (یا 9:30).",
        "۷) اعداد را با ارقام لاتین (انگلیسی) وارد کنید تا فرمول‌ها درست کار کنند.",
    ]
    for i, g in enumerate(golden):
        rr = r + 1 + i
        merge_put(ws, f"A{rr}:H{rr}", g, fnt(9), fl(C["white"] if i % 2 == 0 else C["zebra"]),
                  al("right", wrap=True), BORDER_ALL)
        ws.row_dimensions[rr].height = 18

    set_widths(ws, {"A": 13, "B": 13, "C": 11, "D": 11, "E": 11, "F": 13, "G": 13, "H": 13})
    setup_print(ws, "portrait", header_center="راهنمای فایل مادر")

    # order sheets exactly INFO_SHEET_ORDER
    order = INFO_SHEET_ORDER
    wb._sheets = sorted(wb._sheets, key=lambda s: order.index(s.title))
    wb.active = 0
    wb.save(out_path)
    return out_path


# ============================================================================
# 5. TRANSACTIONS WORKBOOK
# ============================================================================

# Entry columns (1-based): A..P
ENTRY_COLS = [
    ("ردیف", "No."), ("شناسه", "TxnID"), ("تاریخ", "Date"), ("ساعت", "Time"),
    ("بانک", "Bank"), ("نام بانک", "Bank Name"), ("کد طرف حساب", "Party Code"),
    ("نام طرف حساب", "Party"), ("نوع طرف حساب", "Party Type"),
    ("نوع تراکنش", "Txn Type"), ("مبلغ واریز (ریال)", "Deposit"),
    ("مبلغ برداشت (ریال)", "Withdrawal"), ("موجودی بانک (ریال)", "Bank Balance"),
    ("موجودی سیستم (ریال)", "System Balance"), ("وضعیت", "Status"),
    ("توضیحات", "Notes"),
]
ENTRY_WIDTHS = [6, 10, 11, 8, 27, 24, 10, 24, 16, 13, 15, 15, 16, 16, 9, 38]
ENTRY_INPUT_COLS = {3, 4, 5, 7, 9, 10, 11, 12, 13, 16}     # C D E G I J K L M P
ENTRY_FMT_COLS = {11: "#,##0", 12: "#,##0", 13: "#,##0", 14: "#,##0"}

BANK_COLS = [
    ("ردیف", "No."), ("تاریخ", "Date"), ("ساعت", "Time"), ("کد طرف", "Party Code"),
    ("طرف حساب", "Party"), ("نوع طرف حساب", "Party Type"), ("نوع تراکنش", "Txn Type"),
    ("واریز (ریال)", "Deposit"), ("برداشت (ریال)", "Withdrawal"),
    ("مانده سیستم (ریال)", "System Balance"), ("موجودی بانک (ریال)", "Bank Balance"),
    ("وضعیت", "Status"), ("کنترل تکرار", "Dup Check"), ("توضیحات", "Notes"),
    ("شناسه", "TxnID"), ("مرجع", "Ref"),
]
BANK_WIDTHS = [5.5, 11, 7.5, 10, 23, 15, 13, 14, 14, 16, 16, 13, 10, 27, 11, 6]
BANK_FMT_COLS = {8: "#,##0", 9: "#,##0", 10: "#,##0", 11: "#,##0"}


def sync_ref(sheet, cell, external, link=1):
    """External/local reference used by the Sync mirrors."""
    return f"[{link}]{sheet}!{cell}" if external else f"{sheet}!{cell}"


def build_sync(wb, cap, external):
    n0 = SY_DATA0
    ws = wb.create_sheet("Sync")
    ws.sheet_properties.tabColor = C["tab_gray"]
    ws.sheet_view.rightToLeft = False
    ws.sheet_view.showGridLines = False
    brand_band(ws, 46, mode="info")
    merge_put(ws, "A3:AT3",
              "برگه فنی — آینه‌ی فایل مادر (INFO). فرمول‌ها را تغییر ندهید. "
              "اگر همه سلول‌ها خالی/صفر شدند یعنی لینک به‌روز نشده: Data ← Edit Links ← Update Values.",
              fnt(8, True, C["red"]), fl(C["amber_l"]), al("left", wrap=True))

    def mirror(col, sheet, info_col, rows, text=True, num=False):
        """Write mirror formulas for a column block; info rows start at S_DATA0."""
        for i in range(rows):
            r = n0 + i
            ir = S_DATA0 + i
            ref = sync_ref(sheet, f"${info_col}${ir}", external)
            if num:
                ws[f"{col}{r}"] = f"={ref}"
            else:
                ws[f"{col}{r}"] = f'=IF({ref}="","",{ref})'

    def block_header(col, title):
        put(ws, f"{col}5", title, fnt(8, True, C["white"]), fl(C["green"]),
            al("center", wrap=True), BORDER_ALL)

    # --- Banks (A:D) ---
    for cl, t in zip("ABCD", ["Code", "Title", "Opening", "Active"]):
        block_header(cl, f"Banks · {t}")
    mirror("A", "Banks", "A", cap.banks)
    mirror("B", "Banks", "B", cap.banks)
    mirror("C", "Banks", "C", cap.banks, num=True)
    mirror("D", "Banks", "D", cap.banks)
    # --- Customers (F:G) ---
    block_header("F", "Customers · Code"); block_header("G", "Customers · Name")
    mirror("F", "Customers", "A", cap.customers)
    mirror("G", "Customers", "C", cap.customers)
    # --- Suppliers (I:J) ---
    block_header("I", "Suppliers · Code"); block_header("J", "Suppliers · Name")
    mirror("I", "Suppliers", "A", cap.suppliers)
    mirror("J", "Suppliers", "C", cap.suppliers)
    # --- Accounts (L:M) ---
    block_header("L", "Accounts · Code"); block_header("M", "Accounts · Name")
    mirror("L", "Accounts", "A", cap.accounts)
    mirror("M", "Accounts", "C", cap.accounts)
    # --- Staff (O:P) ---
    block_header("O", "Staff · Code"); block_header("P", "Staff · Name")
    mirror("O", "Staff", "A", cap.staff)
    mirror("P", "Staff", "B", cap.staff)
    # --- PartyTypes (R:S) ---
    block_header("R", "PartyTypes · Code"); block_header("S", "PartyTypes · Title")
    mirror("R", "PartyTypes", "A", cap.party_types)
    mirror("S", "PartyTypes", "B", cap.party_types)
    # --- TxnTypes (U:V) ---
    block_header("U", "TxnTypes · Code"); block_header("V", "TxnTypes · Title")
    mirror("U", "TxnTypes", "A", cap.txn_types)
    mirror("V", "TxnTypes", "B", cap.txn_types)
    # --- Products_Buy (X:Y) ---
    block_header("X", "Products_Buy · Code"); block_header("Y", "Products_Buy · Name")
    mirror("X", "Products_Buy", "A", cap.products_buy)
    mirror("Y", "Products_Buy", "B", cap.products_buy)
    # --- Products_Sell (AA:AB) ---
    block_header("AA", "Products_Sell · Code"); block_header("AB", "Products_Sell · Name")
    mirror("AA", "Products_Sell", "A", cap.products_sell)
    mirror("AB", "Products_Sell", "B", cap.products_sell)
    # --- Baskets (AD:AE) ---
    block_header("AD", "Baskets · Code"); block_header("AE", "Baskets · Name")
    mirror("AD", "Baskets", "A", cap.baskets)
    mirror("AE", "Baskets", "B", cap.baskets)
    # --- Warehouses (AG:AH) ---
    block_header("AG", "Warehouses · Code"); block_header("AH", "Warehouses · Name")
    mirror("AG", "Warehouses", "A", cap.warehouses)
    mirror("AH", "Warehouses", "B", cap.warehouses)
    # --- Company (AJ) ---
    block_header("AJ", "Company")
    comp_cells = [("$C$5", "name"), ("$C$8", "address"), ("$C$9", "phone"), ("$C$7", "activity")]
    for i, (ref, _nm) in enumerate(comp_cells):
        ext = sync_ref("Company", ref, external)
        ws[f"AJ{n0 + i}"] = f'=IF({ext}="","",{ext})'
    put(ws, "AJ10", "→ نام شرکت", fnt(7, False, C["gray"]))
    put(ws, "AJ11", "→ نشانی", fnt(7, False, C["gray"]))
    put(ws, "AJ12", "→ تلفن", fnt(7, False, C["gray"]))
    put(ws, "AJ13", "→ زمینه فعالیت", fnt(7, False, C["gray"]))

    # --- packed dropdown lists ---
    def packed(col_cum, col_out, src_col, rows):
        """Cumulative-count packing of a mirror column into a gapless list."""
        for i in range(rows):
            r = n0 + i
            ws[f"{col_cum}{r}"] = f'=COUNTIF(${src_col}${n0}:${src_col}{r},"?*")'
            m = f'MATCH(ROW()-{n0 - 1},${col_cum}${n0}:${col_cum}${n0 + rows - 1},0)'
            ws[f"{col_out}{r}"] = (
                f'=IFERROR(INDEX(${src_col}${n0}:${src_col}${n0 + rows - 1},{m}),"")'
            )

    # party type titles packed: cum AN, list AO
    block_header("AN", "cum·PT"); block_header("AO", "LIST · PartyTypes")
    packed("AN", "AO", "S", cap.party_types)
    # txn type titles packed: cum AP, list AQ
    block_header("AP", "cum·TT"); block_header("AQ", "LIST · TxnTypes")
    packed("AP", "AQ", "V", cap.txn_types)
    # bank codes packed: cum AR (on codes), code AS
    block_header("AR", "cum·B"); block_header("AS", "LIST · Bank codes"); block_header("AT", "LIST · Banks")
    packed("AR", "AS", "A", cap.banks)
    for i in range(cap.banks):
        r = n0 + i
        ws[f"AT{r}"] = (
            f'=IF($AS{r}="","",$AS{r}&" – "&IF(IFERROR(VLOOKUP($AS{r},$A${n0}:$D${n0 + cap.banks - 1},2,0),"")="",'
            f'"(بدون عنوان)",VLOOKUP($AS{r},$A${n0}:$D${n0 + cap.banks - 1},2,0)))'
        )

    # --- sellers mirror (BA:BB) — sellers can be bank parties too ---
    block_header("BA", "Sellers · Code"); block_header("BB", "Sellers · Name")
    mirror("BA", "Sellers", "A", cap.sellers)
    mirror("BB", "Sellers", "B", cap.sellers)

    widths = {"A": 9, "B": 26, "C": 14, "D": 8, "F": 9, "G": 26, "I": 9, "J": 26,
              "L": 9, "M": 26, "O": 9, "P": 24, "R": 9, "S": 22, "U": 9, "V": 22,
              "X": 9, "Y": 24, "AA": 9, "AB": 24, "AD": 9, "AE": 20, "AG": 9,
              "AH": 20, "AJ": 30, "AN": 7, "AO": 22, "AP": 7, "AQ": 22,
              "AR": 7, "AS": 22, "AT": 34, "BA": 9, "BB": 28}
    set_widths(ws, widths)
    ws.freeze_panes = "A6"
    return ws


def build_calc(wb, cap):
    n0, n1 = E_DATA0, e_data1(cap)
    ws = wb.create_sheet("Calc")
    ws.sheet_properties.tabColor = C["tab_gray"]
    ws.sheet_view.rightToLeft = False
    ws.sheet_view.showGridLines = False
    brand_band(ws, 13, mode="info")
    merge_put(ws, "A3:M3",
              "برگه فنی — موتور محاسبات (نرمال‌سازی تاریخ/ساعت، مرتب‌سازی و رتبه‌بندی). فرمول‌ها را تغییر ندهید.",
              fnt(8, True, C["red"]), fl(C["amber_l"]), al("left"))
    hdrs = [("تاریخ خام", "Raw Date"), ("سال", "Y"), ("ماه", "M"), ("روز", "D"),
            ("تاریخ استاندارد", "Norm Date"), ("ساعت استاندارد", "Norm Time"),
            ("کد بانک", "Bank Code"), ("کلید مرتب‌سازی", "Sort Key"),
            ("رتبه در بانک", "Rank"), ("کلید رتبه", "Rank Key"),
            ("سال/ماه", "Year-Month"), ("واریز (عدد)", "Dep N"),
            ("برداشت (عدد)", "Wit N")]
    table_headers(ws, 8, 8, [])  # dummy to set heights not needed; use manual
    for i, (fa, en) in enumerate(hdrs):
        col = get_column_letter(i + 1)
        put(ws, f"{col}8", f"{fa} · {en}", fnt(7.5, True, C["white"]), fl(C["gray"]),
            al("center", wrap=True), BORDER_ALL)
    ws.row_dimensions[8].height = 24

    A = "$A{r}"; fmt = None
    for r in range(n0, n1 + 1):
        # A raw date mirror
        ws[f"A{r}"] = f'=IF(Entry!$C{r}="","",Entry!$C{r})'
        # B year
        ws[f"B{r}"] = (
            f'=IF($A{r}="","",IF(ISNUMBER($A{r}),'
            f'IF(AND($A{r}>=10000101,$A{r}<=99991231),INT($A{r}/10000),""),'
            f'IFERROR(VALUE(TRIM(LEFT($A{r},FIND("/",$A{r})-1))),"")))'
        )
        # C month
        ws[f"C{r}"] = (
            f'=IF(OR($A{r}="",$B{r}=""),"",IF(ISNUMBER($A{r}),'
            f'INT(MOD($A{r},10000)/100),'
            f'IFERROR(VALUE(TRIM(LEFT(MID($A{r},FIND("/",$A{r})+1,20),'
            f'FIND("/",MID($A{r},FIND("/",$A{r})+1,20))-1))),"")))'
        )
        # D day
        ws[f"D{r}"] = (
            f'=IF(OR($A{r}="",$B{r}=""),"",IF(ISNUMBER($A{r}),'
            f'MOD($A{r},100),'
            f'IFERROR(VALUE(TRIM(MID(MID($A{r},FIND("/",$A{r})+1,20),'
            f'FIND("/",MID($A{r},FIND("/",$A{r})+1,20))+1,20))),"")))'
        )
        # E normalized date
        ws[f"E{r}"] = (
            f'=IF($A{r}="","",IF(OR($B{r}="",$C{r}="",$D{r}="",$B{r}<1300,$B{r}>1500,'
            f'$C{r}<1,$C{r}>12,$D{r}<1,$D{r}>31),"نامعتبر",'
            f'RIGHT("0000"&$B{r},4)&"/"&RIGHT("00"&$C{r},2)&"/"&RIGHT("00"&$D{r},2)))'
        )
        # F normalized time (blank time on a valid-date row -> "00:00")
        ws[f"F{r}"] = (
            f'=IF(Entry!$D{r}<>"",'
            f'IF(AND(ISNUMBER(Entry!$D{r}),Entry!$D{r}>=0,Entry!$D{r}<1),'
            f'RIGHT("00"&HOUR(Entry!$D{r}),2)&":"&RIGHT("00"&MINUTE(Entry!$D{r}),2),'
            f'IFERROR(RIGHT("00"&VALUE(TRIM(LEFT(Entry!$D{r},FIND(":",Entry!$D{r})-1))),2)&":"&'
            f'RIGHT("00"&VALUE(TRIM(MID(Entry!$D{r},FIND(":",Entry!$D{r})+1,2))),2),"نامعتبر")),'
            f'IF(OR($E{r}="",$E{r}="نامعتبر"),"","00:00"))'
        )
        # G bank code (gated on valid date)
        ws[f"G{r}"] = (
            f'=IF(OR($E{r}="",$E{r}="نامعتبر"),"",'
            f'IF(Entry!$E{r}="","",LEFT(UPPER(TRIM(Entry!$E{r})),6)))'
        )
        # H sort key
        ws[f"H{r}"] = (
            f'=IF($G{r}="","",LEFT($E{r},4)&MID($E{r},6,2)&RIGHT($E{r},2)&'
            f'IF($F{r}="","0000",IF($F{r}="نامعتبر","9999",LEFT($F{r},2)&RIGHT($F{r},2)))&'
            f'"|"&RIGHT("00000"&{r - E_DATA0 + 1},5))'
        )
        # I rank within bank
        ws[f"I{r}"] = f'=IF($H{r}="","",COUNTIFS(CalcBank,$G{r},CalcKey,"<"&$H{r})+1)'
        # J rank key
        ws[f"J{r}"] = f'=IF($I{r}="","",$G{r}&"|"&RIGHT("000000"&$I{r},6))'
        # K year-month
        ws[f"K{r}"] = f'=IF(OR($E{r}="",$E{r}="نامعتبر"),"",LEFT($E{r},4)&"/"&MID($E{r},6,2))'
        # L numeric deposit
        ws[f"L{r}"] = f'=IF(Entry!$K{r}="",0,Entry!$K{r})'
        # M numeric withdrawal
        ws[f"M{r}"] = f'=IF(Entry!$L{r}="",0,Entry!$L{r})'

    set_widths(ws, {"A": 12, "B": 6, "C": 6, "D": 6, "E": 12, "F": 10, "G": 10,
                    "H": 20, "I": 9, "J": 16, "K": 9, "L": 15, "M": 15})
    ws.freeze_panes = "A9"
    return ws


def build_entry(wb, cap):
    n0, n1 = E_DATA0, e_data1(cap)
    ws = wb.create_sheet("Entry")
    ws.sheet_properties.tabColor = C["tab_orange"]
    rtl(ws, 80)
    ncols = len(ENTRY_COLS)
    last = get_column_letter(ncols)
    brand_band(ws, ncols, mode="txn")
    merge_put(ws, f"A3:{last}3", "ثبت تراکنش‌های بانکی — ورود داده (Entry)",
              fnt(13, True, C["orange_d"]), fl(C["green_xl"]), al("center"),
              bd(b=sd("medium", C["orange"])))
    ws.row_dimensions[3].height = 24
    merge_put(ws, f"A4:{last}4",
              "ستون‌های زرد = ورود دستی شما | ستون‌های خاکستری = خودکار.  تاریخ: 1405/07/10 (یا 14050710) — ساعت: 09:30 (یا 9:30) — "
              "بانک و نوع‌ها از کشویی انتخاب شوند — کد طرف حساب را تایپ کنید (مثل C001 یا S001 یا X001) تا نامش بیاید.",
              fnt(8, False, C["gray"], italic=True), fl(C["white"]), al("right", wrap=True))
    ws.row_dimensions[4].height = 24

    # KPI mini-cards (rows 5-6)
    kpis = [
        ("تعداد ردیف ثبت‌شده", f"=COUNTA($C${n0}:$C${n1})", "0"),
        ("جمع واریز (ریال)", f"=SUM($K${n0}:$K${n1})", "#,##0"),
        ("جمع برداشت (ریال)", f"=SUM($L${n0}:$L${n1})", "#,##0"),
        (f"تعداد {OK_MARK} منطبق", f'=COUNTIF($O${n0}:$O${n1},"{OK_MARK}")', "0"),
        (f"تعداد {BAD_MARK} مغایر", f'=COUNTIF($O${n0}:$O${n1},"{BAD_MARK}")', "0"),
        ("ظرفیت باقی‌مانده", f"={cap.entry_rows}-COUNTA($C${n0}:$C${n1})", "0"),
    ]
    groups = [(1, 3), (4, 6), (7, 9), (10, 11), (12, 13), (14, 16)]
    for (c0, c1), (lbl, formula, fmt) in zip(groups, kpis):
        card(ws, 5, 6, c0, c1, lbl, formula, fmt)
    ws.row_dimensions[5].height = 13
    ws.row_dimensions[6].height = 20

    table_headers(ws, 7, 8, ENTRY_COLS)
    kinds = {c: ("in" if c in ENTRY_INPUT_COLS else "auto") for c in range(1, ncols + 1)}
    data_area(ws, n0, n1, 1, ncols, kinds, ENTRY_FMT_COLS)

    # demo rows
    for i, row in enumerate(SEED_ENTRY):
        if i >= cap.entry_rows:
            break
        r = n0 + i
        date, time_, bank, party, ptype, ttype, dep, wit, man, note = row
        ws[f"C{r}"] = date
        ws[f"D{r}"] = time_
        ws[f"E{r}"] = bank
        ws[f"G{r}"] = party
        ws[f"I{r}"] = ptype
        ws[f"J{r}"] = ttype
        if dep != "": ws[f"K{r}"] = dep
        if wit != "": ws[f"L{r}"] = wit
        if man != "": ws[f"M{r}"] = man
        ws[f"P{r}"] = note

    # formulas (auto columns)
    for r in range(n0, n1 + 1):
        ws[f"A{r}"] = f'=IF($C{r}="","",COUNTA($C${n0}:$C{r}))'
        ws[f"B{r}"] = f'=IF($C{r}="","","TXN-"&RIGHT("0000"&{r - n0 + 1},4))'
        ws[f"F{r}"] = (
            f'=IF($E{r}="","",IFERROR(IF(VLOOKUP(LEFT(UPPER(TRIM($E{r})),6),BankTbl,2,0)="",'
            f'"(بدون عنوان)",VLOOKUP(LEFT(UPPER(TRIM($E{r})),6),BankTbl,2,0)),"بانک ناشناخته"))'
        )
        ws[f"H{r}"] = (
            f'=IF($G{r}="","",IFERROR(VLOOKUP($G{r},CustTbl,2,0),'
            f'IFERROR(VLOOKUP($G{r},SuppTbl,2,0),IFERROR(VLOOKUP($G{r},AcctTbl,2,0),'
            f'IFERROR(VLOOKUP($G{r},StaffTbl,2,0),IFERROR(VLOOKUP($G{r},SellerTbl,2,0),'
            f'"کد ناشناخته"))))))'
        )
        code = f"LEFT(UPPER(TRIM($E{r})),6)"
        ws[f"N{r}"] = (
            f'=IF(OR($C{r}="",$E{r}="",Calc!$H{r}=""),"",'
            f'IFERROR(VLOOKUP({code},BankTbl,3,0),0)'
            f'+SUMIFS(CalcDep,CalcBank,{code},CalcKey,"<="&Calc!$H{r})'
            f'-SUMIFS(CalcWit,CalcBank,{code},CalcKey,"<="&Calc!$H{r}))'
        )
        ws[f"O{r}"] = (
            f'=IF(OR($C{r}="",$M{r}=""),"",IF($N{r}="","{WARN_MARK}",'
            f'IF(ROUND($M{r}-$N{r},0)=0,"{OK_MARK}","{BAD_MARK}")))'
        )

    # column formats for text-input columns
    for col in ("C", "D", "G"):
        for r in range(n0, n1 + 1):
            ws[f"{col}{r}"].number_format = "@"
    set_widths(ws, {get_column_letter(i + 1): w for i, w in enumerate(ENTRY_WIDTHS)})

    # data validations
    dvs = []
    dv_bank = DataValidation(type="list", formula1="ListBanks", allow_blank=True,
                             promptTitle="بانک", prompt="از فهرست انتخاب کنید یا کد (مثل BANK03) را تایپ کنید.")
    dv_bank.add(f"E{n0}:E{n1}")
    dv_party = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR($G{n0}="",COUNTIF(CustCodes,$G{n0})>0,COUNTIF(SuppCodes,$G{n0})>0,'
                 f'COUNTIF(AcctCodes,$G{n0})>0,COUNTIF(StaffCodes,$G{n0})>0,'
                 f'COUNTIF(SellerCodes,$G{n0})>0)',
        errorTitle="کد ناشناخته",
        error="این کد در فایل INFO یافت نشد (C… مشتری، S… تأمین‌کننده، X… سایر حساب‌ها، E… کارمند، SL… فروشنده).",
        promptTitle="کد طرف حساب", prompt="کد را از فایل INFO تایپ کنید؛ نام به‌صورت خودکار می‌آید.")
    dv_party.add(f"G{n0}:G{n1}")
    dv_ptype = DataValidation(type="list", formula1="ListPartyTypes", allow_blank=True)
    dv_ptype.add(f"I{n0}:I{n1}")
    dv_ttype = DataValidation(type="list", formula1="ListTxnTypes", allow_blank=True)
    dv_ttype.add(f"J{n0}:J{n1}")
    dv_date = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'AND(LEN($C{n0})=10,MID($C{n0},5,1)="/",MID($C{n0},8,1)="/",'
                 f'ISNUMBER(VALUE(LEFT($C{n0},4))),ISNUMBER(VALUE(MID($C{n0},6,2))),'
                 f'ISNUMBER(VALUE(MID($C{n0},9,2)))),AND(ISNUMBER($C{n0}),LEN($C{n0})=8))',
        errorTitle="قالب تاریخ", error="تاریخ شمسی با قالب 1405/07/10 یا 14050710 وارد کنید (ارقام لاتین).")
    dv_date.add(f"C{n0}:C{n1}")
    dv_time = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR(AND(ISNUMBER($D{n0}),$D{n0}>=0,$D{n0}<1),AND(LEN($D{n0})=5,'
                 f'MID($D{n0},3,1)=":",ISNUMBER(VALUE(LEFT($D{n0},2))),'
                 f'ISNUMBER(VALUE(MID($D{n0},4,2)))))',
        errorTitle="قالب ساعت", error="ساعت را با قالب 24 ساعته 09:30 یا 9:30 وارد کنید.")
    dv_time.add(f"D{n0}:D{n1}")
    dv_dep = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR($K{n0}="",AND(ISNUMBER($K{n0}),$K{n0}>=0,$L{n0}=""))',
        errorTitle="واریز/برداشت", error="در هر ردیف فقط یکی از ستون‌های واریز یا برداشت پر شود.")
    dv_dep.add(f"K{n0}:K{n1}")
    dv_wit = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR($L{n0}="",AND(ISNUMBER($L{n0}),$L{n0}>=0,$K{n0}=""))',
        errorTitle="واریز/برداشت", error="در هر ردیف فقط یکی از ستون‌های واریز یا برداشت پر شود.")
    dv_wit.add(f"L{n0}:L{n1}")
    dvs += [dv_bank, dv_party, dv_ptype, dv_ttype, dv_date, dv_time, dv_dep, dv_wit]
    for dv in dvs:
        dv.showErrorMessage = True
        if dv.prompt:
            dv.showInputMessage = True
        ws.add_data_validation(dv)

    # conditional formatting
    ok_fill = fl(C["ok_l"]); ok_font = fnt(10, True, C["ok"])
    bad_fill = fl(C["red_l"]); bad_font = fnt(10, True, C["red"])
    warn_fill = fl(C["amber_l"]); warn_font = fnt(10, True, C["amber"])
    o_rng = f"O{n0}:O{n1}"
    ws.conditional_formatting.add(o_rng, CellIsRule(operator="equal", formula=[f'"{OK_MARK}"'], fill=ok_fill, font=ok_font))
    ws.conditional_formatting.add(o_rng, CellIsRule(operator="equal", formula=[f'"{BAD_MARK}"'], fill=bad_fill, font=bad_font))
    ws.conditional_formatting.add(o_rng, CellIsRule(operator="equal", formula=[f'"{WARN_MARK}"'], fill=warn_fill, font=warn_font))
    for col in ("F", "H"):
        rng = f"{col}{n0}:{col}{n1}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'OR(ISNUMBER(SEARCH("ناشناخته",{col}{n0})),ISNUMBER(SEARCH("بدون عنوان",{col}{n0})))'],
            fill=warn_fill, font=warn_font))
    # databars on amounts
    ws.conditional_formatting.add(f"K{n0}:K{n1}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="81C784", showValue=True))
    ws.conditional_formatting.add(f"L{n0}:L{n1}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="FF8A65", showValue=True))

    # header comments
    tips = {
        "C7": "تاریخ شمسی: 1405/07/10 یا 14050710 (ارقام لاتین).",
        "D7": "ساعت 24 ساعته: 09:30 یا 9:30 — برای خالی بگذارید 00:00 فرض می‌شود.",
        "E7": "بانک را از کشویی انتخاب کنید (BANK01 تا BANK10). عنوان از فایل INFO می‌آید.",
        "G7": "کد طرف حساب: C… مشتری، S… تأمین‌کننده، X… سایر حساب‌ها، E… کارمند، SL… فروشنده. نام به‌صورت خودکار می‌آید — واریز فروشنده در شیت SLxxx فایل فروش هم دیده می‌شود.",
        "I7": "نوع طرف حساب — لیست آن در INFO ← PartyTypes تعریف می‌شود.",
        "J7": "نوع تراکنش (واریز/برداشت/…) — لیست آن در INFO ← TxnTypes. در آینده ماهیت‌های تازه اضافه کنید.",
        "K7": "مبلغ واریز به ریال — در هر ردیف فقط واریز «یا» برداشت پر شود.",
        "L7": "مبلغ برداشت به ریال.",
        "M7": "موجودی واقعی بانک (طبق پیامک/صورت‌حساب) — اختیاری؛ اگر بنویسید، سیستم آن را با مانده محاسباتی مقایسه می‌کند.",
    }
    for ref, text in tips.items():
        ws[ref].comment = Comment(text, "مازندشت", height=90, width=260)

    ws.freeze_panes = "A9"
    setup_print(ws, "landscape", title_rows="7:8",
                print_area=f"A1:{last}{n1}",
                header_center="ثبت تراکنش‌های بانکی — Entry")
    return ws


def build_bank_sheet(wb, idx, cap):
    """idx: 1..10 -> sheet Bank{idx} code BANK{idx:02d}"""
    n0, n1 = B_DATA0, b_data1(cap)
    code = f"BANK{idx:02d}"
    ws = wb.create_sheet(f"Bank{idx}")
    ws.sheet_properties.tabColor = C["tab_green"]
    rtl(ws, 85)
    ncols = 15  # A..O visible (P hidden helper)
    last = get_column_letter(ncols)
    brand_band(ws, ncols, mode="txn")
    merge_put(ws, f"A3:H3",
              f'=IFERROR("صورت حساب بانکی — "&IF(VLOOKUP("{code}",BankTbl,2,0)="","(بدون عنوان)",VLOOKUP("{code}",BankTbl,2,0)),"صورت حساب بانکی")',
              fnt(13, True, C["orange_d"]), fl(C["green_xl"]), al("right"),
              bd(b=sd("medium", C["orange"])))
    merge_put(ws, f"I3:{last}3",
              f'="کد حساب: {code}"', fnt(11, True, C["green_d"]), fl(C["orange_l"]),
              al("center"), bd(b=sd("medium", C["orange"])))
    ws.row_dimensions[3].height = 24
    # row 4: capacity warning + manual-details hint
    merge_put(ws, f"A4:H4",
              f'=IF(COUNTIFS(CalcBank,"{code}")>{cap.bank_rows},"{BAD_MARK} ظرفیت {cap.bank_rows} ردیف این برگه پر شده — آخرین ردیف را به پایین کپی کنید","")',
              fnt(8, True, C["red"]), fl(C["white"]), al("right"))
    merge_put(ws, f"I4:{last}4", "مشخصات رسمی زیر، دستی وارد می‌شود (زردرنگ).",
              fnt(8, False, C["gray"], italic=True), fl(C["white"]), al("right"))
    ws.row_dimensions[4].height = 14

    # ---- details block rows 5-10 ----
    # right: auto from INFO (labels A:B merged, values C:E merged)
    merge_put(ws, "A5:B5", "کد حساب:", fnt(8.5, True, C["gray"]), fl(C["green_l"]),
              al("right"), BORDER_ALL)
    merge_put(ws, "C5:E5", code, fnt(10, True, C["orange_d"]), fl(C["auto"]),
              al("center"), BORDER_BOX_D)
    merge_put(ws, "A6:B6", "عنوان نمایشی:", fnt(8.5, True, C["gray"]), fl(C["green_l"]),
              al("right"), BORDER_ALL)
    merge_put(ws, "C6:E6",
              f'=IFERROR(IF(VLOOKUP("{code}",BankTbl,2,0)="","(بدون عنوان)",VLOOKUP("{code}",BankTbl,2,0)),"—")',
              fnt(10, True, C["green_d"]), fl(C["auto"]), al("right"), BORDER_ALL)
    merge_put(ws, "A7:B7", "مانده اولیه (ریال):", fnt(8.5, True, C["gray"]), fl(C["green_l"]),
              al("right"), BORDER_ALL)
    merge_put(ws, "C7:E7", f'=IFERROR(VLOOKUP("{code}",BankTbl,3,0),0)',
              fnt(10, True, C["txt"]), fl(C["auto"]), al("center"), BORDER_ALL, "#,##0")
    merge_put(ws, "A8:B8", "وضعیت فعال:", fnt(8.5, True, C["gray"]), fl(C["green_l"]),
              al("right"), BORDER_ALL)
    merge_put(ws, "C8:E8",
              f'=IFERROR(IF(VLOOKUP("{code}",BankTbl,4,0)=0,"—",VLOOKUP("{code}",BankTbl,4,0)),"—")',
              fnt(9), fl(C["auto"]), al("center"), BORDER_ALL)
    merge_put(ws, "A9:E10",
              "این چهار سلول از فایل مادر خوانده می‌شوند؛ برای تغییر: Mazandasht_INFO.xlsx ← شیت Banks. مشخصات رسمی (زردرنگ) دستی وارد می‌شود.",
              fnt(7.5, False, C["gray"], italic=True), fl(C["white"]), al("right", wrap=True), BORDER_ALL)
    # left: manual legal details (labels H:I merged, values J:O merged)
    manual = ["نام رسمی بانک:", "شعبه:", "شماره حساب:", "شماره شبا (۲۴ رقم):",
              "صاحب حساب:", "تلفن شعبه:"]
    for i, lbl in enumerate(manual):
        r = 5 + i
        merge_put(ws, f"H{r}:I{r}", lbl, fnt(8.5, True, C["gray"]), fl(C["green_l"]),
                  al("right"), BORDER_ALL)
        merge_put(ws, f"J{r}:{last}{r}", None, fnt(9.5), fl(C["input"]), al("right"), BORDER_ALL)
        ws.row_dimensions[r].height = 18
    for r in (7, 8):  # account & sheba: LTR text
        style_range(ws, f"J{r}:{last}{r}", alignment=al("left"))

    # ---- summary cards rows 12-13 ----
    cards = [
        (1, 2, "جمع واریز (ریال)", f"=SUM($H${n0}:$H${n1})"),
        (3, 4, "جمع برداشت (ریال)", f"=SUM($I${n0}:$I${n1})"),
        (5, 6, "مانده نهایی سیستم", f"=IFERROR(LOOKUP(9.99E+307,$J${n0}:$J${n1}),$C$7)"),
        (7, 8, "آخرین موجودی بانک", f'=IFERROR(LOOKUP(9.99E+307,$K${n0}:$K${n1}),"{NONE_MARK}")'),
        (9, 10, "وضعیت مغایرت",
         f'=IF(NOT(ISNUMBER($G$13)),"{NONE_MARK} بدون کنترل",IF(ROUND($G$13-$E$13,0)=0,'
         f'"{OK_MARK} منطبق","{BAD_MARK} مغایرت ("&TEXT(ROUND($G$13-$E$13,0),"#,##0")&" ریال)"))'),
        (11, 12, "تعداد تراکنش", f"=COUNT($A${n0}:$A${n1})"),
        (13, 14, "ردیف تکراری", f'=COUNTIF($M${n0}:$M${n1},"{DUP_MARK}")'),
        (15, 15, "آخرین تاریخ",
         f'=IF(COUNT($A${n0}:$A${n1})=0,"{NONE_MARK}",'
         f'INDEX($B${n0}:$B${n1},COUNT($A${n0}:$A${n1})))'),
    ]
    for c0, c1, lbl, formula in cards:
        card(ws, 12, 13, c0, c1, lbl, formula)
    ws.row_dimensions[12].height = 13
    ws.row_dimensions[13].height = 22
    ws.row_dimensions[14].height = 6

    # ---- ledger headers 15-16 ----
    table_headers(ws, 15, 16, BANK_COLS[:15])
    kinds = {c: "plain" for c in range(1, 16)}
    data_area(ws, n0, n1, 1, 16, kinds, BANK_FMT_COLS)

    # ---- formulas ----
    for k in range(1, cap.bank_rows + 1):
        r = n0 + k - 1
        ws[f"P{r}"] = (f'=IFERROR(MATCH($C$5&"|"&RIGHT("000000"&{k},6),CalcRankKey,0),"")')
        ws[f"A{r}"] = f'=IF($B{r}="","",{k})'
        def disp(col_name, col):
            return (f'=IF($P{r}="","",IF(INDEX({col_name},$P{r})=0,"",INDEX({col_name},$P{r})))')
        ws[f"B{r}"] = disp("CalcNormDate", 2)
        ws[f"C{r}"] = disp("CalcTime", 3)
        ws[f"D{r}"] = disp("EntryParty", 4)
        ws[f"E{r}"] = disp("EntryPartyName", 5)
        ws[f"F{r}"] = disp("EntryPartyType", 6)
        ws[f"G{r}"] = disp("EntryTxnType", 7)
        ws[f"H{r}"] = disp("EntryDep", 8)
        ws[f"I{r}"] = disp("EntryWit", 9)
        ws[f"J{r}"] = (f'=IF($B{r}="","",IF($A{r}=1,$C$7,IF(J{r-1}="",0,J{r-1}))'
                       f'+IF($H{r}="",0,$H{r})-IF($I{r}="",0,$I{r}))')
        ws[f"K{r}"] = disp("EntryMan", 11)
        ws[f"L{r}"] = (f'=IF($B{r}="","",IF($K{r}="","{NONE_MARK}",'
                       f'IF(ROUND($K{r}-$J{r},0)=0,"{OK_MARK}","{BAD_MARK}")))')
        ws[f"M{r}"] = (f'=IF($B{r}="","",IF(COUNTIFS(EntryDates,$B{r},CalcTime,$C{r},'
                       f'CalcBank,$C$5,EntryParty,$D{r},EntryDep,$H{r},EntryWit,$I{r})>1,'
                       f'"{DUP_MARK}",""))')
        ws[f"N{r}"] = disp("EntryNotes", 14)
        ws[f"O{r}"] = disp("EntryTxnID", 15)

    # status fonts (bold marks) via CF
    ok_fill = fl(C["ok_l"]); ok_font = fnt(10, True, C["ok"])
    bad_fill = fl(C["red_l"]); bad_font = fnt(10, True, C["red"])
    l_rng = f"L{n0}:L{n1}"
    ws.conditional_formatting.add(l_rng, CellIsRule(operator="equal", formula=[f'"{OK_MARK}"'], fill=ok_fill, font=ok_font))
    ws.conditional_formatting.add(l_rng, CellIsRule(operator="equal", formula=[f'"{BAD_MARK}"'], fill=bad_fill, font=bad_font))
    ws.conditional_formatting.add(l_rng, CellIsRule(operator="equal", formula=[f'"{NONE_MARK}"'], font=fnt(9, False, C["gray"])))
    ws.conditional_formatting.add(f"M{n0}:M{n1}", CellIsRule(
        operator="equal", formula=[f'"{DUP_MARK}"'], fill=fl(C["amber_l"]), font=fnt(9, True, C["amber"])))
    ws.conditional_formatting.add(f"J{n0}:J{n1}", CellIsRule(
        operator="lessThan", formula=["0"], font=fnt(9, True, C["red"])))
    ws.conditional_formatting.add(f"H{n0}:H{n1}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="81C784", showValue=True))
    ws.conditional_formatting.add(f"I{n0}:I{n1}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="FF8A65", showValue=True))
    # capacity warning row
    ws.conditional_formatting.add("A4:H4", FormulaRule(
        formula=['LEN($A$4)>0'], fill=fl(C["red_l"]), font=fnt(8, True, C["red"])))
    # card status coloring
    ws.conditional_formatting.add("I13", FormulaRule(
        formula=[f'ISNUMBER(SEARCH("{BAD_MARK}",$I$13))'], fill=bad_fill, font=fnt(11, True, C["red"])))
    ws.conditional_formatting.add("I13", FormulaRule(
        formula=[f'ISNUMBER(SEARCH("{OK_MARK}",$I$13))'], fill=ok_fill, font=fnt(11, True, C["ok"])))

    set_widths(ws, {get_column_letter(i + 1): w for i, w in enumerate(BANK_WIDTHS)})
    ws.column_dimensions["P"].hidden = True
    ws.freeze_panes = "A17"
    setup_print(ws, "landscape", title_rows="15:16",
                print_area=f"A1:{last}{n1}",
                header_center=f"صورت حساب بانکی — {code}")
    return ws


def build_summary(wb, cap):
    ws = wb.create_sheet("Summary")
    ws.sheet_properties.tabColor = C["green_d"]
    rtl(ws, 85)
    ncols = 11  # A..K
    last = "K"
    brand_band(ws, ncols, mode="txn")
    merge_put(ws, "A3:K3", "خلاصه و تحلیل کل حساب‌های بانکی", fnt(13, True, C["orange_d"]),
              fl(C["green_xl"]), al("center"), bd(b=sd("medium", C["orange"])))
    ws.row_dimensions[3].height = 24
    merge_put(ws, "A4:K4",
              '=IF(ISERROR(VLOOKUP("BANK01",BankTbl,2,0)),'
              f'"× اتصال به فایل مادر (INFO) برقرار نیست — مسیر: Data ← Edit Links ← Update Values",'
              f'"√ متصل به فایل مادر (Mazandasht_INFO.xlsx)")',
              fnt(9, True, C["txt"]), fl(C["white"]), al("center"), BORDER_ALL)
    ws.conditional_formatting.add("A4", FormulaRule(
        formula=['ISNUMBER(SEARCH("√",$A$4))'], fill=fl(C["ok_l"]), font=fnt(9, True, C["ok"])))
    ws.conditional_formatting.add("A4", FormulaRule(
        formula=['ISNUMBER(SEARCH("×",$A$4))'], fill=fl(C["red_l"]), font=fnt(9, True, C["red"])))
    ws.row_dimensions[4].height = 18

    # bank table rows: header 11, banks 12..21, totals 22
    t_hdr, t0 = 11, 12
    t1 = t0 + cap.banks - 1
    t_tot = t1 + 1
    hdrs = [("کد", "Code"), ("بانک", "Bank"), ("مانده اولیه", "Opening"),
            ("واریز", "Deposits"), ("برداشت", "Withdrawals"),
            ("مانده نهایی سیستم", "Closing (System)"), ("آخرین موجودی بانک", "Last Bank Balance"),
            ("مغایرت (ریال)", "Difference"), ("وضعیت", "Status"),
            ("تعداد تراکنش", "Txns"), ("تکراری", "Dups")]
    table_headers(ws, t_hdr, t_hdr + 1, hdrs)
    for i in range(cap.banks):
        r = t0 + i
        code = f"BANK{i + 1:02d}"
        sheet = f"Bank{i + 1}"
        ws[f"A{r}"] = code
        ws[f"B{r}"] = f'=IFERROR(IF(VLOOKUP("{code}",BankTbl,2,0)="","(بدون عنوان)",VLOOKUP("{code}",BankTbl,2,0)),"—")'
        ws[f"C{r}"] = f'=IFERROR(VLOOKUP("{code}",BankTbl,3,0),0)'
        ws[f"D{r}"] = f'=SUMIFS(CalcDep,CalcBank,$A{r})'
        ws[f"E{r}"] = f'=SUMIFS(CalcWit,CalcBank,$A{r})'
        ws[f"F{r}"] = f'=$C{r}+$D{r}-$E{r}'
        ws[f"G{r}"] = f'=IFERROR(LOOKUP(9.99E+307,{sheet}!$K${B_DATA0}:$K${b_data1(cap)}),"")'
        ws[f"H{r}"] = f'=IF($G{r}="","{NONE_MARK}",ROUND($G{r}-$F{r},0))'
        ws[f"I{r}"] = f'=IF($G{r}="","{NONE_MARK}",IF(ROUND($H{r},0)=0,"{OK_MARK}","{BAD_MARK}"))'
        ws[f"J{r}"] = f'=COUNTIFS(CalcBank,$A{r})'
        ws[f"K{r}"] = f'={sheet}!$M$13'
    # totals
    ws[f"A{t_tot}"] = "جمع"
    merge_put(ws, f"A{t_tot}:B{t_tot}", "جمع کل", fnt(10, True, C["white"]), fl(C["green"]),
              al("center"), bd(b=sd("medium", C["orange"])))
    for col, f_ in [("C", "SUM"), ("D", "SUM"), ("E", "SUM"), ("F", "SUM"), ("G", "SUM")]:
        ws[f"{col}{t_tot}"] = f"={f_}({col}{t0}:{col}{t1})"
    ws[f"H{t_tot}"] = f'=IF($G{t_tot}="","{NONE_MARK}",ROUND($G{t_tot}-$F{t_tot},0))'
    ws[f"I{t_tot}"] = f'=IF($G{t_tot}="","{NONE_MARK}",IF(ROUND($H{t_tot},0)=0,"{OK_MARK}","{BAD_MARK}"))'
    ws[f"J{t_tot}"] = f"=SUM(J{t0}:J{t1})"
    ws[f"K{t_tot}"] = f"=SUM(K{t0}:K{t1})"
    data_area(ws, t0, t1, 1, ncols, {}, {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0",
                                          7: "#,##0", 8: "#,##0"})
    style_range(ws, f"A{t_tot}:K{t_tot}", font=fnt(10, True, C["green_d"]),
                fill=fl(C["green_l"]), alignment=al("center"), border=BORDER_ALL)
    for col in "CDEFGH":
        ws[f"{col}{t_tot}"].number_format = "#,##0"
    # CF
    ok_fill = fl(C["ok_l"]); bad_fill = fl(C["red_l"])
    ws.conditional_formatting.add(f"I{t0}:I{t_tot}", CellIsRule(operator="equal", formula=[f'"{OK_MARK}"'], fill=ok_fill, font=fnt(9, True, C["ok"])))
    ws.conditional_formatting.add(f"I{t0}:I{t_tot}", CellIsRule(operator="equal", formula=[f'"{BAD_MARK}"'], fill=bad_fill, font=fnt(9, True, C["red"])))
    ws.conditional_formatting.add(f"H{t0}:H{t_tot}", FormulaRule(
        formula=[f'AND(ISNUMBER($H{t0}),$H{t0}<>0)'], fill=fl(C["amber_l"]), font=fnt(9, True, C["amber"])))
    ws.conditional_formatting.add(f"H{t0}:H{t_tot}", FormulaRule(
        formula=[f'AND(ISNUMBER($H{t0}),$H{t0}=0)'], fill=ok_fill, font=fnt(9, True, C["ok"])))
    ws.conditional_formatting.add(f"F{t0}:F{t1}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="A5D6A7", showValue=True))

    # KPI cards rows 6-9 (4 + 4)
    kpi1 = [("جمع مانده اولیه (ریال)", f"=$C${t_tot}"),
            ("جمع واریز کل (ریال)", f"=$D${t_tot}"),
            ("جمع برداشت کل (ریال)", f"=$E${t_tot}"),
            ("جمع مانده نهایی سیستم (ریال)", f"=$F${t_tot}")]
    kpi2 = [("جمع آخرین موجودی بانک‌ها (ریال)", f"=SUM($G${t0}:$G${t1})"),
            ("تعداد کل تراکنش‌ها", f"=$J${t_tot}", "0"),
            ("بانک‌های دارای مغایرت", f'=COUNTIF($I${t0}:$I${t1},"{BAD_MARK}")', "0"),
            ("ردیف‌های تکراری کل", f"=$K${t_tot}", "0")]
    groups = [(1, 3), (4, 6), (7, 9), (10, 11)]
    for (c0, c1), item in zip(groups, kpi1):
        card(ws, 6, 7, c0, c1, item[0], item[1])
    for (c0, c1), item in zip(groups, kpi2):
        card(ws, 8, 9, c0, c1, item[0], item[1], item[2] if len(item) > 2 else "#,##0")
    ws.row_dimensions[6].height = 13; ws.row_dimensions[7].height = 20
    ws.row_dimensions[8].height = 13; ws.row_dimensions[9].height = 20

    # ---- monthly analysis ----
    m_title = t_tot + 2
    merge_put(ws, f"A{m_title}:K{m_title}", "تحلیل ماهانه — سال را در خانه زردرنگ انتخاب/تایپ کنید",
              fnt(11, True, C["green_d"]), None, al("right"))
    yr_row = m_title + 1
    put(ws, f"A{yr_row}", "سال:", fnt(9, True, C["gray"]), fl(C["green_l"]), al("center"), BORDER_ALL)
    put(ws, f"B{yr_row}", "1405", fnt(10, True, C["orange_d"]), fl(C["input"]), al("center"),
        BORDER_BOX_D, "@")
    dv_year = DataValidation(type="list",
                             formula1='"1400,1401,1402,1403,1404,1405,1406,1407,1408,1409,1410"',
                             allow_blank=False)
    dv_year.showErrorMessage = False
    ws.add_data_validation(dv_year); dv_year.add(f"B{yr_row}")
    m_hdr = yr_row + 1
    table_headers(ws, m_hdr, m_hdr + 1,
                  [("ماه", "Month"), ("نام ماه", "Month Name"), ("واریز (ریال)", "Deposits"),
                   ("برداشت (ریال)", "Withdrawals"), ("خالص (ریال)", "Net"), ("تعداد تراکنش", "Txns")])
    m0 = m_hdr + 2
    for i in range(12):
        r = m0 + i
        mm = str(i + 1).zfill(2)
        ws[f"A{r}"] = mm
        ws[f"B{r}"] = MONTHS_FA[i]
        ws[f"C{r}"] = f'=SUMIFS(CalcDep,CalcMonth,$B${yr_row}&"/"&$A{r})'
        ws[f"D{r}"] = f'=SUMIFS(CalcWit,CalcMonth,$B${yr_row}&"/"&$A{r})'
        ws[f"E{r}"] = f'=$C{r}-$D{r}'
        ws[f"F{r}"] = f'=COUNTIFS(CalcMonth,$B${yr_row}&"/"&$A{r})'
    m1 = m0 + 11
    data_area(ws, m0, m1, 1, 6, {}, {3: "#,##0", 4: "#,##0", 5: "#,##0"})
    ws.conditional_formatting.add(f"E{m0}:E{m1}", CellIsRule(
        operator="lessThan", formula=["0"], font=fnt(9, True, C["red"])))
    # monthly totals row
    tr = m1 + 1
    merge_put(ws, f"A{tr}:B{tr}", "جمع سال", fnt(9.5, True, C["white"]), fl(C["green"]),
              al("center"), bd(b=sd("medium", C["orange"])))
    for col in "CDEF":
        ws[f"{col}{tr}"] = f"=SUM({col}{m0}:{col}{m1})"
        ws[f"{col}{tr}"].number_format = "#,##0"
        ws[f"{col}{tr}"].font = fnt(9.5, True, C["green_d"])
        ws[f"{col}{tr}"].fill = fl(C["green_l"]); ws[f"{col}{tr}"].border = BORDER_ALL
        ws[f"{col}{tr}"].alignment = al("center")

    # ---- charts ----
    ch_row = tr + 2
    put(ws, f"A{ch_row}", "نمودارها", fnt(11, True, C["green_d"]), None, al("right"))
    bar = BarChart()
    bar.type = "col"
    bar.title = "مانده نهایی سیستم بانک‌ها (ریال)"
    bar.style = 10
    bar.height = 8.5
    bar.width = 17
    data = Reference(ws, min_col=6, min_row=t_hdr, max_row=t1)   # F header + values
    cats = Reference(ws, min_col=2, min_row=t0, max_row=t1)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.legend = None
    bar.y_axis.title = "ریال"
    ws.add_chart(bar, f"A{ch_row + 1}")

    line = LineChart()
    line.title = "روند ماهانه واریز / برداشت (سال انتخابی)"
    line.style = 12
    line.height = 8.5
    line.width = 17
    ldata = Reference(ws, min_col=3, max_col=4, min_row=m_hdr, max_row=m1)
    lcats = Reference(ws, min_col=2, min_row=m0, max_row=m1)
    line.add_data(ldata, titles_from_data=True)
    line.set_categories(lcats)
    line.y_axis.title = "ریال"
    ws.add_chart(line, f"G{ch_row + 1}")

    set_widths(ws, {"A": 10, "B": 26, "C": 16, "D": 16, "E": 16, "F": 18, "G": 19,
                    "H": 16, "I": 10, "J": 12, "K": 10})
    ws.freeze_panes = "A5"
    setup_print(ws, "landscape", print_area=f"A1:K{m1 + 15}",
                header_center="خلاصه و تحلیل حساب‌های بانکی")
    return ws


def build_help(wb, cap):
    ws = wb.create_sheet("Help")
    ws.sheet_properties.tabColor = C["tab_blue"]
    rtl(ws, 100)
    brand_band(ws, 8, mode="info")
    title_row(ws, 8, 3, "راهنمای فایل تراکنش‌ها (Mazandasht_Transactions.xlsx)")
    sections = [
        ("۱) ساختار فایل", [
            "Entry — ورود داده: همه تراکنش‌ها را همین‌جا، به هر ترتیبی، زیر هم ثبت کنید (مرتب بودن لازم نیست).",
            "Bank1 تا Bank10 — صورت‌حساب هر بانک: به‌صورت خودکار مرتب‌شده به تاریخ و ساعت + مانده سیستم + کنترل مغایرت.",
            "Summary — خلاصه کل حساب‌ها، تحلیل ماهانه و نمودارها.",
            "Sync — آینه فایل مادر (INFO)؛ منبع کشوها و کدها. دست نزنید.",
            "Calc — موتور محاسبات (نرمال‌سازی و رتبه‌بندی). دست نزنید.",
        ]),
        ("۲) رنگ‌ها و نشانه‌ها", [
            "زرد = سلول ورودی شما | خاکستری = محاسباتی خودکار.",
            f"{OK_MARK} = موجودی بانک با محاسبات سیستم منطبق است.",
            f"{BAD_MARK} = مغایرت! ردیف‌های قبل از آن و خودش را با صورت‌حساب بانک چک کنید.",
            f"{NONE_MARK} = موجودی بانک برای آن ردیف ثبت نشده (بدون کنترل).",
            f"{WARN_MARK} = تاریخ/بانک نامعتبر است و سیستم نمی‌تواند مانده بسازد.",
            f"{DUP_MARK} = این ردیف با ردیف دیگری (تاریخ، ساعت، بانک، طرف و مبالغ یکسان) تکراری است.",
        ]),
        ("۳) اتصال به فایل مادر", [
            "هر دو فایل همیشه در یک پوشه باشند (لینک نسبی است).",
            "پس از تغییر INFO، در این فایل: Data ← Edit Links ← Update Values (یا هنگام باز شدن، Update را بزنید).",
            "اگر فایل INFO جابه‌جا شد: Edit Links ← Change Source و مسیر جدید را بدهید.",
            "کد بانک‌ها (BANK01…) ثابت است؛ با تغییر «عنوان نمایشی» در INFO، عنوان برگه‌ها خودکار عوض می‌شود (نام زبانه Bank1.. ثابت می‌ماند).",
        ]),
        ("۴) نکات ثبت", [
            "تاریخ: 1405/07/10 یا 14050710 — ساعت: 09:30 یا 9:30 (ارقام لاتین).",
            "واریزکننده/برداشت‌کننده = کد طرف حساب (C…/S…/X…/E…/SL… فروشنده)؛ نامش خودکار می‌آید.",
            "در هر ردیف فقط یکی از «واریز» یا «برداشت» پر شود.",
            "موجودی بانک (M) را همان‌طور که از پیامک/صورت‌حساب می‌خوانید بنویسید؛ سیستم مقایسه می‌کند.",
            "انتقال بین بانکی = دو ردیف: یک برداشت از بانک مبدا، یک واریز به بانک مقصد (نوع تراکنش: انتقال بین بانکی).",
            "برای حذف ردیف، محتوای سلول‌ها را پاک کنید؛ کل ردیف را حذف نکنید.",
        ]),
        ("۵) چاپ", [
            "همه برگه‌ها A4 و آماده چاپ‌اند؛ سرستون‌ها در هر صفحه تکرار می‌شوند.",
            "برگه‌های بانکی: افقی (Landscape) — سربرگ و پاصفحه با نام شرکت و شماره صفحه.",
        ]),
        ("۶) پیوند با فایل فروش", [
            "واریز/تسویه هر فروشنده را در این فایل با کد SLxxx ثبت کنید؛ صورت‌حساب او در شیت SLxxx فایل فروش به‌صورت خودکار می‌آید.",
        ]),
        ("۷) گسترش در آینده", [
            "ظرفیت Entry در این نسخه: %d ردیف و هر برگه بانک: %d ردیف. برای بیشتر، آخرین ردیف را انتخاب و تا پایین کپی کنید." % (cap.entry_rows, cap.bank_rows),
            "فایل‌های آینده (فروش هر فروشنده، انبار، خرید و…) به همین سبک روی کدهای INFO ساخته می‌شوند.",
        ]),
    ]
    r = 5
    for title, lines in sections:
        merge_put(ws, f"A{r}:H{r}", title, fnt(10.5, True, C["white"]), fl(C["green"]),
                  al("right"), bd(b=sd("medium", C["orange"])))
        ws.row_dimensions[r].height = 19
        r += 1
        for i, ln in enumerate(lines):
            merge_put(ws, f"A{r}:H{r}", "• " + ln, fnt(9),
                      fl(C["white"] if i % 2 == 0 else C["zebra"]), al("right", wrap=True), BORDER_ALL)
            ws.row_dimensions[r].height = 17
            r += 1
        r += 1
    set_widths(ws, {"A": 12, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12})
    setup_print(ws, "portrait", header_center="راهنما")
    return ws


def build_transactions(cap, out_path, external=True, info_path=None):
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.title = "Mazandasht Transactions — تراکنش‌های بانکی"
    wb.properties.creator = BRAND_NAME
    wb.properties.company = "Mazandasht"
    wb.calculation.fullCalcOnLoad = True

    # internal QA mode: embed INFO registers so links are local
    if not external:
        for name in INFO_SHEET_ORDER[1:]:
            wb.create_sheet(name).sheet_properties.tabColor = C["tab_gray"]
        build_info_sheets_into(wb, cap)

    build_entry(wb, cap)
    for i in range(1, cap.banks + 1):
        build_bank_sheet(wb, i, cap)
    build_summary(wb, cap)
    build_sync(wb, cap, external)
    build_calc(wb, cap)
    build_help(wb, cap)

    order = (["Entry"] + [f"Bank{i}" for i in range(1, cap.banks + 1)] +
             ["Summary", "Sync", "Calc", "Help"])
    if not external:
        order = (["Entry"] + [f"Bank{i}" for i in range(1, cap.banks + 1)] +
                 ["Summary", "Sync", "Calc", "Help"] + INFO_SHEET_ORDER[1:])
    wb._sheets = sorted(wb._sheets, key=lambda s: order.index(s.title))
    wb.active = 0

    # ---- defined names ----
    n0e, n1e = E_DATA0, e_data1(cap)
    n0b, n1b = B_DATA0, b_data1(cap)
    y0 = SY_DATA0
    names = {
        "BankTbl": f"Sync!$A${y0}:$D${y0 + cap.banks - 1}",
        "CustTbl": f"Sync!$F${y0}:$G${y0 + cap.customers - 1}",
        "SuppTbl": f"Sync!$I${y0}:$J${y0 + cap.suppliers - 1}",
        "AcctTbl": f"Sync!$L${y0}:$M${y0 + cap.accounts - 1}",
        "StaffTbl": f"Sync!$O${y0}:$P${y0 + cap.staff - 1}",
        "SellerTbl": f"Sync!$BA${y0}:$BB${y0 + cap.sellers - 1}",
        "SellerCodes": f"Sync!$BA${y0}:$BA${y0 + cap.sellers - 1}",
        "CustCodes": f"Sync!$F${y0}:$F${y0 + cap.customers - 1}",
        "SuppCodes": f"Sync!$I${y0}:$I${y0 + cap.suppliers - 1}",
        "AcctCodes": f"Sync!$L${y0}:$L${y0 + cap.accounts - 1}",
        "StaffCodes": f"Sync!$O${y0}:$O${y0 + cap.staff - 1}",
        "ListBanks": (f"OFFSET(Sync!$AT${y0},0,0,"
                      f"MAX(COUNTIF(Sync!$AT${y0}:$AT${y0 + cap.banks - 1},\"?*\"),1),1)"),
        "ListPartyTypes": (f"OFFSET(Sync!$AO${y0},0,0,"
                           f"MAX(COUNTIF(Sync!$AO${y0}:$AO${y0 + cap.party_types - 1},\"?*\"),1),1)"),
        "ListTxnTypes": (f"OFFSET(Sync!$AQ${y0},0,0,"
                         f"MAX(COUNTIF(Sync!$AQ${y0}:$AQ${y0 + cap.txn_types - 1},\"?*\"),1),1)"),
        "CalcBank": f"Calc!$G${n0e}:$G${n1e}",
        "CalcKey": f"Calc!$H${n0e}:$H${n1e}",
        "CalcRankKey": f"Calc!$J${n0e}:$J${n1e}",
        "CalcNormDate": f"Calc!$E${n0e}:$E${n1e}",
        "CalcTime": f"Calc!$F${n0e}:$F${n1e}",
        "CalcMonth": f"Calc!$K${n0e}:$K${n1e}",
        "CalcDep": f"Calc!$L${n0e}:$L${n1e}",
        "CalcWit": f"Calc!$M${n0e}:$M${n1e}",
        "EntryDates": f"Entry!$C${n0e}:$C${n1e}",
        "EntryTimes": f"Entry!$D${n0e}:$D${n1e}",
        "EntryBanks": f"Entry!$E${n0e}:$E${n1e}",
        "EntryParty": f"Entry!$G${n0e}:$G${n1e}",
        "EntryPartyName": f"Entry!$H${n0e}:$H${n1e}",
        "EntryPartyType": f"Entry!$I${n0e}:$I${n1e}",
        "EntryTxnType": f"Entry!$J${n0e}:$J${n1e}",
        "EntryDep": f"Entry!$K${n0e}:$K${n1e}",
        "EntryWit": f"Entry!$L${n0e}:$L${n1e}",
        "EntryMan": f"Entry!$M${n0e}:$M${n1e}",
        "EntryTxnID": f"Entry!$B${n0e}:$B${n1e}",
        "EntryNotes": f"Entry!$P${n0e}:$P${n1e}",
    }
    for n, ref in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)

    wb.save(out_path)
    return out_path


def build_info_sheets_into(wb, cap):
    """Populate pre-created INFO sheets inside the workbook (QA internal mode)."""
    for name in INFO_SHEET_ORDER[1:]:
        del wb[name]
    # Lists
    ws = wb.create_sheet("Lists")
    rtl(ws)
    lists_headers = [("واحد شمارش", "Units"), ("سمت‌ها", "Roles"),
                     ("دسته محصولات", "Categories"), ("بله/خیر", "Yes/No")]
    table_headers(ws, 5, 6, lists_headers)
    lists_data = [SEED_LISTS_UNITS, SEED_LISTS_ROLES, SEED_LISTS_CATS, SEED_LISTS_YESNO]
    for ci, values in enumerate(lists_data):
        for ri, v in enumerate(values):
            ws.cell(row=S_DATA0 + ri, column=ci + 1, value=v)
    # Company
    ws = wb.create_sheet("Company")
    rtl(ws)
    for i, (label, value) in enumerate(SEED_COMPANY):
        row = 5 + i
        ws[f"B{row}"] = label
        ws.merge_cells(f"C{row}:G{row}")
        ws[f"C{row}"] = value
    # registers
    dv_map = {
        "Banks": {"D": "ListYesNo"}, "Customers": {"B": "ListPT"},
        "Sellers": {"C": "ListPT"},
        "Suppliers": {"B": "ListPT"}, "Accounts": {"B": "ListPT"},
        "Staff": {"C": "ListRoles"},
        "Products_Buy": {"C": "ListCategories", "D": "ListUnits"},
        "Products_Sell": {"C": "ListCategories", "D": "ListUnits", "E": "ListBaskets"},
        "Warehouses": {"D": "ListRoles"},
    }
    for name in ["Banks", "Customers", "Sellers", "Suppliers", "Accounts", "Staff",
                 "PartyTypes", "TxnTypes", "Products_Buy", "Products_Sell",
                 "Baskets", "Warehouses"]:
        build_register_sheet(wb, name, cap, None)
    info_names = {
        "ListUnits": f"Lists!$A${S_DATA0}:$A${S_DATA0 + 25}",
        "ListRoles": f"Lists!$B${S_DATA0}:$B${S_DATA0 + 25}",
        "ListCategories": f"Lists!$C${S_DATA0}:$C${S_DATA0 + 25}",
        "ListYesNo": f"Lists!$D${S_DATA0}:$D${S_DATA0 + 1}",
        "ListPT": f"PartyTypes!$B${S_DATA0}:$B${S_DATA0 + cap.party_types - 1}",
        "ListBaskets": (f"OFFSET(Baskets!$E${S_DATA0},0,0,"
                        f"MAX(COUNTIF(Baskets!$E${S_DATA0}:$E${S_DATA0 + cap.baskets - 1},\"?*\"),1),1)"),
    }
    for n, ref in info_names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
    # names will be re-added by build_transactions; register here temporarily


# ============================================================================
# 5b. SALES WORKBOOK (Mazandasht_Sales.xlsx)
# ============================================================================

SALES_FILE = "Mazandasht_Sales.xlsx"
SALES_SLOTS = 15                     # product slots per invoice row
SALES_DATA0 = 10                     # first Entry data row
SLOT_W = 5                           # cols per slot: product|basket|citrus|count|weight


def s_data1(cap):
    return SALES_DATA0 + cap.entry_rows - 1


def slot_col(i, off):
    """column index of slot i (1..15), offset 0..4"""
    return 12 + SLOT_W * (i - 1) + 1 + off


def parse_code_expr(ref):
    """classic formula: extract code from 'SL001 – name' or bare 'SL001'"""
    return (f'IF(ISNUMBER(FIND(" – ",{ref})),LEFT({ref},FIND(" – ",{ref})-1),{ref})')


def build_sales_sync(wb, cap, external):
    n0 = SY_DATA0
    ws = wb.create_sheet("Sync")
    ws.sheet_properties.tabColor = C["tab_gray"]
    ws.sheet_view.rightToLeft = False
    ws.sheet_view.showGridLines = False
    brand_band(ws, 32, mode="info")
    merge_put(ws, "A3:AF3",
              "برگه فنی — آینه‌ی فایل مادر (INFO) و فایل واریز/برداشت (Transactions) برای فایل فروش. فرمول‌ها را تغییر ندهید. "
              "اگر همه سلول‌ها خالی شدند یعنی لینک به‌روز نشده: Data ← Edit Links ← Update Values.",
              fnt(8, True, C["red"]), fl(C["amber_l"]), al("left", wrap=True))

    def mirror(col, sheet, info_col, rows, num=False):
        for i in range(rows):
            r = n0 + i
            ir = S_DATA0 + i
            ref = sync_ref(sheet, f"${info_col}${ir}", external)
            if num:
                ws[f"{col}{r}"] = f"={ref}"
            else:
                ws[f"{col}{r}"] = f'=IF({ref}="","",{ref})'

    def block_header(col, title):
        put(ws, f"{col}5", title, fnt(8, True, C["white"]), fl(C["green"]),
            al("center", wrap=True), BORDER_ALL)

    # Sellers (A:C) — code, name, type
    for cl, t in zip("ABC", ["Sellers · Code", "Sellers · Name", "Sellers · Type"]):
        block_header(cl, t)
    mirror("A", "Sellers", "A", cap.sellers)
    mirror("B", "Sellers", "B", cap.sellers)
    mirror("C", "Sellers", "C", cap.sellers)
    # Products_Sell (E:I) — code, name, category, unit, basket
    for cl, t in zip("EFGHI", ["Products_Sell · Code", "· Name", "· Category",
                               "· Unit", "· Basket"]):
        block_header(cl, t)
    mirror("E", "Products_Sell", "A", cap.products_sell)
    mirror("F", "Products_Sell", "B", cap.products_sell)
    mirror("G", "Products_Sell", "C", cap.products_sell)
    mirror("H", "Products_Sell", "D", cap.products_sell)
    mirror("I", "Products_Sell", "E", cap.products_sell)
    # Baskets (K:L) — code, name
    block_header("K", "Baskets · Code"); block_header("L", "Baskets · Name")
    mirror("K", "Baskets", "A", cap.baskets)
    mirror("L", "Baskets", "B", cap.baskets)
    # Company (N)
    block_header("N", "Company")
    comp_cells = [("$C$5", "name"), ("$C$8", "address"), ("$C$9", "phone"), ("$C$7", "activity")]
    for i, (ref, _nm) in enumerate(comp_cells):
        ext = sync_ref("Company", ref, external)
        ws[f"N{n0 + i}"] = f'=IF({ext}="","",{ext})'

    # packed seller list: cum P, code Q, label R
    block_header("P", "cum·S"); block_header("Q", "LIST · Seller codes"); block_header("R", "LIST · Sellers")
    for i in range(cap.sellers):
        r = n0 + i
        ws[f"P{r}"] = f'=COUNTIF($A${n0}:$A{r},"?*")'
        m = f'MATCH(ROW()-{n0 - 1},$P${n0}:$P${n0 + cap.sellers - 1},0)'
        ws[f"Q{r}"] = f'=IFERROR(INDEX($A${n0}:$A${n0 + cap.sellers - 1},{m}),"")'
        ws[f"R{r}"] = (
            f'=IF($Q{r}="","",$Q{r}&" – "&IF(IFERROR(VLOOKUP($Q{r},$A${n0}:$B${n0 + cap.sellers - 1},2,0),"")="",'
            f'"(بدون عنوان)",VLOOKUP($Q{r},$A${n0}:$B${n0 + cap.sellers - 1},2,0)))'
        )
    # packed product list: cum T, code U, label V
    block_header("T", "cum·P"); block_header("U", "LIST · Product codes"); block_header("V", "LIST · Products")
    for i in range(cap.products_sell):
        r = n0 + i
        ws[f"T{r}"] = f'=COUNTIF($E${n0}:$E{r},"?*")'
        m = f'MATCH(ROW()-{n0 - 1},$T${n0}:$T${n0 + cap.products_sell - 1},0)'
        ws[f"U{r}"] = f'=IFERROR(INDEX($E${n0}:$E${n0 + cap.products_sell - 1},{m}),"")'
        ws[f"V{r}"] = (
            f'=IF($U{r}="","",$U{r}&" – "&IF(IFERROR(VLOOKUP($U{r},$E${n0}:$F${n0 + cap.products_sell - 1},2,0),"")="",'
            f'"(بدون عنوان)",VLOOKUP($U{r},$E${n0}:$F${n0 + cap.products_sell - 1},2,0)))'
        )

    # --- extended seller info (BA:BC): phone, mobile, address ---
    for cl, t, info_col in zip(("BA", "BB", "BC"),
                               ("Sellers · Phone", "· Mobile", "· Address"),
                               ("D", "E", "F")):
        block_header(cl, t)
        mirror(cl, "Sellers", info_col, cap.sellers)
    # --- bank Entry mirror (X..AF) — payments source, link 2 ---
    bank_sheet = "Entry" if external else "BankEntry"
    bank_cols = {"X": "C", "Y": "D", "Z": "E", "AA": "G", "AB": "I",
                 "AC": "J", "AD": "K", "AE": "L", "AF": "P"}
    for cl, src in bank_cols.items():
        block_header(cl, f"Bank {bank_sheet} · {src}")
    for i in range(cap.entry_rows):
        r = n0 + i
        br = E_DATA0 + i
        for cl, src in bank_cols.items():
            ref = sync_ref(bank_sheet, f"${src}${br}", external, link=2)
            ws[f"{cl}{r}"] = f'=IF({ref}="","",{ref})'

    set_widths(ws, {"A": 9, "B": 28, "C": 20, "E": 9, "F": 26, "G": 14, "H": 10,
                    "I": 12, "K": 9, "L": 20, "N": 30, "P": 7, "Q": 22, "R": 36,
                    "T": 7, "U": 22, "V": 36, "X": 12, "Y": 8, "Z": 28, "AA": 10,
                    "AB": 18, "AC": 14, "AD": 14, "AE": 14, "AF": 30,
                    "BA": 14, "BB": 14, "BC": 30})
    ws.freeze_panes = "A6"
    return ws


def build_sales_calc(wb, cap):
    n0, n1 = SALES_DATA0, s_data1(cap)
    ws = wb.create_sheet("Calc")
    ws.sheet_properties.tabColor = C["tab_gray"]
    ws.sheet_view.rightToLeft = False
    ws.sheet_view.showGridLines = False
    brand_band(ws, 13, mode="info")
    merge_put(ws, "A3:M3",
              "برگه فنی — موتور محاسبات فایل فروش (نرمال‌سازی تاریخ، کد فروشنده و آمار ردیف). فرمول‌ها را تغییر ندهید.",
              fnt(8, True, C["red"]), fl(C["amber_l"]), al("left"))
    hdrs = ["تاریخ خام · Raw", "سال · Y", "ماه · M", "روز · D", "تاریخ استاندارد · Norm",
            "سال/ماه · YM", "کد فروشنده · Seller", "شماره فاکتور · Invoice",
            "وزن ارسالی · Sent", "وزن فروش · Sold", "مبلغ صافی · Amount",
            "جمع وزن محصولات · ProdWt", "تعداد اقلام · Items"]
    for i, h in enumerate(hdrs):
        col = get_column_letter(i + 1)
        put(ws, f"{col}8", h, fnt(7.5, True, C["white"]), fl(C["gray"]),
            al("center", wrap=True), BORDER_ALL)
    ws.row_dimensions[8].height = 24

    for r in range(n0, n1 + 1):
        ws[f"A{r}"] = f'=IF(Entry!$B{r}="","",Entry!$B{r})'
        ws[f"B{r}"] = (
            f'=IF($A{r}="","",IF(ISNUMBER($A{r}),'
            f'IF(AND($A{r}>=10000101,$A{r}<=99991231),INT($A{r}/10000),""),'
            f'IFERROR(VALUE(TRIM(LEFT($A{r},FIND("/",$A{r})-1))),"")))'
        )
        ws[f"C{r}"] = (
            f'=IF(OR($A{r}="",$B{r}=""),"",IF(ISNUMBER($A{r}),'
            f'INT(MOD($A{r},10000)/100),'
            f'IFERROR(VALUE(TRIM(LEFT(MID($A{r},FIND("/",$A{r})+1,20),'
            f'FIND("/",MID($A{r},FIND("/",$A{r})+1,20))-1))),"")))'
        )
        ws[f"D{r}"] = (
            f'=IF(OR($A{r}="",$B{r}=""),"",IF(ISNUMBER($A{r}),'
            f'MOD($A{r},100),'
            f'IFERROR(VALUE(TRIM(MID(MID($A{r},FIND("/",$A{r})+1,20),'
            f'FIND("/",MID($A{r},FIND("/",$A{r})+1,20))+1,20))),"")))'
        )
        ws[f"E{r}"] = (
            f'=IF($A{r}="","",IF(OR($B{r}="",$C{r}="",$D{r}="",$B{r}<1300,$B{r}>1500,'
            f'$C{r}<1,$C{r}>12,$D{r}<1,$D{r}>31),"نامعتبر",'
            f'RIGHT("0000"&$B{r},4)&"/"&RIGHT("00"&$C{r},2)&"/"&RIGHT("00"&$D{r},2)))'
        )
        ws[f"F{r}"] = f'=IF(OR($E{r}="",$E{r}="نامعتبر"),"",LEFT($E{r},4)&"/"&MID($E{r},6,2))'
        ws[f"G{r}"] = (
            f'=IF(OR($E{r}="",$E{r}="نامعتبر",Entry!$D{r}=""),"",'
            f'{parse_code_expr("Entry!$D" + str(r))})'
        )
        ws[f"H{r}"] = f'=IF(Entry!$C{r}="","",Entry!$C{r})'
        ws[f"I{r}"] = f'=IF(Entry!$F{r}="",0,Entry!$F{r})'
        ws[f"J{r}"] = f'=IF(Entry!$G{r}="",0,Entry!$G{r})'
        ws[f"K{r}"] = f'=IF(Entry!$H{r}="",0,Entry!$H{r})'
        ws[f"L{r}"] = f'=IF(Entry!$K{r}="",0,Entry!$K{r})'
        cnt_args = ",".join(f'Entry!${get_column_letter(slot_col(i, 0))}{r}'
                            for i in range(1, SALES_SLOTS + 1))
        ws[f"M{r}"] = f'=IF($G{r}="","",COUNTA({cnt_args}))'

    set_widths(ws, {"A": 12, "B": 6, "C": 6, "D": 6, "E": 12, "F": 9, "G": 10,
                    "H": 11, "I": 13, "J": 13, "K": 15, "L": 14, "M": 9})
    ws.freeze_panes = "A9"
    return ws


# seed rows: (date, invoice, seller, sent, sold, amount, [(product, count, weight), ...])
SEED_SALES = [
    ("1405/07/05", 1001, "SL001 – حاج رضا محمدی (نمونه — حجره‌دار تهران)", 12000, 11950,
     950000000, [("PS01 – پرتقال درجه ۱ — بسته‌بندی", 600, 6000),
                 ("PS02 – پرتقال درجه ۲ — عمومی", 595, 5950)]),
    ("1405/07/06", 1002, "SL002 – شرکت گلسرای شمال (نمونه — صادراتی)", 20000, 19500,
     1600000000, [("PS03 – نارنگی درجه ۱ — بسته‌بندی", 1000, 10000),
                  ("PS04 – نارنگی عمومی", 950, 9500)]),
    ("1405/07/08", 1003, "SL003 – عباس نیک‌پی (نمونه — واسطه)", 8000, 8000,
     640000000, [("PS02 – پرتقال درجه ۲ — عمومی", 800, 8000)]),
    ("1405/07/10", 1004, "SL004", 5000, 4990,
     400000000, [("PS01", 499, 4990)]),
]


def build_sales_entry(wb, cap):
    n0, n1 = SALES_DATA0, s_data1(cap)
    ws = wb.create_sheet("Entry")
    ws.sheet_properties.tabColor = C["tab_orange"]
    rtl(ws, 80)
    ncols = 12 + SALES_SLOTS * SLOT_W
    last = get_column_letter(ncols)
    print_last = "H"
    brand_band(ws, ncols, mode="txn", comp_col="N", print_ncols=8)
    merge_put(ws, f"A3:H3", "فروش محصولات به فروشندگان — ورود داده (Entry)",
              fnt(13, True, C["orange_d"]), fl(C["green_xl"]), al("center"),
              bd(b=sd("medium", C["orange"])))
    merge_put(ws, f"I3:{last}3", "۱۵ ردیف محصول — ناحیه فقط-اکسل (خارج از چاپ)",
              fnt(9, True, C["green"]), fl(C["green_xl"]), al("center"),
              bd(b=sd("medium", C["green"])))
    ws.row_dimensions[3].height = 24
    merge_put(ws, f"A4:H4",
              "ستون‌های زرد = ورود دستی شما | ستون‌های خاکستری = خودکار.  تاریخ: 1405/07/10 (یا 14050710) — "
              "کد فروشنده را تایپ کنید (مثل SL001) یا از کشویی انتخاب کنید؛ نامش خودکار می‌آید.",
              fnt(8, False, C["gray"], italic=True), fl(C["white"]), al("right", wrap=True))
    merge_put(ws, f"I4:{last}4",
              "برای هر فاکتور تا ۱۵ نوع محصول: محصول را از کشویی انتخاب کنید تا نوع سبد و نوع مرکبات خودکار نوشته شود، "
              "سپس تعداد سبد و وزن را وارد کنید. این ستون‌ها با دکمه +/─ بالای ستون‌ها جمع می‌شوند.",
              fnt(8, False, C["gray"], italic=True), fl(C["white"]), al("right", wrap=True))
    ws.row_dimensions[4].height = 24

    # KPI cards (rows 5-6): 4 printed + 2 excel-only
    kpis = [
        (1, 2, "تعداد فاکتور", f"=COUNTA($B${n0}:$B${n1})", "0"),
        (3, 4, "جمع وزن ارسالی (kg)", f"=SUM($F${n0}:$F${n1})", "#,##0.0"),
        (5, 6, "جمع وزن فروش (kg)", f"=SUM($G${n0}:$G${n1})", "#,##0.0"),
        (7, 8, "جمع مبلغ صافی (ریال)", f"=SUM($H${n0}:$H${n1})", "#,##0"),
        (9, 10, f"تعداد {BAD_MARK} اختلاف وزن", f'=COUNTIF($J${n0}:$J${n1},"{BAD_MARK}")', "0"),
        (11, 12, "ظرفیت باقی‌مانده", f"={cap.entry_rows}-COUNTA($B${n0}:$B${n1})", "0"),
    ]
    for c0, c1, lbl, formula, fmt in kpis:
        card(ws, 5, 6, c0, c1, lbl, formula, fmt)
    ws.row_dimensions[5].height = 13
    ws.row_dimensions[6].height = 20

    # band row 7
    merge_put(ws, "A7:H7", "مشخصات فاکتور — محدوده چاپ",
              fnt(9, True, C["green_d"]), fl(C["green_l"]), al("center"),
              bd(b=sd("medium", C["green"]), t=sd("medium", C["green"])))
    merge_put(ws, "I7:L7", "کنترل وزن — فقط اکسل (چاپ نمی‌شود)",
              fnt(9, True, C["amber"]), fl(C["amber_l"]), al("center"),
              bd(b=sd("medium", C["amber"]), t=sd("medium", C["amber"])))
    for i in range(1, SALES_SLOTS + 1):
        c0 = slot_col(i, 0)
        c1 = c0 + SLOT_W - 1
        fill = fl(C["white"] if i % 2 else C["zebra"])
        merge_put(ws, f"{get_column_letter(c0)}7:{get_column_letter(c1)}7",
                  f"ردیف محصول {i}", fnt(8, True, C["olive"] if "olive" in C else C["green_d"]),
                  fill, al("center"), BORDER_ALL)
    ws.row_dimensions[7].height = 16

    # headers rows 8-9
    headers = [
        ("ردیف", "No."), ("تاریخ", "Date"), ("شماره فاکتور", "Invoice No."),
        ("کد فروشنده", "Seller Code"), ("نام فروشنده", "Seller"),
        ("وزن ارسالی (kg)", "Sent Wt"), ("وزن فروش (kg)", "Sold Wt"),
        ("مبلغ صافی (ریال)", "Net Amount"),
        ("اختلاف وزن", "Wt Diff"), ("وضعیت ۱٪", "1% Status"),
        ("جمع وزن محصولات", "Products Wt"), ("اختلاف با ارسالی", "Sent−Prod"),
    ]
    for i in range(1, SALES_SLOTS + 1):
        headers += [
            (f"محصول {i}", f"Product {i}"), ("نوع سبد", "Basket"),
            ("نوع مرکبات", "Citrus"), ("تعداد سبد", "Count"), ("وزن (kg)", "Weight"),
        ]
    table_headers(ws, 8, 9, headers)

    # data area styling
    input_cols = {2, 3, 4, 6, 7, 8}
    auto_cols = {1, 5, 9, 10, 11, 12}
    fmts = {6: "#,##0.0", 7: "#,##0.0", 8: "#,##0", 9: "#,##0.0", 11: "#,##0.0", 12: "#,##0.0"}
    for i in range(1, SALES_SLOTS + 1):
        input_cols |= {slot_col(i, 0), slot_col(i, 3), slot_col(i, 4)}
        auto_cols |= {slot_col(i, 1), slot_col(i, 2)}
        fmts[slot_col(i, 3)] = "#,##0"
        fmts[slot_col(i, 4)] = "#,##0.0"
    kinds = {c: ("in" if c in input_cols else "auto") for c in range(1, ncols + 1)}
    data_area(ws, n0, n1, 1, ncols, kinds, fmts)
    # text formats
    for col in ("B", "D"):
        for r in range(n0, n1 + 1):
            ws[f"{col}{r}"].number_format = "@"

    # ---- formulas ----
    weight_refs_tpl = ",".join(f"${get_column_letter(slot_col(i, 4))}{{r}}" for i in range(1, SALES_SLOTS + 1))
    for r in range(n0, n1 + 1):
        ws[f"A{r}"] = f'=IF($B{r}="","",COUNTA($B${n0}:$B{r}))'
        code = parse_code_expr(f"$D{r}")
        ws[f"E{r}"] = (f'=IF($D{r}="","",IFERROR(VLOOKUP({code},SellerTbl,2,0),"کد ناشناخته"))')
        ws[f"I{r}"] = f'=IF(OR($F{r}="",$G{r}=""),"",$F{r}-$G{r})'
        ws[f"J{r}"] = f'=IF($I{r}="","",IF(ABS($I{r})<=0.01*$F{r},"{OK_MARK}","{BAD_MARK}"))'
        wrefs = weight_refs_tpl.format(r=r)
        ws[f"K{r}"] = f'=IF(COUNT({wrefs})=0,"",SUM({wrefs}))'
        ws[f"L{r}"] = f'=IF(OR($F{r}="",$K{r}=""),"",$F{r}-$K{r})'
        for i in range(1, SALES_SLOTS + 1):
            pc = f"${get_column_letter(slot_col(i, 0))}{r}"
            pcode = parse_code_expr(pc)
            ws.cell(row=r, column=slot_col(i, 1)).value = (
                f'=IF({pc}="","",IFERROR(VLOOKUP(IFERROR(VLOOKUP({pcode},ProdTbl,5,0),""),'
                f'BasketTbl,2,0),"کد ناشناخته"))'
            )
            ws.cell(row=r, column=slot_col(i, 2)).value = (
                f'=IF({pc}="","",IFERROR(VLOOKUP({pcode},ProdTbl,3,0),"کد ناشناخته"))'
            )

    # ---- seed rows ----
    for ri, (date, inv, seller, sent, sold, amount, prods) in enumerate(SEED_SALES):
        r = n0 + ri
        ws[f"B{r}"] = date
        ws[f"C{r}"] = inv
        ws[f"D{r}"] = seller
        ws[f"F{r}"] = sent
        ws[f"G{r}"] = sold
        ws[f"H{r}"] = amount
        for pi, (prod, cnt, wt) in enumerate(prods):
            ws.cell(row=r, column=slot_col(pi + 1, 0)).value = prod
            ws.cell(row=r, column=slot_col(pi + 1, 3)).value = cnt
            ws.cell(row=r, column=slot_col(pi + 1, 4)).value = wt

    # ---- widths & grouping ----
    widths = {"A": 5.5, "B": 11, "C": 11, "D": 38, "E": 30, "F": 12, "G": 12,
              "H": 16, "I": 12, "J": 9, "K": 13, "L": 13}
    for i in range(1, SALES_SLOTS + 1):
        widths[get_column_letter(slot_col(i, 0))] = 30
        widths[get_column_letter(slot_col(i, 1))] = 16
        widths[get_column_letter(slot_col(i, 2))] = 12
        widths[get_column_letter(slot_col(i, 3))] = 10
        widths[get_column_letter(slot_col(i, 4))] = 12
    set_widths(ws, widths)
    for c in range(slot_col(1, 0), ncols + 1):
        ws.column_dimensions[get_column_letter(c)].outlineLevel = 1

    # ---- data validation ----
    dvs = []
    dv_date = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'AND(LEN($B{n0})=10,MID($B{n0},5,1)="/",MID($B{n0},8,1)="/",'
                 f'ISNUMBER(VALUE(LEFT($B{n0},4))),ISNUMBER(VALUE(MID($B{n0},6,2))),'
                 f'ISNUMBER(VALUE(MID($B{n0},9,2))))',
        errorTitle="قالب تاریخ", error="تاریخ شمسی با قالب 1405/07/10 یا 14050710 وارد کنید (ارقام لاتین).")
    dv_date.add(f"B{n0}:B{n1}")
    dv_seller = DataValidation(
        type="custom", allow_blank=True, errorStyle="stop",
        formula1=f'OR($D{n0}="",COUNTIF(SellerCodes,$D{n0})>0,COUNTIF(SellerList,$D{n0})>0)',
        errorTitle="کد فروشنده نامعتبر",
        error="کد فروشنده در فایل INFO (شیت Sellers) یافت نشد. از کشویی انتخاب کنید یا کد صحیح (مثل SL001) را وارد کنید.",
        promptTitle="کد فروشنده",
        prompt="کد را تایپ کنید (SL001…) یا از کشویی انتخاب کنید؛ نام فروشنده خودکار نمایش داده می‌شود.")
    dv_seller.showInputMessage = True
    dv_seller.add(f"D{n0}:D{n1}")
    dv_sent = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR($F{n0}="",AND(ISNUMBER($F{n0}),$F{n0}>=0))',
        errorTitle="وزن ارسالی", error="وزن باید عدد بزرگ‌تر یا مساوی صفر باشد.")
    dv_sent.add(f"F{n0}:F{n1}")
    dv_sold = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR($G{n0}="",AND(ISNUMBER($G{n0}),$G{n0}>=0))',
        errorTitle="وزن فروش", error="وزن باید عدد بزرگ‌تر یا مساوی صفر باشد.")
    dv_sold.add(f"G{n0}:G{n1}")
    dv_amount = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR($H{n0}="",AND(ISNUMBER($H{n0}),$H{n0}>=0))',
        errorTitle="مبلغ صافی", error="مبلغ باید عدد بزرگ‌تر یا مساوی صفر باشد (ریال).")
    dv_amount.add(f"H{n0}:H{n1}")
    dv_prod = DataValidation(
        type="custom", allow_blank=True, errorStyle="stop",
        formula1=f'OR({get_column_letter(slot_col(1, 0))}{n0}="",'
                 f'COUNTIF(ProdCodes,{get_column_letter(slot_col(1, 0))}{n0})>0,'
                 f'COUNTIF(ProdList,{get_column_letter(slot_col(1, 0))}{n0})>0)',
        errorTitle="کد محصول نامعتبر",
        error="این محصول در فایل INFO (شیت Products_Sell) تعریف نشده است. از کشویی انتخاب کنید.",
        promptTitle="محصول",
        prompt="محصول فروش را از کشویی انتخاب کنید (یا کد آن مثل PS01 را تایپ کنید)؛ نوع سبد و نوع مرکبات خودکار نوشته می‌شود.")
    dv_prod.showInputMessage = True
    for i in range(1, SALES_SLOTS + 1):
        dv_prod.add(f"{get_column_letter(slot_col(i, 0))}{n0}:{get_column_letter(slot_col(i, 0))}{n1}")
    dv_cnt = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR({get_column_letter(slot_col(1, 3))}{n0}="",'
                 f'AND(ISNUMBER({get_column_letter(slot_col(1, 3))}{n0}),'
                 f'{get_column_letter(slot_col(1, 3))}{n0}>=0,'
                 f'INT({get_column_letter(slot_col(1, 3))}{n0})={get_column_letter(slot_col(1, 3))}{n0}))',
        errorTitle="تعداد سبد", error="تعداد سبد باید عدد صحیح بزرگ‌تر یا مساوی صفر باشد.")
    for i in range(1, SALES_SLOTS + 1):
        dv_cnt.add(f"{get_column_letter(slot_col(i, 3))}{n0}:{get_column_letter(slot_col(i, 3))}{n1}")
    dv_wt = DataValidation(
        type="custom", allow_blank=True, errorStyle="warning",
        formula1=f'OR({get_column_letter(slot_col(1, 4))}{n0}="",'
                 f'AND(ISNUMBER({get_column_letter(slot_col(1, 4))}{n0}),'
                 f'{get_column_letter(slot_col(1, 4))}{n0}>=0))',
        errorTitle="وزن", error="وزن باید عدد بزرگ‌تر یا مساوی صفر باشد.")
    for i in range(1, SALES_SLOTS + 1):
        dv_wt.add(f"{get_column_letter(slot_col(i, 4))}{n0}:{get_column_letter(slot_col(i, 4))}{n1}")
    dvs += [dv_date, dv_seller, dv_sent, dv_sold, dv_amount, dv_prod, dv_cnt, dv_wt]
    for dv in dvs:
        dv.showErrorMessage = True
        ws.add_data_validation(dv)

    # ---- conditional formatting ----
    ok_fill = fl(C["ok_l"]); ok_font = fnt(10, True, C["ok"])
    bad_fill = fl(C["red_l"]); bad_font = fnt(10, True, C["red"])
    warn_fill = fl(C["amber_l"]); warn_font = fnt(9, True, C["amber"])
    j_rng = f"J{n0}:J{n1}"
    ws.conditional_formatting.add(j_rng, CellIsRule(operator="equal", formula=[f'"{OK_MARK}"'], fill=ok_fill, font=ok_font))
    ws.conditional_formatting.add(j_rng, CellIsRule(operator="equal", formula=[f'"{BAD_MARK}"'], fill=bad_fill, font=bad_font))
    ws.conditional_formatting.add(f"I{n0}:I{n1}", FormulaRule(
        formula=[f'AND($I{n0}<>"",ABS($I{n0})>0.01*$F{n0})'], font=fnt(9, True, C["red"])))
    ws.conditional_formatting.add(f"E{n0}:E{n1}", FormulaRule(
        formula=[f'ISNUMBER(SEARCH("ناشناخته",$E{n0}))'], fill=warn_fill, font=warn_font))
    # unknown product marks on basket/citrus columns (multi-range)
    auto_rngs = " ".join(
        f"{get_column_letter(slot_col(i, o))}{n0}:{get_column_letter(slot_col(i, o))}{n1}"
        for i in range(1, SALES_SLOTS + 1) for o in (1, 2))
    ws.conditional_formatting.add(auto_rngs, FormulaRule(
        formula=[f'ISNUMBER(SEARCH("ناشناخته",{get_column_letter(slot_col(1, 1))}{n0}))'],
        fill=warn_fill, font=warn_font))
    ws.conditional_formatting.add(f"H{n0}:H{n1}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color="FFB74D", showValue=True))

    # ---- header comments ----
    tips = {
        "B8": "تاریخ شمسی: 1405/07/10 یا 14050710 (ارقام لاتین).",
        "C8": "شماره فاکتور — عدد آزاد؛ نیازی به تعریف قبلی ندارد.",
        "D8": "کد فروشنده: از کشویی (SL001…) انتخاب یا تایپ کنید. کد غلط = خطا (باید در INFO ← Sellers تعریف شده باشد).",
        "F8": "وزن ارسالی به کیلوگرم.",
        "G8": "وزن فروش به کیلوگرم.",
        "H8": "مبلغ صافی فاکتور به ریال.",
        "I8": "اختلاف وزن (ارسالی − فروش) — فقط در اکسل، در چاپ نمی‌آید.",
        "J8": "کنترل سقف ۱٪: اگر اختلاف وزن بیش از ۱٪ وزن ارسالی باشد ضربدر قرمز (√ = منطبق). فقط در اکسل.",
        "K8": "جمع وزن ۱۵ ردیف محصول — فقط در اکسل.",
        "L8": "وزن ارسالی منهای جمع وزن محصولات — فقط در اکسل (کنترل تطبیق).",
    }
    first_prod = f"{get_column_letter(slot_col(1, 0))}8"
    tips[first_prod] = ("محصول را از کشویی انتخاب کنید؛ نوع سبد و نوع مرکبات خودکار می‌آید. "
                        "برای هر فاکتور تا ۱۵ نوع محصول. این ستون‌ها در چاپ نمی‌آیند.")
    for ref, text in tips.items():
        ws[ref].comment = Comment(text, "مازندشت", height=100, width=280)

    ws.freeze_panes = "F10"
    setup_print(ws, "landscape", title_rows="7:9",
                print_area=f"A1:{print_last}{n1}",
                header_center="فروش محصولات به فروشندگان — Entry")
    return ws



# ---- seller statement sheets (SL001..SLnnn) ----
SELL_HDR1, SELL_HDR2, SELL_D0 = 15, 16, 17


def build_sales_ledger(wb, cap):
    """Ledger engine: merges sales invoices (Calc) + seller bank payments
    (Sync mirror of the Transactions Entry) into per-seller ranked events."""
    n0 = SY_DATA0
    rows = cap.entry_rows
    n1 = n0 + rows - 1
    ws = wb.create_sheet("Ledger")
    ws.sheet_properties.tabColor = C["tab_gray"]
    ws.sheet_view.rightToLeft = False
    ws.sheet_view.showGridLines = False
    brand_band(ws, 14, mode="info")
    merge_put(ws, "A3:N3",
              "برگه فنی — موتور صورت‌حساب فروشندگان (ترکیب فاکتورهای فروش و پرداخت‌های بانکی + رتبه‌بندی). فرمول‌ها را تغییر ندهید.",
              fnt(8, True, C["red"]), fl(C["amber_l"]), al("left"))
    for cl, t in zip("ABCD", ["A·Seller", "A·Key", "A·Rank", "A·RankKey"]):
        put(ws, f"{cl}5", t, fnt(8, True, C["white"]), fl(C["green"]), al("center"), BORDER_ALL)
    for cl, t in zip("FGHIJKLMN", ["B·Seller", "B·Y", "B·M", "B·D", "B·Date",
                                   "B·Key", "B·Rank", "B·RankKey", "B·Time"]):
        put(ws, f"{cl}5", t, fnt(8, True, C["white"]), fl(C["olive"] if "olive" in C else C["gray"]),
            al("center"), BORDER_ALL)
    for i in range(rows):
        r = n0 + i
        er = SALES_DATA0 + i     # Calc / sales-Entry row
        sr = n0 + i              # Sync row
        # --- block A: invoices ---
        ws[f"A{r}"] = (f'=IF(OR(Calc!$E{er}="",Calc!$E{er}="نامعتبر",Calc!$G{er}=""),"",'
                       f'IF(LEFT(Calc!$G{er},2)="SL",Calc!$G{er},""))')
        ws[f"B{r}"] = (f'=IF($A{r}="","",LEFT(Calc!$E{er},4)&MID(Calc!$E{er},6,2)&'
                       f'RIGHT(Calc!$E{er},2)&"0000"&"|"&"1"&RIGHT("00000"&{i + 1},5))')
        ws[f"C{r}"] = (f'=IF($B{r}="","",COUNTIFS(LdgSellerA,$A{r},LdgKeyA,"<"&$B{r})'
                       f'+COUNTIFS(LdgSellerB,$A{r},LdgKeyB,"<"&$B{r})+1)')
        ws[f"D{r}"] = f'=IF($C{r}="","",$A{r}&"|"&RIGHT("000000"&$C{r},6))'
        # --- block B: bank payments ---
        ws[f"F{r}"] = (f'=IF(Sync!$AA{sr}="","",IF(LEFT(UPPER(TRIM(Sync!$AA{sr})),2)="SL",'
                       f'UPPER(TRIM(Sync!$AA{sr})),""))')
        ws[f"G{r}"] = (f'=IF($F{r}="","",IF(Sync!$X{sr}="","",IF(ISNUMBER(Sync!$X{sr}),'
                       f'IF(AND(Sync!$X{sr}>=10000101,Sync!$X{sr}<=99991231),INT(Sync!$X{sr}/10000),""),'
                       f'IFERROR(VALUE(TRIM(LEFT(Sync!$X{sr},FIND("/",Sync!$X{sr})-1))),""))))')
        ws[f"H{r}"] = (f'=IF(OR($F{r}="",$G{r}=""),"",IF(ISNUMBER(Sync!$X{sr}),'
                       f'INT(MOD(Sync!$X{sr},10000)/100),'
                       f'IFERROR(VALUE(TRIM(LEFT(MID(Sync!$X{sr},FIND("/",Sync!$X{sr})+1,20),'
                       f'FIND("/",MID(Sync!$X{sr},FIND("/",Sync!$X{sr})+1,20))-1))),"")))')
        ws[f"I{r}"] = (f'=IF(OR($F{r}="",$G{r}=""),"",IF(ISNUMBER(Sync!$X{sr}),'
                       f'MOD(Sync!$X{sr},100),'
                       f'IFERROR(VALUE(TRIM(MID(MID(Sync!$X{sr},FIND("/",Sync!$X{sr})+1,20),'
                       f'FIND("/",MID(Sync!$X{sr},FIND("/",Sync!$X{sr})+1,20))+1,20))),"")))')
        ws[f"J{r}"] = (f'=IF($F{r}="","",IF(OR($G{r}="",$H{r}="",$I{r}="",$G{r}<1300,$G{r}>1500,'
                       f'$H{r}<1,$H{r}>12,$I{r}<1,$I{r}>31),"نامعتبر",'
                       f'RIGHT("0000"&$G{r},4)&"/"&RIGHT("00"&$H{r},2)&"/"&RIGHT("00"&$I{r},2)))')
        ws[f"N{r}"] = (f'=IF($F{r}="","",IF(Sync!$Y{sr}="","0000",'
                       f'IF(AND(ISNUMBER(Sync!$Y{sr}),Sync!$Y{sr}>=0,Sync!$Y{sr}<1),'
                       f'RIGHT("00"&HOUR(Sync!$Y{sr}),2)&RIGHT("00"&MINUTE(Sync!$Y{sr}),2),'
                       f'IFERROR(RIGHT("00"&VALUE(LEFT(Sync!$Y{sr},FIND(":",Sync!$Y{sr})-1)),2)&'
                       f'RIGHT("00"&VALUE(MID(Sync!$Y{sr},FIND(":",Sync!$Y{sr})+1,2)),2),"0000"))))')
        ws[f"K{r}"] = (f'=IF(OR($F{r}="",$J{r}="",$J{r}="نامعتبر"),"",'
                       f'LEFT($J{r},4)&MID($J{r},6,2)&RIGHT($J{r},2)&$N{r}&"|"&"2"&RIGHT("00000"&{i + 1},5))')
        ws[f"L{r}"] = (f'=IF($K{r}="","",COUNTIFS(LdgSellerA,$F{r},LdgKeyA,"<"&$K{r})'
                       f'+COUNTIFS(LdgSellerB,$F{r},LdgKeyB,"<"&$K{r})+1)')
        ws[f"M{r}"] = f'=IF($L{r}="","",$F{r}&"|"&RIGHT("000000"&$L{r},6))'
    set_widths(ws, {"A": 9, "B": 18, "C": 7, "D": 15, "F": 9, "G": 6, "H": 6,
                    "I": 6, "J": 12, "K": 18, "L": 7, "M": 15, "N": 7})
    ws.freeze_panes = "A6"
    return ws


def build_seller_sheet(wb, idx, cap):
    """One statement sheet per seller: SL001.. — auto ledger from Ledger engine."""
    code = f"SL{idx:03d}"
    n0 = SELL_D0
    n1 = n0 + cap.seller_rows - 1
    ws = wb.create_sheet(code)
    ws.sheet_properties.tabColor = C["tab_green"]
    rtl(ws, 85)
    last = "J"
    brand_band(ws, 10, mode="txn", comp_col="N")
    merge_put(ws, "A3:J3",
              f'="صورت حساب فروشنده — "&IFERROR(IF(VLOOKUP("{code}",SellerTbl,2,0)="","(بدون عنوان)",'
              f'VLOOKUP("{code}",SellerTbl,2,0)),"—")',
              fnt(13, True, C["orange_d"]), fl(C["green_xl"]), al("center"),
              bd(b=sd("medium", C["orange"])))
    ws.row_dimensions[3].height = 24

    # ---- account info block rows 5-10 ----
    m_ph = f'MATCH("{code}",SellerCodes,0)'
    info_rows = [
        ("کد سازمانی حساب:", code, None),
        ("نام صاحب حساب:", f'=IFERROR(IF(VLOOKUP("{code}",SellerTbl,2,0)="","(بدون عنوان)",VLOOKUP("{code}",SellerTbl,2,0)),"—")', None),
        ("نوع مشتری:", f'=IFERROR(IF(VLOOKUP("{code}",SellerTbl,3,0)="","—",VLOOKUP("{code}",SellerTbl,3,0)),"—")', None),
        ("تلفن / همراه:", f'=IFERROR(IF(AND(INDEX(SellerPhone,{m_ph})="",INDEX(SellerMobile,{m_ph})=""),"—",'
                          f'TRIM(INDEX(SellerPhone,{m_ph})&"  /  "&INDEX(SellerMobile,{m_ph}))),"—")', None),
        ("نشانی:", f'=IFERROR(IF(INDEX(SellerAddr,{m_ph})="","—",INDEX(SellerAddr,{m_ph})),"—")', None),
    ]
    for i, (lbl, val, _x) in enumerate(info_rows):
        r = 5 + i
        merge_put(ws, f"A{r}:B{r}", lbl, fnt(8.5, True, C["gray"]), fl(C["green_l"]),
                  al("right"), BORDER_ALL)
        merge_put(ws, f"C{r}:J{r}", val, fnt(10, True if i == 0 else False,
                                             C["orange_d"] if i == 0 else C["txt"]),
                  fl(C["auto"] if i < 5 else C["input"]), al("right"), BORDER_ALL)
        ws.row_dimensions[r].height = 17
    # opening balance (manual)
    merge_put(ws, "A10:B10", "مانده اولیه حساب (دستی):", fnt(8.5, True, C["gray"]),
              fl(C["green_l"]), al("right"), BORDER_ALL)
    merge_put(ws, "C10:D10", 0, fnt(10, True, C["txt"]), fl(C["input"]), al("center"),
              BORDER_ALL, "#,##0")
    merge_put(ws, "E10:J10", "اگر حساب از قبل مانده دارد اینجا بنویسید (مثبت = بدهکار فروشنده) — در مانده‌ها اثر می‌گذارد.",
              fnt(7.5, False, C["gray"], italic=True), fl(C["white"]), al("right"), BORDER_ALL)

    # ---- KPI cards rows 11-14 ----
    cards1 = [
        (1, 2, "جمع بدهکار (ریال)", f"=SUM($H${n0}:$H${n1})", "#,##0"),
        (3, 4, "جمع بستانکار (ریال)", f"=SUM($I${n0}:$I${n1})", "#,##0"),
        (5, 6, "مانده نهایی حساب", f"=IFERROR(LOOKUP(9.99E+307,$J${n0}:$J${n1}),$C$10)", "#,##0"),
        (7, 10, "ماهیت حساب",
         f'=IF($E$12>0,"▲ بدهکار — فروشنده به ما بدهکار است",IF($E$12<0,'
         f'"▼ بستانکار — ما به فروشنده بدهکاریم","= تسویه شده"))', None),
    ]
    cards2 = [
        (1, 2, "تعداد فاکتور فروش", f'=COUNTIF($C${n0}:$C${n1},"فروش")', "0"),
        (3, 4, "تعداد پرداخت بانکی", f'=COUNTIF($C${n0}:$C${n1},"پرداخت بانکی")', "0"),
        (5, 6, "جمع وزن ارسالی (kg)", f'=SUMIF($C${n0}:$C${n1},"فروش",$F${n0}:$F${n1})', "#,##0.0"),
        (7, 8, "جمع وزن فروش (kg)", f'=SUMIF($C${n0}:$C${n1},"فروش",$G${n0}:$G${n1})', "#,##0.0"),
        (9, 10, "آخرین تراکنش", f'=IF(COUNT($J${n0}:$J${n1})=0,"—",INDEX($B${n0}:$B${n1},COUNT($J${n0}:$J${n1})))', None),
    ]
    for (c0, c1, lbl, formula, fmt) in cards1:
        card(ws, 11, 12, c0, c1, lbl, formula, fmt)
    for (c0, c1, lbl, formula, fmt) in cards2:
        card(ws, 13, 14, c0, c1, lbl, formula, fmt)
    ws.row_dimensions[11].height = 13
    ws.row_dimensions[12].height = 22
    ws.row_dimensions[13].height = 13
    ws.row_dimensions[14].height = 22

    # ---- ledger headers ----
    table_headers(ws, SELL_HDR1, SELL_HDR2, [
        ("ردیف", "No."), ("تاریخ", "Date"), ("نوع عملیات", "Type"),
        ("مرجع", "Ref"), ("توضیحات", "Description"),
        ("وزن ارسالی (kg)", "Sent Wt"), ("وزن فروش (kg)", "Sold Wt"),
        ("بدهکار (ریال)", "Debit"), ("بستانکار (ریال)", "Credit"), ("مانده (ریال)", "Balance"),
    ])

    # ---- data rows (fast shared styles) ----
    F_BODY = fnt(9)
    F_NO = fnt(9, False, C["gray"])
    B_ALL = BORDER_ALL
    AL_C = al("center")
    AL_R = al("right")
    FILL_W = fl(C["white"])
    FILL_Z = fl(C["zebra"])
    NF_INT = "#,##0"
    NF_W = "#,##0.0"
    for k in range(1, cap.seller_rows + 1):
        r = n0 + k - 1
        zeb = FILL_W if k % 2 else FILL_Z
        cells = [
            ("A", k, F_NO, zeb, None, AL_C),
            ("B", f'=IF(AND($L{r}="",$M{r}=""),"",IF($L{r}<>"",INDEX(SCNormDate,$L{r}),INDEX(LdgDateB,$M{r})))', F_BODY, zeb, None, AL_C),
            ("C", f'=IF($L{r}<>"","فروش",IF($M{r}<>"","پرداخت بانکی",""))', F_BODY, zeb, None, AL_C),
            ("D", f'=IF($L{r}<>"",INDEX(SCInv,$L{r}),IF($M{r}<>"",INDEX(LdgBankB,$M{r}),""))', F_BODY, zeb, None, AL_C),
            ("E", f'=IF($L{r}<>"",IF(INDEX(SCItems,$L{r})=0,"","فروش محصولات — "&INDEX(SCItems,$L{r})&" قلم"),'
                  f'IF($M{r}<>"",IF(INDEX(LdgNoteB,$M{r})="","—",INDEX(LdgNoteB,$M{r})),""))', F_BODY, zeb, None, AL_R),
            ("F", f'=IF($L{r}<>"",IF(INDEX(SCWSent,$L{r})=0,"",INDEX(SCWSent,$L{r})),"")', F_BODY, zeb, NF_W, AL_C),
            ("G", f'=IF($L{r}<>"",IF(INDEX(SCWSold,$L{r})=0,"",INDEX(SCWSold,$L{r})),"")', F_BODY, zeb, NF_W, AL_C),
            ("H", f'=IF($L{r}<>"",IF(INDEX(SCAmount,$L{r})=0,"",INDEX(SCAmount,$L{r})),'
                  f'IF($M{r}<>"",IF(INDEX(LdgWitB,$M{r})=0,"",INDEX(LdgWitB,$M{r})),""))', F_BODY, zeb, NF_INT, AL_C),
            ("I", f'=IF($M{r}<>"",IF(INDEX(LdgDepB,$M{r})=0,"",INDEX(LdgDepB,$M{r})),"")', F_BODY, zeb, NF_INT, AL_C),
        ]
        if k == 1:
            jf = f'=IF($B{r}="","",$C$10+IF($H{r}="",0,$H{r})-IF($I{r}="",0,$I{r}))'
        else:
            jf = (f'=IF($B{r}="","",IF($J{r - 1}="",$C$10,$J{r - 1})'
                  f'+IF($H{r}="",0,$H{r})-IF($I{r}="",0,$I{r}))')
        cells.append(("J", jf, fnt(9, True), zeb, NF_INT, AL_C))
        for col, val, f_, fill, nf, alg in cells:
            c = ws[f"{col}{r}"]
            c.value = val
            c.font = f_
            c.fill = fill
            c.border = B_ALL
            c.alignment = alg
            if nf:
                c.number_format = nf
        # hidden helpers L/M
        ws[f"L{r}"] = f'=IFERROR(MATCH("{code}|"&RIGHT("000000"&{k},6),LdgRankKeyA,0),"")'
        ws[f"M{r}"] = f'=IFERROR(MATCH("{code}|"&RIGHT("000000"&{k},6),LdgRankKeyB,0),"")'

    # ---- conditional formatting ----
    ok_fill = fl(C["ok_l"]); ok_font = fnt(9, True, C["ok"])
    bad_fill = fl(C["red_l"]); bad_font = fnt(9, True, C["red"])
    amber_fill = fl(C["amber_l"]); amber_font = fnt(9, True, C["amber"])
    c_rng = f"C{n0}:C{n1}"
    ws.conditional_formatting.add(c_rng, CellIsRule(operator="equal", formula=['"فروش"'], fill=amber_fill, font=amber_font))
    ws.conditional_formatting.add(c_rng, CellIsRule(operator="equal", formula=['"پرداخت بانکی"'], fill=ok_fill, font=ok_font))
    ws.conditional_formatting.add(f"H{n0}:H{n1}", CellIsRule(operator="greaterThan", formula=["0"], font=fnt(9, True, C["amber"])))
    ws.conditional_formatting.add(f"I{n0}:I{n1}", CellIsRule(operator="greaterThan", formula=["0"], font=fnt(9, True, C["ok"])))
    ws.conditional_formatting.add(f"J{n0}:J{n1}", CellIsRule(operator="lessThan", formula=["0"], font=fnt(9, True, C["red"])))
    ws.conditional_formatting.add("G12", FormulaRule(formula=['ISNUMBER(SEARCH("▲",$G$12))'], fill=amber_fill, font=fnt(10, True, C["amber"])))
    ws.conditional_formatting.add("G12", FormulaRule(formula=['ISNUMBER(SEARCH("▼",$G$12))'], fill=bad_fill, font=fnt(10, True, C["red"])))
    ws.conditional_formatting.add("G12", FormulaRule(formula=['ISNUMBER(SEARCH("تسویه",$G$12))'], fill=ok_fill, font=fnt(10, True, C["ok"])))
    ws.conditional_formatting.add("E12", CellIsRule(operator="greaterThan", formula=["0"], font=fnt(11, True, C["amber"])))
    ws.conditional_formatting.add("E12", CellIsRule(operator="lessThan", formula=["0"], font=fnt(11, True, C["red"])))
    ws.conditional_formatting.add("E12", CellIsRule(operator="equal", formula=["0"], font=fnt(11, True, C["ok"])))

    set_widths(ws, {"A": 5.5, "B": 13, "C": 13, "D": 24, "E": 30, "F": 12, "G": 12,
                    "H": 15, "I": 15, "J": 16, "L": 8, "M": 8})
    ws.column_dimensions["L"].hidden = True
    ws.column_dimensions["M"].hidden = True
    ws.freeze_panes = "A17"
    setup_print(ws, "landscape", title_rows="15:16", print_area=f"A1:J{n1}",
                header_center=f"صورت حساب فروشنده — {code}")
    return ws

def build_sales_help(wb, cap):
    ws = wb.create_sheet("Help")
    ws.sheet_properties.tabColor = C["tab_blue"]
    rtl(ws, 100)
    brand_band(ws, 8, mode="info")
    title_row(ws, 8, 3, "راهنمای فایل فروش (Mazandasht_Sales.xlsx)")
    sections = [
        ("۱) ساختار فایل", [
            "Entry — ورود داده: هر ردیف = یک فاکتور فروش (یک ماشین‌بار). ردیف‌ها به هر ترتیبی زیر هم ثبت شوند.",
            "۱۵ ردیف محصول در هر ردیف (ستون‌های بعد از «اختلاف با ارسالی»، خارج از محدوده چاپ): محصول (کشویی) | نوع سبد (خودکار) | نوع مرکبات (خودکار) | تعداد سبد | وزن.",
            "SL001 … SL100 — صورت‌حساب خودکار هر فروشنده: فاکتورها (بدهکار) و پرداخت‌های بانکی (بستانکار) به ترتیب تاریخ + مانده و ماهیت حساب.",
            "Sync — آینه فایل مادر (INFO) و فایل واریز/برداشت (Transactions)؛ منبع کشوها و کدها. دست نزنید.",
            "Calc / Ledger — موتورهای محاسبات (نرمال‌سازی تاریخ، رتبه‌بندی رویدادها). دست نزنید.",
        ]),
        ("۲) رنگ‌ها و نشانه‌ها", [
            "زرد = ورودی شما | خاکستری = خودکار.",
            f"{OK_MARK} = اختلاف وزن حداکثر ۱٪ وزن ارسالی است.",
            f"{BAD_MARK} = اختلاف وزن بیش از ۱٪ وزن ارسالی — کنترل کنید.",
            "«کد ناشناخته» = کد در INFO تعریف نشده (کشویی ورودی غیرمجاز را متوقف می‌کند؛ این علامت برای موارد چسبانده‌شده است).",
            "ستون‌های «فقط اکسل» (اختلاف وزن تا پایان ۱۵ ردیف محصول) در چاپ نمی‌آیند؛ با دکمه +/− بالای ستون‌ها قابل جمع‌شدن‌اند.",
        ]),
        ("۳) اتصال به فایل مادر", [
            "این فایل و Mazandasht_INFO.xlsx همیشه در یک پوشه باشند.",
            "پس از تغییر INFO (فروشنده/محصول جدید و…): Data ← Edit Links ← Update Values.",
            "فروشندگان در INFO ← شیت Sellers تعریف می‌شوند (کد SL…، نام، نوع مشتری که خودش کشویی است و در PartyTypes گسترش می‌یابد: حجره‌دار، صادراتی، واسطه، نقدی و…).",
            "محصولات فروش در INFO ← Products_Sell — هر محصول «نوع سبد» و «دسته (نوع مرکبات)» دارد که اینجا خودکار نمایش داده می‌شود.",
        ]),
        ("۴) نکات ثبت", [
            "تاریخ: 1405/07/10 یا 14050710 — شماره فاکتور: عدد آزاد.",
            "کد فروشنده: تایپ (SL001) یا کشویی؛ اگر کد غلط باشد اکسل خطا می‌دهد (Stop).",
            "وزن‌ها کیلوگرم (اعشار مجاز) — مبلغ صافی ریال.",
            "کنترل ۱٪: اختلاف وزن ارسالی/فروش بیش از ۱٪ ارسالی باشد، ضربدر قرمز می‌گیرید.",
            "ظرفیت: %d ردیف فاکتور × ۱۵ ردیف محصول. برای بیشتر، فرمول‌های ردیف آخر را به پایین کپی کنید." % cap.entry_rows,
        ]),
        ("۵) شیت‌های صورت‌حساب فروشنده (SL001…)", [
            "سربرگ هر شیت: کد سازمانی، نام، نوع مشتری، تلفن و نشانی — خودکار از INFO؛ «مانده اولیه» دستی است.",
            "کارت‌های بالای صفحه: جمع بدهکار، جمع بستانکار، مانده نهایی، ماهیت حساب (▲ بدهکار / ▼ بستانکار / = تسویه)، تعداد فاکتور و پرداخت، جمع وزن‌ها و آخرین تراکنش.",
            "هر ردیف: تاریخ | نوع عملیات | مرجع | توضیحات | وزن ارسالی و فروش (فقط ردیف‌های فروش) | بدهکار | بستانکار | مانده.",
            "فروش = مبلغ صافی در ستون بدهکار؛ پرداخت بانکی فروشنده (ثبت‌شده در فایل واریز/برداشت با کد SLxxx) = ستون بستانکار.",
            "ردیف‌ها خودکار به تاریخ (و ساعت پرداخت‌ها) مرتب می‌شوند؛ چیزی در این شیت‌ها تایپ نکنید.",
            "ظرفیت هر شیت: %d ردیف." % cap.seller_rows,
        ]),
        ("۶) آینده", [
            "دیتای این Entry برای فایل‌ها/شیت‌های تحلیلی آینده (فروش هر فروشنده، آنالیز محصول و سبد، مطابقت آماری و حسابرسی) آماده است — از طریق Calc و نام‌های تعریف‌شده (SC…).",
        ]),
    ]
    r = 5
    for title, lines in sections:
        merge_put(ws, f"A{r}:H{r}", title, fnt(10.5, True, C["white"]), fl(C["green"]),
                  al("right"), bd(b=sd("medium", C["orange"])))
        ws.row_dimensions[r].height = 19
        r += 1
        for i, ln in enumerate(lines):
            merge_put(ws, f"A{r}:H{r}", "• " + ln, fnt(9),
                      fl(C["white"] if i % 2 == 0 else C["zebra"]), al("right", wrap=True), BORDER_ALL)
            ws.row_dimensions[r].height = 17
            r += 1
        r += 1
    set_widths(ws, {"A": 12, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12})
    setup_print(ws, "portrait", header_center="راهنما — فایل فروش")
    return ws


def build_sales(cap, out_path, external=True):
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.title = "Mazandasht Sales — فروش به فروشندگان"
    wb.properties.creator = BRAND_NAME
    wb.properties.company = "Mazandasht"

    if not external:
        for name in INFO_SHEET_ORDER[1:]:
            wb.create_sheet(name).sheet_properties.tabColor = C["tab_gray"]
        build_info_sheets_into(wb, cap)
        # local stand-in for the bank Transactions Entry (QA mode)
        bws = wb.create_sheet("BankEntry")
        rtl(bws)
        bws.sheet_properties.tabColor = C["tab_gray"]
        set_widths(bws, {"C": 12, "D": 8, "E": 28, "G": 10, "I": 18, "J": 14,
                         "K": 14, "L": 14, "P": 32})
        seed_bank = list(SEED_ENTRY) + [
            ("1405/07/12", "09:00", "BANK01 – ملت — جاری شرکت", "SL005", "سایر", "واریز",
             100000000, "", "", "پرداخت بدون فاکتور (نمونه)"),
        ]
        for i, row in enumerate(seed_bank):
            r = E_DATA0 + i
            date, time_, bank, party, ptype, ttype, dep, wit, man, note = row
            bws[f"C{r}"] = date
            bws[f"D{r}"] = time_
            bws[f"E{r}"] = bank
            bws[f"G{r}"] = party
            bws[f"I{r}"] = ptype
            bws[f"J{r}"] = ttype
            if dep != "":
                bws[f"K{r}"] = dep
            if wit != "":
                bws[f"L{r}"] = wit
            bws[f"P{r}"] = note

    build_sales_entry(wb, cap)
    build_sales_sync(wb, cap, external)
    build_sales_calc(wb, cap)
    build_sales_ledger(wb, cap)
    for i in range(1, cap.sellers + 1):
        build_seller_sheet(wb, i, cap)
    build_sales_help(wb, cap)

    order = (["Entry", "Sync", "Calc", "Ledger"] +
             [f"SL{i:03d}" for i in range(1, cap.sellers + 1)] + ["Help"])
    if not external:
        order += ["BankEntry"] + INFO_SHEET_ORDER[1:]
    wb._sheets = sorted(wb._sheets, key=lambda s: order.index(s.title))
    wb.active = 0

    n0, n1 = SALES_DATA0, s_data1(cap)
    y0 = SY_DATA0
    c0, c1 = SALES_DATA0, SALES_DATA0 + cap.entry_rows - 1
    names = {
        "SellerTbl": f"Sync!$A${y0}:$C${y0 + cap.sellers - 1}",
        "SellerCodes": f"Sync!$A${y0}:$A${y0 + cap.sellers - 1}",
        "SellerList": (f"OFFSET(Sync!$R${y0},0,0,"
                       f"MAX(COUNTIF(Sync!$R${y0}:$R${y0 + cap.sellers - 1},\"?*\"),1),1)"),
        "ProdTbl": f"Sync!$E${y0}:$I${y0 + cap.products_sell - 1}",
        "ProdCodes": f"Sync!$E${y0}:$E${y0 + cap.products_sell - 1}",
        "ProdList": (f"OFFSET(Sync!$V${y0},0,0,"
                     f"MAX(COUNTIF(Sync!$V${y0}:$V${y0 + cap.products_sell - 1},\"?*\"),1),1)"),
        "BasketTbl": f"Sync!$K${y0}:$L${y0 + cap.baskets - 1}",
        "SellerPhone": f"Sync!$BA${y0}:$BA${y0 + cap.sellers - 1}",
        "SellerMobile": f"Sync!$BB${y0}:$BB${y0 + cap.sellers - 1}",
        "SellerAddr": f"Sync!$BC${y0}:$BC${y0 + cap.sellers - 1}",
        "LdgSellerA": f"Ledger!$A${y0}:$A${y0 + cap.entry_rows - 1}",
        "LdgKeyA": f"Ledger!$B${y0}:$B${y0 + cap.entry_rows - 1}",
        "LdgRankKeyA": f"Ledger!$D${y0}:$D${y0 + cap.entry_rows - 1}",
        "LdgSellerB": f"Ledger!$F${y0}:$F${y0 + cap.entry_rows - 1}",
        "LdgKeyB": f"Ledger!$K${y0}:$K${y0 + cap.entry_rows - 1}",
        "LdgRankKeyB": f"Ledger!$M${y0}:$M${y0 + cap.entry_rows - 1}",
        "LdgDateB": f"Ledger!$J${y0}:$J${y0 + cap.entry_rows - 1}",
        "LdgDepB": f"Sync!$AD${y0}:$AD${y0 + cap.entry_rows - 1}",
        "LdgWitB": f"Sync!$AE${y0}:$AE${y0 + cap.entry_rows - 1}",
        "LdgBankB": f"Sync!$Z${y0}:$Z${y0 + cap.entry_rows - 1}",
        "LdgNoteB": f"Sync!$AF${y0}:$AF${y0 + cap.entry_rows - 1}",
        "SCNormDate": f"Calc!$E${c0}:$E${c1}",
        "SCYM": f"Calc!$F${c0}:$F${c1}",
        "SCSeller": f"Calc!$G${c0}:$G${c1}",
        "SCInv": f"Calc!$H${c0}:$H${c1}",
        "SCWSent": f"Calc!$I${c0}:$I${c1}",
        "SCWSold": f"Calc!$J${c0}:$J${c1}",
        "SCAmount": f"Calc!$K${c0}:$K${c1}",
        "SCWProd": f"Calc!$L${c0}:$L${c1}",
        "SCItems": f"Calc!$M${c0}:$M${c1}",
        "SalesDate": f"Entry!$B${n0}:$B${n1}",
        "SalesSellerCell": f"Entry!$D${n0}:$D${n1}",
        "SalesSellerName": f"Entry!$E${n0}:$E${n1}",
        "SalesWSent": f"Entry!$F${n0}:$F${n1}",
        "SalesWSold": f"Entry!$G${n0}:$G${n1}",
        "SalesAmount": f"Entry!$H${n0}:$H${n1}",
    }
    for n, ref in names.items():
        wb.defined_names[n] = DefinedName(n, attr_text=ref)

    wb.save(out_path)
    return out_path


# ============================================================================
# 6. EXTERNAL LINK INJECTION (ECMA-376 externalLink part)
# ============================================================================

def mirror_map(cap):
    """INFO ranges mirrored into consumer workbooks: sheet -> (cols, row0, row1)."""
    rng = lambda r0, n: (S_DATA0, S_DATA0 + n - 1)
    return {
        "Company": ([3], 5, 16),
        "Banks": (list(range(1, 5)), *rng(0, cap.banks)),
        "Customers": (list(range(1, 4)), *rng(0, cap.customers)),
        "Sellers": (list(range(1, 7)), *rng(0, cap.sellers)),
        "Suppliers": (list(range(1, 4)), *rng(0, cap.suppliers)),
        "Accounts": (list(range(1, 4)), *rng(0, cap.accounts)),
        "Staff": (list(range(1, 3)), *rng(0, cap.staff)),
        "PartyTypes": (list(range(1, 3)), *rng(0, cap.party_types)),
        "TxnTypes": (list(range(1, 3)), *rng(0, cap.txn_types)),
        "Products_Buy": (list(range(1, 3)), *rng(0, cap.products_buy)),
        "Products_Sell": (list(range(1, 6)), *rng(0, cap.products_sell)),
        "Baskets": (list(range(1, 3)), *rng(0, cap.baskets)),
        "Warehouses": (list(range(1, 3)), *rng(0, cap.warehouses)),
    }


TXN_SHEET_ORDER = (["Entry"] + [f"Bank{i}" for i in range(1, 11)] +
                   ["Summary", "Sync", "Calc", "Help"])


def txn_mirror(cap):
    """Transactions Entry ranges mirrored into the sales workbook (input columns)."""
    return {"Entry": ([3, 4, 5, 7, 9, 10, 11, 12, 16],
                      E_DATA0, E_DATA0 + cap.entry_rows - 1)}


def _build_link_cache(source_path, sheet_order, mirror):
    """Build sheetNames + cached sheetData XML for one external link."""
    src = load_workbook(source_path, data_only=False)
    blocks = []
    for sid, name in enumerate(sheet_order):
        if name not in mirror or name not in src.sheetnames:
            continue
        cols, r0, r1 = mirror[name]
        ws = src[name]
        rows_xml = []
        for r in range(r0, r1 + 1):
            cells = []
            for c in cols:
                v = ws.cell(row=r, column=c).value
                if v is None or v == "" or (isinstance(v, str) and v.startswith("=")):
                    continue
                ref = f"{get_column_letter(c)}{r}"
                if isinstance(v, bool):
                    cells.append(f'<cell r="{ref}" t="b"><v>{1 if v else 0}</v></cell>')
                elif isinstance(v, (int, float)):
                    val = int(v) if float(v).is_integer() else v
                    cells.append(f'<cell r="{ref}"><v>{val}</v></cell>')
                else:
                    v_esc = (str(v).replace("&", "&amp;").replace("<", "&lt;")
                             .replace(">", "&gt;").replace('"', "&quot;"))
                    cells.append(f'<cell r="{ref}" t="str"><v>{v_esc}</v></cell>')
            if cells:
                rows_xml.append(f'<row r="{r}">' + "".join(cells) + "</row>")
        blocks.append((sid, "".join(rows_xml)))
    sheet_names_xml = "".join(f'<sheetName val="{n}"/>' for n in sheet_order)
    dataset_xml = "".join(
        f'<sheetData sheetId="{sid}">{rows}</sheetData>' if rows
        else f'<sheetData sheetId="{sid}"/>' for sid, rows in blocks)
    return sheet_names_xml, dataset_xml


def inject_external_links(path, links):
    """
    Add real external-workbook links to a saved xlsx.
    links: list of (filename, source_path, sheet_order, mirror) — link numbers
    are assigned in order ([1], [2], ...).
    """
    parts = {}          # part name -> bytes
    rel_entries = []    # (link_no)
    wb_rels_add = []
    ext_ref_ids = []
    content_add = []

    zin = zipfile.ZipFile(path)
    items = [(i.filename, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    data = dict(items)

    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
    next_id = (max(ids) + 1) if ids else 1

    for no, (filename, source_path, sheet_order, mirror) in enumerate(links, start=1):
        names_xml, dataset_xml = _build_link_cache(source_path, sheet_order, mirror)
        ext_link = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<externalBook r:id="rId1">'
            f'<sheetNames>{names_xml}</sheetNames>'
            f'<sheetDataSet>{dataset_xml}</sheetDataSet>'
            '</externalBook></externalLink>'
        )
        fname = filename.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        ext_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
            f'Target="{fname}" TargetMode="External"/>'
            '</Relationships>'
        )
        parts[f"xl/externalLinks/externalLink{no}.xml"] = ext_link.encode("utf-8")
        parts[f"xl/externalLinks/_rels/externalLink{no}.xml.rels"] = ext_rels.encode("utf-8")
        wb_rels_add.append(
            f'<Relationship Id="rId{next_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" '
            f'Target="externalLinks/externalLink{no}.xml"/>')
        ext_ref_ids.append(f"rId{next_id}")
        content_add.append(
            f'<Override PartName="/xl/externalLinks/externalLink{no}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>')
        next_id += 1

    # workbook rels
    rels = rels.replace("</Relationships>", "".join(wb_rels_add) + "</Relationships>")
    data["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")

    # workbook.xml: <externalReferences> with all children
    wbxml = data["xl/workbook.xml"].decode("utf-8")
    extrefs = "<externalReferences>" + "".join(
        f'<externalReference r:id="{rid}"/>' for rid in ext_ref_ids) + "</externalReferences>"
    if "<externalReferences>" in wbxml:
        wbxml = re.sub(r"<externalReferences>.*?</externalReferences>", extrefs, wbxml, flags=re.S)
    elif "<definedNames>" in wbxml:
        wbxml = wbxml.replace("<definedNames>", extrefs + "<definedNames>", 1)
    elif "<calcPr" in wbxml:
        wbxml = wbxml.replace("<calcPr", extrefs + "<calcPr", 1)
    else:
        wbxml = wbxml.replace("</workbook>", extrefs + "</workbook>")
    data["xl/workbook.xml"] = wbxml.encode("utf-8")

    # content types
    ct = data["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>", "".join(content_add) + "</Types>")
    data["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, _ in items:
            z.writestr(name, data[name])
        for name, blob in parts.items():
            z.writestr(name, blob)
    return path


def inject_external_link(txn_path, info_filename, info_source_path, cap):
    """Single external link (INFO) — used for the Transactions workbook."""
    return inject_external_links(
        txn_path,
        [(info_filename, info_source_path, INFO_SHEET_ORDER, mirror_map(cap))])


# ============================================================================
# 7. MAIN
# ============================================================================

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true", help="build full deliverables")
    ap.add_argument("--qa", action="store_true", help="small QA build")
    ap.add_argument("--external", action="store_true",
                    help="(QA) build externally-linked pair instead of embedded")
    ap.add_argument("--out", default=".", help="output directory")
    args = ap.parse_args()

    out = os.path.abspath(args.out)
    os.makedirs(out, exist_ok=True)
    cap = Cap() if args.all else QA_CAP

    if args.qa:
        if args.external:
            info_path = os.path.join(out, INFO_FILE)
            build_info(QA_CAP, info_path)
            txn_path = os.path.join(out, TXN_FILE)
            build_transactions(QA_CAP, txn_path, external=True)
            inject_external_link(txn_path, INFO_FILE, info_path, QA_CAP)
            sales_path = os.path.join(out, SALES_FILE)
            build_sales(QA_CAP, sales_path, external=True)
            inject_external_links(sales_path, [
                (INFO_FILE, info_path, INFO_SHEET_ORDER, mirror_map(QA_CAP)),
                (TXN_FILE, txn_path, TXN_SHEET_ORDER, txn_mirror(QA_CAP)),
            ])
            print("QA external set:", info_path, txn_path, sales_path)
        else:
            txn_path = os.path.join(out, "QA_internal.xlsx")
            build_transactions(QA_CAP, txn_path, external=False)
            sales_path = os.path.join(out, "QA_sales_internal.xlsx")
            build_sales(QA_CAP, sales_path, external=False)
            print("QA internal:", txn_path, sales_path)
        return

    info_path = os.path.join(out, INFO_FILE)
    txn_path = os.path.join(out, TXN_FILE)
    sales_path = os.path.join(out, SALES_FILE)
    build_info(cap, info_path)
    build_transactions(cap, txn_path, external=True)
    inject_external_link(txn_path, INFO_FILE, info_path, cap)
    build_sales(cap, sales_path, external=True)
    inject_external_links(sales_path, [
        (INFO_FILE, info_path, INFO_SHEET_ORDER, mirror_map(cap)),
        (TXN_FILE, txn_path, TXN_SHEET_ORDER, txn_mirror(cap)),
    ])
    print("INFO :", info_path)
    print("TXN  :", txn_path)
    print("SALES:", sales_path)


if __name__ == "__main__":
    main()
