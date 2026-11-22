import collections
import io
import math
import os
import re
import sys
from datetime import datetime, timedelta, timezone

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from playwright.sync_api import sync_playwright

BASE = "https://www.boclife.com.hk/tc/support/financial-information/"
PAGES = [("總現金價值比率", BASE + "total-cash-value-ratio.html", False),
         ("分紅實現率", BASE + "dividend-fulfillment-ratios.html", True)]
COMPANY = "中銀人壽"
OUT_DIR = "output"
NOTION_VERSION = "2025-09-03"
AUTO = "自動生成"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
NAME_MAX = 200

JS_TABLES = """() => {
  const clean = s => (s || '').replace(/\\s+/g, ' ').trim();
  return Array.from(document.querySelectorAll('table')).map(t => {
    const raw = Array.from(t.rows).map(tr => Array.from(tr.cells).map(c => clean(c.textContent)));
    const grid = [];
    Array.from(t.rows).forEach((tr, r) => {
      grid[r] = grid[r] || [];
      let c = 0;
      Array.from(tr.cells).forEach(cell => {
        while (grid[r][c] !== undefined) c++;
        const txt = clean(cell.textContent);
        const rs = cell.rowSpan || 1, cs = cell.colSpan || 1;
        for (let i = 0; i < rs; i++) {
          grid[r + i] = grid[r + i] || [];
          for (let j = 0; j < cs; j++) grid[r + i][c + j] = txt;
        }
        c += cs;
      });
    });
    const ctx = [];
    let node = t;
    for (let d = 1; d <= 9 && node.parentElement && node.parentElement !== document.body; d++) {
      let sib = node.previousElementSibling;
      for (let k = 0; sib && k < 4; k++) {
        const s = clean(sib.textContent);
        if (s && s.length <= 200 && sib.tagName !== 'TABLE' && !sib.querySelector('table')) ctx.push(s);
        sib = sib.previousElementSibling;
      }
      node = node.parentElement;
      if (node.id) {
        const id = node.id;
        const sel = '[href="#' + id + '"],[data-bs-target="#' + id + '"],[data-target="#' + id +
                    '"],[aria-controls="' + id + '"]';
        try {
          document.querySelectorAll(sel).forEach(x => {
            const s = clean(x.textContent);
            if (s && s.length <= 200) ctx.push(s);
          });
        } catch (e) {}
      }
    }
    return {raw: raw, ctx: ctx,
            grid: grid.map(row => Array.from(row, x => x === undefined ? '' : x))};
  });
}"""
JS_TEXT = "() => document.body ? document.body.textContent : ''"
JS_NOTES = """() => {
  const clean = s => (s || '').replace(/\\s+/g, ' ').trim();
  const out = [];
  document.querySelectorAll('p, li, span, div, small').forEach(e => {
    if (e.closest('table') || e.children.length > 3) return;
    const s = clean(e.textContent);
    if (s.length >= 8 && s.length <= 300 && /N[/]A|不適用/.test(s) && !out.includes(s)) out.push(s);
  });
  return out;
}"""

PCT = re.compile(r"^(\d+(?:\.\d+)?)\s*[%％]$")
NA = re.compile(r"^(?:N\s*/?\s*A|不適用)\s*[\(（]?[0-9A-Za-z\*#^]{0,3}[\)）]?$", re.I)
CJK = re.compile(r"[\u4e00-\u9fff]")
TYPE = re.compile(r"[\u4e00-\u9fff]{0,6}(?:紅利|非保證入息|非保證年金)")
NAME = re.compile(r"保|計劃|年金|系列")
SKIP = re.compile(r"實現率|比率|保單年度|報告年度|短片|按此|請選擇|未有數字")

NF = 3
FIXED = ["產品名稱", "產品類別", "類型"]
HEAD_ROW = 3
FIRST = 4
ROW_H = 24
SEP_H = 8
YEAR_W = 13
LOW_LIMIT = 0.8
HIGH_LIMIT = 1.0
NAVY_HEX = "1F4E79"


def solid(argb):
    return PatternFill(fill_type="solid", start_color=argb, end_color=argb)


DARK = solid("FF" + NAVY_HEX)
LIGHT = solid("FFDDEBF7")
PALE = solid("FFF4F7FB")
WHITE = solid("FFFFFFFF")
SEP_FILL = solid("FFE9ECEF")
GOOD_FILL = solid("FFE2F0D9")
LOW_FILL = solid("FFFCE4E4")
WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
BLACK_BOLD = Font(bold=True, color="FF000000")
NAME_FONT = Font(bold=True, color="FF1F2933")
TYPE_FONT = Font(color="FF2F5597")
LINK_FONT = Font(bold=True, underline="single", color="FF" + NAVY_HEX)
GREY_FONT = Font(color="FFBFBFBF")
GOOD_FONT = Font(color="FF375623")
LOW_FONT = Font(color="FFC00000")
TITLE_FONT = Font(bold=True, size=16, color="FFFFFFFF")
NOTE_FONT = Font(size=10, color="FF44546A")
SECTION_FONT = Font(bold=True, size=12, color="FF" + NAVY_HEX)
LINE = Side(style="thin", color="FFD0D7E2")
GROUP = Side(style="thin", color="FF7F9BBF")
BORDER = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
CENTER_FIT = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
LEFT_FIT = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)


def one_line(t):
    return re.sub(r"\s+", " ", t or "").strip()


def norm(s):
    return re.sub(r"\s+", "", s or "").replace("（", "(").replace("）", ")").replace("％", "%")


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def reason_of(x):
    if NA.match(x):
        return "N/A"
    if "尚未推出" in x:
        return "尚未推出"
    if "已停售" in x:
        return "已停售"
    if re.search(r"尚未達至|未有數字|保單周年日|保單週年日", x):
        return "未達第一個保單周年日"
    return ""


def clean_kind(k):
    if "的" in k:
        k = k.rsplit("的", 1)[1]
    return re.sub(r"^.*年度", "", k)


# ---------- 1. 讀取 ----------

def fetch():
    out = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                    args=["--disable-blink-features=AutomationControlled"])
        try:
            ctx = browser.new_context(locale="zh-HK", timezone_id="Asia/Hong_Kong",
                                      viewport={"width": 1366, "height": 900})
            for label, url, is_div in PAGES:
                page = ctx.new_page()
                tables = []
                for attempt in range(3):
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=60000)
                        page.wait_for_selector("table", timeout=45000, state="attached")
                        page.wait_for_timeout(4000)
                        tables = page.evaluate(JS_TABLES)
                    except Exception as e:
                        print("   %s 載入出錯：%s" % (label, one_line(str(e))[:200]))
                    if tables:
                        break
                if not tables:
                    print("%s：三次嘗試都讀不到表格，停止。" % label)
                    sys.exit(1)
                out.append({"label": label, "url": url, "is_div": is_div, "tables": tables,
                            "text": page.evaluate(JS_TEXT), "notes": page.evaluate(JS_NOTES)})
                page.close()
        finally:
            browser.close()
    return out


# ---------- 2. 解析 ----------

def col_label(parts, k):
    joined = " ".join(parts)
    m = re.search(r"(?<!\d)(\d{1,2})\s*(\+|或以上)?\s*[\(（]\s*([^\)）]*)[\)）]", joined)
    if m:
        num = m.group(1) + ("+" if m.group(2) else "")
        extra = one_line(m.group(3))
    else:
        m = re.search(r"(?<!\d)(\d{1,2})(?!\d)\s*(\+|或以上)?", joined)
        num = (m.group(1) + ("+" if m.group(2) else "")) if m else str(k)
        y = re.search(r"(?:19|20)\d{2}", joined)
        extra = y.group(0) if y else ""
    label = "保單年度%s" % num
    if extra:
        label += "\n" + extra
    return label, num


def pick(t, cats, is_div):
    kind = ""
    if is_div:
        heads = [c for r in t["grid"][:-1] for c in r]
        for s in heads + t["ctx"]:
            if len(s) <= 30:
                m = TYPE.search(s)
                if m:
                    kind = clean_kind(m.group(0))
                    break
    name = ""
    for s in t["ctx"]:
        if len(s) > NAME_MAX or s in cats or TYPE.search(s) or SKIP.search(s):
            continue
        if NAME.search(s) and CJK.search(s):
            name = s
            break
    return name, kind


def parse_page(pg, problems):
    label, is_div, tables = pg["label"], pg["is_div"], pg["tables"]
    print("\n=== %s ===" % label)
    desk = [i for i, t in enumerate(tables) if max([len(r) for r in t["raw"]] or [0]) >= 11]
    if not desk:
        print("找不到橫排表格，停止。")
        sys.exit(1)
    deskset = set(desk)
    cats = set(tables[i]["grid"][-1][0] for i in desk if tables[i]["grid"])
    rows, labels, reasons, seen = [], None, collections.Counter(), set()
    same = diff = shown = 0
    for i in desk:
        t = tables[i]
        tag = "%s 表格 %d" % (label, i + 1)
        width = max(len(r) for r in t["grid"])
        grid = [list(r) + [""] * (width - len(r)) for r in t["grid"]]
        if len(grid) < 2:
            problems.append("%s：只有 %d 行" % (tag, len(grid)))
            continue
        data = grid[-1]
        head = [r for r in grid[:-1] if len(set(r)) > 1]
        labs = []
        for c in range(1, width):
            parts = []
            for r in head:
                if r[c] and r[c] not in parts:
                    parts.append(r[c])
            labs.append(col_label(parts, c))
        nums = [n for _, n in labs]
        if len(set(nums)) != len(nums):
            problems.append("%s：年度欄重複 %s" % (tag, " ".join(nums)))
            continue
        if labels is None:
            labels = labs
        elif nums != [n for _, n in labels]:
            problems.append("%s：年度欄 %s 與第一個表格不同" % (tag, " ".join(nums)))
            continue

        a = t["raw"][-1][1:]
        if i + 1 < len(tables) and (i + 1) not in deskset:
            b = [r[-1] for r in tables[i + 1]["raw"][1:] if r]
            if [norm(x) for x in a] == [norm(x) for x in b]:
                same += 1
            else:
                diff += 1
                problems.append("%s：桌面版與手機版不同｜%s｜%s" % (tag, " ".join(a), " ".join(b)))
        else:
            problems.append("%s：沒有手機版表格可核對" % tag)

        vals = []
        for x in data[1:]:
            m = PCT.match(x)
            if m:
                vals.append(round(float(m.group(1)) / 100.0, 4))
                continue
            why = reason_of(x)
            if why:
                reasons[why] += 1
            else:
                problems.append("%s：無法辨認的值「%s」" % (tag, x[:30]))
            vals.append("/")

        name, kind = pick(t, cats, is_div)
        bad = False
        if not name:
            bad = True
            problems.append("%s：找不到產品名稱" % tag)
            name = "未知產品%d" % (i + 1)
        if not is_div:
            kind = "總現金價值"
        elif not kind:
            bad = True
            problems.append("%s：找不到紅利類型" % tag)
            kind = "未知類型"
        if (name, kind) in seen:
            bad = True
            problems.append("%s：「%s／%s」重複出現" % (tag, name, kind))
        seen.add((name, kind))
        if (bad and shown < 6) or i == desk[0]:
            shown += 1
            print("   表格 %d 的上文：%s" % (i + 1, " ‖ ".join(x[:200] for x in t["ctx"][:8])))
        rows.append([name, data[0], kind] + vals)

    if not rows or labels is None:
        print("沒有讀到任何資料，停止。")
        sys.exit(1)
    year = "未知年度"
    for lab, num in labels:
        y = re.search(r"(?:19|20)\d{2}", lab.split("\n", 1)[1] if "\n" in lab else "")
        if num == "1" and y:
            year = str(int(y.group(0)) + 1)
    if year == "未知年度":
        problems.append("%s：讀不到報告年度" % label)
    for m in re.finditer(r"(20\d{2})\s*年?度?\s*的?\s*報告年度", pg["text"]):
        if m.group(1) != year:
            problems.append("%s：頁面文字的報告年度 %s 與表頭推算的 %s 不同" % (label, m.group(1), year))
            break
    kinds = collections.Counter(r[2] for r in rows)
    if len(kinds) > 6:
        problems.append("%s：紅利類型有 %d 種，可能讀錯" % (label, len(kinds)))
    names = []
    for r in rows:
        if r[0] not in names:
            names.append(r[0])
    ns = numbers(rows)
    print("   表格 %d｜橫排 %d｜手機版核對：相同 %d、不同 %d" % (len(tables), len(desk), same, diff))
    print("   報告年度 %s｜年度欄 %s" % (year, " ".join(n for _, n in labels)))
    print("   產品 %d 個｜%d 行｜%d 個數值｜平均 %s｜最高 %s｜最低 %s" % (
        len(names), len(rows), len(ns), fmt_pct(mean_or_slash(ns)),
        fmt_pct(max(ns) if ns else "/"), fmt_pct(min(ns) if ns else "/")))
    print("   類別：%s" % "｜".join("%s × %d" % x for x in collections.Counter(r[1] for r in rows).items()))
    print("   類型：%s" % "｜".join("%s × %d" % x for x in kinds.items()))
    print("   沒有數字的格：%s" % "｜".join("%s × %d" % x for x in reasons.most_common()))
    print("   最長的產品名稱：%s" % max(names, key=len))
    low = [r for r in rows if any(is_num(v) and v < LOW_LIMIT for v in r[NF:])]
    print("   有低於 80%% 的行 %d 行：" % len(low))
    for r in low[:20]:
        print("      %s／%s：%s" % (r[0], r[2], " ".join(
            "%.0f%%" % (v * 100) if is_num(v) else "/" for v in r[NF:])))
    notes = [s for s in pg["notes"] if not any(o != s and o in s for o in pg["notes"])][:5]
    print("   官網有關 N/A 的說明文字 %d 條：" % len(notes))
    for s in notes:
        print("      %s" % s[:300])
    return {"label": label, "is_div": is_div, "rows": rows, "labels": labels, "year": year,
            "reasons": reasons, "notes": notes}


def numbers(rows):
    return [v for r in rows for v in r[NF:] if is_num(v)]


def mean_or_slash(nums):
    return sum(nums) / len(nums) if nums else "/"


def fmt_pct(v):
    return "%.2f%%" % (v * 100) if is_num(v) else "/"


def make_groups(res):
    col = 2 if res["is_div"] else 1
    groups = [("所有分紅" if res["is_div"] else "所有產品", res["rows"])]
    keys = []
    for r in res["rows"]:
        if r[col] not in keys:
            keys.append(r[col])
    for k in keys:
        sub = [r for r in res["rows"] if r[col] == k]
        if numbers(sub):
            groups.append((k, sub))
    return groups


def build_summary(labels, groups):
    width = NF + len(labels)
    nums = [num for _, num in labels]
    head = ["所有年份平均值"] + ["持滿%s年平均值" % x for x in nums[1:]]
    cum = []
    for label, sub in groups:
        line = [label]
        for start in range(NF, width):
            line.append(mean_or_slash(
                [r[c] for r in sub for c in range(start, width) if is_num(r[c])]))
        cum.append(line)
    all_label, everything = groups[0]
    year_line = [all_label] + [mean_or_slash([r[c] for r in everything if is_num(r[c])])
                               for c in range(NF, width)]
    return {"head": head, "cum": cum,
            "year_head": ["第%s年" % x for x in nums], "year": [year_line]}


# ---------- 3. Excel（參考保誠格式） ----------

def write_row(ws, r, col, values):
    for i, v in enumerate(values):
        ws.cell(row=r, column=col + i, value=v)


def dark(cell):
    cell.fill = DARK
    cell.font = WHITE_BOLD


def light(cell):
    cell.fill = LIGHT
    cell.font = BLACK_BOLD


def heat(cell, v, base):
    if not is_num(v):
        cell.fill = base
        cell.font = GREY_FONT
    elif v >= HIGH_LIMIT:
        cell.fill = GOOD_FILL
        cell.font = GOOD_FONT
    elif v < LOW_LIMIT:
        cell.fill = LOW_FILL
        cell.font = LOW_FONT
    else:
        cell.fill = base


def pct_format(v):
    p = v * 100
    return "0%" if abs(p - round(p)) < 1e-6 else "0.0%"


def units(s):
    return sum(2 if ord(ch) > 255 else 1 for ch in str(s))


def fit_width(values, low, cap):
    longest = max([units(v) for v in values] or [0])
    return min(max(math.ceil(longest * 1.1) + 4, low), cap)


def add_sheet(wb, title, heading, notes, header, rows, summary, head_label, tab):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    last_col = len(header)
    widths = []
    for k in range(NF):
        low, cap = (30, 56) if k == 0 else (14, 30)
        widths.append(fit_width([header[k]] + [row[k] for row in rows], low, cap))

    ws.cell(row=1, column=1, value=heading)
    ws.cell(row=2, column=1, value=notes)
    for rr in (1, 2):
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
        ws.cell(row=rr, column=1).alignment = LEFT
    ws.cell(row=1, column=1).font = TITLE_FONT
    ws.cell(row=1, column=1).fill = DARK
    ws.cell(row=2, column=1).font = NOTE_FONT
    ws.cell(row=2, column=1).fill = PALE
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 66

    write_row(ws, HEAD_ROW, 1, header)
    for c in range(1, last_col + 1):
        cell = ws.cell(row=HEAD_ROW, column=c)
        dark(cell)
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[HEAD_ROW].height = 42

    rr = FIRST
    for i, row in enumerate(rows):
        first = i == 0 or rows[i - 1][0] != row[0]
        last = i == len(rows) - 1 or rows[i + 1][0] != row[0]
        if first and i > 0:
            ws.cell(row=rr, column=1).fill = SEP_FILL
            ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=last_col)
            ws.row_dimensions[rr].height = SEP_H
            rr += 1
        border = Border(left=LINE, right=LINE,
                        top=GROUP if first else LINE,
                        bottom=GROUP if last else LINE)
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c == 1:
                cell.font = NAME_FONT
            elif c <= NF:
                cell.font = TYPE_FONT
            else:
                if is_num(v):
                    cell.number_format = pct_format(v)
                heat(cell, v, WHITE)
        ws.row_dimensions[rr].height = ROW_H
        rr += 1
    data_end = rr - 1
    ws.auto_filter.ref = "A%d:%s%d" % (HEAD_ROW, get_column_letter(last_col), data_end)

    lc = NF
    r = data_end + 2
    ws.cell(row=r, column=lc, value="平均值摘要").font = SECTION_FONT
    ws.row_dimensions[r].height = ROW_H
    r += 1
    table_start = r
    write_row(ws, r, lc, [head_label] + summary["head"])
    r += 1
    value_rows = []
    for line in summary["cum"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    year_head_row = r
    write_row(ws, r, lc, ["當年平均值"] + summary["year_head"])
    r += 1
    for line in summary["year"]:
        write_row(ws, r, lc, line)
        value_rows.append((r, line))
        r += 1
    table_end = r - 1
    for row_i in range(table_start, table_end + 1):
        for c in range(lc, last_col + 1):
            cell = ws.cell(row=row_i, column=c)
            cell.alignment = CENTER
            cell.border = BORDER
    for c in range(lc, last_col + 1):
        dark(ws.cell(row=table_start, column=c))
    dark(ws.cell(row=year_head_row, column=lc))
    for c in range(lc + 1, last_col + 1):
        light(ws.cell(row=year_head_row, column=c))
    for row_i in (table_start, year_head_row):
        ws.row_dimensions[row_i].height = 32
    for vr, line in value_rows:
        label_cell = ws.cell(row=vr, column=lc)
        light(label_cell)
        label_cell.alignment = CENTER_FIT
        ws.row_dimensions[vr].height = ROW_H
        for k, v in enumerate(line[1:]):
            cell = ws.cell(row=vr, column=lc + 1 + k)
            cell.alignment = CENTER_FIT
            cell.number_format = "0.00%"
            heat(cell, v, WHITE)
    note_row = table_end + 1
    ws.cell(row=note_row, column=lc,
            value="「持滿N年平均值」＝保單年度N及之後所有數字的平均；"
                  "「當年平均值」＝該年度一欄所有數字的平均。平均值為固定數字，不會隨篩選變動。")
    ws.merge_cells(start_row=note_row, start_column=lc, end_row=note_row, end_column=last_col)
    ws.cell(row=note_row, column=lc).font = NOTE_FONT
    ws.cell(row=note_row, column=lc).alignment = LEFT
    ws.row_dimensions[note_row].height = 32

    for i in range(1, last_col + 1):
        w = widths[i - 1] if i <= NF else YEAR_W
        if i == lc:
            w = max(w, 18)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=FIRST, column=2)
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "%d:%d" % (HEAD_ROW, HEAD_ROW)


def add_overview(wb, year, today, sheets):
    ws = wb.create_sheet("概覽", 0)
    ws.sheet_properties.tabColor = "FFC000"
    ws.sheet_view.showGridLines = False
    last_col = 6
    for i, w in enumerate([36, 18, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def banner(r, value, font, fill, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = font
        cell.fill = fill
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    def section(r, value):
        ws.cell(row=r, column=1, value=value).font = SECTION_FONT
        ws.row_dimensions[r].height = ROW_H

    def table_head(r, values):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            dark(cell)
            cell.alignment = CENTER_FIT
            cell.border = BORDER
        ws.row_dimensions[r].height = ROW_H

    def body(r, values, pct_cols):
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
            cell.alignment = LEFT_FIT if c == 1 else CENTER_FIT
            cell.fill = WHITE
            if c in pct_cols:
                cell.number_format = pct_cols[c]
                heat(cell, v, WHITE)
        ws.row_dimensions[r].height = ROW_H

    def wide(r, value, height):
        cell = ws.cell(row=r, column=1, value=value)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
        cell.font = NOTE_FONT
        cell.alignment = LEFT
        ws.row_dimensions[r].height = height

    banner(1, "%s分紅表現概覽（報告年度 %s）" % (COMPANY, year), TITLE_FONT, DARK, 40)
    banner(2, "資料來源：中銀人壽網站｜更新日期：%s" % today, NOTE_FONT, PALE, ROW_H)

    r = 4
    section(r, "資料概況")
    r += 1
    table_head(r, ["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"])
    for title, res, groups in sheets:
        rows = res["rows"]
        nums = numbers(rows)
        low = sum(1 for v in nums if v < LOW_LIMIT) / len(nums) if nums else "/"
        r += 1
        body(r, [title, len({row[0] for row in rows}), len(rows), len(nums),
                 mean_or_slash(nums), low], {5: "0.00%"})
        ws.cell(row=r, column=6).number_format = "0.0%"
        link = ws.cell(row=r, column=1)
        link.font = LINK_FONT
        link.hyperlink = Hyperlink(ref=link.coordinate, location="'%s'!A1" % title)

    for title, res, groups in sheets:
        r += 2
        section(r, "%s：分組統計" % title)
        r += 1
        table_head(r, ["組別", "行數", "有數字的格數", "平均值", "最高", "最低"])
        for label, sub in groups:
            ns = numbers(sub)
            r += 1
            body(r, [label, len(sub), len(ns), mean_or_slash(ns),
                     max(ns) if ns else "/", min(ns) if ns else "/"],
                 {4: "0.00%", 5: "0%", 6: "0%"})
            ws.cell(row=r, column=1).font = BLACK_BOLD

    r += 2
    section(r, "顏色說明")
    samples = [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
               (0.79, "低於 80%：明顯低於銷售時的說明"),
               ("/", "沒有數字：官方網站顯示 N/A、尚未推出、已停售或未有數字")]
    for v, desc in samples:
        r += 1
        cell = ws.cell(row=r, column=1, value=v)
        cell.border = BORDER
        cell.alignment = CENTER_FIT
        if is_num(v):
            cell.number_format = "0%"
        heat(cell, v, WHITE)
        d = ws.cell(row=r, column=2, value=desc)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=last_col)
        d.alignment = LEFT_FIT
        ws.row_dimensions[r].height = ROW_H

    r += 2
    section(r, "「/」的原因（官方網站原來的寫法及格數）")
    for title, res, groups in sheets:
        r += 1
        wide(r, "%s：%s" % (title, "；".join(
            "%s %d 格" % x for x in res["reasons"].most_common())), ROW_H)
    notes = []
    for title, res, groups in sheets:
        for s in res["notes"]:
            if s not in notes:
                notes.append(s)
    for s in notes[:6]:
        r += 1
        wide(r, "官方網站原文：" + s, 48)

    r += 2
    section(r, "使用提示")
    tips = ["點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
            "工作表第 3 行有篩選箭嘴，可按產品、類別或類型篩選。",
            "向右或向下捲動時，標題列和產品名稱會固定不動。",
            "平均值摘要在工作表最底部。",
            "手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。"]
    for t in tips:
        r += 1
        wide(r, "・" + t, ROW_H)


def build_excel(year, results, now):
    today = now.strftime("%Y-%m-%d")
    wb = Workbook()
    wb.remove(wb.active)
    sheets = []
    print("\n=== 平均值摘要 ===")
    for res in results:
        groups = make_groups(res)
        summary = build_summary(res["labels"], groups)
        for line in summary["cum"]:
            print("   %s｜累計 %s：%s" % (res["label"], line[0],
                                        " ".join(fmt_pct(v) for v in line[1:])))
        title = str(year) + res["label"]
        if res["is_div"]:
            text = ("分紅實現率＝實際派發的非保證紅利 ÷ 銷售時所示的金額。"
                    "100% 即與銷售時一致，高於 100% 即比銷售時多。\n")
        else:
            text = ("總現金價值比率＝實際總現金價值 ÷ 銷售時所示的金額，"
                    "包括保證及非保證部分。\n")
        text += ("綠色＝100% 或以上；紅色＝低於 80%；「/」＝官方網站沒有數字。\n"
                 "資料來源：中銀人壽網站（報告年度 " + str(year) + "）｜更新日期：" + today)
        header = FIXED + [label for label, _ in res["labels"]]
        add_sheet(wb, title, "%s%s（報告年度 %s）" % (COMPANY, res["label"], year), text,
                  header, res["rows"], summary,
                  "分紅類型" if res["is_div"] else "產品類別",
                  "2E75B6" if res["is_div"] else "548235")
        sheets.append((title, res, groups))
    add_overview(wb, year, today, sheets)
    buf = io.BytesIO()
    wb.save(buf)
    content = buf.getvalue()
    os.makedirs(OUT_DIR, exist_ok=True)
    name = "%s分紅實現率（%s年）%s.xlsx" % (COMPANY, year, AUTO)
    path = os.path.join(OUT_DIR, name)
    with open(path, "wb") as fh:
        fh.write(content)
    print("\n=== Excel ===")
    print("   %s｜%d KB｜工作表：%s" % (path, len(content) // 1024, "、".join(wb.sheetnames)))
    return name, content


# ---------- 4. Notion ----------

def notion_upload(name, content, year):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = (os.environ.get("NOTION_DATABASE_URL") or os.environ.get("NOTION_DATABASE_ID") or "").strip()
    print("\n=== Notion ===")
    if os.environ.get("UPLOAD", "").strip().lower() != "true":
        print("   這次執行沒有勾選「上傳到 Notion」，只生成 Excel。")
        return
    if not token or not raw:
        print("   未設定 Notion Secrets，只生成 Excel。")
        return
    m = re.search(r"([0-9a-fA-F]{32})$", raw.split("?")[0].rstrip("/").replace("-", ""))
    if not m:
        print("   資料庫網址內找不到 ID，停止。")
        sys.exit(1)
    db_id = m.group(1)
    auth = {"Authorization": "Bearer " + token, "Notion-Version": NOTION_VERSION}

    def api(method, url, body=None):
        r = requests.request(method, "https://api.notion.com/v1" + url, headers=auth,
                             json=body, timeout=60)
        if r.status_code >= 300:
            raise RuntimeError("Notion %s %s：%s %s" % (
                method, url.split("?")[0], r.status_code, r.text[:300]))
        return r.json()

    sources = api("GET", "/databases/" + db_id).get("data_sources") or []
    if not sources:
        print("   資料庫沒有 data source，停止。")
        sys.exit(1)
    ds = sources[0]["id"]
    props = api("GET", "/data_sources/" + ds).get("properties", {})
    title_prop = next((k for k, v in props.items() if v.get("type") == "title"), None)
    if not title_prop:
        print("   找不到標題欄，停止。")
        sys.exit(1)

    page_id, cursor = None, None
    while True:
        body = {"page_size": 100}
        if cursor:
            body["start_cursor"] = cursor
        res = api("POST", "/data_sources/%s/query" % ds, body)
        for pg in res.get("results", []):
            if pg.get("in_trash") or pg.get("archived"):
                continue
            parts = pg.get("properties", {}).get(title_prop, {}).get("title", [])
            title = "".join(x.get("plain_text", "") for x in parts)
            if str(year) in title:
                page_id = pg["id"]
                print("   使用現有頁面：%s" % title)
                break
        if page_id or not res.get("has_more"):
            break
        cursor = res.get("next_cursor")
    if not page_id:
        pg = api("POST", "/pages", {
            "parent": {"type": "data_source_id", "data_source_id": ds},
            "properties": {title_prop: {"title": [{"text": {"content": str(year)}}]}}})
        page_id = pg["id"]
        print("   已建立頁面：%s" % year)

    old, cursor = [], None
    while True:
        url = "/blocks/%s/children?page_size=100" % page_id
        if cursor:
            url += "&start_cursor=" + cursor
        res = api("GET", url)
        for blk in res.get("results", []):
            if blk.get("type") != "file":
                continue
            f = blk.get("file", {})
            cap = "".join(x.get("plain_text", "") for x in f.get("caption", []))
            if AUTO in (f.get("name") or "") or AUTO in cap:
                old.append(blk["id"])
        if not res.get("has_more"):
            break
        cursor = res.get("next_cursor")

    up = api("POST", "/file_uploads", {"mode": "single_part", "filename": name,
                                       "content_type": XLSX})
    r = requests.post("https://api.notion.com/v1/file_uploads/%s/send" % up["id"],
                      headers=auth, timeout=120,
                      files={"file": (name, content, XLSX)})
    if r.status_code >= 300:
        raise RuntimeError("傳送檔案失敗：%s %s" % (r.status_code, r.text[:300]))
    api("PATCH", "/blocks/%s/children" % page_id, {"children": [{
        "object": "block", "type": "file",
        "file": {"type": "file_upload", "file_upload": {"id": up["id"]}}}]})
    for bid in old:
        api("DELETE", "/blocks/" + bid)
    print("   已上傳：%s（刪除舊的自動生成檔案 %d 個）" % (name, len(old)))


def main():
    now = datetime.now(timezone(timedelta(hours=8)))
    problems = []
    print("=== 讀取 ===")
    pages = fetch()
    results = [parse_page(pg, problems) for pg in pages]
    years = [r["year"] for r in results]
    if len(set(years)) > 1:
        problems.append("兩個頁面的報告年度不同：%s" % "、".join(years))
    year = years[0]
    name, content = build_excel(year, results, now)
    if problems:
        print("\n=== 發現 %d 個問題，不上傳 Notion ===" % len(problems))
        for x in problems[:40]:
            print("   ✗ %s" % x)
        sys.exit(1)
    notion_upload(name, content, year)
    print("\n=== 完成 ===")


if __name__ == "__main__":
    main()
