"""
私募基金净值分析工具 - Mac App 版本
使用 Tkinter 构建，可打包为 .app 直接双击运行
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import math
from datetime import datetime, date
from pathlib import Path
import os
import sys


# ============== 配置区 ==============
RF_ANNUAL = 1.82  # 年化无风险利率（%）
SUPPORTED_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
# ============== 配置区 ==============


def parse_nav(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).strip().replace(",", ""))
    except ValueError:
        return None


def parse_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return datetime.strptime(str(val).strip()[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def get_nav_on_or_before(data, target_date):
    result = None
    for record in data:
        if len(record) < 2:
            continue
        d = record[0]
        if d <= target_date:
            result = record
        else:
            break
    return result


def shift_years(base_date, years):
    try:
        return date(base_date.year - years, base_date.month, base_date.day)
    except ValueError:
        return date(base_date.year - years, 2, 28)


def get_trailing_start_record(data, target_date):
    start_record = get_nav_on_or_before(data, target_date)
    if start_record:
        return start_record, False
    return data[0], True


def normalize_header(value):
    return str(value or "").strip().replace("\n", "").replace(" ", "")


def find_column(header, predicate):
    for idx, title in enumerate(header):
        if predicate(title):
            return idx
    return None


def parse_fund_rows(rows, fallback_name, sheet_name):
    header_row_idx = None
    for i, row in enumerate(rows):
        joined = " ".join(str(c) if c else "" for c in row)
        if "净值日期" in joined and "单位净值" in joined:
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("未找到包含'净值日期'和'单位净值'的表头行")

    header = [normalize_header(c) for c in rows[header_row_idx]]
    date_col = find_column(header, lambda h: "净值日期" in h)
    nav_col = find_column(header, lambda h: h == "单位净值")
    if nav_col is None:
        nav_col = find_column(header, lambda h: "单位净值" in h and "累计" not in h)
    name_col = find_column(header, lambda h: "产品名称" in h)
    cum_nav_col = find_column(header, lambda h: "累计净值" in h or "累计单位净值" in h or ("累计" in h and "净值" in h))
    if date_col is None or nav_col is None:
        raise ValueError(f"表头中未找到净值日期或单位净值列: {header}")

    fund_name = fallback_name
    if name_col is not None:
        for row in rows[header_row_idx + 1:]:
            if len(row) > name_col and row[name_col]:
                fund_name = str(row[name_col]).strip()
                break

    data = []
    for row in rows[header_row_idx + 1:]:
        if len(row) <= max(date_col, nav_col):
            continue
        d = parse_date(row[date_col])
        nav = parse_nav(row[nav_col])
        cum_nav = None
        if cum_nav_col is not None and len(row) > cum_nav_col:
            cum_nav = parse_nav(row[cum_nav_col])
        if d and nav is not None and nav > 0:
            data.append((d, nav, cum_nav))
    data.sort(key=lambda x: x[0])
    if not data:
        raise ValueError("已找到表头，但未解析出有效净值数据，请确认日期和单位净值列有数值")

    return fund_name, data, sheet_name


def load_fund_data(filepath):
    ext = Path(filepath).suffix.lower()
    if ext not in SUPPORTED_EXCEL_EXTENSIONS:
        raise ValueError("当前版本支持 .xlsx/.xlsm 文件；如果是 .xls，请先另存为 .xlsx 后再上传")

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    errors = []
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            try:
                fund_name, data, _ = parse_fund_rows(rows, Path(filepath).stem, ws.title)
                return fund_name, data
            except ValueError as exc:
                errors.append(f"{ws.title}: {exc}")
    finally:
        wb.close()

    sheet_names = ", ".join(wb.sheetnames)
    raise ValueError(
        "无法读取净值数据。请确认某个工作表里有'净值日期'和'单位净值'两列。"
        f"\n已检查工作表: {sheet_names}"
        f"\n详细原因: {'; '.join(errors)}"
    )


def calc_annualized_volatility(daily_returns):
    if len(daily_returns) < 2:
        return 0.0
    mean = sum(daily_returns) / len(daily_returns)
    variance = sum((r - mean) ** 2 for r in daily_returns) / (len(daily_returns) - 1)
    return math.sqrt(variance) * math.sqrt(252) * 100


def calc_sharpe_ratio(ann_return, ann_volatility, rf):
    if ann_volatility == 0:
        return 0.0
    return (ann_return - rf) / ann_volatility


def calc_sortino_ratio(daily_returns, rf_daily, rf_annual, ann_return):
    if len(daily_returns) < 2:
        return 0.0
    squared_diffs = [min(r - rf_daily, 0) ** 2 for r in daily_returns]
    downside_vol = math.sqrt(sum(squared_diffs) / len(daily_returns)) * math.sqrt(252) * 100
    if downside_vol == 0:
        return 0.0
    return (ann_return - rf_annual) / downside_vol


def calc_calmar_ratio(ann_return, max_drawdown):
    if max_drawdown == 0:
        return 0.0
    return ann_return / abs(max_drawdown)


def calc_metrics(data, start_nav_record=None, rf_annual=RF_ANNUAL):
    if not data:
        return {}

    if start_nav_record:
        first_nav = start_nav_record[1]
        start_date = start_nav_record[0]
        all_data = [start_nav_record] + data
    else:
        first_nav = data[0][1]
        start_date = data[0][0]
        all_data = data

    last_nav = data[-1][1]
    end_date = data[-1][0]
    cum_nav = data[-1][2] if len(data[-1]) > 2 else None
    days = (end_date - start_date).days

    total_return = (last_nav / first_nav - 1) * 100
    ann_return = ((last_nav / first_nav) ** (365.0 / days) - 1) * 100 if days > 0 else 0.0

    peak = first_nav
    max_dd = 0.0
    for record in all_data:
        nav = record[1]
        if nav > peak:
            peak = nav
        dd = (peak - nav) / peak
        if dd > max_dd:
            max_dd = dd

    daily_returns = []
    for i in range(1, len(all_data)):
        prev_nav = all_data[i - 1][1]
        curr_nav = all_data[i][1]
        if prev_nav > 0:
            daily_returns.append(curr_nav / prev_nav - 1)

    ann_vol = calc_annualized_volatility(daily_returns)
    sharpe = calc_sharpe_ratio(ann_return, ann_vol, rf_annual)
    rf_daily = (1 + rf_annual / 100) ** (1.0 / 252) - 1
    sortino = calc_sortino_ratio(daily_returns, rf_daily, rf_annual, ann_return)
    calmar = calc_calmar_ratio(ann_return, -max_dd * 100)

    return {
        "start_date": start_date,
        "end_date": end_date,
        "days": days,
        "start_nav": first_nav,
        "end_nav": last_nav,
        "cum_nav": cum_nav,
        "total_return": total_return,
        "ann_return": ann_return,
        "max_drawdown": -max_dd * 100,
        "ann_volatility": ann_vol,
        "sharpe_ratio": sharpe,
        "sortino_ratio": sortino,
        "calmar_ratio": calmar,
    }


def calc_one_year_return(data):
    if not data:
        return None

    latest_date = data[-1][0]
    one_year_ago = shift_years(latest_date, 1)

    start_record, _ = get_trailing_start_record(data, one_year_ago)

    end_record = data[-1]

    if start_record[0] == end_record[0]:
        return None

    one_year_return = (end_record[1] / start_record[1] - 1) * 100
    return {
        "start_date": start_record[0],
        "end_date": end_record[0],
        "start_nav": start_record[1],
        "end_nav": end_record[1],
        "one_year_return": one_year_return,
        "cum_nav": end_record[2] if len(end_record) > 2 else None
    }


def build_trailing_report(fund_name, data, rf_annual):
    results = []
    if not data:
        return results

    latest_date = data[-1][0]
    for years in (1, 2, 3):
        target_date = shift_years(latest_date, years)
        start_record, insufficient = get_trailing_start_record(data, target_date)
        if start_record[0] == latest_date:
            continue

        period_data = [record for record in data if start_record[0] < record[0] <= latest_date]
        if not period_data:
            continue

        m = calc_metrics(period_data, start_nav_record=start_record, rf_annual=rf_annual)
        base_label = "近一年" if years == 1 else f"近{years}年"
        m["label"] = f"{base_label}(不足{years}年)" if insufficient else base_label
        m["fund"] = fund_name
        m["note"] = f"数据不足{years}年，按可用区间计算" if insufficient else ""
        results.append(m)

    return results


def build_report(fund_name, data, rf_annual, use_prev_year_end=True):
    results = build_trailing_report(fund_name, data, rf_annual)
    year = data[-1][0].year

    for y in range(year, 2020, -1):
        if use_prev_year_end:
            prev_dec31 = get_nav_on_or_before(data, date(y - 1, 12, 31))
        else:
            prev_dec31 = None

        period_data = [(d, n, c) for d, n, c in data if date(y - 1, 12, 31) < d <= date(y, 12, 31)]
        if not period_data:
            continue

        if prev_dec31 and use_prev_year_end:
            m = calc_metrics(period_data, start_nav_record=prev_dec31, rf_annual=rf_annual)
        else:
            m = calc_metrics(period_data, rf_annual=rf_annual)

        label = f"{y}年初至今" if y == year else f"{y}年"
        m["label"] = label
        m["fund"] = fund_name
        m["note"] = ""
        results.append(m)

    m = calc_metrics(data, rf_annual=rf_annual)
    m["label"] = "成立以来"
    m["fund"] = fund_name
    m["note"] = ""
    results.append(m)

    return results


def autosize_worksheet(ws):
    for column_cells in ws.columns:
        column_letter = openpyxl.utils.get_column_letter(column_cells[0].column)
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 30)


def save_results_workbook(filepath, fund_name, data, results, one_year_data):
    latest = data[-1]
    wb = openpyxl.Workbook()

    summary_ws = wb.active
    summary_ws.title = "汇总"
    summary_ws.append(["项目", "值"])
    summary_rows = [
        ["基金名称", fund_name],
        ["最新净值日期", str(latest[0])],
        ["单位净值", f"{latest[1]:.4f}"],
        ["累计净值", f"{latest[2]:.4f}" if latest[2] is not None else "无"],
        ["近一年收益率", f"{one_year_data['one_year_return']:+.2f}%" if one_year_data else "数据不足"],
        ["近一年起始日期", str(one_year_data["start_date"]) if one_year_data else "--"],
        ["近一年起始净值", f"{one_year_data['start_nav']:.4f}" if one_year_data else "--"],
    ]
    for row in summary_rows:
        summary_ws.append(row)

    metrics_ws = wb.create_sheet("业绩指标")
    headers = ["基金", "区间", "起始日期", "结束日期", "天数", "期初净值", "期末净值", "期末累计净值",
               "收益率%", "年化收益%", "最大回撤%",
               "年化波动率%", "Sharpe", "Sortino", "Calmar", "备注"]
    metrics_ws.append(headers)
    for m in results:
        metrics_ws.append([
            m["fund"],
            m["label"],
            str(m["start_date"]),
            str(m["end_date"]),
            m["days"],
            m["start_nav"],
            m["end_nav"],
            m["cum_nav"],
            m["total_return"],
            m["ann_return"],
            m["max_drawdown"],
            m["ann_volatility"],
            m["sharpe_ratio"],
            m["sortino_ratio"],
            m["calmar_ratio"],
            m.get("note", ""),
        ])

    header_fill = openpyxl.styles.PatternFill("solid", fgColor="D9EAF7")
    header_font = openpyxl.styles.Font(bold=True)
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        autosize_worksheet(ws)

    for row in metrics_ws.iter_rows(min_row=2, min_col=6, max_col=8):
        for cell in row:
            cell.number_format = "0.0000"
    for row in metrics_ws.iter_rows(min_row=2, min_col=9, max_col=12):
        for cell in row:
            cell.number_format = "0.00"
    for row in metrics_ws.iter_rows(min_row=2, min_col=13, max_col=15):
        for cell in row:
            cell.number_format = "0.0000"

    wb.save(filepath)


def find_result(results, label_prefix):
    for item in results:
        if item.get("label", "").startswith(label_prefix):
            return item
    return None


def get_product_compare_values(product):
    data = product["data"]
    latest = data[-1]
    results = product.get("results", [])
    near_one = find_result(results, "近一年")
    near_two = find_result(results, "近2年")
    near_three = find_result(results, "近3年")
    since = find_result(results, "成立以来")
    notes = [m.get("note", "") for m in (near_one, near_two, near_three) if m and m.get("note")]

    return {
        "fund": product["fund_name"],
        "file": product["filename"],
        "start_date": data[0][0],
        "latest_date": latest[0],
        "records": len(data),
        "latest_nav": latest[1],
        "latest_cum_nav": latest[2] if len(latest) > 2 else None,
        "near_one_return": near_one["total_return"] if near_one else None,
        "near_two_return": near_two["total_return"] if near_two else None,
        "near_three_return": near_three["total_return"] if near_three else None,
        "since_return": since["total_return"] if since else None,
        "since_max_drawdown": since["max_drawdown"] if since else None,
        "since_sharpe": since["sharpe_ratio"] if since else None,
        "note": "；".join(notes),
    }


def save_all_products_workbook(filepath, products):
    wb = openpyxl.Workbook()
    summary_ws = wb.active
    summary_ws.title = "产品对比"
    summary_headers = [
        "产品", "文件", "数据起点", "最新日期", "记录数", "最新单位净值", "最新累计净值",
        "近一年收益%", "近2年收益%", "近3年收益%", "成立以来收益%",
        "成立以来最大回撤%", "成立以来Sharpe", "备注"
    ]
    summary_ws.append(summary_headers)

    for product in products:
        row = get_product_compare_values(product)
        summary_ws.append([
            row["fund"],
            row["file"],
            str(row["start_date"]),
            str(row["latest_date"]),
            row["records"],
            row["latest_nav"],
            row["latest_cum_nav"],
            row["near_one_return"],
            row["near_two_return"],
            row["near_three_return"],
            row["since_return"],
            row["since_max_drawdown"],
            row["since_sharpe"],
            row["note"],
        ])

    metrics_ws = wb.create_sheet("业绩指标")
    metrics_headers = [
        "产品", "文件", "区间", "起始日期", "结束日期", "天数", "期初净值", "期末净值",
        "期末累计净值", "收益率%", "年化收益%", "最大回撤%", "年化波动率%",
        "Sharpe", "Sortino", "Calmar", "备注"
    ]
    metrics_ws.append(metrics_headers)
    for product in products:
        for m in product.get("results", []):
            metrics_ws.append([
                product["fund_name"],
                product["filename"],
                m["label"],
                str(m["start_date"]),
                str(m["end_date"]),
                m["days"],
                m["start_nav"],
                m["end_nav"],
                m["cum_nav"],
                m["total_return"],
                m["ann_return"],
                m["max_drawdown"],
                m["ann_volatility"],
                m["sharpe_ratio"],
                m["sortino_ratio"],
                m["calmar_ratio"],
                m.get("note", ""),
            ])

    header_fill = openpyxl.styles.PatternFill("solid", fgColor="D9EAF7")
    header_font = openpyxl.styles.Font(bold=True)
    for ws in wb.worksheets:
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        autosize_worksheet(ws)

    for row in summary_ws.iter_rows(min_row=2, min_col=6, max_col=7):
        for cell in row:
            cell.number_format = "0.0000"
    for row in summary_ws.iter_rows(min_row=2, min_col=8, max_col=12):
        for cell in row:
            cell.number_format = "0.00"
    for row in summary_ws.iter_rows(min_row=2, min_col=13, max_col=13):
        for cell in row:
            cell.number_format = "0.0000"

    for row in metrics_ws.iter_rows(min_row=2, min_col=7, max_col=9):
        for cell in row:
            cell.number_format = "0.0000"
    for row in metrics_ws.iter_rows(min_row=2, min_col=10, max_col=13):
        for cell in row:
            cell.number_format = "0.00"
    for row in metrics_ws.iter_rows(min_row=2, min_col=14, max_col=16):
        for cell in row:
            cell.number_format = "0.0000"

    wb.save(filepath)


# ============== Tkinter GUI ==============
class FundAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("私募基金净值分析工具")
        self.root.geometry("1280x760")
        self.root.minsize(1120, 680)
        self.root.configure(bg="#f4f6f8")

        self.fund_name = ""
        self.filepath = ""
        self.data = []
        self.products = []
        self.selected_index = None
        self.rf_annual = RF_ANNUAL
        self.use_prev_year_end = True
        self.results = []
        self.one_year_data = None

        self.setup_ui()

    def setup_ui(self):
        self.colors = {
            "bg": "#f4f6f8",
            "panel": "#ffffff",
            "navy": "#172033",
            "text": "#172033",
            "muted": "#667085",
            "line": "#d9dee8",
            "blue": "#0b3a75",
            "green": "#14532d",
            "red": "#b42318",
            "orange": "#7c2d12",
            "disabled": "#94a3b8",
        }

        try:
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Treeview", rowheight=28, font=("Arial", 11), borderwidth=0)
            style.configure("Treeview.Heading", font=("Arial", 11, "bold"), padding=(6, 8))
            style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", self.colors["text"])])
        except tk.TclError:
            pass

        header_frame = tk.Frame(self.root, bg=self.colors["navy"], height=78)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        tk.Label(
            header_frame,
            text="私募基金净值分析工具",
            font=("Arial", 22, "bold"),
            bg=self.colors["navy"],
            fg="white",
        ).pack(side="left", padx=24)
        tk.Label(
            header_frame,
            text="上传净值 Excel，查看产品摘要、最新净值和区间业绩指标",
            font=("Arial", 12),
            bg=self.colors["navy"],
            fg="#d1d7e0",
        ).pack(side="left", padx=(8, 0), pady=(6, 0))

        main_frame = tk.Frame(self.root, bg=self.colors["bg"], padx=18, pady=18)
        main_frame.pack(fill="both", expand=True)

        left_frame = tk.Frame(main_frame, bg=self.colors["panel"], padx=16, pady=16, width=292)
        left_frame.pack(side="left", fill="y", padx=(0, 16))
        left_frame.pack_propagate(False)

        tk.Label(
            left_frame,
            text="文件与参数",
            font=("Arial", 16, "bold"),
            bg=self.colors["panel"],
            fg=self.colors["text"],
        ).pack(anchor="w")

        tk.Label(
            left_frame,
            text="支持 .xlsx/.xlsm；表头需包含净值日期、单位净值，可选累计净值或累计单位净值。",
            font=("Arial", 10),
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            justify="left",
            wraplength=244,
        ).pack(anchor="w", pady=(6, 16))

        self.upload_btn = tk.Button(
            left_frame,
            text="批量上传净值 Excel 文件",
            command=self.upload_file,
            font=("Arial", 12, "bold"),
            bg=self.colors["blue"],
            fg="white",
            activebackground="#082f63",
            activeforeground="white",
            relief="flat",
            bd=0,
            highlightthickness=0,
            highlightbackground=self.colors["blue"],
            padx=14,
            pady=10,
            cursor="hand2",
        )
        self.upload_btn.pack(fill="x")

        self.file_label = tk.Label(
            left_frame,
            text="未选择文件",
            font=("Arial", 10),
            fg=self.colors["muted"],
            bg=self.colors["panel"],
            justify="left",
            wraplength=244,
        )
        self.file_label.pack(anchor="w", fill="x", pady=(10, 18))

        tk.Label(
            left_frame,
            text="产品列表",
            font=("Arial", 12, "bold"),
            bg=self.colors["panel"],
            fg=self.colors["text"],
        ).pack(anchor="w", pady=(0, 6))
        product_list_frame = tk.Frame(left_frame, bg=self.colors["panel"])
        product_list_frame.pack(fill="x", pady=(0, 16))
        self.product_listbox = tk.Listbox(
            product_list_frame,
            height=8,
            font=("Arial", 10),
            activestyle="none",
            relief="solid",
            bd=1,
            selectbackground="#dbeafe",
            selectforeground=self.colors["text"],
            exportselection=False,
        )
        self.product_listbox.pack(side="left", fill="both", expand=True)
        product_scrollbar = ttk.Scrollbar(product_list_frame, orient="vertical", command=self.product_listbox.yview)
        product_scrollbar.pack(side="right", fill="y")
        self.product_listbox.configure(yscrollcommand=product_scrollbar.set)
        self.product_listbox.bind("<<ListboxSelect>>", self.on_product_select)

        separator = tk.Frame(left_frame, bg=self.colors["line"], height=1)
        separator.pack(fill="x", pady=(0, 16))

        tk.Label(
            left_frame,
            text="计算设置",
            font=("Arial", 13, "bold"),
            bg=self.colors["panel"],
            fg=self.colors["text"],
        ).pack(anchor="w")

        tk.Label(
            left_frame,
            text="年化无风险利率 (%)",
            font=("Arial", 10),
            bg=self.colors["panel"],
            fg=self.colors["muted"],
        ).pack(anchor="w", pady=(12, 4))
        self.rf_entry = tk.Entry(left_frame, width=15, font=("Arial", 12), relief="solid", bd=1)
        self.rf_entry.insert(0, str(self.rf_annual))
        self.rf_entry.pack(anchor="w", fill="x", ipady=6)

        self.rf_var = tk.BooleanVar(value=True)
        self.rf_check = tk.Checkbutton(
            left_frame,
            text="年初基准使用前一年 12-31 净值",
            variable=self.rf_var,
            bg=self.colors["panel"],
            fg=self.colors["text"],
            activebackground=self.colors["panel"],
            font=("Arial", 10),
            anchor="w",
        )
        self.rf_check.pack(anchor="w", fill="x", pady=(10, 16))

        self.recalc_btn = tk.Button(
            left_frame,
            text="重新计算",
            command=self.recalculate,
            bg=self.colors["green"],
            fg="white",
            activebackground="#0f3f22",
            activeforeground="white",
            relief="flat",
            bd=0,
            highlightthickness=0,
            highlightbackground=self.colors["green"],
            font=("Arial", 11, "bold"),
            cursor="hand2",
            padx=14,
            pady=9,
        )
        self.recalc_btn.pack(fill="x", pady=(0, 8))

        self.export_btn = tk.Button(
            left_frame,
            text="导出 Excel 分析结果",
            command=self.export_results,
            bg=self.colors["orange"],
            fg="white",
            activebackground="#5f220d",
            activeforeground="white",
            disabledforeground="#f8fafc",
            relief="flat",
            bd=0,
            highlightthickness=0,
            highlightbackground=self.colors["orange"],
            font=("Arial", 11, "bold"),
            cursor="hand2",
            state="disabled",
            padx=14,
            pady=9,
        )
        self.export_btn.pack(fill="x")
        self.set_export_button_state(False)

        tk.Label(
            left_frame,
            text="提示：上传后右侧会直接显示计算结果；导出前会按当前参数重新计算。",
            font=("Arial", 10),
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            justify="left",
            wraplength=244,
        ).pack(side="bottom", anchor="w")

        right_frame = tk.Frame(main_frame, bg=self.colors["bg"])
        right_frame.pack(side="left", fill="both", expand=True)

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill="both", expand=True)
        self.detail_tab = tk.Frame(self.notebook, bg=self.colors["bg"], padx=2, pady=2)
        self.compare_tab = tk.Frame(self.notebook, bg=self.colors["bg"], padx=2, pady=2)
        self.notebook.add(self.detail_tab, text="单产品分析")
        self.notebook.add(self.compare_tab, text="产品对比")

        product_frame = tk.Frame(self.detail_tab, bg=self.colors["panel"], padx=18, pady=16)
        product_frame.pack(fill="x", pady=(0, 14))

        self.product_name_label = tk.Label(
            product_frame,
            text="请上传净值 Excel 文件",
            font=("Arial", 18, "bold"),
            bg=self.colors["panel"],
            fg=self.colors["text"],
            anchor="w",
            justify="left",
            wraplength=820,
        )
        self.product_name_label.pack(anchor="w", fill="x")

        self.product_meta_label = tk.Label(
            product_frame,
            text="上传后会在这里显示产品名称、数据区间、记录数和计算参数。",
            font=("Arial", 11),
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            anchor="w",
            justify="left",
            wraplength=920,
        )
        self.product_meta_label.pack(anchor="w", fill="x", pady=(6, 12))

        summary_grid = tk.Frame(product_frame, bg=self.colors["panel"])
        summary_grid.pack(fill="x")
        self.summary_labels = {}
        summary_items = [
            ("file", "文件"),
            ("period", "数据区间"),
            ("count", "记录数"),
            ("rf", "无风险利率"),
        ]
        for col, (key, title) in enumerate(summary_items):
            cell = tk.Frame(summary_grid, bg=self.colors["panel"])
            cell.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 16, 0))
            summary_grid.columnconfigure(col, weight=1)
            tk.Label(cell, text=title, font=("Arial", 10), bg=self.colors["panel"], fg=self.colors["muted"]).pack(anchor="w")
            self.summary_labels[key] = tk.Label(
                cell,
                text="--",
                font=("Arial", 11, "bold"),
                bg=self.colors["panel"],
                fg=self.colors["text"],
                anchor="w",
                justify="left",
                wraplength=200,
            )
            self.summary_labels[key].pack(anchor="w")

        metric_frame = tk.Frame(self.detail_tab, bg=self.colors["bg"])
        metric_frame.pack(fill="x", pady=(0, 14))
        self.metric_value_labels = {}
        self.metric_hint_labels = {}
        metric_items = [
            ("latest_nav", "最新单位净值"),
            ("cum_nav", "最新累计净值"),
            ("one_year", "近一年收益率"),
            ("latest_date", "最新净值日期"),
        ]
        for col, (key, title) in enumerate(metric_items):
            card = tk.Frame(metric_frame, bg=self.colors["panel"], padx=16, pady=14)
            card.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 12, 0))
            metric_frame.columnconfigure(col, weight=1, uniform="metric")
            tk.Label(card, text=title, font=("Arial", 11), bg=self.colors["panel"], fg=self.colors["muted"]).pack(anchor="w")
            self.metric_value_labels[key] = tk.Label(
                card,
                text="--",
                font=("Arial", 22, "bold"),
                bg=self.colors["panel"],
                fg=self.colors["text"],
                anchor="w",
            )
            self.metric_value_labels[key].pack(anchor="w", pady=(6, 2))
            self.metric_hint_labels[key] = tk.Label(
                card,
                text="上传后显示",
                font=("Arial", 10),
                bg=self.colors["panel"],
                fg=self.colors["muted"],
                anchor="w",
            )
            self.metric_hint_labels[key].pack(anchor="w")

        self.metrics_frame = tk.Frame(self.detail_tab, bg=self.colors["panel"], padx=14, pady=12)
        self.metrics_frame.pack(fill="both", expand=True)
        table_header = tk.Frame(self.metrics_frame, bg=self.colors["panel"])
        table_header.pack(fill="x", pady=(0, 8))
        tk.Label(
            table_header,
            text="区间业绩指标",
            font=("Arial", 15, "bold"),
            bg=self.colors["panel"],
            fg=self.colors["text"],
        ).pack(side="left")
        self.table_note_label = tk.Label(
            table_header,
            text="上传后显示逐年和成立以来指标",
            font=("Arial", 10),
            bg=self.colors["panel"],
            fg=self.colors["muted"],
        )
        self.table_note_label.pack(side="left", padx=(12, 0), pady=(4, 0))

        table_container = tk.Frame(self.metrics_frame, bg=self.colors["panel"])
        table_container.pack(fill="both", expand=True)

        self.tree_columns = [
            ("label", "区间", 120),
            ("start_date", "起始日期", 92),
            ("end_date", "结束日期", 92),
            ("start_nav", "期初单位净值", 112),
            ("end_nav", "期末单位净值", 112),
            ("cum_nav", "期末累计净值", 112),
            ("total_return", "区间收益%", 92),
            ("ann_return", "年化收益%", 92),
            ("max_drawdown", "最大回撤%", 92),
            ("ann_volatility", "年化波动率%", 100),
            ("sharpe_ratio", "Sharpe", 80),
            ("sortino_ratio", "Sortino", 80),
            ("calmar_ratio", "Calmar", 80),
            ("note", "备注", 170),
        ]
        columns = [col_id for col_id, _, _ in self.tree_columns]
        self.tree = ttk.Treeview(table_container, columns=columns, show="headings", height=14)

        for col_id, heading, width in self.tree_columns:
            self.tree.heading(col_id, text=heading)
            self.tree.column(col_id, width=width, minwidth=width, anchor="center", stretch=False)

        y_scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        x_scrollbar = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        self.tree.tag_configure("positive", foreground=self.colors["red"])
        self.tree.tag_configure("negative", foreground=self.colors["green"])

        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        table_container.rowconfigure(0, weight=1)
        table_container.columnconfigure(0, weight=1)

        compare_frame = tk.Frame(self.compare_tab, bg=self.colors["panel"], padx=14, pady=12)
        compare_frame.pack(fill="both", expand=True)
        compare_header = tk.Frame(compare_frame, bg=self.colors["panel"])
        compare_header.pack(fill="x", pady=(0, 8))
        tk.Label(
            compare_header,
            text="产品对比",
            font=("Arial", 15, "bold"),
            bg=self.colors["panel"],
            fg=self.colors["text"],
        ).pack(side="left")
        self.compare_note_label = tk.Label(
            compare_header,
            text="批量上传后显示所有产品核心指标",
            font=("Arial", 10),
            bg=self.colors["panel"],
            fg=self.colors["muted"],
        )
        self.compare_note_label.pack(side="left", padx=(12, 0), pady=(4, 0))

        compare_table_container = tk.Frame(compare_frame, bg=self.colors["panel"])
        compare_table_container.pack(fill="both", expand=True)
        self.compare_columns = [
            ("fund", "产品", 260),
            ("file", "文件", 150),
            ("start_date", "数据起点", 92),
            ("latest_date", "最新日期", 92),
            ("records", "记录数", 72),
            ("latest_nav", "单位净值", 82),
            ("latest_cum_nav", "累计净值", 82),
            ("near_one_return", "近一年%", 82),
            ("near_two_return", "近2年%", 82),
            ("near_three_return", "近3年%", 82),
            ("since_return", "成立以来%", 92),
            ("since_max_drawdown", "最大回撤%", 92),
            ("since_sharpe", "Sharpe", 78),
            ("note", "备注", 240),
        ]
        compare_ids = [col_id for col_id, _, _ in self.compare_columns]
        self.compare_tree = ttk.Treeview(compare_table_container, columns=compare_ids, show="headings", height=18)
        for col_id, heading, width in self.compare_columns:
            self.compare_tree.heading(col_id, text=heading)
            self.compare_tree.column(col_id, width=width, minwidth=width, anchor="center", stretch=False)
        compare_y = ttk.Scrollbar(compare_table_container, orient="vertical", command=self.compare_tree.yview)
        compare_x = ttk.Scrollbar(compare_table_container, orient="horizontal", command=self.compare_tree.xview)
        self.compare_tree.configure(yscrollcommand=compare_y.set, xscrollcommand=compare_x.set)
        self.compare_tree.grid(row=0, column=0, sticky="nsew")
        compare_y.grid(row=0, column=1, sticky="ns")
        compare_x.grid(row=1, column=0, sticky="ew")
        compare_table_container.rowconfigure(0, weight=1)
        compare_table_container.columnconfigure(0, weight=1)

    def upload_file(self):
        filepaths = filedialog.askopenfilenames(
            title="选择一个或多个净值 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if not filepaths:
            return

        loaded_products = []
        errors = []
        for filepath in filepaths:
            try:
                fund_name, data = load_fund_data(filepath)
                loaded_products.append({
                    "fund_name": fund_name,
                    "filepath": filepath,
                    "filename": Path(filepath).name,
                    "data": data,
                    "one_year_data": None,
                    "results": [],
                })
            except Exception as exc:
                errors.append(f"{Path(filepath).name}: {exc}")

        if not loaded_products:
            messagebox.showerror("错误", "没有成功读取任何文件。\n\n" + "\n".join(errors[:5]))
            return

        self.products = loaded_products
        self.selected_index = 0
        self.update_product_list()
        self.update_display()
        self.set_export_button_state(True)
        self.file_label.config(
            text=(
                f"已读取 {len(self.products)} 个产品 / {len(filepaths)} 个文件\n"
                f"当前：{self.products[0]['filename']}\n"
                f"{self.products[0]['data'][0][0]} 至 {self.products[0]['data'][-1][0]}"
            )
        )
        if len(self.products) > 1:
            self.notebook.select(self.compare_tab)

        if errors:
            messagebox.showwarning(
                "部分文件读取失败",
                f"成功读取 {len(self.products)} 个文件，失败 {len(errors)} 个。\n\n" + "\n".join(errors[:8])
            )

    def update_product_list(self):
        self.product_listbox.delete(0, tk.END)
        for idx, product in enumerate(self.products, start=1):
            label = f"{idx}. {product['fund_name']} ({product['filename']})"
            self.product_listbox.insert(tk.END, label)
        if self.products and self.selected_index is not None:
            self.product_listbox.selection_clear(0, tk.END)
            self.product_listbox.selection_set(self.selected_index)
            self.product_listbox.activate(self.selected_index)

    def on_product_select(self, _event=None):
        selection = self.product_listbox.curselection()
        if not selection:
            return
        self.selected_index = selection[0]
        self.update_display()
        self.notebook.select(self.detail_tab)

    def set_export_button_state(self, enabled):
        if enabled:
            self.export_btn.config(
                state="normal",
                bg=self.colors["orange"],
                activebackground="#5f220d",
                cursor="hand2",
            )
        else:
            self.export_btn.config(
                state="disabled",
                bg=self.colors["disabled"],
                activebackground=self.colors["disabled"],
                cursor="arrow",
            )

    def get_selected_product(self):
        if not self.products:
            return None
        if self.selected_index is None or self.selected_index >= len(self.products):
            self.selected_index = 0
        return self.products[self.selected_index]

    def parse_settings(self):
        try:
            self.rf_annual = float(self.rf_entry.get())
        except ValueError as exc:
            raise ValueError("无风险利率必须填写数字，例如 1.82") from exc
        self.use_prev_year_end = self.rf_var.get()

    def recalculate_products(self):
        self.parse_settings()
        for product in self.products:
            product["one_year_data"] = calc_one_year_return(product["data"])
            product["results"] = build_report(
                product["fund_name"],
                product["data"],
                self.rf_annual,
                self.use_prev_year_end,
            )

    def sync_current_product(self, product):
        self.fund_name = product["fund_name"]
        self.filepath = product["filepath"]
        self.data = product["data"]
        self.one_year_data = product.get("one_year_data")
        self.results = product.get("results", [])

    def format_nav(self, value):
        return "--" if value is None else f"{value:.4f}"

    def format_percent(self, value, signed=True):
        if value is None:
            return "--"
        return f"{value:+.2f}" if signed else f"{value:.2f}"

    def update_display(self):
        if not self.products:
            return

        self.recalculate_products()
        product = self.get_selected_product()
        if not product:
            return
        self.sync_current_product(product)

        latest = self.data[-1]

        self.product_name_label.config(text=self.fund_name)
        self.product_meta_label.config(
            text=(
                f"这是产品“{self.fund_name}”的净值分析结果。"
                f"当前为第 {self.selected_index + 1} / {len(self.products)} 个产品，"
                f"读取 {len(self.data)} 条净值记录，最新净值日期为 {latest[0]}。"
            )
        )
        self.summary_labels["file"].config(text=Path(self.filepath).name if self.filepath else "--")
        self.summary_labels["period"].config(text=f"{self.data[0][0]} 至 {self.data[-1][0]}")
        self.summary_labels["count"].config(text=f"{len(self.data)} 条")
        self.summary_labels["rf"].config(text=f"{self.rf_annual:.2f}%")

        self.metric_value_labels["latest_nav"].config(text=f"{latest[1]:.4f}", fg=self.colors["text"])
        self.metric_hint_labels["latest_nav"].config(text="单位净值，来自最新一条净值记录")

        if latest[2] is not None:
            self.metric_value_labels["cum_nav"].config(text=f"{latest[2]:.4f}", fg=self.colors["text"])
            self.metric_hint_labels["cum_nav"].config(text="累计单位净值，来自累计净值列")
        else:
            self.metric_value_labels["cum_nav"].config(text="无", fg=self.colors["muted"])
            self.metric_hint_labels["cum_nav"].config(text="文件中未识别到累计净值列")

        self.metric_value_labels["latest_date"].config(text=str(latest[0]), fg=self.colors["text"])
        self.metric_hint_labels["latest_date"].config(text=f"数据起点：{self.data[0][0]}")

        if self.one_year_data:
            one_year_return = self.one_year_data["one_year_return"]
            return_color = self.colors["red"] if one_year_return >= 0 else self.colors["green"]
            self.metric_value_labels["one_year"].config(text=f"{one_year_return:+.2f}%", fg=return_color)
            self.metric_hint_labels["one_year"].config(
                text=(
                    f"{self.one_year_data['start_date']} "
                    f"{self.one_year_data['start_nav']:.4f} 至 "
                    f"{self.one_year_data['end_nav']:.4f}"
                )
            )
        else:
            self.metric_value_labels["one_year"].config(text="数据不足", fg=self.colors["muted"])
            self.metric_hint_labels["one_year"].config(text="净值历史不足以形成区间对比")

        for item in self.tree.get_children():
            self.tree.delete(item)

        basis_text = "自然年度使用前一年 12-31 净值作年初基准" if self.use_prev_year_end else "自然年度使用当年首条净值作年初基准"
        self.table_note_label.config(text=f"含近1/2/3年、自然年度、成立以来；{basis_text}")

        for m in self.results:
            sign_return = f"{m['total_return']:+.2f}" if m['total_return'] >= 0 else f"{m['total_return']:.2f}"
            sign_ann = f"{m['ann_return']:+.2f}" if m['ann_return'] >= 0 else f"{m['ann_return']:.2f}"
            cum_nav_text = f"{m['cum_nav']:.4f}" if m.get("cum_nav") is not None else "--"
            values = [
                m["label"],
                str(m["start_date"]),
                str(m["end_date"]),
                f"{m['start_nav']:.4f}",
                f"{m['end_nav']:.4f}",
                cum_nav_text,
                sign_return,
                sign_ann,
                f"{m['max_drawdown']:.2f}",
                f"{m['ann_volatility']:.2f}",
                f"{m['sharpe_ratio']:.4f}",
                f"{m['sortino_ratio']:.4f}",
                f"{m['calmar_ratio']:.4f}",
                m.get("note", ""),
            ]
            row_tag = "positive" if m["total_return"] >= 0 else "negative"
            self.tree.insert("", "end", values=values, tags=(row_tag,))

        self.update_compare_display()
        self.file_label.config(
            text=(
                f"已读取 {len(self.products)} 个产品\n"
                f"当前：{product['filename']}\n"
                f"{self.data[0][0]} 至 {self.data[-1][0]}"
            )
        )

    def update_compare_display(self):
        for item in self.compare_tree.get_children():
            self.compare_tree.delete(item)

        if not self.products:
            self.compare_note_label.config(text="批量上传后显示所有产品核心指标")
            return

        self.compare_note_label.config(
            text=f"共 {len(self.products)} 个产品；近1/2/3年不足区间会按可用数据计算并在备注说明"
        )
        for product in self.products:
            row = get_product_compare_values(product)
            values = [
                row["fund"],
                row["file"],
                str(row["start_date"]),
                str(row["latest_date"]),
                row["records"],
                self.format_nav(row["latest_nav"]),
                self.format_nav(row["latest_cum_nav"]),
                self.format_percent(row["near_one_return"]),
                self.format_percent(row["near_two_return"]),
                self.format_percent(row["near_three_return"]),
                self.format_percent(row["since_return"]),
                self.format_percent(row["since_max_drawdown"], signed=False),
                self.format_nav(row["since_sharpe"]),
                row["note"],
            ]
            self.compare_tree.insert("", "end", values=values)

    def recalculate(self):
        if self.products:
            try:
                self.update_display()
            except Exception as e:
                messagebox.showerror("错误", f"重新计算失败：{str(e)}")

    def export_results(self):
        if not self.products:
            messagebox.showwarning("警告", "请先加载数据")
            return

        try:
            self.update_display()
        except Exception as e:
            messagebox.showerror("错误", f"导出前计算失败：{str(e)}")
            return

        filepath = filedialog.asksaveasfilename(
            title="导出结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="多产品净值分析结果.xlsx" if len(self.products) > 1 else f"{self.fund_name}_分析结果.xlsx"
        )
        if filepath:
            try:
                save_all_products_workbook(filepath, self.products)
                messagebox.showinfo("成功", f"结果已导出至：\n{filepath}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{str(e)}")


def main():
    root = tk.Tk()
    app = FundAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
