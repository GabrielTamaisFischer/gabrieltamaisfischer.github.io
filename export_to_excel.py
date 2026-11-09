import csv
import sys
from pathlib import Path
from collections import Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

folder = Path(r"c:\Users\jenny\OneDrive\桌面\歷史作業212")
csv_path = folder / "R14CDB_cleaned.csv"
output_xlsx = folder / "歷史作業212_分組統計結果.xlsx"

def bp_to_calendar(bp_age):
    cal_year = 1950 - bp_age
    if cal_year > 0:
        return f"西元 {cal_year} 年 (AD {cal_year})"
    else:
        bc_year = abs(cal_year) + 1
        return f"西元前 {bc_year} 年 ({bc_year} BC)"

# 讀取清洗後資料
with open(csv_path, "r", encoding="utf-8-sig") as f:
    rows = list(csv.DictReader(f))

total_samples = len(rows)

# 依區域分組資料
grouped = {}
region_samples = Counter()
region_sites = {}

for r in rows:
    reg = r['Region']
    age = int(r['C14Age'])
    err = int(r['C14Error'])
    site = f"{r['SiteName_jp']} ({r['SiteName_en']})"

    if reg not in grouped:
        grouped[reg] = []
        region_sites[reg] = Counter()
    
    grouped[reg].append({
        'age': age,
        'err': err,
        'site_jp': r['SiteName_jp'],
        'site_en': r['SiteName_en'],
        'pref': r['Prefecture'],
        'id': r['ID']
    })
    region_samples[reg] += 1
    region_sites[reg][site] += 1

region_order = ["Kyushu", "Shikoku", "Chugoku", "Kansai", "Chubu", "Kanto", "Tohoku"]
sorted_regions = sorted(grouped.keys(), key=lambda x: region_order.index(x) if x in region_order else 99)

# 建立 Excel 工作簿
wb = openpyxl.Workbook()
# 預設工作表
ws_overview = wb.active
ws_overview.title = "分組綜合統計總表"

# 工作表 2: 早晚年代統計
ws_dates = wb.create_sheet(title="區域早晚種植時間")

# 工作表 3: 採樣次數統計
ws_counts = wb.create_sheet(title="區域採集次數與遺址統計")

# 樣式定義
header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
title_font = Font(name="微軟正黑體", size=14, bold=True, color="1F4E79")
sub_font = Font(name="微軟正黑體", size=10, italic=True, color="595959")
data_font = Font(name="微軟正黑體", size=10)
bold_font = Font(name="微軟正黑體", size=10, bold=True)
total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
top_thick_bottom_double = Border(
    top=Side(style='thin', color='000000'),
    bottom=Side(style='double', color='000000')
)

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

# ==========================================
# 1. 填寫 Sheet 1: 分組綜合統計總表
# ==========================================
ws_overview["A1"] = "日本史前水稻農業各地區傳播時間與採樣統計彙整表"
ws_overview["A1"].font = title_font
ws_overview["A2"] = "資料來源：Science Advances (Crema et al., 2022) 炭化米 14C 定年數據庫 (R14CDB.csv 清洗版，共 296 筆樣本 / 132 處遺址)"
ws_overview["A2"].font = sub_font

overview_headers = [
    "地理區域 (Region)", "採集次數 (筆數)", "佔比 (%)", "涵蓋遺址數",
    "最早種植年代 (BP)", "最早西元換算", "最早出土遺址", "最早所在都道府縣",
    "最晚種植年代 (BP)", "最晚西元換算", "最晚出土遺址", "最晚所在都道府縣",
    "採集次數最多遺址 (採集次數)"
]

ws_overview.append([])  # Row 3 empty
ws_overview.append(overview_headers)  # Row 4 header

row_idx = 5
for reg in sorted_regions:
    samples = grouped[reg]
    samples.sort(key=lambda x: x['age'], reverse=True)
    earliest = samples[0]
    latest = samples[-1]
    
    cnt = len(samples)
    pct = (cnt / total_samples) * 100
    num_sites = len(region_sites[reg])
    top_site, top_cnt = region_sites[reg].most_common(1)[0]

    row_data = [
        reg,
        cnt,
        f"{pct:.2f}%",
        num_sites,
        f"{earliest['age']} ± {earliest['err']}",
        bp_to_calendar(earliest['age']),
        f"{earliest['site_jp']} ({earliest['site_en']})",
        earliest['pref'],
        f"{latest['age']} ± {latest['err']}",
        bp_to_calendar(latest['age']),
        f"{latest['site_jp']} ({latest['site_en']})",
        latest['pref'],
        f"{top_site} ({top_cnt}次)"
    ]
    ws_overview.append(row_data)
    row_idx += 1

# 合計列
total_row = [
    "全國合計", total_samples, "100.00%", sum(len(v) for v in region_sites.values()),
    "2710 ± 24", "西元前 761 年 (761 BC)", "宇木汲田貝塚 (佐賀縣)", "Saga",
    "1000 ± 15", "西元 950 年 (AD 950)", "鞠智城跡 (熊本縣)", "Kumamoto",
    "下老子笹川遺跡 (18次)"
]
ws_overview.append(total_row)

# ==========================================
# 2. 填寫 Sheet 2: 區域早晚種植時間
# ==========================================
ws_dates["A1"] = "日本各地區炭化米出土最早與最晚年代統計"
ws_dates["A1"].font = title_font
ws_dates["A2"] = "註：西元年換算公式為 1950 - C14Age（若小於等於 0 則為西元前，已消除西元 0 年差）"
ws_dates["A2"].font = sub_font

dates_headers = [
    "地理區域 (Region)", "樣本數",
    "最早年代 (BP)", "最早誤差 (±)", "最早西元年", "最早出土遺址 (日文/英文)", "最早所在縣",
    "最晚年代 (BP)", "最晚誤差 (±)", "最晚西元年", "最晚出土遺址 (日文/英文)", "最晚所在縣"
]
ws_dates.append([])
ws_dates.append(dates_headers)

for reg in sorted_regions:
    samples = grouped[reg]
    samples.sort(key=lambda x: x['age'], reverse=True)
    earliest = samples[0]
    latest = samples[-1]

    ws_dates.append([
        reg,
        len(samples),
        earliest['age'],
        earliest['err'],
        bp_to_calendar(earliest['age']),
        f"{earliest['site_jp']} ({earliest['site_en']})",
        earliest['pref'],
        latest['age'],
        latest['err'],
        bp_to_calendar(latest['age']),
        f"{latest['site_jp']} ({latest['site_en']})",
        latest['pref']
    ])

# ==========================================
# 3. 填寫 Sheet 3: 區域採集次數與遺址統計
# ==========================================
ws_counts["A1"] = "日本各地區炭化米資料採集次數與考古遺址統計"
ws_counts["A1"].font = title_font
ws_counts["A2"] = "採樣數反映各地區考古研究發掘密度與水稻農耕中心聚落分布"
ws_counts["A2"].font = sub_font

counts_headers = [
    "地理區域 (Region)", "採集次數 (樣本數)", "全國佔比 (%)", "涵蓋遺址數量",
    "平均每遺址採集次數", "採樣最多的代表遺址", "該遺址採樣次數"
]
ws_counts.append([])
ws_counts.append(counts_headers)

for reg in sorted_regions:
    cnt = len(grouped[reg])
    pct = (cnt / total_samples) * 100
    num_sites = len(region_sites[reg])
    top_site, top_cnt = region_sites[reg].most_common(1)[0]
    avg_per_site = cnt / num_sites if num_sites else 0

    ws_counts.append([
        reg,
        cnt,
        f"{pct:.2f}%",
        num_sites,
        round(avg_per_site, 2),
        top_site,
        top_cnt
    ])

ws_counts.append([
    "全國合計", total_samples, "100.00%", sum(len(v) for v in region_sites.values()),
    round(total_samples / 132, 2), "下老子笹川遺跡", 18
])

# ==========================================
# 4. 美化所有工作表
# ==========================================
for ws in [ws_overview, ws_dates, ws_counts]:
    ws.views.sheetView[0].showGridLines = True
    
    # 標題行格式化 (Row 4)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=4, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    # 內容列格式化 (Row 5 to max_row)
    for row in range(5, ws.max_row + 1):
        is_total = (row == ws.max_row and ws in [ws_overview, ws_counts])
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            if is_total:
                cell.font = bold_font
                cell.fill = total_fill
                cell.border = top_thick_bottom_double
            else:
                cell.font = data_font
            
            # 對齊處理
            if isinstance(cell.value, (int, float)):
                cell.alignment = align_right
            elif isinstance(cell.value, str) and (cell.value.endswith('%') or cell.value in sorted_regions or cell.value == "全國合計"):
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # 自動調整欄寬
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            # 針對中文字元計算適當寬度
            length = sum(2 if ord(char) > 127 else 1 for char in val_str)
            if cell.row in [1, 2]: # 跳過主標題欄長
                continue
            if length > max_len:
                max_len = length
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# 儲存 Excel
wb.save(output_xlsx)
print(f"成功產出 Excel 檔案：{output_xlsx}")
