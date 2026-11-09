import csv
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

folder = Path(r"c:\Users\jenny\OneDrive\桌面\歷史作業212")
input_csv = folder / "R14CDB.csv"
output_valid = folder / "R14CDB_valid.csv"
output_paper = folder / "R14CDB_cleaned.csv"

with open(input_csv, "r", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    fieldnames = reader.fieldnames
    rows = list(reader)

total = len(rows)
print(f"原始資料筆數: {total}")

# 1. 基礎清理：僅保留 UseForAnalyses == 'TRUE' 且 C14Age 非 NA
valid_rows = [r for r in rows if r['UseForAnalyses'] == 'TRUE' and r['C14Age'] != 'NA']
print(f"基礎清理 (排除無效/標記FALSE樣本): 保留 {len(valid_rows)} 筆，剔除 {total - len(valid_rows)} 筆")

with open(output_valid, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(valid_rows)

# 2. 論文研究嚴格清理（依據 sciadv.adc9171.pdf 第 7 頁標準）：
# - UseForAnalyses == 'TRUE'
# - 排除北海道 (Region != 'Hokkaido')
# - 排除小於 1000 BP 之較現代樣本 (C14Age >= 1000)
paper_rows = []
removed_breakdown = {
    "UseForAnalyses is FALSE (非稻米/污染/現代)": 0,
    "Hokkaido 地區樣本": 0,
    "C14Age < 1000 BP (過近/非史前擴散期)": 0,
}

for r in rows:
    if r['UseForAnalyses'] != 'TRUE':
        removed_breakdown["UseForAnalyses is FALSE (非稻米/污染/現代)"] += 1
        continue
    if r['Region'] == 'Hokkaido':
        removed_breakdown["Hokkaido 地區樣本"] += 1
        continue
    try:
        age = int(r['C14Age'])
        if age < 1000:
            removed_breakdown["C14Age < 1000 BP (過近/非史前擴散期)"] += 1
            continue
    except ValueError:
        continue
    paper_rows.append(r)

print(f"\n論文標準清理 (保留彌生稻作傳播核心數據): 保留 {len(paper_rows)} 筆")
for reason, count in removed_breakdown.items():
    print(f"  - 剔除原因 [{reason}]: {count} 筆")

with open(output_paper, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(paper_rows)

print(f"\n已成功輸出檔案：")
print(f"1. {output_valid.name} (基礎清洗版，414 筆)")
print(f"2. {output_paper.name} (論文研究核心版，{len(paper_rows)} 筆，涵蓋 132 個遺址)")
