import csv
import sys
from pathlib import Path
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')

folder = Path(r"c:\Users\jenny\OneDrive\桌面\歷史作業212")
cleaned_csv = folder / "R14CDB_cleaned.csv"
raw_csv = folder / "R14CDB.csv"
output_csv = folder / "region_sampling_counts.csv"

# 1. 讀取清洗後的核心資料 (296 筆)
with open(cleaned_csv, "r", encoding="utf-8-sig") as f:
    cleaned_rows = list(csv.DictReader(f))

# 2. 讀取原始資料 (439 筆) 作為對照
with open(raw_csv, "r", encoding="utf-8") as f:
    raw_rows = list(csv.DictReader(f))

# 統計清洗後數據
cleaned_region_counts = Counter(r['Region'] for r in cleaned_rows)
raw_region_counts = Counter(r['Region'] for r in raw_rows)

# 統計每個區域內的遺址點位數量 (Unique Sites) 與每個遺址的採集次數
region_sites = {}
for r in cleaned_rows:
    reg = r['Region']
    site = f"{r['SiteName_jp']} ({r['SiteName_en']})"
    if reg not in region_sites:
        region_sites[reg] = Counter()
    region_sites[reg][site] += 1

region_order = ["Kyushu", "Shikoku", "Chugoku", "Kansai", "Chubu", "Kanto", "Tohoku"]
sorted_regions = sorted(cleaned_region_counts.keys(), key=lambda x: region_order.index(x) if x in region_order else 99)

total_cleaned = len(cleaned_rows)

print("=" * 85)
print(f"{'區域 (Region)':<10} | {'採集次數(筆數)':<12} | {'佔比 (%)':<8} | {'涵蓋遺址數':<10} | {'採集次數最多的遺址 (樣本數)'}")
print("=" * 85)

results = []
for reg in sorted_regions:
    cnt = cleaned_region_counts[reg]
    pct = (cnt / total_cleaned) * 100
    num_sites = len(region_sites[reg])
    top_site, top_cnt = region_sites[reg].most_common(1)[0]
    
    print(f"{reg:<10} | {cnt:<12} | {pct:>6.2f}%  | {num_sites:<10} | {top_site} ({top_cnt} 次)")
    
    results.append({
        "Region": reg,
        "Cleaned_SampleCount": cnt,
        "Percentage": f"{pct:.2f}%",
        "UniqueSitesCount": num_sites,
        "Raw_SampleCount": raw_region_counts.get(reg, 0),
        "TopSite": top_site,
        "TopSite_Count": top_cnt
    })

print("=" * 85)
print(f"{'合計 (Total)':<10} | {total_cleaned:<12} | 100.00%  | {sum(len(v) for v in region_sites.values()):<10} | 全國 132 個點位")
print("=" * 85)

# 儲存至 CSV
with open(output_csv, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=list(results[0].keys()))
    writer.writeheader()
    writer.writerows(results)

print(f"\n已匯出統計結果至檔案: {output_csv.name}")
