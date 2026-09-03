# -*- coding: utf-8 -*-
import json, openpyxl

PATH = r"analysis\output\茂物高层建筑清单.xlsx"
W,S,E,N = 106.76, -6.64, 106.85, -6.54

wb = openpyxl.load_workbook(PATH, read_only=True, data_only=True)
sheets = wb.sheetnames

res = {"city":"茂物","anomalies":[]}
ano = res["anomalies"]

# 1 sheets
expected_sheets = ["高层清单","地标对照","说明"]
sheets_ok = sheets == expected_sheets
if not sheets_ok:
    ano.append(f"sheet名称/数量不符: 实际={sheets}")

ws = wb[sheets[0]]
rows = list(ws.iter_rows(values_only=True))
header = list(rows[0]) if rows else []
data = rows[1:] if len(rows)>1 else []
# 去掉全空尾行
data = [r for r in data if any(c is not None and str(c).strip()!="" for c in r)]

# 2 columns
expected_cols = ["序号","楼宇名称","名称来源","纬度","经度","高度(米)","高度来源","层数","层数来源","高度分档","用途分类","建筑面积","备注"]
columns_ok = (len(header)==13) and ([(c if c is None else str(c).strip()) for c in header]==expected_cols)
if not columns_ok:
    ano.append(f"主sheet列名/列数不符: 实际列={header}")

# 列索引按预期映射(若列名匹配)，否则按位置
def col_idx(name):
    try:
        return expected_cols.index(name)
    except ValueError:
        return None

idx_name = 1; idx_lat=3; idx_lon=4; idx_h=5; idx_band=9

count = len(data)
res["count"]=count

# 4 tier_sum
from collections import Counter
band_counter = Counter()
for r in data:
    b = r[idx_band] if len(r)>idx_band else None
    band_counter[b]+=1
tier_sum = sum(band_counter.values())
tier_sum_ok = (tier_sum==count)

def to_float(x):
    try:
        return float(x)
    except:
        return None

# 5 band_consistent + 6 all_height_ge32
def expected_band(h):
    # 返回该高度应落入的档(取满足的最高档)
    if h>=80: return ">=80"
    if h>=64: return ">=64"
    if h>=48: return ">=48"
    if h>=40: return ">=40"
    if h>=32: return ">=32"
    return None

def norm_band(b):
    if b is None: return None
    s=str(b).replace("≥",">=").replace("米","").replace("m","").strip()
    # 提取数字
    import re
    m=re.search(r"(\d+)",s)
    if m:
        return ">="+m.group(1)
    return s

inconsistent=0
inconsistent_examples=[]
below32=0
heights=[]
for r in data:
    h=to_float(r[idx_h]) if len(r)>idx_h else None
    b=norm_band(r[idx_band]) if len(r)>idx_band else None
    if h is not None:
        heights.append(h)
        if h<32: below32+=1
        eb=expected_band(h)
        if eb is None:
            inconsistent+=1
        elif b!=eb:
            inconsistent+=1
            if len(inconsistent_examples)<8:
                inconsistent_examples.append({"name":r[idx_name] if len(r)>idx_name else None,"h":h,"band":r[idx_band] if len(r)>idx_band else None,"expect":eb})
    else:
        inconsistent+=1
band_consistent = (inconsistent==0)
all_height_ge32 = (below32==0)
if inconsistent>0:
    ano.append(f"高度分档不自洽行数={inconsistent}, 示例={inconsistent_examples}")
if below32>0:
    ano.append(f"高度<32米行数={below32}")
if heights:
    hmax=max(heights)
    if hmax>250:
        big=[(r[idx_name] if len(r)>idx_name else None, to_float(r[idx_h])) for r in data if (to_float(r[idx_h]) or 0)>250]
        ano.append(f"高度>250米可疑(二三线城市): {big}")

# 7 coords bbox
inb=0; total_coord=0
for r in data:
    lat=to_float(r[idx_lat]) if len(r)>idx_lat else None
    lon=to_float(r[idx_lon]) if len(r)>idx_lon else None
    if lat is not None and lon is not None:
        total_coord+=1
        if (S-0.05)<=lat<=(N+0.05) and (W-0.05)<=lon<=(E+0.05):
            inb+=1
coords_pct = round(inb/total_coord*100,1) if total_coord else 0.0
if coords_pct<80:
    ano.append(f"坐标落在bbox(±0.05)内占比偏低={coords_pct}% ({inb}/{total_coord})")

# 8 landmark
landmark_rows=0; landmark_has_ctbuh=False
if "地标对照" in sheets:
    lws=wb["地标对照"]
    lrows=list(lws.iter_rows(values_only=True))
    lhdr=lrows[0] if lrows else []
    ldata=[r for r in lrows[1:] if any(c is not None and str(c).strip()!="" for c in r)]
    landmark_rows=len(ldata)
    blob=" ".join(str(c) for r in lrows for c in r if c is not None).upper()
    landmark_has_ctbuh = "CTBUH" in blob
    if landmark_rows==0:
        ano.append("地标对照sheet为空")
    if not landmark_has_ctbuh:
        ano.append("地标对照sheet未含CTBUH来源")

# 9 disclosure
disclosure_ok=False
if "说明" in sheets:
    dws=wb["说明"]
    dblob=" ".join(str(c) for r in dws.iter_rows(values_only=True) for c in r if c is not None)
    disclosure_ok = any(k in dblob for k in ["下限","估算","捕获率"])
    if not disclosure_ok:
        ano.append("说明sheet缺诚实标注关键词(下限/估算/捕获率)")

# 10 duplicates
seen=Counter()
for r in data:
    key=(r[idx_name] if len(r)>idx_name else None, r[idx_lat] if len(r)>idx_lat else None, r[idx_lon] if len(r)>idx_lon else None)
    seen[key]+=1
duplicate_rows=sum(c-1 for c in seen.values() if c>1)
if duplicate_rows>0:
    ano.append(f"楼名+坐标完全重复行数={duplicate_rows}")

# 11 named
named_count=sum(1 for r in data if (len(r)>idx_name and r[idx_name] is not None and str(r[idx_name]).strip()!=""))

checks={
 "sheets_ok":sheets_ok,
 "columns_ok":columns_ok,
 "tier_sum_ok":tier_sum_ok,
 "band_consistent":band_consistent,
 "all_height_ge32":all_height_ge32,
 "coords_in_bbox_pct":coords_pct,
 "landmark_rows":landmark_rows,
 "landmark_has_ctbuh":landmark_has_ctbuh,
 "disclosure_ok":disclosure_ok,
 "duplicate_rows":duplicate_rows,
 "named_count":named_count,
}
res["checks"]=checks

format_ok = sheets_ok and columns_ok and tier_sum_ok and band_consistent and all_height_ge32
severe = (not landmark_has_ctbuh and landmark_rows>0) or landmark_rows==0 or (not disclosure_ok) or duplicate_rows>0 or (heights and max(heights)>250)
res["pass"] = bool(format_ok and not severe)

res["band_counter"]={str(k):v for k,v in band_counter.items()}
res["named_count"]=named_count
print(json.dumps(res,ensure_ascii=True))
