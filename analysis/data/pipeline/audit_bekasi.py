# -*- coding: utf-8 -*-
# 验收审计：勿加泗高层建筑清单.xlsx
import json, openpyxl

PATH = r"analysis\output\勿加泗高层建筑清单.xlsx"
W, S, E, N = 106.94, -6.3, 107.06, -6.18

wb = openpyxl.load_workbook(PATH, read_only=True, data_only=True)
sheet_names = wb.sheetnames
anomalies = []

# 1 sheets_ok
expected_sheets = ["高层清单", "地标对照", "说明"]
sheets_ok = sheet_names == expected_sheets
if not sheets_ok:
    anomalies.append(f"sheet名称/数量不符: 实际={sheet_names}")

ws = wb[sheet_names[0]]
rows = list(ws.iter_rows(values_only=True))
header = list(rows[0]) if rows else []
data = rows[1:] if len(rows) > 1 else []
# 去掉全空尾行
data = [r for r in data if any(c is not None and str(c).strip() != "" for c in r)]

# 2 columns_ok
expected_cols = ["序号","楼宇名称","名称来源","纬度","经度","高度(米)","高度来源","层数","层数来源","高度分档","用途分类","建筑面积","备注"]
header_clean = [str(h).strip() if h is not None else "" for h in header]
columns_ok = (len(header_clean) == 13) and (header_clean == expected_cols)
if not columns_ok:
    anomalies.append(f"列名不符: 实际({len(header_clean)}列)={header_clean}")

# 列索引
def idx(name):
    return expected_cols.index(name) if name in header_clean and header_clean==expected_cols else (header_clean.index(name) if name in header_clean else None)

i_name = idx("楼宇名称")
i_lat = idx("纬度")
i_lon = idx("经度")
i_h = idx("高度(米)")
i_band = idx("高度分档")

count = len(data)

# 4 tier_sum_ok
band_counts = {}
for r in data:
    b = r[i_band] if i_band is not None and i_band < len(r) else None
    b = str(b).strip() if b is not None else "(空)"
    band_counts[b] = band_counts.get(b, 0) + 1
tier_sum_ok = sum(band_counts.values()) == count

def to_f(v):
    try:
        return float(v)
    except:
        return None

# 5 band_consistent  口径4米/层： >=32=>8层档, >=40=>10, >=48=>12, >=64=>16, >=80=>20
def expected_band(h):
    if h >= 80: return 20
    if h >= 64: return 16
    if h >= 48: return 12
    if h >= 40: return 10
    if h >= 32: return 8
    return None

import re
def band_num(bandstr):
    # 从分档文本中提取期望的层档或高度阈值
    if bandstr is None: return None
    s = str(bandstr)
    nums = re.findall(r"\d+", s)
    return [int(x) for x in nums] if nums else []

band_inconsistent = 0
band_examples = []
for r in data:
    h = to_f(r[i_h]) if i_h is not None and i_h < len(r) else None
    bstr = r[i_band] if i_band is not None and i_band < len(r) else None
    if h is None:
        band_inconsistent += 1
        if len(band_examples) < 5: band_examples.append(f"高度非数值:{r[i_h] if i_h is not None else '?'}")
        continue
    exp_layer = expected_band(h)
    nums = band_num(bstr)
    # 分档文本可能是 "≥80" / "≥80(20层)" / "20层" 等。判定：取文本中的高度阈值(>=32)匹配
    # 期望高度阈值
    if h >= 80: exp_thr = 80
    elif h >= 64: exp_thr = 64
    elif h >= 48: exp_thr = 48
    elif h >= 40: exp_thr = 40
    elif h >= 32: exp_thr = 32
    else: exp_thr = None
    ok = False
    if exp_thr is not None and exp_thr in nums:
        ok = True
    elif exp_layer is not None and exp_layer in nums:
        ok = True
    if not ok:
        band_inconsistent += 1
        if len(band_examples) < 8:
            band_examples.append(f"h={h} 分档='{bstr}'")
band_consistent = (band_inconsistent == 0)
if band_inconsistent > 0:
    anomalies.append(f"分档不自洽行数={band_inconsistent}; 示例={band_examples}")

# 6 all_height_ge32
below32 = 0
heights = []
for r in data:
    h = to_f(r[i_h]) if i_h is not None and i_h < len(r) else None
    if h is not None:
        heights.append(h)
        if h < 32:
            below32 += 1
all_height_ge32 = (below32 == 0)
if below32 > 0:
    anomalies.append(f"高度<32米的行数={below32}")
# 高度异常大
big = [h for h in heights if h > 250]
if big:
    anomalies.append(f"高度>250米可疑行数={len(big)}; 值={sorted(big, reverse=True)[:5]}")

# 7 coords_in_bbox_pct
in_box = 0
valid_coord = 0
for r in data:
    lat = to_f(r[i_lat]) if i_lat is not None and i_lat < len(r) else None
    lon = to_f(r[i_lon]) if i_lon is not None and i_lon < len(r) else None
    if lat is None or lon is None:
        continue
    valid_coord += 1
    if (S-0.05) <= lat <= (N+0.05) and (W-0.05) <= lon <= (E+0.05):
        in_box += 1
coords_in_bbox_pct = round(100.0*in_box/count, 1) if count else 0.0
if count and (in_box/count) < 0.7:
    anomalies.append(f"坐标越界严重: 仅{coords_in_bbox_pct}%在bbox内")

# 8 landmark
ws_lm = wb[sheet_names[1]] if len(sheet_names) > 1 else None
landmark_rows = 0
landmark_has_ctbuh = False
if ws_lm is not None:
    lm_rows = list(ws_lm.iter_rows(values_only=True))
    lm_data = lm_rows[1:] if len(lm_rows) > 1 else []
    lm_data = [r for r in lm_data if any(c is not None and str(c).strip() != "" for c in r)]
    landmark_rows = len(lm_data)
    allcells = " ".join(str(c) for r in lm_rows for c in r if c is not None)
    landmark_has_ctbuh = ("CTBUH" in allcells.upper())
if landmark_rows == 0:
    anomalies.append("地标对照sheet为空")
if not landmark_has_ctbuh:
    anomalies.append("地标对照不含CTBUH来源")

# 9 disclosure_ok
ws_doc = wb[sheet_names[2]] if len(sheet_names) > 2 else None
disclosure_ok = False
if ws_doc is not None:
    doc_text = " ".join(str(c) for r in ws_doc.iter_rows(values_only=True) for c in r if c is not None)
    disclosure_ok = any(k in doc_text for k in ["下限", "估算", "捕获率"])
if not disclosure_ok:
    anomalies.append("说明sheet缺诚实标注关键词(下限/估算/捕获率)")

# 10 duplicate_rows
seen = {}
dup = 0
for r in data:
    nm = r[i_name] if i_name is not None and i_name < len(r) else None
    lat = r[i_lat] if i_lat is not None and i_lat < len(r) else None
    lon = r[i_lon] if i_lon is not None and i_lon < len(r) else None
    key = (str(nm).strip() if nm else "", lat, lon)
    seen[key] = seen.get(key, 0) + 1
for k, v in seen.items():
    if v > 1:
        dup += (v - 1)
if dup > 0:
    anomalies.append(f"楼名+坐标完全重复行数={dup}")

# 11 named_count
named_count = 0
for r in data:
    nm = r[i_name] if i_name is not None and i_name < len(r) else None
    if nm is not None and str(nm).strip() != "":
        named_count += 1

wb.close()

# pass 判定
format_ok = sheets_ok and columns_ok and tier_sum_ok and band_consistent and all_height_ge32
severe = (not landmark_has_ctbuh) or landmark_rows == 0 or (not disclosure_ok) or bool(big) or dup > 0 or (count and in_box/count < 0.7)
passed = bool(format_ok and not severe)

result = {
    "city": "勿加泗（bekasi）",
    "pass": passed,
    "count": count,
    "checks": {
        "sheets_ok": sheets_ok,
        "columns_ok": columns_ok,
        "tier_sum_ok": tier_sum_ok,
        "band_consistent": band_consistent,
        "all_height_ge32": all_height_ge32,
        "coords_in_bbox_pct": coords_in_bbox_pct,
        "landmark_rows": landmark_rows,
        "landmark_has_ctbuh": landmark_has_ctbuh,
        "disclosure_ok": disclosure_ok,
        "duplicate_rows": dup,
        "named_count": named_count,
    },
    "anomalies": anomalies,
    "_extra": {
        "sheet_names": sheet_names,
        "header": header_clean,
        "band_counts": band_counts,
        "band_inconsistent": band_inconsistent,
        "band_examples": band_examples,
        "valid_coord": valid_coord,
        "in_box": in_box,
        "height_max": max(heights) if heights else None,
        "height_min": min(heights) if heights else None,
        "below32": below32,
    }
}
print(json.dumps(result, ensure_ascii=False))
