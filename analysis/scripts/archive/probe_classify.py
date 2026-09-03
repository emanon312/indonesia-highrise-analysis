# -*- coding: utf-8 -*-
"""探查 jns_bgn / fungsi 取值分布，为剔除规则建映射表"""
import json, os, sys, collections
sys.stdout.reconfigure(encoding="utf-8")

RAW = os.path.join(os.path.dirname(__file__), "..", "data", "jakarta_official_geom_raw.json")
data = json.load(open(RAW, encoding="utf-8"))

def norm(v):
    if v is None: return "<空>"
    v = str(v).strip()
    if v in ("", "<Null>", "-", ".", "null"): return "<空>"
    return v

print(f"总记录: {len(data)}\n")

print("===== jns_bgn 取值分布（用途细类，降序）=====")
c = collections.Counter(norm(a.get("jns_bgn")) for a in data)
for k, n in c.most_common():
    print(f"  {n:5d}  {k}")

print("\n===== fungsi 取值分布（用途大类）=====")
c2 = collections.Counter(norm(a.get("fungsi")) for a in data)
for k, n in c2.most_common():
    print(f"  {n:5d}  {k}")

print("\n===== kegiatan 取值分布（活动，前30）=====")
c3 = collections.Counter(norm(a.get("kegiatan")) for a in data)
for k, n in c3.most_common(30):
    print(f"  {n:5d}  {k}")

# 楼名填充率
def hasname(a):
    return norm(a.get("nama_tower")) != "<空>" or norm(a.get("nama_bgn")) != "<空>"
n_name = sum(1 for a in data if hasname(a))
print(f"\n楼名填充率: {n_name}/{len(data)} = {n_name/len(data)*100:.1f}%")

# 层数填充率
n_lapis = sum(1 for a in data if a.get("jml_lapis") and a.get("jml_lapis") > 0)
print(f"层数填充率: {n_lapis}/{len(data)} = {n_lapis/len(data)*100:.1f}%")

# 高度各档（剔除前）
for thr, lbl in [(32,"≥8层"),(40,"≥10层"),(48,"≥12层"),(64,"≥16层"),(80,"≥20层")]:
    n = sum(1 for a in data if (a.get("bldgheight") or 0) >= thr)
    print(f"≥{thr}m ({lbl}): {n}")
