# -*- coding: utf-8 -*-
"""审查剔除质量, 揪出误伤(尤其机构名的公寓/住宅)"""
import json, os, re, sys, collections
sys.stdout.reconfigure(encoding="utf-8")
from build_excel import classify, clean, normj, bname

RAW = os.path.join(os.path.dirname(__file__), "..", "data", "jakarta_official_geom_raw.json")
data = json.load(open(RAW, encoding="utf-8"))

RESI = re.compile(r"apartemen|apartment|residence|rumah susun|rumah tinggal|flat|hunian|kost|kondominium|condominium", re.I)

name_hits, suspect = [], []
for rec in data:
    status, cat, basis, hit = classify(rec)
    if status != "剔除":
        continue
    nm = bname(rec); jns = clean(rec.get("jns_bgn")); fungsi = clean(rec.get("fungsi"))
    if basis == "楼名关键词":
        name_hits.append((cat, hit, nm, jns, fungsi))
    # 可疑误伤: 被剔但带居住属性
    blob = f"{nm} {jns} {fungsi}"
    if RESI.search(blob):
        suspect.append((cat, basis, hit, nm, jns, fungsi))

print(f"===== 楼名兜底命中全部 {len(name_hits)} 条 =====")
for cat, hit, nm, jns, fungsi in name_hits:
    print(f"  [{cat}] '{hit}' | {nm} | jns={jns} | fungsi={fungsi}")

print(f"\n===== 可疑误伤: 被剔但带居住属性 {len(suspect)} 条 =====")
for cat, basis, hit, nm, jns, fungsi in suspect:
    print(f"  [{cat}/{basis}] '{hit}' | {nm} | jns={jns} | fungsi={fungsi}")
