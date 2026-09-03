# -*- coding: utf-8 -*-
"""解析雅加达建筑 JSON，统计高层建筑数据"""
import json
import re

RAW = r"analysis/data/jakarta_buildings_raw.json"

with open(RAW, "r", encoding="utf-8") as f:
    data = json.load(f)

elements = data.get("elements", [])
total_elements = len(elements)


def parse_levels(val):
    """容错解析 building:levels，返回 float 或 None。
    支持 "20"、"20.5"、"3;4"（取最大值）等格式。"""
    if val is None:
        return None
    s = str(val).strip()
    # 提取所有数字（含小数）
    nums = re.findall(r"-?\d+(?:\.\d+)?", s)
    if not nums:
        return None
    try:
        # 多个值（如 "3;4"）取最大值，代表该建筑的最高楼层数
        return max(float(n) for n in nums)
    except ValueError:
        return None


buildings = []          # 成功解析的建筑 (levels, tags)
unparsable = 0          # 有标签但无法解析
with_levels_tag = 0     # 有 building:levels 标签的总数

for el in elements:
    tags = el.get("tags", {})
    if "building:levels" not in tags:
        continue
    with_levels_tag += 1
    lv = parse_levels(tags.get("building:levels"))
    if lv is None:
        unparsable += 1
        continue
    buildings.append((lv, tags))

# ---------- A. 各阈值数量 ----------
cnt_8 = sum(1 for lv, _ in buildings if lv >= 8)
cnt_10 = sum(1 for lv, _ in buildings if lv >= 10)
cnt_12 = sum(1 for lv, _ in buildings if lv >= 12)

# ---------- B. ≥8 层按建筑类型分类 ----------
office = apartments = hotel = commercial = other = 0
for lv, tags in buildings:
    if lv < 8:
        continue
    bt = str(tags.get("building", "")).lower()
    if bt == "office":
        office += 1
    elif bt in ("apartments", "residential"):
        apartments += 1
    elif bt == "hotel":
        hotel += 1
    elif bt in ("commercial", "retail", "mixed_use"):
        commercial += 1
    else:
        other += 1

# ---------- C. 数据完整度 ----------
# 楼层数直方图区间
bins = [
    ("8-9", lambda x: 8 <= x <= 9),
    ("10-11", lambda x: 10 <= x <= 11),
    ("12-14", lambda x: 12 <= x <= 14),
    ("15-19", lambda x: 15 <= x <= 19),
    ("20-29", lambda x: 20 <= x <= 29),
    ("30-49", lambda x: 30 <= x <= 49),
    ("50+", lambda x: x >= 50),
]
hist = {}
for label, fn in bins:
    hist[label] = sum(1 for lv, _ in buildings if lv >= 8 and fn(lv))

# 最高楼层
max_lv = 0
max_name = None
for lv, tags in buildings:
    if lv > max_lv:
        max_lv = lv
        max_name = tags.get("name") or tags.get("name:en") or "(无名称)"

# ---------- D. 楼层最高的 20 栋 ----------
top20 = sorted(buildings, key=lambda x: x[0], reverse=True)[:20]

# ========== 输出 ==========
print("=" * 60)
print("总抓取元素数:", total_elements)
print("有 building:levels 标签的建筑总数:", with_levels_tag)
print("成功解析:", len(buildings), " | 无法解析:", unparsable)
print("=" * 60)
print("\n[A] 各楼层阈值建筑数量")
print("  >= 8 层:", cnt_8)
print("  >= 10 层:", cnt_10)
print("  >= 12 层:", cnt_12)

print("\n[B] >=8 层建筑类型分类")
print("  office(写字楼):", office)
print("  apartments/residential(高层公寓):", apartments)
print("  hotel(酒店):", hotel)
print("  commercial/retail/mixed_use(综合体):", commercial)
print("  其他:", other)
print("  合计:", office + apartments + hotel + commercial + other)

print("\n[C] 数据完整度")
print("  楼层分布直方图(仅 >=8 层):")
for label, _ in bins:
    print(f"    {label:>6}: {hist[label]}")
print(f"  最高楼层: {max_lv:g} 层  ->  {max_name}")

print("\n[D] 楼层最高的 20 栋建筑")
print(f"  {'排名':<4}{'楼层':<6}{'类型':<14}{'名称'}")
for i, (lv, tags) in enumerate(top20, 1):
    name = tags.get("name") or tags.get("name:en") or "(无名称)"
    bt = tags.get("building", "")
    print(f"  {i:<4}{lv:<6g}{bt:<14}{name}")
