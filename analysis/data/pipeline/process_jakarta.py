# -*- coding: utf-8 -*-
"""
雅加达 OSM 建筑数据清洗脚本
铁律：只用真实数据，绝不推断、绝不估算、绝不补全。
"""
import json
import re
import math
from collections import Counter

RAW = "jakarta_buildings_raw.json"
OUT = "cities/jakarta.json"

# OSM building 标签 -> 中文类别
TYPE_MAP = {
    "office": "写字楼",
    "apartments": "公寓",
    "residential": "公寓",
    "hotel": "酒店",
    "commercial": "综合体",
    "retail": "综合体",
    "mixed_use": "综合体",
}


def map_type(tag_value):
    """根据 OSM building 标签值映射中文类别，无法判定归未分类（铁律：不推断）"""
    return TYPE_MAP.get(tag_value, "未分类")


def parse_levels(v):
    """健壮解析 building:levels。
    支持："20"、"20.5"、"3;4"（多值取最大）。
    无法解析的返回 None（跳过并计数）。
    注意：像 '10+'、'4-7'、'>7'、'Ground'、'-1'、'5+' 这类无法确定具体真实楼层数的，
    一律按无法解析处理（铁律：不推断、不估算）。
    """
    if v is None:
        return None
    v = v.strip()
    if not v:
        return None
    # 多值，取最大
    if ";" in v:
        nums = []
        for part in v.split(";"):
            n = parse_levels(part)
            if n is not None:
                nums.append(n)
        return max(nums) if nums else None
    # 纯整数
    if re.fullmatch(r"\d+", v):
        return int(v)
    # 小数（向下取整为楼层数；真实标注，无推断）
    if re.fullmatch(r"\d+\.\d+", v):
        return int(float(v))
    # 其它格式（10+、4-7、>7、Ground、-1、5+ 等）无法确定真实值，跳过
    return None


def haversine(lat1, lon1, lat2, lon2):
    """两点距离（米）"""
    R = 6371000.0
    p1 = math.radians(lat1)
    p2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def main():
    with open(RAW, encoding="utf-8") as f:
        raw = json.load(f)
    elements = raw["elements"]

    parse_failed = 0
    total_with_levels = 0  # 成功解析出真实楼层数的建筑总数
    existing = []          # 已建成（去重前）
    under_construction = []

    for el in elements:
        tags = el.get("tags", {})
        raw_levels = tags.get("building:levels")
        levels = parse_levels(raw_levels)
        if levels is None:
            parse_failed += 1
            continue
        total_with_levels += 1

        building = tags.get("building", "")
        # 中心点坐标
        center = el.get("center") or {}
        lat = center.get("lat", el.get("lat"))
        lon = center.get("lon", el.get("lon"))
        name = tags.get("name")

        # 在建判定：building=construction
        if building == "construction":
            # 在建建筑真实类型在 construction 键里
            cons_type = tags.get("construction", "")
            rec = {
                "levels": levels,
                "type": map_type(cons_type),
                "name": name,
                "lat": lat,
                "lon": lon,
            }
            under_construction.append(rec)
        else:
            rec = {
                "levels": levels,
                "type": map_type(building),
                "name": name,
                "lat": lat,
                "lon": lon,
            }
            existing.append(rec)

    # ---- 保守去重（仅对已建成） ----
    # 规则：同 name 且中心点距离极近 (<30m)，或完全相同坐标。无名建筑不因坐标近而删。
    removed = 0
    kept = []
    # 按 name 分组处理同名近距离重复
    by_name = {}
    no_name = []
    for r in existing:
        if r["name"]:
            by_name.setdefault(r["name"], []).append(r)
        else:
            no_name.append(r)

    for name, group in by_name.items():
        survivors = []
        for r in group:
            dup = False
            for s in survivors:
                if r["lat"] is None or s["lat"] is None:
                    continue
                if haversine(r["lat"], r["lon"], s["lat"], s["lon"]) < 30:
                    dup = True
                    break
            if dup:
                removed += 1
                # 保留楼层更高者（同名同位置，取更完整信息）
                # 找到匹配的 survivor，若当前更高则替换
                for i, s in enumerate(survivors):
                    if s["lat"] is not None and r["lat"] is not None and \
                       haversine(r["lat"], r["lon"], s["lat"], s["lon"]) < 30:
                        if r["levels"] > s["levels"]:
                            survivors[i] = r
                        break
            else:
                survivors.append(r)
        kept.extend(survivors)

    # 无名建筑：仅删完全相同坐标
    seen_coords = {}
    for r in no_name:
        key = (r["lat"], r["lon"])
        if r["lat"] is not None and key in seen_coords:
            # 完全相同坐标 -> 重复
            removed += 1
            s = seen_coords[key]
            if r["levels"] > s["levels"]:
                # 替换为更高的
                idx = kept.index(s)
                kept[idx] = r
                seen_coords[key] = r
            continue
        kept.append(r)
        if r["lat"] is not None:
            seen_coords[key] = r

    existing = kept

    # ---- 统计已建成 ge8/ge10/ge12 ----
    def empty_by_type():
        return {"写字楼": 0, "公寓": 0, "酒店": 0, "综合体": 0, "未分类": 0}

    def stat_band(items, threshold):
        sub = [r for r in items if r["levels"] >= threshold]
        bt = empty_by_type()
        for r in sub:
            bt[r["type"]] += 1
        return {"total": len(sub), "by_type": bt}

    ge8 = stat_band(existing, 8)
    ge10 = stat_band(existing, 10)
    ge12 = stat_band(existing, 12)

    # ---- 楼层分布直方图（基于 >=8 的已建成）----
    bands = [
        ("8-9", 8, 9),
        ("10-11", 10, 11),
        ("12-14", 12, 14),
        ("15-19", 15, 19),
        ("20-29", 20, 29),
        ("30-49", 30, 49),
        ("50+", 50, 10 ** 9),
    ]
    floor_dist = []
    for label, lo, hi in bands:
        cnt = sum(1 for r in existing if lo <= r["levels"] <= hi)
        floor_dist.append({"range": label, "count": cnt})

    # ---- Top25 最高已建成 ----
    top_sorted = sorted(existing, key=lambda r: r["levels"], reverse=True)[:25]
    top_buildings = []
    for i, r in enumerate(top_sorted, 1):
        top_buildings.append({
            "rank": i,
            "name": r["name"] if r["name"] else "（无名称）",
            "levels": r["levels"],
            "type": r["type"],
        })

    # ---- 在建统计 ----
    uc_ge8 = sum(1 for r in under_construction if r["levels"] >= 8)
    uc_ge10 = sum(1 for r in under_construction if r["levels"] >= 10)
    uc_ge12 = sum(1 for r in under_construction if r["levels"] >= 12)
    uc_top = sorted(under_construction, key=lambda r: r["levels"], reverse=True)[:10]
    uc_top_buildings = []
    for r in uc_top:
        uc_top_buildings.append({
            "name": r["name"] if r["name"] else "（无名称）",
            "levels": r["levels"],
            "type": r["type"],
        })

    result = {
        "city_id": "jakarta",
        "city_name_cn": "雅加达",
        "city_name_en": "Jakarta",
        "data_source": "OpenStreetMap / Overpass API (maps.mail.ru 镜像)",
        "fetch_date": "2026-06-18",
        "disclaimer": "本数据100%来自OSM真实building:levels标注，未做任何推断或补全",
        "summary": {
            "total_with_levels": total_with_levels,
            "duplicates_removed": removed,
            "under_construction_count": len(under_construction),
        },
        "existing": {
            "ge8": ge8,
            "ge10": ge10,
            "ge12": ge12,
            "floor_distribution": floor_dist,
            "top_buildings": top_buildings,
        },
        "under_construction": {
            "ge8": uc_ge8,
            "ge10": uc_ge10,
            "ge12": uc_ge12,
            "top_buildings": uc_top_buildings,
        },
    }

    import os
    os.makedirs("cities", exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # 控制台核验摘要
    print("=== 处理摘要 ===")
    print("解析失败(跳过)的楼层值:", parse_failed)
    print("有真实楼层数的建筑总数:", total_with_levels)
    print("在建建筑数:", len(under_construction))
    print("去重删除数:", removed)
    print("已建成(去重后):", len(existing))
    print("ge8:", ge8["total"], "ge10:", ge10["total"], "ge12:", ge12["total"])
    print("ge8 by_type:", ge8["by_type"])
    print("floor_dist:", floor_dist)
    print("在建 ge8/ge10/ge12:", uc_ge8, uc_ge10, uc_ge12)
    print("写入:", OUT)


if __name__ == "__main__":
    main()
