# -*- coding: utf-8 -*-
"""
雅加达官方 DPMPTSP 三维建筑库 全量拉取（带坐标）
- 五区 FeatureServer，拉 ≥32米 建筑的属性
- 坐标来自 gmaps 字段内嵌的经纬度（几何是 3D MultiPatch，不直接用）
- 带重试 + 分页，存原始 JSON，并交叉校验五区计数
"""
import requests, json, re, time, os

BASE = "https://gis-dpmptsp.jakarta.go.id/arcgis/rest/services/Hosted"

# (中文区名, 服务/图层路径, 预期≥32m计数)
DISTRICTS = [
    ("南雅加达", "Bangunan_Jaksel_Validasi/FeatureServer/311", 2023),
    ("中雅加达", "Bangunan_Jakpus_Validasi1/FeatureServer/234", 1802),
    ("北雅加达", "Bangunan_Jakut_Validasi/FeatureServer/235", 1804),
    ("西雅加达", "Bangunan_Jakbar_Validasi/FeatureServer/0", 1153),
    ("东雅加达", "Bangunan_Jaktim_Validasi1/FeatureServer/0", 438),
]

FIELDS = ("objectid,nama_tower,nama_bgn,bldgheight,jml_lapis,fungsi,jns_bgn,"
          "gmaps,osm,alamat,nama_jln,wadmkc,wadmkd,luas_bgn,kegiatan")

OUT = os.path.join(os.path.dirname(__file__), "..", "data", "jakarta_official_geom_raw.json")
PAGE = 2000


def fetch(url, retries=6):
    last = None
    for i in range(retries):
        try:
            r = requests.get(url, timeout=120)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last = e
            print(f"    重试 {i+1}/{retries}: {e}")
            time.sleep(3)
    raise RuntimeError(f"拉取失败: {url}\n{last}")


def extract_lonlat(gmaps):
    """从 https://www.google.com/maps/search/-6.2285,106.8272 提取 (lon, lat)"""
    if not gmaps or gmaps in ("<Null>", "-", "."):
        return (None, None)
    m = re.search(r"(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)", gmaps)
    if not m:
        return (None, None)
    lat, lon = float(m.group(1)), float(m.group(2))
    # 雅加达大致范围: lat -6.4~-5.9, lon 106.6~107.0；防止经纬度写反
    if not (-7 < lat < -5 and 106 < lon < 108):
        return (None, None)
    return (lon, lat)


def main():
    all_feats = []
    for dname, path, expect in DISTRICTS:
        got = []
        offset = 0
        while True:
            q = (f"{BASE}/{path}/query?where=bldgheight%3E=32&outFields={FIELDS}"
                 f"&returnGeometry=false&resultOffset={offset}&resultRecordCount={PAGE}&f=json")
            d = fetch(q)
            feats = d.get("features", [])
            for f in feats:
                a = f["attributes"]
                a["_dist"] = dname
                lon, lat = extract_lonlat(a.get("gmaps"))
                a["_lon"] = lon
                a["_lat"] = lat
                got.append(a)
            print(f"  {dname} offset={offset} 取回 {len(feats)} 条")
            if len(feats) < PAGE and not d.get("exceededTransferLimit"):
                break
            if len(feats) == 0:
                break
            offset += PAGE
        n_coord = sum(1 for a in got if a["_lon"] is not None)
        flag = "OK" if len(got) == expect else "!! 不符"
        print(f"==> {dname}: 拉到 {len(got)} 条 (预期 {expect} {flag})，有坐标 {n_coord} 条")
        all_feats.extend(got)

    # 全局校验
    print("\n========== 汇总 ==========")
    print(f"总记录数: {len(all_feats)} (预期 7220)")
    n_coord = sum(1 for a in all_feats if a["_lon"] is not None)
    print(f"有坐标: {n_coord} ({n_coord/len(all_feats)*100:.1f}%)")
    # 抽查最高楼坐标
    tops = sorted(all_feats, key=lambda a: a.get("bldgheight") or 0, reverse=True)[:5]
    print("最高5栋坐标抽查:")
    for a in tops:
        print(f"  {a.get('bldgheight'):.1f}m  {a.get('nama_tower') or a.get('nama_bgn')}  "
              f"lat={a['_lat']} lon={a['_lon']}  [{a['_dist']}]")

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(all_feats, f, ensure_ascii=False)
    print(f"\n已保存: {os.path.abspath(OUT)}")


if __name__ == "__main__":
    main()
