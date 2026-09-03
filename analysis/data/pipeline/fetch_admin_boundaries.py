# -*- coding: utf-8 -*-
"""从 OSM Overpass 抓取各城的 kecamatan 行政区边界（admin_level=6）。
输出 {city}_admin_boundaries.json，包含每个 kecamatan 的名称和多边形坐标。
不需要 shapely/geopandas 依赖。

用法: python fetch_admin_boundaries.py <city>
"""
import sys, json, requests, time, math
sys.path.insert(0, '.')
from googleob_city import CITIES

city = sys.argv[1]
c = CITIES[city]
W, S, E, N = c['bbox']
# 扩大一点 bbox 确保覆盖
W -= 0.02; S -= 0.02; E += 0.02; N += 0.02
bbox = '%s,%s,%s,%s' % (S, W, N, E)

# 查询 admin_level=6 (kecamatan) 和 admin_level=5 (kota/kabupaten)
q = ('[out:json][timeout:120];'
     '(rel["boundary"="administrative"]["admin_level"="6"](%s);'
     ' rel["boundary"="administrative"]["admin_level"="5"](%s););'
     'out tags;'
     'out geom;') % (bbox, bbox)

url = 'https://maps.mail.ru/osm/tools/overpass/api/interpreter'
r = requests.post(url, data={'data': q}, timeout=180)
r.raise_for_status()
data = r.json()

boundaries = []
for el in data.get('elements', []):
    tags = el.get('tags', {})
    admin_level = tags.get('admin_level', '')
    name = tags.get('name', '')
    if not name:
        continue

    # 提取多边形坐标：将同一 relation 中 role=outer 的 way 拼接成完整环
    members = el.get('members', [])
    outer_ways = []
    for member in members:
        if member.get('type') == 'way' and member.get('role') in ('outer', ''):
            geom = member.get('geometry', [])
            if geom:
                outer_ways.append([(pt['lon'], pt['lat']) for pt in geom])

    # 拼接 ways 成完整环（OSM relation 的 ways 按顺序排列）
    polygons = []
    if outer_ways:
        # 尝试直接拼接：将所有 way 的坐标串联
        all_coords = []
        for way in outer_ways:
            if all_coords and way:
                # 如果首点与前一个 way 的末点相同，跳过重复点
                if abs(all_coords[-1][0] - way[0][0]) < 0.00001 and abs(all_coords[-1][1] - way[0][1]) < 0.00001:
                    all_coords.extend(way[1:])
                else:
                    all_coords.extend(way)
            else:
                all_coords.extend(way)
        if len(all_coords) >= 3:
            polygons.append(all_coords)

    if polygons:
        boundaries.append({
            'name': name,
            'admin_level': admin_level,
            'polygons': polygons,
        })

# 按 admin_level 分组
kecamatans = [b for b in boundaries if b['admin_level'] == '6']
kotas = [b for b in boundaries if b['admin_level'] == '5']

print('%s: %d 个 kecamatan, %d 个 kota' % (city, len(kecamatans), len(kotas)), flush=True)

out = {
    'city': city,
    'cn': c['cn'],
    'bbox': c['bbox'],
    'kecamatans': kecamatans,
    'kotas': kotas,
}
json.dump(out, open('%s_admin_boundaries.json' % city, 'w', encoding='utf-8'), ensure_ascii=False)
print('已保存 %s_admin_boundaries.json' % city, flush=True)
