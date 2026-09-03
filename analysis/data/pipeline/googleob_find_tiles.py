# -*- coding: utf-8 -*-
"""查找 Google Open Buildings 2.5D Temporal (2023) 中覆盖雅加达 bbox 的瓦片。

数据来源：公开 GCS 桶 open-buildings-temporal-data（匿名可读）。
manifest 命名：v1/manifests/{s2token}_EPSG_{epsg}_{year}_06_30.json
瓦片波段：band1=fractional_building_count, band2=building_height(米), band3=building_presence
"""
import requests, json

# 雅加达金标准 bbox（经纬度 WGS84）
WEST, SOUTH, EAST, NORTH = 106.6, -6.4, 107.0, -6.1

BASE = 'https://storage.googleapis.com/open-buildings-temporal-data/'

import pyproj
# 雅加达在 UTM zone 48S = EPSG:32748
EPSG = 32748
tr = pyproj.Transformer.from_crs('EPSG:4326', f'EPSG:{EPSG}', always_xy=True)
# bbox 四角转 UTM，取外包矩形
xs, ys = [], []
for lon in (WEST, EAST):
    for lat in (SOUTH, NORTH):
        x, y = tr.transform(lon, lat)
        xs.append(x); ys.append(y)
bx0, bx1 = min(xs), max(xs)
by0, by1 = min(ys), max(ys)
print(f'雅加达 bbox 在 EPSG:{EPSG} 下的范围: X[{bx0:.0f},{bx1:.0f}] Y[{by0:.0f},{by1:.0f}]')

# 该 EPSG 对应的 manifest
name = '2f_EPSG_32748_2023_06_30.json'
m = requests.get(BASE + 'v1/manifests/' + name, timeout=120).json()
uri_prefix = m['uriPrefix']  # gs://...
# 转成 https
https_prefix = uri_prefix.replace('gs://', 'https://storage.googleapis.com/')

hits = []
for ts in m['tilesets']:
    for src in ts['sources']:
        a = src['affineTransform']
        d = src['dimensions']
        x0 = a['translateX']; y0 = a['translateY']
        sx = a['scaleX']; sy = a['scaleY']  # sy 为负
        w = d['width']; h = d['height']
        # 瓦片四角 -> UTM 范围
        tx0 = x0
        tx1 = x0 + sx * w
        ty0 = y0
        ty1 = y0 + sy * h
        tile_x0, tile_x1 = min(tx0, tx1), max(tx0, tx1)
        tile_y0, tile_y1 = min(ty0, ty1), max(ty0, ty1)
        # 矩形相交判断
        if tile_x1 < bx0 or tile_x0 > bx1 or tile_y1 < by0 or tile_y0 > by1:
            continue
        url = https_prefix + src['uris'][0]
        hits.append({
            'url': url,
            'x0': tile_x0, 'x1': tile_x1, 'y0': tile_y0, 'y1': tile_y1,
            'affine': a, 'w': w, 'h': h
        })

print(f'相交瓦片数: {len(hits)}')
for hh in hits:
    print(hh['url'])
    print(f"   UTM X[{hh['x0']:.0f},{hh['x1']:.0f}] Y[{hh['y0']:.0f},{hh['y1']:.0f}]")

# 保存结果
out = {
    'bbox_wgs84': [WEST, SOUTH, EAST, NORTH],
    'epsg': EPSG,
    'bbox_utm': [bx0, by0, bx1, by1],
    'manifest': name,
    'https_prefix': https_prefix,
    'tiles': hits,
}
with open('印尼楼房分析/analysis/data/googleob_jakarta_tiles.json', 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
print('已保存 googleob_jakarta_tiles.json')
