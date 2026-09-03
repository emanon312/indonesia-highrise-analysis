# -*- coding: utf-8 -*-
"""抓取某城 OSM 中"有名字/有楼层/商业类型"的建筑（用于给 V3 足迹最近邻补名/用途/真层）。
用 maps.mail.ru Overpass 镜像（国内可达，已验证）。
用法: python fetch_osm.py <city>  ->  {city}_osm.json: [{name,btype,levels,lat,lon}]
"""
import sys, json, requests
sys.path.insert(0, '.')
from googleob_city import CITIES

BTYPE_CN = {'office': '写字楼', 'apartments': '公寓', 'residential': '住宅', 'house': '住宅',
            'hotel': '酒店', 'commercial': '商业', 'retail': '商场', 'mixed': '综合体',
            'mixed_use': '综合体', 'dormitory': '宿舍', 'hospital': '医院', 'school': '学校',
            'university': '高校', 'government': '政府', 'public': '公共'}

city = sys.argv[1]; c = CITIES[city]; W, S, E, N = c['bbox']
bbox = '%s,%s,%s,%s' % (S, W, N, E)
q = ('[out:json][timeout:300];'
     '(way["building"]["name"](%s);'
     ' way["building:levels"](%s);'
     ' way["building"~"office|apartments|hotel|commercial|retail|residential|mixed"](%s);'
     ' relation["building"]["name"](%s););'
     'out tags center;') % (bbox, bbox, bbox, bbox)
url = 'https://maps.mail.ru/osm/tools/overpass/api/interpreter'
r = requests.post(url, data={'data': q}, timeout=320)
els = r.json()['elements']
out = []
for e in els:
    t = e.get('tags', {}); ctr = e.get('center') or {}
    lat = ctr.get('lat'); lon = ctr.get('lon')
    if lat is None:
        continue
    lv = t.get('building:levels')
    try:
        lv = int(float(lv)) if lv else None
    except Exception:
        lv = None
    bt = t.get('building')
    bt = None if bt in ('yes', 'true', None) else BTYPE_CN.get(bt, bt)
    out.append({'name': t.get('name'), 'btype': bt, 'levels': lv, 'lat': lat, 'lon': lon})
json.dump(out, open('%s_osm.json' % city, 'w', encoding='utf-8'), ensure_ascii=False)
print('%s OSM enrich candidates %d' % (city, len(out)))
