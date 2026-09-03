# -*- coding: utf-8 -*-
"""给 CTBUH 地标清单补坐标：agent 抓回 {city}_ctbuh_landmarks.json (name/height/floors[/lat/lon])，
对缺坐标的用 {city}_osm.json 名字匹配补全 -> {city}_ctbuh.json (只保留有坐标的)。
用法: python geocode_ctbuh.py <city>
"""
import sys, json, os, re
city = sys.argv[1]
lm_path = '%s_ctbuh_landmarks.json' % city
if not os.path.exists(lm_path):
    print('%s 无地标清单文件，跳过' % city); sys.exit()
lm = json.load(open(lm_path, encoding='utf-8'))
osm = json.load(open('%s_osm.json' % city, encoding='utf-8')) if os.path.exists('%s_osm.json' % city) else []
named = [o for o in osm if o.get('name')]


def norm(s):
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


out = []
for l in lm:
    lat = l.get('lat'); lon = l.get('lon')
    if lat is None or lon is None:
        nm = norm(l.get('name'))
        if nm:
            for o in named:
                on = norm(o['name'])
                if on and (nm in on or on in nm):
                    lat = o['lat']; lon = o['lon']; break
    out.append({'name': l.get('name'), 'height_m': l.get('height_m'),
                'floors': l.get('floors'), 'lat': lat, 'lon': lon})
geo = sum(1 for o in out if o['lat'] is not None)
json.dump(out, open('%s_ctbuh.json' % city, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('%s ctbuh %d 栋, 有坐标 %d' % (city, len(out), geo))
