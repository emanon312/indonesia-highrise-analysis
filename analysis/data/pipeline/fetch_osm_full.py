# -*- coding: utf-8 -*-
"""抓取某城 OSM 建筑数据（完整标签版），用于公共设施筛选和地址补全。
相比 fetch_osm.py，额外抓取：
- amenity 标签（place_of_worship/school/hospital/police 等）用于公共设施识别
- addr:* 标签（addr:full/addr:street/addr:city）用于地址补全
- building:use / office / religion 等辅助分类标签

用 maps.mail.ru Overpass 镜像（国内可达，已验证）。
用法: python fetch_osm_full.py <city>  ->  {city}_osm_full.json
输出: [{name,btype,levels,lat,lon,amenity,addr_full,addr_street,addr_city,building_use,office,religion,tags}]
"""
import sys, json, requests, time
sys.path.insert(0, '.')
from googleob_city import CITIES

BTYPE_CN = {'office': '写字楼', 'apartments': '公寓', 'residential': '住宅', 'house': '住宅',
            'hotel': '酒店', 'commercial': '商业', 'retail': '商场', 'mixed': '综合体',
            'mixed_use': '综合体', 'dormitory': '宿舍', 'hospital': '医院', 'school': '学校',
            'university': '高校', 'government': '政府', 'public': '公共',
            'kindergarten': '学校', 'college': '高校', 'religious': '宗教',
            'industrial': '工业', 'warehouse': '仓库', 'garage': '交通/停车',
            'sports_centre': '体育设施', 'stadium': '体育设施', 'train_station': '交通/停车',
            'transportation': '交通/停车', 'parking': '交通/停车'}

city = sys.argv[1]
c = CITIES[city]
W, S, E, N = c['bbox']
bbox = '%s,%s,%s,%s' % (S, W, N, E)

# 查询 1：带名字/楼层/商业类型的建筑（同原版）
q1 = ('[out:json][timeout:300];'
      '(way["building"]["name"](%s);'
      ' way["building:levels"](%s);'
      ' way["building"~"office|apartments|hotel|commercial|retail|residential|mixed|'
      'school|hospital|university|college|government|public|religious|dormitory|industrial|'
      'warehouse|train_station|transportation|parking|sports_centre|kindergarten"](%s);'
      ' relation["building"]["name"](%s););'
      'out tags center;') % (bbox, bbox, bbox, bbox)

# 查询 2：带 amenity 标签的建筑（补充公共设施识别）
q2 = ('[out:json][timeout:300];'
      '(way["amenity"]["building"](%s);'
      ' way["amenity"](%s);'
      ' relation["amenity"]["building"](%s););'
      'out tags center;') % (bbox, bbox, bbox)

url = 'https://maps.mail.ru/osm/tools/overpass/api/interpreter'

all_els = {}
t0 = time.time()

for label, q in [('building', q1), ('amenity', q2)]:
    try:
        r = requests.post(url, data={'data': q}, timeout=320)
        r.raise_for_status()
        els = r.json().get('elements', [])
        for e in els:
            eid = str(e.get('type', '')) + '/' + str(e.get('id', ''))
            if eid not in all_els:
                all_els[eid] = e
        print('%s 查询 %s: %d 条 (累计去重 %d)' % (city, label, len(els), len(all_els)), flush=True)
    except Exception as ex:
        print('%s 查询 %s 失败: %s' % (city, label, ex), flush=True)

out = []
for e in all_els.values():
    t = e.get('tags', {})
    ctr = e.get('center') or {}
    lat = ctr.get('lat')
    lon = ctr.get('lon')
    if lat is None:
        continue

    lv = t.get('building:levels')
    try:
        lv = int(float(lv)) if lv else None
    except Exception:
        lv = None

    bt = t.get('building')
    bt = None if bt in ('yes', 'true', None) else BTYPE_CN.get(bt, bt)

    # 提取 amenity 标签（公共设施识别核心）
    amenity = t.get('amenity')

    # 提取地址标签
    addr_full = t.get('addr:full')
    addr_street = t.get('addr:street')
    addr_city = t.get('addr:city')

    # 辅助分类标签
    building_use = t.get('building:use')
    office = t.get('office')
    religion = t.get('religion')

    rec = {
        'name': t.get('name'),
        'btype': bt,
        'levels': lv,
        'lat': lat,
        'lon': lon,
        'amenity': amenity,
        'addr_full': addr_full,
        'addr_street': addr_street,
        'addr_city': addr_city,
        'building_use': building_use,
        'office': office,
        'religion': religion,
    }
    out.append(rec)

out_path = '%s_osm_full.json' % city
json.dump(out, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False)

# 统计
n_amenity = sum(1 for r in out if r['amenity'])
n_addr = sum(1 for r in out if r['addr_full'] or r['addr_street'])
print('%s OSM 完整: %d 栋, 有amenity=%d, 有地址=%d, 耗时 %.0fs' %
      (city, len(out), n_amenity, n_addr, time.time() - t0), flush=True)
