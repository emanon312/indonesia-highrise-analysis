# -*- coding: utf-8 -*-
"""生成KML —— Google Earth/My Maps用，按高度分档5色，5个Folder分组。用法：python build_kml.py"""
import os, sys, csv
sys.stdout.reconfigure(encoding='utf-8')
HERE=os.path.dirname(os.path.abspath(__file__)); OUT=os.path.join(HERE,'..','output')

# 高度分档 → (KML颜色aabbggrr, 标签, 排序键)
def band(h):
    if h>=80: return 'b80'
    if h>=64: return 'b64'
    if h>=48: return 'b48'
    if h>=40: return 'b40'
    return 'b32'
STYLES=[  # id, KML色(aabbggrr), 显示标签
    ('b80','ff2730d7','≥80米 超高层'),
    ('b64','ff598dfc','≥64米'),
    ('b48','ff00a3f4','≥48米'),
    ('b40','ff63bd66','≥40米'),
    ('b32','ffb44575','≥32米'),
]
def esc(s): return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

# 读现有CSV
pts=[]
with open(os.path.join(OUT,'建筑坐标.csv'),encoding='utf-8-sig') as f:
    for r in csv.DictReader(f):
        pts.append(r)
print('读入 %d 点'%len(pts))

# 分组
groups={sid:[] for sid,_,_ in STYLES}
for p in pts:
    groups[band(float(p['高度(米)']))].append(p)

# 拼KML
buf=['<?xml version="1.0" encoding="UTF-8"?>',
     '<kml xmlns="http://www.opengis.net/kml/2.2"><Document>',
     '<name>印尼12城高层建筑分布（≥32米，共%d栋）</name>'%len(pts)]
CIRCLE='http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
for sid,color,_ in STYLES:
    buf.append(f'<Style id="{sid}"><IconStyle><color>{color}</color><scale>0.7</scale>'
               f'<Icon><href>{CIRCLE}</href></Icon></IconStyle></Style>')
for sid,_,label in STYLES:
    g=groups[sid]
    buf.append(f'<Folder><name>{label} ({len(g)})</name>')
    for p in g:
        desc=(f'城市：{esc(p["城市"])}｜高度：{p["高度(米)"]}米｜用途：{esc(p["用途分类"])}'
              f'｜数据来源：{esc(p["数据来源"])}｜地址：{esc(p["地址"])}')
        buf.append(f'<Placemark><name>{esc(p["楼宇名称"] or "(无名)")}</name>'
                   f'<description><![CDATA[{desc}]]></description>'
                   f'<styleUrl>#{sid}</styleUrl>'
                   f'<Point><coordinates>{p["经度"]},{p["纬度"]},0</coordinates></Point></Placemark>')
    buf.append('</Folder>')
buf.append('</Document></kml>')
path=os.path.join(OUT,'建筑分布地图.kml')
open(path,'w',encoding='utf-8').write('\n'.join(buf))
print('已生成: %s (%.1f MB)'%(path, os.path.getsize(path)/1024/1024))
for sid,_,label in STYLES: print('  %s: %d 栋'%(label,len(groups[sid])))
