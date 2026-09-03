# -*- coding: utf-8 -*-
"""从 Google Open Buildings V3 (GCS csv.gz) 按城市 bbox 提取建筑足迹质心+面积。
用法: python extract_v3.py <city>
输出: {city}_v3.parquet (纬度/经度/面积/置信度)
csv.gz 是 gzip 无法随机读，duckdb 会流式下载整份再按 bbox 过滤，输出只保留城市范围。
"""
import sys, duckdb, time
sys.path.insert(0, '.')
from googleob_city import CITIES

# 各城 S2 level-4 token（s2sphere 算出，已与用户给的 2dd/303 校验一致）
TOKEN = {'surabaya': '2dd', 'bandung': '2e7', 'bekasi': '2e7', 'tangerang': '2e7',
         'bogor': '2e7', 'medan': '303', 'semarang': '2e7', 'makassar': '2db',
         'batam': '31d', 'palembang': '2e3', 'depok': '2e7'}

city = sys.argv[1]
c = CITIES[city]; W, S, E, N = c['bbox']; tk = TOKEN[city]
url = 'https://storage.googleapis.com/open-buildings-data/v3/polygons_s2_level_4_gzip/%s_buildings.csv.gz' % tk
out = '%s_v3.parquet' % city
con = duckdb.connect(); con.execute('INSTALL httpfs; LOAD httpfs;')
t = time.time()
sql = ("COPY (SELECT latitude,longitude,area_in_meters,confidence "
       "FROM read_csv_auto('%s') "
       "WHERE longitude BETWEEN %s AND %s AND latitude BETWEEN %s AND %s) "
       "TO '%s' (FORMAT parquet)") % (url, W, E, S, N, out)
con.execute(sql)
n = con.execute("SELECT count(*) FROM read_parquet('%s')" % out).fetchone()[0]
print('%s V3 footprints %d  %.0fs  token=%s' % (city, n, time.time()-t, tk))
