# -*- coding: utf-8 -*-
"""按 S2 cell token 一次性下载 V3 csv.gz，提取该 cell 覆盖的所有城市（避免重复下载）。
用法: python extract_v3_cell.py <token>
"""
import sys, os, duckdb, time
sys.path.insert(0, '.')
from googleob_city import CITIES

TOKEN_CITIES = {
    '2dd': ['surabaya'],
    '2e7': ['bandung', 'bekasi', 'tangerang', 'bogor', 'semarang', 'depok'],
    '303': ['medan'], '2db': ['makassar'], '31d': ['batam'], '2e3': ['palembang'],
}

token = sys.argv[1]
cities = TOKEN_CITIES[token]
Ws = [CITIES[c]['bbox'][0] for c in cities]; Ss = [CITIES[c]['bbox'][1] for c in cities]
Es = [CITIES[c]['bbox'][2] for c in cities]; Ns = [CITIES[c]['bbox'][3] for c in cities]
uW, uS, uE, uN = min(Ws), min(Ss), max(Es), max(Ns)
# 优先读本地完整文件（curl 续传下好的），否则流式读远程
local = '%s.csv.gz' % token
src = local if os.path.exists(local) else \
    'https://storage.googleapis.com/open-buildings-data/v3/polygons_s2_level_4_gzip/%s_buildings.csv.gz' % token
con = duckdb.connect(); con.execute('INSTALL httpfs; LOAD httpfs;')
t = time.time()
con.execute("CREATE TABLE cell AS SELECT latitude,longitude,area_in_meters,confidence "
            "FROM read_csv_auto('%s') "
            "WHERE longitude BETWEEN %s AND %s AND latitude BETWEEN %s AND %s" % (src, uW, uE, uS, uN))
print('%s cell 读取完成 %.0fs (源:%s)' % (token, time.time()-t, 'local' if src == local else 'remote'), flush=True)
for c in cities:
    W, S, E, N = CITIES[c]['bbox']
    con.execute("COPY (SELECT * FROM cell WHERE longitude BETWEEN %s AND %s AND latitude BETWEEN %s AND %s) "
                "TO '%s_v3.parquet' (FORMAT parquet)" % (W, E, S, N, c))
    n = con.execute("SELECT count(*) FROM read_parquet('%s_v3.parquet')" % c).fetchone()[0]
    print('  %s: %d 栋' % (c, n), flush=True)
