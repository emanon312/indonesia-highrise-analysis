# -*- coding: utf-8 -*-
"""测试：从 page0(全分辨率 0.5m) 用 tifffile 读取一个子窗口，验证只拉取重叠的内部512瓦片。"""
import sys
sys.path.insert(0, r'analysis\data')
from googleob_http_cog import HttpRangeFile
import tifffile, numpy as np, time

url = 'https://storage.googleapis.com/open-buildings-temporal-data/v1/geotiffs/2e69c_2023_06_30/tile_ySHmjn0JYAs.tif'
f = HttpRangeFile(url)
tif = tifffile.TiffFile(f)
p0 = tif.pages[0]
print('page0 shape', p0.shape, 'tiled', p0.is_tiled, 'tile', p0.tilewidth, p0.tilelength)
# 用 zarr 接口做窗口读取
store = p0.aszarr()
import zarr
z = zarr.open(store, mode='r')
print('zarr shape', z.shape, 'dtype', z.dtype, 'chunks', z.chunks)
t0 = time.time()
# 读 band1 的一个 2048x2048 窗口
sub = z[1, 5000:7048, 5000:7048]
print('子窗口 shape', sub.shape, 'max', float(sub.max()), 'mean', float(sub.mean()))
print('>=32m像元', int((sub>=32).sum()), '>=48m像元', int((sub>=48).sum()))
print('耗时', round(time.time()-t0,1),'s  HTTP请求', f.nreq, '字节MB', round(f.nbytes/1e6,1))
