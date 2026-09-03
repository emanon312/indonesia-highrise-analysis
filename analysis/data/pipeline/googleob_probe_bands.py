# -*- coding: utf-8 -*-
"""抽样校验：读取一个瓦片的中等分辨率层级，确认 band2=building_height 的取值范围与口径。"""
import sys
sys.path.insert(0, r'analysis\data')
from googleob_http_cog import HttpRangeFile
import tifffile, numpy as np

url = 'https://storage.googleapis.com/open-buildings-temporal-data/v1/geotiffs/2e69c_2023_06_30/tile_ySHmjn0JYAs.tif'
f = HttpRangeFile(url)
tif = tifffile.TiffFile(f)
# 读 page11 (1563x1563, 约8m/像元) 整页，看三个波段分布
arr = tif.pages[11].asarray()  # shape (3, H, W)
print('整页 shape', arr.shape)
for b in range(3):
    band = arr[b]
    print(f'band{b}: min={np.nanmin(band):.3f} max={np.nanmax(band):.3f} '
          f'mean={np.nanmean(band):.3f} 非零像元={np.count_nonzero(band)}')
# band2 即 building_height（按 README band 顺序 1=count,2=height,3=presence，0-based 即 index1）
h = arr[1]
print('--- 按 band index1 当作 height ---')
for thr in (12, 32, 40, 48):
    print(f'  >= {thr}m 像元数(8m分辨率层级): {(h>=thr).sum()}')
print('HTTP请求', f.nreq, '字节', f.nbytes)
