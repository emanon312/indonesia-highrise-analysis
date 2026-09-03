# -*- coding: utf-8 -*-
"""统计 Google OB 2.5D Temporal(2023) 雅加达高层 —— 直读 COG 内部块偏移并行解码。

关键：planarconfig=2(波段分离), 49x49=2401块/波段, deflate+float预测器(predictor=3)。
只取 band1(building_height) 中与窗口重叠的内部块，用线程池并发 Range GET 原始字节，
本地用 imagecodecs 解 deflate + 反 float 预测器，避免任何 zarr/重复元数据开销。
工作分辨率 2m（4x4 最大池化保留峰值）。
"""
import sys, json, time, math, threading
sys.path.insert(0, r'analysis\data')
from googleob_http_cog import HttpRangeFile
import tifffile, numpy as np, requests
import pyproj
from scipy import ndimage
from concurrent.futures import ThreadPoolExecutor
import imagecodecs

WEST, SOUTH, EAST, NORTH = 106.6, -6.4, 107.0, -6.1
EPSG=32748; POOL=4; RES=0.5; CELL=RES*POOL; CHUNK=512; NTHREADS=48
BANDS_PER=2401; TGRID=49

tr=pyproj.Transformer.from_crs('EPSG:4326',f'EPSG:{EPSG}',always_xy=True)
xs,ys=[],[]
for lon in (WEST,EAST):
    for lat in (SOUTH,NORTH):
        x,y=tr.transform(lon,lat); xs.append(x); ys.append(y)
BX0,BX1=min(xs),max(xs); BY0,BY1=min(ys),max(ys)
GX0=np.floor(BX0/CELL)*CELL; GX1=np.ceil(BX1/CELL)*CELL
GY0=np.floor(BY0/CELL)*CELL; GY1=np.ceil(BY1/CELL)*CELL
W2=int(round((GX1-GX0)/CELL)); H2=int(round((GY1-GY0)/CELL))
print(f'主网格 {H2}x{W2} (2m)', flush=True)
master=np.zeros((H2,W2),dtype=np.float32)

tiles=json.load(open(r'analysis\data\googleob_jakarta_tiles.json',encoding='utf-8'))['tiles']
tiles=[t for t in tiles if '2e69c_2023' in t['url']]
print(f'瓦片 {len(tiles)} 个', flush=True)

def decode_tile(raw, h, w):
    """deflate 解压 + 反 float 预测器(predictor=3) -> (h,w) float32（已与 tifffile 校验一致）"""
    buf=imagecodecs.deflate_decode(raw)
    data=np.frombuffer(buf,dtype=np.float32).reshape(h,w).copy()
    return imagecodecs.floatpred_decode(data, axis=-1)

t_start=time.time()
for ti,t in enumerate(tiles):
    tt0=time.time()
    url=t['url']; a=t['affine']
    x0=a['translateX']; y0=a['translateY']; sx=a['scaleX']; sy=a['scaleY']; tw,th=t['w'],t['h']
    c_lo=max(0,int(np.floor((BX0-x0)/sx))); c_hi=min(tw,int(np.ceil((BX1-x0)/sx)))
    r_lo=max(0,int(np.floor((y0-BY1)/(-sy)))); r_hi=min(th,int(np.ceil((y0-BY0)/(-sy))))
    c_lo-=c_lo%POOL; r_lo-=r_lo%POOL; c_hi-=c_hi%POOL; r_hi-=r_hi%POOL
    if c_hi<=c_lo or r_hi<=r_lo:
        print(f'  瓦片{ti+1} 无重叠', flush=True); continue
    f=HttpRangeFile(url); p=tifffile.TiffFile(f).pages[0]
    offs=p.dataoffsets; cnts=p.databytecounts
    H=r_hi-r_lo; Wd=c_hi-c_lo
    sub=np.zeros((H,Wd),dtype=np.float32)
    rb0,rb1=r_lo//CHUNK,(r_hi-1)//CHUNK
    cb0,cb1=c_lo//CHUNK,(c_hi-1)//CHUNK
    blocks=[(rb,cb) for rb in range(rb0,rb1+1) for cb in range(cb0,cb1+1)]
    sess_local=threading.local()
    def sess():
        if not hasattr(sess_local,'s'): sess_local.s=requests.Session()
        return sess_local.s
    def fetch(block):
        rb,cb=block
        idx=BANDS_PER + rb*TGRID + cb   # band1 起始 2401
        off=offs[idx]; cnt=cnts[idx]
        r=sess().get(url, headers={'Range':f'bytes={off}-{off+cnt-1}'}, timeout=120)
        raw=r.content
        rs=rb*CHUNK; re=min((rb+1)*CHUNK,th); cs=cb*CHUNK; ce=min((cb+1)*CHUNK,tw)
        bh=re-rs; bw=ce-cs
        # COG 边缘块仍按完整 512 编码，解码后裁剪
        full=decode_tile(raw,CHUNK,CHUNK)
        return (rs,re,cs,ce, full[:bh,:bw], len(raw))
    nb=0
    with ThreadPoolExecutor(max_workers=NTHREADS) as ex:
        for rs,re,cs,ce,data,nbytes in ex.map(fetch,blocks):
            nb+=nbytes
            ar0=max(rs,r_lo); ar1=min(re,r_hi); ac0=max(cs,c_lo); ac1=min(ce,c_hi)
            if ar1<=ar0 or ac1<=ac0: continue
            sub[ar0-r_lo:ar1-r_lo, ac0-c_lo:ac1-c_lo]=data[ar0-rs:ar1-rs, ac0-cs:ac1-cs]
    sp=sub.reshape(H//POOL,POOL,Wd//POOL,POOL).max(axis=(1,3))
    sub_x0=x0+c_lo*sx; sub_y0=y0+r_lo*sy
    col_off=int(round((sub_x0-GX0)/CELL)); row_off=int(round((GY1-sub_y0)/CELL))
    r0=max(0,row_off); c0=max(0,col_off); r1=min(H2,row_off+sp.shape[0]); c1=min(W2,col_off+sp.shape[1])
    if r1>r0 and c1>c0:
        sr0=r0-row_off; sc0=c0-col_off
        blk=sp[sr0:sr0+(r1-r0),sc0:sc0+(c1-c0)]
        np.maximum(master[r0:r1,c0:c1],blk,out=master[r0:r1,c0:c1])
    print(f'  瓦片{ti+1}/{len(tiles)} 窗口{sub.shape} {len(blocks)}块 {nb/1e6:.0f}MB {time.time()-tt0:.0f}s 累计{time.time()-t_start:.0f}s max={master.max():.1f}', flush=True)

print(f'读取完成 耗时{time.time()-t_start:.0f}s 最高{master.max():.1f}m', flush=True)
results={'max_height_m':float(master.max()),'work_res_m':CELL,
         'source':'Google Open Buildings 2.5D Temporal v1 year2023, GCS open-buildings-temporal-data, tileset 2e69c'}
cell_area=CELL**2; struct=np.ones((3,3),dtype=int); gold={32:7220,40:5255,48:4222}
for thr in (32,40,48):
    mask=master>=thr; npix=int(mask.sum())
    lbl,ncomp=ndimage.label(mask,structure=struct)
    if ncomp:
        sizes=np.bincount(lbl.ravel())[1:]; n16=int((sizes>=4).sum()); n40=int((sizes>=10).sum())
    else: n16=n40=0
    results[f'ge{thr}']={'pixels_2m':npix,'area_km2':round(npix*cell_area/1e6,4),
                         'components_all':int(ncomp),'components_ge16m2':n16,'components_ge40m2':n40}
    print(f'>= {thr}m: 像元={npix} 面积={npix*cell_area/1e6:.3f}km2 块全部={ncomp} >=16m2={n16} >=40m2={n40}', flush=True)
print('--- 捕获率(连通块>=16m2/金标准) ---', flush=True)
for thr in (32,40,48):
    c=results[f'ge{thr}']['components_ge16m2']; print(f'  >= {thr}m: {c}/{gold[thr]} = {100*c/gold[thr]:.1f}%', flush=True)
results['gold']=gold
np.save(r'analysis\data\googleob_jakarta_height_2m.npy', master)
json.dump(results,open(r'analysis\data\googleob_jakarta_results.json','w',encoding='utf-8'),ensure_ascii=False,indent=2)
print('已保存 npy 与 results.json', flush=True)
