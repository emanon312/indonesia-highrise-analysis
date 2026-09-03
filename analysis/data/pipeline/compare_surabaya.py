# -*- coding: utf-8 -*-
"""一次性：按用户 GEE 的阈值重算泗水，便于和"逐栋 GEE"结果对比。
复用 googleob_city.py 已验证的瓦片读取逻辑，额外保存 npy 并按多阈值统计。
"""
import sys, math, threading, time
DATA = r'analysis\data'
sys.path.insert(0, DATA)
import googleob_city as G
import numpy as np, requests, tifffile
from scipy import ndimage
from concurrent.futures import ThreadPoolExecutor

CHUNK = G.CHUNK; POOL = G.POOL; CELL = G.CELL

city = 'surabaya'
c = G.CITIES[city]; bbox = c['bbox']; epsg = c['epsg']
(bx0, by0, bx1, by1), tiles = G.find_tiles(bbox, epsg, c['manifests'])
print('泗水相交瓦片', len(tiles), flush=True)

GX0 = np.floor(bx0/CELL)*CELL; GX1 = np.ceil(bx1/CELL)*CELL
GY0 = np.floor(by0/CELL)*CELL; GY1 = np.ceil(by1/CELL)*CELL
W2 = int(round((GX1-GX0)/CELL)); H2 = int(round((GY1-GY0)/CELL))
master = np.zeros((H2, W2), dtype=np.float32)

t0 = time.time()
for ti, t in enumerate(tiles):
    url = t['url']; a = t['affine']
    x0 = a['translateX']; y0 = a['translateY']; sx = a['scaleX']; sy = a['scaleY']; tw, th = t['w'], t['h']
    c_lo = max(0, int(np.floor((bx0-x0)/sx))); c_hi = min(tw, int(np.ceil((bx1-x0)/sx)))
    r_lo = max(0, int(np.floor((y0-by1)/(-sy)))); r_hi = min(th, int(np.ceil((y0-by0)/(-sy))))
    c_lo -= c_lo % POOL; r_lo -= r_lo % POOL; c_hi -= c_hi % POOL; r_hi -= r_hi % POOL
    if c_hi <= c_lo or r_hi <= r_lo:
        continue
    f = G.HttpRangeFile(url); p = tifffile.TiffFile(f).pages[0]
    offs = p.dataoffsets; cnts = p.databytecounts
    TGRID = math.ceil(tw/CHUNK); BANDS_PER = TGRID*math.ceil(th/CHUNK)
    H = r_hi-r_lo; Wd = c_hi-c_lo
    sub = np.zeros((H, Wd), dtype=np.float32)
    rb0, rb1 = r_lo//CHUNK, (r_hi-1)//CHUNK; cb0, cb1 = c_lo//CHUNK, (c_hi-1)//CHUNK
    blocks = [(rb, cb) for rb in range(rb0, rb1+1) for cb in range(cb0, cb1+1)]
    sess_local = threading.local()
    def sess():
        if not hasattr(sess_local, 's'): sess_local.s = requests.Session()
        return sess_local.s
    def fetch(block):
        rb, cb = block; idx = BANDS_PER + rb*TGRID + cb
        off = offs[idx]; cnt = cnts[idx]
        r = sess().get(url, headers={'Range': f'bytes={off}-{off+cnt-1}'}, timeout=120)
        raw = r.content
        rs = rb*CHUNK; re = min((rb+1)*CHUNK, th); cs = cb*CHUNK; ce = min((cb+1)*CHUNK, tw)
        return (rs, re, cs, ce, G.decode_tile(raw, CHUNK, CHUNK)[:re-rs, :ce-cs])
    with ThreadPoolExecutor(max_workers=G.NTHREADS) as ex:
        for rs, re, cs, ce, data in ex.map(fetch, blocks):
            ar0 = max(rs, r_lo); ar1 = min(re, r_hi); ac0 = max(cs, c_lo); ac1 = min(ce, c_hi)
            if ar1 <= ar0 or ac1 <= ac0: continue
            sub[ar0-r_lo:ar1-r_lo, ac0-c_lo:ac1-c_lo] = data[ar0-rs:ar1-rs, ac0-cs:ac1-cs]
    sp = sub.reshape(H//POOL, POOL, Wd//POOL, POOL).max(axis=(1, 3))
    sub_x0 = x0+c_lo*sx; sub_y0 = y0+r_lo*sy
    col_off = int(round((sub_x0-GX0)/CELL)); row_off = int(round((GY1-sub_y0)/CELL))
    r0 = max(0, row_off); c0 = max(0, col_off); r1 = min(H2, row_off+sp.shape[0]); c1 = min(W2, col_off+sp.shape[1])
    if r1 > r0 and c1 > c0:
        blk = sp[r0-row_off:r0-row_off+(r1-r0), c0-col_off:c0-col_off+(c1-c0)]
        np.maximum(master[r0:r1, c0:c1], blk, out=master[r0:r1, c0:c1])
    print('  瓦片%d/%d %ds max=%.1f' % (ti+1, len(tiles), time.time()-t0, master.max()), flush=True)

np.save(DATA+r'\googleob_surabaya_height_2m.npy', master)
struct = np.ones((3, 3), int)
user = {28: 1319, 35: 845, 42: 558, 56: 248, 70: 117}
print('\n泗水对比：我(连通块>=16m2) vs 你(GEE逐栋)', flush=True)
print('%5s %14s %12s %10s' % ('阈值', '我-像元面积km2', '我-连通块', '你-逐栋'), flush=True)
for thr in [28, 32, 35, 40, 42, 48, 56, 70]:
    m = master >= thr; npix = int(m.sum())
    lbl, n = ndimage.label(m, struct)
    sizes = np.bincount(lbl.ravel())[1:] if n else np.array([])
    n16 = int((sizes >= 4).sum())
    u = user.get(thr, '-')
    print('%4dm %14.3f %12d %10s' % (thr, npix*4/1e6, n16, u), flush=True)
print('我的泗水最高: %.1f m' % master.max(), flush=True)
