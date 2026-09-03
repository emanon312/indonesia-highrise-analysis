# -*- coding: utf-8 -*-
"""通用：统计任一印尼城市 Google OB 2.5D Temporal(2023) 高层"下限"。

复用已在雅加达验证过的 COG 内部块解码（planarconfig=2 波段分离、deflate+float
预测器 predictor=3）+ 2m 工作分辨率（0.5m 做 4x4 最大池化保峰值）+ 8 连通域计数。
不保存全栅格（省磁盘），只输出 googleob_<city>_results.json。

用法: python googleob_city.py <city_id>
口径与雅加达完全一致，便于横向对比。数据为 ML 估算，仅作"下限/定性"。
"""
import sys, json, time, math, threading, os
DATA = r'analysis\data'
sys.path.insert(0, DATA)
from googleob_http_cog import HttpRangeFile
import tifffile, numpy as np, requests, pyproj
from scipy import ndimage
from concurrent.futures import ThreadPoolExecutor
import imagecodecs

# 各城：中文名、bbox(W,S,E,N 经纬度,覆盖建成区)、EPSG(UTM带)、候选 manifest(同EPSG多个时脚本自动遍历取相交)
CITIES = {
    'surabaya':  dict(cn='泗水',   bbox=(112.62,-7.36,112.84,-7.19), epsg=32749, manifests=['2d_EPSG_32749_2023_06_30.json','2f_EPSG_32749_2023_06_30.json']),
    'bandung':   dict(cn='万隆',   bbox=(107.55,-6.97,107.73,-6.86), epsg=32748, manifests=['2f_EPSG_32748_2023_06_30.json']),
    'bekasi':    dict(cn='勿加泗', bbox=(106.94,-6.30,107.06,-6.18), epsg=32748, manifests=['2f_EPSG_32748_2023_06_30.json']),
    'tangerang': dict(cn='唐格朗', bbox=(106.58,-6.27,106.74,-6.12), epsg=32748, manifests=['2f_EPSG_32748_2023_06_30.json']),
    'bogor':     dict(cn='茂物',   bbox=(106.76,-6.64,106.85,-6.54), epsg=32748, manifests=['2f_EPSG_32748_2023_06_30.json']),
    'medan':     dict(cn='棉兰',   bbox=(98.62,3.52,98.74,3.66),     epsg=32647, manifests=['31_EPSG_32647_2023_06_30.json','37_EPSG_32647_2023_06_30.json']),
    'semarang':  dict(cn='三宝垄', bbox=(110.36,-7.05,110.49,-6.95), epsg=32749, manifests=['2d_EPSG_32749_2023_06_30.json','2f_EPSG_32749_2023_06_30.json']),
    'makassar':  dict(cn='望加锡', bbox=(119.39,-5.19,119.52,-5.07), epsg=32750, manifests=['2d_EPSG_32750_2023_06_30.json']),
    'batam':     dict(cn='巴淡',   bbox=(103.98,1.03,104.15,1.18),   epsg=32648, manifests=['31_EPSG_32648_2023_06_30.json','37_EPSG_32648_2023_06_30.json']),
    'palembang': dict(cn='巨港',   bbox=(104.68,-3.04,104.82,-2.93), epsg=32748, manifests=['2f_EPSG_32748_2023_06_30.json']),
    'depok':     dict(cn='德波',   bbox=(106.76,-6.47,106.86,-6.35), epsg=32748, manifests=['2f_EPSG_32748_2023_06_30.json']),
}

BASE = 'https://storage.googleapis.com/open-buildings-temporal-data/'
RES = 0.5; POOL = 4; CELL = RES*POOL; CHUNK = 512; NTHREADS = 48


def decode_tile(raw, h, w):
    """deflate 解压 + 反 float 预测器(predictor=3) -> (h,w) float32（已与 tifffile 校验一致）"""
    buf = imagecodecs.deflate_decode(raw)
    data = np.frombuffer(buf, dtype=np.float32).reshape(h, w).copy()
    return imagecodecs.floatpred_decode(data, axis=-1)


def find_tiles(bbox, epsg, manifests):
    """遍历候选 manifest，返回 (bbox_utm, 相交瓦片列表)。瓦片含 url/affine/w/h。"""
    W, S, E, N = bbox
    tr = pyproj.Transformer.from_crs('EPSG:4326', f'EPSG:{epsg}', always_xy=True)
    xs, ys = [], []
    for lon in (W, E):
        for lat in (S, N):
            x, y = tr.transform(lon, lat); xs.append(x); ys.append(y)
    bx0, bx1 = min(xs), max(xs); by0, by1 = min(ys), max(ys)
    hits = []
    for name in manifests:
        try:
            m = requests.get(BASE+'v1/manifests/'+name, timeout=120).json()
        except Exception as ex:
            print('  manifest 读取失败', name, ex, flush=True); continue
        https_prefix = m['uriPrefix'].replace('gs://', 'https://storage.googleapis.com/')
        for ts in m['tilesets']:
            for src in ts['sources']:
                a = src['affineTransform']; d = src['dimensions']
                x0 = a['translateX']; y0 = a['translateY']; sx = a['scaleX']; sy = a['scaleY']
                w = d['width']; h = d['height']
                tx0, tx1 = x0, x0+sx*w; ty0, ty1 = y0, y0+sy*h
                t_x0, t_x1 = min(tx0, tx1), max(tx0, tx1); t_y0, t_y1 = min(ty0, ty1), max(ty0, ty1)
                if t_x1 < bx0 or t_x0 > bx1 or t_y1 < by0 or t_y0 > by1:
                    continue
                hits.append({'url': https_prefix+src['uris'][0], 'affine': a, 'w': w, 'h': h})
    return (bx0, by0, bx1, by1), hits


def main(city):
    c = CITIES[city]; bbox = c['bbox']; epsg = c['epsg']
    (bx0, by0, bx1, by1), tiles = find_tiles(bbox, epsg, c['manifests'])
    print(f'{city}({c["cn"]}) EPSG{epsg} 相交瓦片 {len(tiles)} 个', flush=True)
    out_path = os.path.join(DATA, f'googleob_{city}_results.json')
    if not tiles:
        json.dump({'city': city, 'cn': c['cn'], 'error': 'no intersecting tiles', 'bbox_wgs84': list(bbox), 'epsg': epsg},
                  open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        print('无相交瓦片，已记录退出', flush=True); return

    GX0 = np.floor(bx0/CELL)*CELL; GX1 = np.ceil(bx1/CELL)*CELL
    GY0 = np.floor(by0/CELL)*CELL; GY1 = np.ceil(by1/CELL)*CELL
    W2 = int(round((GX1-GX0)/CELL)); H2 = int(round((GY1-GY0)/CELL))
    print(f'主网格 {H2}x{W2} (2m)', flush=True)
    master = np.zeros((H2, W2), dtype=np.float32)

    t_start = time.time()
    for ti, t in enumerate(tiles):
        url = t['url']; a = t['affine']
        x0 = a['translateX']; y0 = a['translateY']; sx = a['scaleX']; sy = a['scaleY']; tw, th = t['w'], t['h']
        c_lo = max(0, int(np.floor((bx0-x0)/sx))); c_hi = min(tw, int(np.ceil((bx1-x0)/sx)))
        r_lo = max(0, int(np.floor((y0-by1)/(-sy)))); r_hi = min(th, int(np.ceil((y0-by0)/(-sy))))
        c_lo -= c_lo % POOL; r_lo -= r_lo % POOL; c_hi -= c_hi % POOL; r_hi -= r_hi % POOL
        if c_hi <= c_lo or r_hi <= r_lo:
            print(f'  瓦片{ti+1} 无重叠', flush=True); continue
        f = HttpRangeFile(url); p = tifffile.TiffFile(f).pages[0]
        offs = p.dataoffsets; cnts = p.databytecounts
        TGRID = math.ceil(tw/CHUNK); BANDS_PER = TGRID*math.ceil(th/CHUNK)  # 按瓦片尺寸动态算（雅加达=49/2401）
        H = r_hi-r_lo; Wd = c_hi-c_lo
        sub = np.zeros((H, Wd), dtype=np.float32)
        rb0, rb1 = r_lo//CHUNK, (r_hi-1)//CHUNK
        cb0, cb1 = c_lo//CHUNK, (c_hi-1)//CHUNK
        blocks = [(rb, cb) for rb in range(rb0, rb1+1) for cb in range(cb0, cb1+1)]
        sess_local = threading.local()
        def sess():
            if not hasattr(sess_local, 's'): sess_local.s = requests.Session()
            return sess_local.s
        def fetch(block):
            rb, cb = block
            idx = BANDS_PER + rb*TGRID + cb   # band1(height) 起始 = 一个波段的块数
            off = offs[idx]; cnt = cnts[idx]
            r = sess().get(url, headers={'Range': f'bytes={off}-{off+cnt-1}'}, timeout=120)
            raw = r.content
            rs = rb*CHUNK; re = min((rb+1)*CHUNK, th); cs = cb*CHUNK; ce = min((cb+1)*CHUNK, tw)
            bh = re-rs; bw = ce-cs
            full = decode_tile(raw, CHUNK, CHUNK)
            return (rs, re, cs, ce, full[:bh, :bw], len(raw))
        nb = 0
        with ThreadPoolExecutor(max_workers=NTHREADS) as ex:
            for rs, re, cs, ce, data, nbytes in ex.map(fetch, blocks):
                nb += nbytes
                ar0 = max(rs, r_lo); ar1 = min(re, r_hi); ac0 = max(cs, c_lo); ac1 = min(ce, c_hi)
                if ar1 <= ar0 or ac1 <= ac0: continue
                sub[ar0-r_lo:ar1-r_lo, ac0-c_lo:ac1-c_lo] = data[ar0-rs:ar1-rs, ac0-cs:ac1-cs]
        sp = sub.reshape(H//POOL, POOL, Wd//POOL, POOL).max(axis=(1, 3))
        sub_x0 = x0+c_lo*sx; sub_y0 = y0+r_lo*sy
        col_off = int(round((sub_x0-GX0)/CELL)); row_off = int(round((GY1-sub_y0)/CELL))
        r0 = max(0, row_off); c0 = max(0, col_off); r1 = min(H2, row_off+sp.shape[0]); c1 = min(W2, col_off+sp.shape[1])
        if r1 > r0 and c1 > c0:
            sr0 = r0-row_off; sc0 = c0-col_off
            blk = sp[sr0:sr0+(r1-r0), sc0:sc0+(c1-c0)]
            np.maximum(master[r0:r1, c0:c1], blk, out=master[r0:r1, c0:c1])
        print(f'  瓦片{ti+1}/{len(tiles)} 窗口{sub.shape} {len(blocks)}块 {nb/1e6:.0f}MB {time.time()-t_start:.0f}s max={master.max():.1f}', flush=True)

    cell_area = CELL**2; struct = np.ones((3, 3), dtype=int)
    res = {'city': city, 'cn': c['cn'], 'bbox_wgs84': list(bbox), 'epsg': epsg,
           'max_height_m': round(float(master.max()), 2), 'work_res_m': CELL,
           'source': 'Google Open Buildings 2.5D Temporal v1 year2023 (GCS, ML估算下限)'}
    for thr in (32, 40, 48):
        mask = master >= thr; npix = int(mask.sum())
        lbl, ncomp = ndimage.label(mask, structure=struct)
        if ncomp:
            sizes = np.bincount(lbl.ravel())[1:]; n16 = int((sizes >= 4).sum())
        else:
            n16 = 0
        res[f'ge{thr}'] = {'pixels_2m': npix, 'area_km2': round(npix*cell_area/1e6, 4), 'components_ge16m2': n16}
        print(f'>= {thr}m: 像元={npix} 面积={npix*cell_area/1e6:.3f}km2 连通块>=16m2={n16}', flush=True)
    res['elapsed_s'] = round(time.time()-t_start, 1)
    json.dump(res, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print(f'已保存 {out_path}  最高{master.max():.1f}m  耗时{res["elapsed_s"]}s', flush=True)


if __name__ == '__main__':
    main(sys.argv[1])
