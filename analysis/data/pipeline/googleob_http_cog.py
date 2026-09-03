# -*- coding: utf-8 -*-
"""测试用 HTTP range 读取 COG 瓦片的元数据（不下载整文件）。"""
import requests, io

class HttpRangeFile(io.RawIOBase):
    """通过 HTTP Range 请求模拟可随机访问的文件对象。"""
    def __init__(self, url, session=None):
        self.url = url
        self.s = session or requests.Session()
        h = self.s.head(url, timeout=60)
        self.size = int(h.headers['content-length'])
        self.pos = 0
        self.nreq = 0
        self.nbytes = 0
    def seekable(self): return True
    def readable(self): return True
    def seek(self, off, whence=0):
        if whence == 0: self.pos = off
        elif whence == 1: self.pos += off
        elif whence == 2: self.pos = self.size + off
        return self.pos
    def tell(self): return self.pos
    def read(self, n=-1):
        if n is None or n < 0:
            n = self.size - self.pos
        if n == 0: return b''
        end = min(self.pos + n - 1, self.size - 1)
        hdr = {'Range': f'bytes={self.pos}-{end}'}
        r = self.s.get(self.url, headers=hdr, timeout=120)
        data = r.content
        self.pos += len(data)
        self.nreq += 1
        self.nbytes += len(data)
        return data
    def readinto(self, b):
        data = self.read(len(b))
        b[:len(data)] = data
        return len(data)

if __name__ == '__main__':
    import tifffile
    url = 'https://storage.googleapis.com/open-buildings-temporal-data/v1/geotiffs/2e69c_2023_06_30/tile_ySHmjn0JYAs.tif'
    f = HttpRangeFile(url)
    tif = tifffile.TiffFile(f)
    print('页数(分辨率层级)', len(tif.pages))
    for i, p in enumerate(tif.pages):
        print(f'  page{i}: shape={p.shape} dtype={p.dtype} tiled={p.is_tiled} '
              f'tile={getattr(p,"tilewidth",None)}x{getattr(p,"tilelength",None)} '
              f'bands(samples)={p.samplesperpixel} compression={p.compression}')
    print('HTTP请求数', f.nreq, '读取字节', f.nbytes)
