# -*- coding: utf-8 -*-
"""多线程分段续传下载器：从已下字节处起，把剩余部分切 N 段并行 Range GET 下载，绕过单连接限速。
用法: python fast_download.py <url> <path> [threads]
"""
import sys, os, requests, threading

url = sys.argv[1]
path = sys.argv[2]
N = int(sys.argv[3]) if len(sys.argv) > 3 else 16

total = int(requests.head(url, timeout=60).headers['content-length'])
start = os.path.getsize(path) if os.path.exists(path) else 0
print('总大小 %.2fG  已下 %.2fG  待下 %.2fG  线程 %d' % (total/1e9, start/1e9, (total-start)/1e9, N), flush=True)
if start >= total:
    print('已完整'); sys.exit()

remaining = total - start
chunk = remaining // N
ranges = []
for i in range(N):
    s = start + i*chunk
    e = (start + (i+1)*chunk - 1) if i < N-1 else (total - 1)
    ranges.append((i, s, e))

done = [0]*N
lock = threading.Lock()


def dl(i, s, e):
    seg = '%s.seg%d' % (path, i)
    have = os.path.getsize(seg) if os.path.exists(seg) else 0  # 段级续传
    if s+have > e:
        done[i] = e-s+1; return
    r = requests.get(url, headers={'Range': 'bytes=%d-%d' % (s+have, e)}, stream=True, timeout=180)
    with open(seg, 'ab') as f:
        for ch in r.iter_content(1 << 20):
            if ch:
                f.write(ch)
                with lock:
                    done[i] += len(ch)


ths = [threading.Thread(target=dl, args=r) for r in ranges]
for t in ths:
    t.start()
for t in ths:
    t.join()

# 校验各段大小后顺序拼接到主文件末尾
ok = True
for i, s, e in ranges:
    seg = '%s.seg%d' % (path, i)
    exp = e-s+1
    got = os.path.getsize(seg) if os.path.exists(seg) else 0
    if got != exp:
        print('段%d 不完整 %d/%d' % (i, got, exp), flush=True); ok = False
if not ok:
    print('有段不完整，重跑本脚本可段级续传'); sys.exit(1)
with open(path, 'ab') as out:
    for i, s, e in ranges:
        seg = '%s.seg%d' % (path, i)
        with open(seg, 'rb') as f:
            while True:
                b = f.read(1 << 22)
                if not b: break
                out.write(b)
        os.remove(seg)
print('完成 %.2fG / %.2fG' % (os.path.getsize(path)/1e9, total/1e9), flush=True)
