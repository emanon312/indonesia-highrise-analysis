# 印尼 12 城高层建筑分析

对印尼 12 座主要城市的高层建筑（≥32 米 / 约 8 层以上）做的逐栋普查，
产出可检索的建筑清单（Excel）、交互式分布地图（HTML/KML）和汇总报告。

**成果规模**：12 城共识别高层建筑 **4801 栋**（每栋单算口径）；
若把紧邻成群的楼合并为综合体计算则为 **4462 个**。
覆盖：雅加达 2556 · 泗水 635 · 万隆 357 · 唐格朗 336 · 望加锡 183 ·
棉兰 169 · 巴淡 154 · 三宝垄 118 · 勿加泗 84 · 茂物 74 · 巨港 70 · 德波 65。

> English: A building-by-building census of high-rises (≥32 m / ~8 floors)
> across 12 Indonesian cities, producing Excel lists, an interactive map,
> and a summary report. 4,801 buildings identified.

## 目录结构

```text
├── analysis/
│   ├── output/                  # ★ 最终成果：12 城 Excel 清单、汇总表、
│   │                            #   交互地图 HTML/KML、建筑坐标 CSV
│   │   └── en/                  # 英文版 Excel 清单
│   ├── scripts/                 # 交付物生成脚本（Excel/地图/KML/交付包打包）
│   │   └── archive/             # 已跑完的一次性处理脚本（去重、清洗、加列等）
│   ├── data/pipeline/           # 数据获取与清洗管线（OSM/官方源/Google OB）
│   ├── reference/               # 20 城 GDP/人口参考表
│   ├── indonesia-highrise-analysis-plan.md   # 方案设计文档
│   ├── 印尼高层建筑分析_中文交付包.zip        # 可直接分发的成品包（中文）
│   └── Indonesia_HighRise_EN_Package.zip     # 可直接分发的成品包（英文）
├── README.md
└── .gitignore
```

## 数据来源与许可

| 来源 | 用途 | 许可 |
|------|------|------|
| [Google Open Buildings v3](https://sites.research.google/open-buildings/) | 建筑足迹 + 2.5D 高度栅格 | [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)，须署名 "© Google Open Buildings, used under CC BY 4.0" |
| [OpenStreetMap](https://openstreetmap.org) | 建筑高度/层数、行政边界 | [ODbL](https://opendatacommons.org/licenses/odbl/) |
| Jakarta Satu / 各市官方 GeoServer | 官方逐栋登记（雅加达等） | 各市公开数据条款 |
| GADM 4.1 | 行政区划边界 | 学术免费，[条款](https://gadm.org/license.html) |

**数据可信度提示**：不同城市数据源质量差异大（雅加达为官方逐栋登记，可信度最高；
其余城市以 OSM/Google OB 为主，高度字段为估计值），使用前请阅读 `analysis/output/12城高层建筑汇总.md` 中的口径说明。

## 复现

原始下载数据（约 4GB）不入库，按以下顺序用脚本可重新生成（在仓库根目录运行）：

```bash
cd analysis/data/pipeline
python fetch_osm_full.py <city>        # OSM 建筑 + 边界
python extract_v3.py <city>            # Google Open Buildings v3 足迹
python googleob_city.py <city>         # Google 2.5D 高度栅格
python build_city_highrise_v2.py <city># 合并、过滤 ≥32m、行政分区
cd ../../scripts
python build_excel.py                  # 生成 12 城 Excel
python build_kml.py && python build_deliverable.py  # 地图与交付包
```

## 状态

项目已结项（2026-07）。本仓库为最终成果与处理脚本的公开存档。
