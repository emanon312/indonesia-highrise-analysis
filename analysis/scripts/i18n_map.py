# -*- coding: utf-8 -*-
"""英文交付物 · 翻译对照表（中文值 → 英文）。
楼名/地址/街道办区是印尼语原文，不在此表，保持原样。
用途分类(CATEGORY_EN)等 clean_category 清洗方案定了再补。"""

# 城市名（中文 → 英文）
CITY_EN = {
    '雅加达':'Jakarta', '泗水':'Surabaya', '万隆':'Bandung', '唐格朗':'Tangerang',
    '望加锡':'Makassar', '棉兰':'Medan', '巴淡':'Batam', '三宝垄':'Semarang',
    '勿加泗':'Bekasi', '德波':'Depok', '茂物':'Bogor', '巨港':'Palembang',
}

# 列名（中文 → 英文）
COLS_EN = {
    '序号':'No.', '楼宇名称':'Building Name', '名称来源':'Name Source',
    '数据来源':'Data Source', 'OSM匹配距离(米)':'OSM Match Dist. (m)',
    '建筑全名(官方)':'Official Full Name', '纬度':'Latitude', '经度':'Longitude',
    '高度(米)':'Height (m)', '层数':'Floors', '高度分档':'Height Tier',
    '用途分类':'Category', '行政区':'District', '街道办区':'Sub-district',
    '地址':'Address', '建筑面积(㎡)':'Floor Area (sqm)', '近距簇':'Cluster',
}

# 名称来源值（中文 → 英文）
NAME_SRC_EN = {
    '官方':'Official', '无':'None', 'OSM补充':'OSM',
    'OSM补充(存疑)':'OSM (uncertain)', 'CTBUH地标':'CTBUH Landmark',
}

# 数据来源值（中文 → 英文）
DATA_SRC_EN = {
    '官方(DPMPTSP)':'Official (DPMPTSP)',
    'GoogleV3坐标+2.5D高度':'Google V3 coord + 2.5D height',
    'GoogleV3坐标+CTBUH真高':'Google V3 coord + CTBUH height',
}

# 高度分档值（中文 → 英文）
TIER_EN = {
    '≥80米(约20层)':'≥80m (~20F)', '≥64米(约16层)':'≥64m (~16F)',
    '≥48米(约12层)':'≥48m (~12F)', '≥40米(约10层)':'≥40m (~10F)',
    '≥32米(约8层)':'≥32m (~8F)',
}

# 行政区含中文的值（中文 → 英文）；其余为印尼语，保留原文
DISTRICT_EN = {
    '南雅加达':'South Jakarta', '中雅加达':'Central Jakarta', '北雅加达':'North Jakarta',
    '西雅加达':'West Jakarta', '东雅加达':'East Jakarta',
}

# 用途分类值（中文 → 英文）—— 覆盖12城全部22种值（雅加达官方变体 + 11城清洗后规范集）
CATEGORY_EN = {
    '未分类':'Unclassified', '综合体':'Mixed-use', '写字楼':'Office', '公寓':'Apartment',
    '酒店':'Hotel', '医院':'Hospital', '商场':'Mall', '商业':'Commercial',
    '非住宅':'Non-residential', '商业/店铺':'Retail/Shophouse', '商业(店铺)':'Retail/Shophouse',
    '住宅':'Residential', '住宅类':'Residential', '工业':'Industrial', '其他':'Other',
    '商业类':'Commercial', '政府公共':'Government/Public', '仓库':'Warehouse', '宿舍':'Dormitory',
    '会展':'Convention', '办公住宅':'Office-Residential', '宗教':'Religious', '旅游休闲':'Leisure',
}

# 界面文案（英文交付物 UI）
UI = {
    'title':'Indonesia High-Rise Buildings — 12 Cities',
    'subtitle':'7,430 buildings ≥32m · Jakarta official (DPMPTSP), other cities Google estimate',
    'tab_overview':'Overview', 'tab_map':'Map', 'tab_list':'Building List',
    'export':'Export to CSV', 'search':'Search building / city / district...',
    'legend':'Height Tier', 'total':'Total Buildings', 'by_city':'By City',
    'by_tier':'By Height Tier', 'by_source':'By Data Source',
}
