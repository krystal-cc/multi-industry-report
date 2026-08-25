#!/usr/bin/env python3
"""将 爆品榜单0825.xlsx 的 Sheet7 数据按商品类目/商品名称归入前面四个行业 sheet。"""
import pandas as pd
import shutil

path = '/Users/krystalcao/Desktop/爆品榜单0825.xlsx'
bak = path + '.bak'
shutil.copy(path, bak)

xl = pd.ExcelFile(path)
sheet_names = xl.sheet_names

# 目标 sheet 的 10 列
TARGET_COLS = ['投放一级行业', '投放二级行业', '投放三级行业', '引流平台',
               '商品名称', '商品主图', '品牌', '落地页URL', '创意标题', '抖音视频链接']

# 读取全部 sheet
sheets = {s: pd.read_excel(path, sheet_name=s) for s in sheet_names}

# 从四个行业 sheet 现有数据推断「二级行业 -> 一级行业」映射
l1_map = {}
for s in ['消电日百', '食品饮料+阿里', '美妆个护', '服饰']:
    df = sheets[s]
    for l2, l1 in zip(df['投放二级行业'], df['投放一级行业']):
        if pd.notna(l2):
            l1_map[str(l2).strip()] = str(l1).strip()

# 行业名 -> 实际 sheet 名
SHEET_NAME = {'消电日百': '消电日百', '食品饮料': '食品饮料+阿里', '美护': '美妆个护', '服饰': '服饰'}

# 商品类目 -> (目标行业, 二级行业)
CAT_MAP = {
    # ===== 服饰 =====
    "时尚饰品": ("服饰", "时尚饰品"),
    "创意饰品": ("服饰", "时尚饰品"),
    "发饰": ("服饰", "时尚饰品"),
    "装扮用品": ("服饰", "时尚饰品"),
    "文玩饰品/把件": ("服饰", "时尚饰品"),
    "女装": ("服饰", "女装"),
    "男装": ("服饰", "男装"),
    "内衣": ("服饰", "内衣"),
    "服饰配件": ("服饰", "服饰配件"),
    "女包": ("服饰", "箱包"),
    "功能箱包": ("服饰", "箱包"),
    "腕表": ("服饰", "钟表眼镜"),
    "眼镜": ("服饰", "钟表眼镜"),
    "时尚女鞋": ("服饰", "女鞋"),
    "运动鞋": ("服饰", "运动户外及功能鞋"),
    "体育用品": ("服饰", "运动户外及功能鞋"),
    "健身训练": ("服饰", "运动户外及功能鞋"),
    "垂钓用品": ("服饰", "运动户外及功能鞋"),
    # ===== 食品饮料 =====
    "药食同源": ("食品饮料", "传统滋补"),
    "增强免疫": ("食品饮料", "传统滋补"),
    "心脑血管养护": ("食品饮料", "传统滋补"),
    "其他营养保健": ("食品饮料", "传统滋补"),
    "养生茶饮": ("食品饮料", "茶叶"),
    "茗茶": ("食品饮料", "茶叶"),
    "节庆食品": ("食品饮料", "休闲食品"),
    "膨化食品": ("食品饮料", "休闲食品"),
    "方便食品": ("食品饮料", "休闲食品"),
    "坚果炒货": ("食品饮料", "休闲食品"),
    "海味零食": ("食品饮料", "休闲食品"),
    "海鲜水产": ("食品饮料", "海鲜水产"),
    "蛋类": ("食品饮料", "肉禽蛋品"),
    "米": ("食品饮料", "粮油干货"),
    # ===== 美护 =====
    "身体护理": ("美护", "身体护理"),
    "面部护肤": ("美护", "面部洗护"),
    "口腔护理": ("美护", "口腔护理"),
    "洗发护发": ("美护", "洗发护发"),
    "美发造型": ("美护", "洗发护发"),
    "家庭环境清洁": ("美护", "家清纸品"),
    "衣物清洁/护理": ("美护", "家清纸品"),
    "驱蚊驱虫": ("美护", "家清纸品"),
    "清洁纸品": ("美护", "家清纸品"),
    "个人护理工具": ("美护", "美妆工具"),
    "女性护理": ("美护", "其他日化用品"),
    # ===== 消电日百 =====
    "厨用工具": ("消电日百", "厨具"),
    "烹饪锅具": ("消电日百", "厨具"),
    "餐具": ("消电日百", "厨具"),
    "一次性用品": ("消电日百", "厨具"),
    "厨房小电": ("消电日百", "厨房电器"),
    "生活电器": ("消电日百", "环境电器"),
    "贴饰": ("消电日百", "居家饰品"),
    "植物/仿真植物": ("消电日百", "居家饰品"),
    "装饰画": ("消电日百", "居家饰品"),
    "保健器械": ("消电日百", "个护健康电器"),
    "防护护理": ("消电日百", "个护健康电器"),
    "个护健康": ("消电日百", "个护健康电器"),
    "冷暖防护": ("消电日百", "个护健康电器"),
    "防尘罩/膜": ("消电日百", "家务工具"),
    "清洁用具": ("消电日百", "家务工具"),
    "电工电料": ("消电日百", "电子/电工"),
    "厨卫配件": ("消电日百", "家装建材"),
    "龙头": ("消电日百", "家装建材"),
    "淋浴花洒": ("消电日百", "家装建材"),
    "厨卫挂件": ("消电日百", "家装建材"),
    "基建材料": ("消电日百", "家装建材"),
    "手机配件": ("消电日百", "手机及配件"),
    "汽车装饰": ("消电日百", "手机及配件"),
    "收纳袋": ("消电日百", "收纳整理工具"),
    "收纳盒": ("消电日百", "收纳整理工具"),
    "收纳箱": ("消电日百", "收纳整理工具"),
    "收纳架": ("消电日百", "收纳整理工具"),
    "网络产品": ("消电日百", "3C数码设备及配件"),
    "维修保养": ("消电日百", "五金/工具"),
    "日用工具": ("消电日百", "五金/工具"),
    "电动工具": ("消电日百", "五金/工具"),
    "影音娱乐": ("消电日百", "影音电器"),
    "智能设备": ("消电日百", "影像与监控设备"),
    "美容清洗": ("消电日百", "其他家居建材"),  # 实为汽车美容清洁用品
}

# OTC 药品类目（不归入四个 sheet，保留在 Sheet7 不动）
OTC_CATS = {"非处方药", "风湿骨外用药OTC", "消化系统用药OTC", "补益用药OTC", "眼部用药OTC"}


def classify_special(name):
    """「特殊商品」类目按商品名称关键词归类"""
    if "水龙头" in name:
        return ("消电日百", "家装建材")
    if any(k in name for k in ["多维片", "虾青素", "酵素", "小橘片"]):
        return ("食品饮料", "传统滋补")
    return ("美护", "身体护理")  # 皮炎膏、外翻膏等皮肤护理


df7 = sheets['Sheet7']
rows_to_add = []   # (实际sheet名, 新行dict)
keep_idx = []      # Sheet7 保留的行（药品等）
unmapped = set()

for idx, row in df7.iterrows():
    cat = str(row['商品类目']).strip() if pd.notna(row['商品类目']) else ''
    name = str(row['商品名称']) if pd.notna(row['商品名称']) else ''
    if cat in OTC_CATS:
        keep_idx.append(idx)
        continue
    if cat == '特殊商品':
        target_ind, l2 = classify_special(name)
    elif cat in CAT_MAP:
        target_ind, l2 = CAT_MAP[cat]
    else:
        unmapped.add(cat)
        keep_idx.append(idx)
        continue
    l1 = l1_map.get(l2, '')
    new_row = {
        '投放一级行业': l1,
        '投放二级行业': l2,
        '投放三级行业': cat,
        '引流平台': str(row['引流平台']) if pd.notna(row['引流平台']) else '',
        '商品名称': name,
        '商品主图': '',
        '品牌': '',
        '落地页URL': str(row['落地页']) if pd.notna(row['落地页']) else '',
        '创意标题': '',
        '抖音视频链接': '',
    }
    rows_to_add.append((SHEET_NAME[target_ind], new_row))

if unmapped:
    print('⚠️ 未映射类目（已保留在 Sheet7）:', unmapped)

# 追加到对应行业 sheet
from collections import defaultdict
add_by_sheet = defaultdict(list)
for sn, r in rows_to_add:
    add_by_sheet[sn].append(r)

for sn, rows in add_by_sheet.items():
    sheets[sn] = pd.concat([sheets[sn], pd.DataFrame(rows, columns=TARGET_COLS)], ignore_index=True)

# Sheet7 只保留药品（及未映射）
sheets['Sheet7'] = df7.loc[keep_idx]

# 写回（保持原 sheet 顺序）
with pd.ExcelWriter(path, engine='openpyxl') as writer:
    for s in sheet_names:
        sheets[s].to_excel(writer, sheet_name=s, index=False)

print('=' * 60)
print('✅ 归并完成，备份文件:', bak)
print('Sheet7 原始行数:', len(df7), '→ 归入四个 sheet:', len(rows_to_add), '→ Sheet7 保留:', len(keep_idx))
for sn, rows in add_by_sheet.items():
    print(f'   追加到 {sn}: +{len(rows)} 行')
