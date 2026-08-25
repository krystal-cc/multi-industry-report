#!/usr/bin/env python3
"""将「爆品数据2.xlsx」的数据按行业归并入「爆品榜单0825.xlsx」。

字段映射（依据用户选择）：
- 商品名称 <- 营销对象ID(翻译后)
- 抖音视频链接 <- 素材URL(创意唯一)
- 引流平台 <- 腾讯广告(gdt)
- 投放一级/二级/三级行业 <- 按商品名称自动判断补齐
- 日均消耗(元) 不保留；商品主图/品牌/落地页URL/创意标题 留空
"""
import pandas as pd
import shutil

SRC = '/Users/krystalcao/Desktop/爆品数据2.xlsx'
DST = '/Users/krystalcao/Desktop/爆品榜单0825.xlsx'
bak = DST + '.bak2'
shutil.copy(DST, bak)

TARGET_COLS = ['投放一级行业', '投放二级行业', '投放三级行业', '引流平台',
               '商品名称', '商品主图', '品牌', '落地页URL', '创意标题', '抖音视频链接']

# 数据2 sheet -> 榜单 sheet
SHEET_MAP = {'消电日百': '消电日百', '食品饮料': '食品饮料+阿里', '美护': '美妆个护', '服饰': '服饰'}

# 商品名 -> (二级行业, 三级行业)
NAME_MAP = {
    # 消电日百
    '双刀头剃须刀': ('个护健康电器', '男士理容'),
    '包挂件': ('居家饰品', '摆件'),
    '家居饰品': ('居家饰品', '摆件'),
    '装饰摆件': ('居家饰品', '摆件'),
    '床单': ('床上用品', '床单'),
    '电动工具': ('五金/工具', '电动工具'),
    '生活用品': ('其他家居建材', '其他家居建材'),
    '装饰画': ('居家饰品', '装饰画'),
    '植物大师艾灸贴': ('个护健康电器', '家用护理辅助器材'),
    # 食品饮料
    '53度5L酱香型白酒': ('酒类', '国产白酒'),
    '菜谱式调料': ('粮油干货', '调味品'),
    '调味酱': ('粮油干货', '调味品'),
    '水产类干货': ('海鲜水产', '干货'),
    '御生堂牌糖脂茶': ('茶叶', '代用/花草/水果/再加工茶'),
    '五仁味月饼': ('休闲食品', '糕点/点心/面包'),
    '袋装方便面': ('休闲食品', '方便速食'),
    '水产罐头': ('海鲜水产', '鲜活水产'),
    '康生泰®西洋参含片': ('传统滋补', '其他传统滋补营养品'),
    # 美护
    '沐浴露': ('身体清洁', '身体清洁'),
    '管道疏通剂': ('家清纸品', '厨卫清洁'),
    '地面清洁剂': ('家清纸品', '厨卫清洁'),
    '白云山/BAIYUNSHAN眼部抗皱紧致植萃精华油': ('眼部护理', '眼部护理'),
    '颜大师美白祛痘洗面奶': ('面部洗护', '面部护肤'),
    '油污清洁剂': ('家清纸品', '厨卫清洁'),
    '透之谜/TOUZHIMI美白补水保湿素颜霜': ('面部洗护', '面部护肤'),
    '植物大师/PLANT MASTER光感美白祛斑霜': ('面部洗护', '祛斑美白'),
    '冰箱清洁剂': ('家清纸品', '厨卫清洁'),
    # 服饰
    '平底低帮时尚休闲鞋': ('女鞋', '低帮鞋'),
    '圆形男女通用老花眼镜': ('钟表眼镜', '眼镜'),
}

# 读取榜单
dst_sheets = {s: pd.read_excel(DST, sheet_name=s) for s in pd.ExcelFile(DST).sheet_names}

# 从榜单现有数据构建「二级行业 -> 一级行业」映射
l1_map = {}
for s in ['消电日百', '食品饮料+阿里', '美妆个护', '服饰']:
    df = dst_sheets[s]
    for l2, l1 in zip(df['投放二级行业'], df['投放一级行业']):
        if pd.notna(l2):
            l1_map.setdefault(str(l2).strip(), str(l1).strip())

unmatched = []
add_by_sheet = {}
for src_sheet in pd.ExcelFile(SRC).sheet_names:
    df = pd.read_excel(SRC, sheet_name=src_sheet)
    dst_sheet = SHEET_MAP[src_sheet]
    rows = []
    for _, r in df.iterrows():
        name = str(r['营销对象ID(翻译后)']) if pd.notna(r['营销对象ID(翻译后)']) else ''
        video = str(r['素材URL(创意唯一)']) if pd.notna(r['素材URL(创意唯一)']) else ''
        if name in NAME_MAP:
            l2, l3 = NAME_MAP[name]
        else:
            unmatched.append((src_sheet, name))
            continue
        l1 = l1_map.get(l2, '')
        rows.append({
            '投放一级行业': l1,
            '投放二级行业': l2,
            '投放三级行业': l3,
            '引流平台': '腾讯广告',
            '商品名称': name,
            '商品主图': '',
            '品牌': '',
            '落地页URL': '',
            '创意标题': '',
            '抖音视频链接': video,
        })
    if rows:
        add_by_sheet[dst_sheet] = add_by_sheet.get(dst_sheet, []) + rows

if unmatched:
    print('⚠️ 未匹配商品（已跳过）:')
    for s, n in unmatched:
        print(f'   [{s}] {n}')

total = 0
for dst_sheet, rows in add_by_sheet.items():
    dst_sheets[dst_sheet] = pd.concat(
        [dst_sheets[dst_sheet], pd.DataFrame(rows, columns=TARGET_COLS)], ignore_index=True)
    total += len(rows)
    print(f'   追加到 {dst_sheet}: +{len(rows)} 行')

# 写回（保持原 sheet 顺序）
with pd.ExcelWriter(DST, engine='openpyxl') as writer:
    for s in pd.ExcelFile(bak).sheet_names:
        dst_sheets[s].to_excel(writer, sheet_name=s, index=False)

print('=' * 60)
print('✅ 归并完成，备份:', bak)
print(f'共并入 {total} 行')
