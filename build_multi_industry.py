#!/usr/bin/env python3
"""
多行业爆品选品报告生成器
读取爆品数据Excel，生成支持行业切换的HTML报告。
四个行业：服饰、消电日百、食品饮料、美护（暂无数据）
"""

import pandas as pd
import json
import os
import re
from datetime import datetime
from collections import Counter

# 行业展示顺序
INDUSTRY_ORDER = ["消电日百", "食品饮料", "美护", "服饰"]

# 数据源配置：每个数据源对应一个时间版本（列表顺序决定默认显示顺序，第一个为默认显示）
# sheets 为「行业名 -> sheet名匹配前缀」映射（新老数据的 sheet 命名不同）
DATA_SOURCES = [
    {
        "key": "0817",
        "label": "8月17日更新",
        "path": "/Users/krystalcao/Desktop/【0817】爆品榜单.xlsx",
        "sheets": {"消电日百": "消电日百", "食品饮料": "食品", "美护": "美妆个护", "服饰": "服饰"},
    },
    {
        "key": "0813",
        "label": "8月13日更新",
        "path": "/Users/krystalcao/Desktop/【0813】爆品数据.xlsx",
        "sheets": {"消电日百": "消电日百", "食品饮料": "食品饮料", "美护": "美护", "服饰": "服饰"},
    },
]

# 二级行业归类映射（合并过小的类目）
INDUSTRY_MERGE = {
    "服饰": {
        "服饰配件": "其他",
        "其他服装配饰": "其他",
        "运动户外及功能鞋": "其他",
        "黄金珠宝": "其他",
        "女鞋": "其他",
    },
    "消电日百": {
        "其他家居建材": "其他",
        "其他3C及电器": "其他",
        "影音设备及配件": "其他",
        "清洁电器": "其他",
        "智能设备": "其他",
        "家具": "其他",
        "影音电器": "其他",
        "电子/电工": "其他",
        "影像与监控设备": "其他",
        "家装灯饰光源": "其他",
        "元器件": "其他",
        "3C数码设备及配件": "其他",
    },
    "食品饮料": {
        "其他食品饮料": "其他",
        "出行用品": "其他",
        "肉禽蛋品": "其他",
        "宠物生活": "宠物相关",
        "喂养用品": "宠物相关",
    },
    "美护": {
        "美甲工具": "其他",
        "眼部彩妆": "其他",
        "美妆工具": "其他",
    },
}

INDUSTRY_CONFIG = {
    "服饰": {
        "emoji": "👗",
        "color": "#c49a6c",           # 柔和暖棕
        "color_dark": "#a77b52",       # 深暖棕
        "color_deep": "#8b6540",       # 深咖
        "color_text": "#a77b52",       # 文字/数字
        "color_light": "#fdf6f0",      # 极浅米色
        "color_bg": "#fefaf6",         # 暖白底
        "color_border": "#e8cfb3",     # 浅暖咖边框
        "color_btn_hover": "#6b4a2c",  # 深咖
        "shadow_color": "rgba(180,140,100,0.15)",
    },
    "消电日百": {
        "emoji": "🏠",
        "color": "#3498db",
        "color_dark": "#2980b9",
        "color_deep": "#2471a3",
        "color_text": "#2980b9",
        "color_light": "#ebf5fb",
        "color_bg": "#f4f9fd",
        "color_border": "#d4e6f9",
        "color_btn_hover": "#1a5276",
        "shadow_color": "rgba(52,152,219,0.15)",
    },
    "食品饮料": {
        "emoji": "🍜",
        "color": "#27ae60",
        "color_dark": "#1e8449",
        "color_deep": "#1b6e3a",
        "color_text": "#426b50",
        "color_light": "#eaf2eb",
        "color_bg": "#f2f7f3",
        "color_border": "#dbe6e0",
        "color_btn_hover": "#224332",
        "shadow_color": "rgba(39,174,96,0.15)",
    },
    "美护": {
        "emoji": "💄",
        "color": "#d4879a",
        "color_dark": "#b87080",
        "color_deep": "#9c5a6a",
        "color_text": "#b87080",
        "color_light": "#fdf2f4",
        "color_bg": "#fef9fa",
        "color_border": "#f0d0d7",
        "color_btn_hover": "#7d4252",
        "shadow_color": "rgba(200,130,150,0.15)",
    },
}

def find_sheet_name(excel_path, target_name):
    """在Excel中查找匹配的sheet名（支持带（数字）后缀的sheet名）"""
    xl = pd.ExcelFile(excel_path)
    for s in xl.sheet_names:
        if s == target_name or s.startswith(target_name):
            return s
    return None

def read_sheet(excel_path, industry_name, sheet_prefix):
    """从指定 Excel 读取某行业 sheet（sheet_prefix 为 sheet 名匹配前缀）"""
    actual_name = find_sheet_name(excel_path, sheet_prefix)
    if actual_name is None:
        return None
    df = pd.read_excel(excel_path, actual_name)
    if df.empty:
        return None
    # 合并过小的二级类目
    merge_map = INDUSTRY_MERGE.get(industry_name, {})
    if merge_map:
        df["投放二级行业"] = df["投放二级行业"].apply(lambda x: merge_map.get(x, x))
    return df


def load_industries_data(source):
    """根据数据源配置读取四个行业的数据（按 INDUSTRY_ORDER 顺序）"""
    industries_data = []
    for industry_name in INDUSTRY_ORDER:
        sheet_prefix = source["sheets"].get(industry_name, industry_name)
        df = read_sheet(source["path"], industry_name, sheet_prefix)
        if df is None:
            cfg = INDUSTRY_CONFIG[industry_name]
            ind_data = {
                "name": industry_name,
                "emoji": cfg["emoji"],
                "color": cfg["color"],
                "color_dark": cfg["color_dark"],
                "color_deep": cfg["color_deep"],
                "color_text": cfg["color_text"],
                "color_light": cfg["color_light"],
                "color_bg": cfg["color_bg"],
                "color_border": cfg["color_border"],
                "color_btn_hover": cfg["color_btn_hover"],
                "shadow_color": cfg["shadow_color"],
                "total": 0,
                "categories": [],
                "keywords": [],
                "insights": [],
                "platforms": {},
            }
        else:
            ind_data = build_industry_data(industry_name, df)
        industries_data.append(ind_data)
    return industries_data

def build_industry_data(sheet_name, df):
    """将一个行业的DataFrame转为报告所需的data结构"""
    total = len(df)
    
    # 二级行业分组（按数量降序，"其他"排最后）
    cat_counts = df.groupby("投放二级行业").size().sort_values(ascending=False)
    cat_list = list(cat_counts.items())
    # "其他"放最后
    cat_list.sort(key=lambda x: (x[0] == "其他", -x[1]))
    
    categories = []
    for cat_name, count in cat_list:
        cat_df = df[df["投放二级行业"] == cat_name]
        products = []
        for _, row in cat_df.iterrows():
            # 处理商品主图（NaN或空字符串都视为无图）
            img = row.get("商品主图")
            img_url = str(img) if pd.notna(img) and str(img).strip() and str(img) != "nan" else ""
            # 处理商品名称
            name = row.get("商品名称")
            name_str = str(name) if pd.notna(name) else ""
            
            products.append({
                "name": name_str,
                "image_url": img_url,
                "tags": [str(row.get("投放三级行业", ""))] if pd.notna(row.get("投放三级行业")) else [],
                "source": str(row.get("引流平台", "")) if pd.notna(row.get("引流平台")) else "",
                "copy": str(row.get("创意标题", "")) if pd.notna(row.get("创意标题")) else "",
                "video_link": str(row.get("抖音视频链接", "")) if pd.notna(row.get("抖音视频链接")) else "#",
                "link": str(row.get("落地页URL", "")) if pd.notna(row.get("落地页URL")) else "#",
            })
        
        categories.append({
            "name": cat_name,
            "emoji": get_cat_emoji(sheet_name, cat_name),
            "count": count,
            "percentage": f"{count/total*100:.1f}%",
            "products": products,
        })
    
    # 高频词统计（从商品名称和创意标题中提取）
    keywords = extract_keywords(df)
    
    # 行业特定的洞察
    insights = get_insights(sheet_name, df)
    
    cfg = INDUSTRY_CONFIG[sheet_name]
    return {
        "name": sheet_name,
        "emoji": cfg["emoji"],
        "color": cfg["color"],
        "color_dark": cfg["color_dark"],
        "color_deep": cfg["color_deep"],
        "color_text": cfg["color_text"],
        "color_light": cfg["color_light"],
        "color_bg": cfg["color_bg"],
        "color_border": cfg["color_border"],
        "color_btn_hover": cfg["color_btn_hover"],
        "shadow_color": cfg["shadow_color"],
        "total": total,
        "categories": categories,
        "keywords": keywords[:20],
        "insights": insights,
        "platforms": dict(df["引流平台"].value_counts()),
    }

CAT_EMOJI_MAP = {
    # 服饰
    "时尚饰品": "💍", "内衣": "🩲", "男装": "👔", "钟表眼镜": "⌚",
    "男鞋": "👞", "女装": "👗", "箱包": "👜", "其他": "📦",
    # 消电日百
    "环境电器": "🌀", "家装建材": "🔨", "厨具": "🍳", "厨房电器": "🔌",
    "大家电": "📺", "五金/工具": "🔧", "个护健康电器": "💇",
    "居家饰品": "🖼️", "手机及配件": "📱", "电器配件": "🔋",
    "居家布艺": "🧵", "收纳整理工具": "📦", "火机烟具": "🔥",
    "床上用品": "🛏️", "家务工具": "🧹", "影音电器": "🎵",
    "电子/电工": "⚡", "影像与监控设备": "📷", "家装灯饰光源": "💡",
    "元器件": "🔩", "3C数码设备及配件": "💻",
    # 食品饮料
    "宠物相关": "🐾", "海外膳食营养品": "💊", "粮油干货": "🌾",
    "膳食营养品": "💪", "休闲食品": "🍪", "传统滋补": "🍯",
    "茶叶": "🍵", "水饮冲调": "🥤",
    "水果蔬菜": "🍓", "酒类": "🍶",
    # 美护
    "面部洗护": "🧴", "美妆特殊化妆品": "✨", "眼部护理": "👁️",
    "面部彩妆": "💄", "香水/香膏": "🌸", "唇部彩妆": "💋",
    "其他": "📦",
}

def get_cat_emoji(industry, cat_name):
    if cat_name in CAT_EMOJI_MAP:
        return CAT_EMOJI_MAP[cat_name]
    return "📦"

def extract_keywords(df):
    """从商品名称中提取高频词"""
    # 常见无用词
    stop_words = {"【", "】", "!", "！", "。", "，", ",", " ", "", "-", "_",
                  "的", "了", "在", "是", "我", "有", "和", "就", "不", "人", "都",
                  "一", "个", "上", "也", "很", "到", "说", "要", "去", "你",
                  "会", "着", "没有", "看", "好", "自己", "这", "【", "】",
                  "抢", "!","拍", "发", "包", "装", "送", "款", "新", "L1", "L2", "L3", "L4", "L5",
                  "L", "J", "Y", "R", "Z", "R1", "JY4", "0371", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
                  "5979", "多送", "正宗", "即食", "新鲜", "包邮", "特价", "热卖", "抢购",
                  "【7天内发货】", "【5天内发货】", "【48小时发货】", "【24小时发货】"}
    
    all_words = []
    for text in df["商品名称"].dropna():
        # 简单分词：按常见分隔符切分
        for seg in text.replace("【", " ").replace("】", " ").replace("!", " ").replace("！", " ").split():
            seg = seg.strip()
            if len(seg) >= 2 and seg not in stop_words and not seg.isdigit():
                all_words.append(seg)
    
    counter = Counter(all_words)
    top = counter.most_common(25)
    return [{"word": w, "count": c} for w, c in top if c >= 2]

def get_insights(sheet_name, df):
    """方案A：根据行业数据自动分析生成洞察密码（动态词频分析，不硬编码）"""
    
    # 合并所有文本用于分析
    all_names = " ".join(df["商品名称"].dropna().astype(str))
    all_copy = ""
    if "创意标题" in df.columns:
        all_copy = " ".join(df["创意标题"].dropna().astype(str))
    all_text = all_names + " " + all_copy
    total_products = len(df)
    
    # 类目分布
    cat_counts = df.groupby("投放二级行业").size().sort_values(ascending=False)
    top_cats = list(cat_counts.items())
    
    # 平台分布
    plat_counts = {}
    top_plat_name = ""
    top_plat_pct = ""
    if "引流平台" in df.columns:
        plat_counts = df["引流平台"].value_counts()
        if len(plat_counts) > 0:
            top_plat_name = str(plat_counts.index[0])
            top_plat_pct = f"{plat_counts.iloc[0]/total_products*100:.1f}%"
    
    # 主题检测规则（按行业通用+特定行业覆盖）
    theme_rules = [
        {
            "title": "情感送礼场景绑定",
            "emoji": "🎁", "color": "#e74c3c",
            "patterns": ["送朋友", "送爱人", "送自己", "礼物", "七夕", "节日", "情人节", "送礼", "生日", "送女友", "送妈妈"],
            "desc_tpl": "送礼/情感场景关键词累计出现{count}次，{top_words}等词高频出现，将商品与情感场景绑定是核心转化驱动力。"
        },
        {
            "title": "极致性价比轰炸",
            "emoji": "💰", "color": "#f39c12",
            "patterns": ["两位数", "工厂", "直销", "便宜", "实惠", "包邮", "平价", "批发", "源头", "白菜价", "清仓", "亏本", "仅需", "只要", "低至", "起"],
            "desc_tpl": "{top_words}等价格锚点话术累计{count}次，用极致性价比降低购买决策门槛。"
        },
        {
            "title": "视觉冲击+穿着效果",
            "emoji": "✨", "color": "#e67e22",
            "patterns": ["显身材", "防走光", "像没穿", "上身", "显瘦", "修身", "舒适", "冰丝", "透气", "亲肤", "不闷", "不扎", "裸感", "无痕", "高弹"],
            "desc_tpl": "{top_words}等穿着体验词累计{count}次，聚焦上身效果和痛点解决，用强烈画面感激发购买欲。"
        },
        {
            "title": "新款/应季追新",
            "emoji": "🔥", "color": "#f39c12",
            "patterns": ["新款", "2026", "限量", "首发", "新品", "爆款", "抢", "手慢无", "库存", "断货"],
            "desc_tpl": "{top_words}等追新紧迫词累计{count}次，利用消费者追新和稀缺心理制造购买紧迫感。"
        },
        {
            "title": "懒人便捷+解放双手",
            "emoji": "⏰", "color": "#2ecc71",
            "patterns": ["免安装", "一键", "全自动", "免煮", "即食", "冲泡", "方便", "速食", "懒人", "省力", "省时", "快捷", "开袋即食", "免动手", "智能"],
            "desc_tpl": "{top_words}等便捷关键词累计{count}次，快节奏生活下\"省时省力\"是爆款第一准则。"
        },
        {
            "title": "健康安全品质升级",
            "emoji": "🌿", "color": "#27ae60",
            "patterns": ["除菌", "净化", "无油", "健康", "无添加", "高纯净", "好吸收", "天然", "有机", "零添加", "非油炸", "母婴", "安全"],
            "desc_tpl": "{top_words}等健康品质关键词累计{count}次，消费者对健康安全关注度持续走高，品质升级是溢价关键。"
        },
        {
            "title": "多功能一物多用",
            "emoji": "🏡", "color": "#8e44ad",
            "patterns": ["多功能", "一物多用", "多用", "两用", "三合一", "多合一", "万能", "百搭"],
            "desc_tpl": "{top_words}等多功能标签累计{count}次，一物多用降低用户决策成本，是提升客单价的有效策略。"
        },
        {
            "title": "宠物经济大爆发",
            "emoji": "🐶", "color": "#27ae60",
            "patterns": ["冻干", "无谷", "宠物", "猫", "狗", "毛孩子", "猫咪", "狗狗", "猫粮", "狗粮", "猫砂"],
            "desc_tpl": "宠物相关关键词累计{count}次，{top_words}是核心卖点，宠物主愿意为毛孩子健康高品质买单。"
        },
        {
            "title": "功效护肤/美白淡斑",
            "emoji": "✨", "color": "#e91e63",
            "patterns": ["美白", "淡斑", "祛斑", "烟酰胺", "377", "祛痘", "淡印", "提亮", "去黄", "防晒", "遮瑕"],
            "desc_tpl": "{top_words}等功效成分词累计{count}次，功效型护肤需求持续爆发，\"成分硬核\"是护肤爆款核心密码。"
        },
        {
            "title": "成分温和护肤",
            "emoji": "🧴", "color": "#9c27b0",
            "patterns": ["氨基酸", "温和", "控油", "玻尿酸", "精华", "水杨酸", "维C", "胶原", "舒缓", "修复", "补水", "保湿", "洁面"],
            "desc_tpl": "{top_words}等成分/功效关键词累计{count}次，成分党驱动的精准护肤需求是稳定基本盘。"
        },
        {
            "title": "零风险信任背书",
            "emoji": "🛡️", "color": "#e74c3c",
            "patterns": ["不满意包退", "不好吃退", "不好用退", "无理由退", "退换", "包退", "试用不满意", "放心买", "随时退"],
            "desc_tpl": "零风险承诺话术累计{count}次，\"不满意包退\"是打消疑虑、促成购买转化的最后一把火。"
        },
        {
            "title": "彩妆品质感进阶",
            "emoji": "💄", "color": "#e040fb",
            "patterns": ["持妆", "显色", "高级感", "哑光", "水光", "丝绒", "雾面", "不脱妆", "不沾杯", "口红", "唇釉"],
            "desc_tpl": "{top_words}等品质感关键词累计{count}次，高级妆效和持妆力是彩妆爆款的标配卖点。"
        },
        {
            "title": "养生滋补升级",
            "emoji": "💊", "color": "#f39c12",
            "patterns": ["膳食", "营养", "滋补", "高蛋白", "养生", "调理", "燕窝", "枸杞", "参", "草本", "药食同源", "传统", "煲汤"],
            "desc_tpl": "{top_words}等养生滋补词累计{count}次，大健康概念升级，传统滋补赛道持续走热。"
        },
        {
            "title": "工艺/材质安心感",
            "emoji": "💎", "color": "#c0392b",
            "patterns": ["浮雕", "活口", "可调节", "纯银", "不锈钢", "高档", "精致", "匠心", "手工", "精工", "真皮"],
            "desc_tpl": "{top_words}等工艺材质词累计{count}次，用工艺细节支撑价格，增强信任感和价值感。"
        },
    ]
    
    # 计算每个主题的得分
    scored = []
    for rule in theme_rules:
        total_count = 0
        matched_words = []
        for pat in rule["patterns"]:
            c = all_text.count(pat)
            if c > 0:
                total_count += c
                matched_words.append((pat, c))
        if total_count >= 2:
            matched_words.sort(key=lambda x: x[1], reverse=True)
            scored.append((total_count, rule, matched_words))
    
    # 按得分降序，取前4
    scored.sort(key=lambda x: x[0], reverse=True)
    
    insights = []
    for i, (score, rule, matched_words) in enumerate(scored[:4]):
        top_words = "、".join([f'"{w}"' for w, _ in matched_words[:3]])
        desc = rule["desc_tpl"].format(count=score, top_words=top_words)
        insights.append({
            "id": i + 1,
            "title": rule["title"],
            "emoji": rule["emoji"],
            "color": rule["color"],
            "desc": desc,
        })
    
    # 如果主题不足4个，用数据特征补充
    idx = len(insights)
    fallback_pool = []
    
    if top_cats and idx < 4:
        top_cat_name = top_cats[0][0]
        top_cat_pct = f"{top_cats[0][1]/total_products*100:.1f}%"
        fallback_pool.append({
            "title": f"头部类目集中度高",
            "emoji": "📊", "color": "#3498db",
            "desc": f"{top_cat_name}类目占比{top_cat_pct}，是{sheet_name}行业绝对主战场，抓住头部类目即抓住大盘。"
        })
    
    if len(cat_counts) >= 4 and idx + len(fallback_pool) < 4:
        fallback_pool.append({
            "title": "细分类目百花齐放",
            "emoji": "🌐", "color": "#9b59b6",
            "desc": f"共{len(cat_counts)}个二级类目入榜，品类覆盖广泛，选品需精细化运营，挖掘细分蓝海机会。"
        })
    
    if top_plat_name and idx + len(fallback_pool) < 4:
        fallback_pool.append({
            "title": "平台引流集中化",
            "emoji": "📱", "color": "#1abc9c",
            "desc": f"{top_plat_name}是主要引流平台（占比{top_plat_pct}），内容创意应重点适配该平台流量特点。"
        })
    
    if idx + len(fallback_pool) < 4:
        fallback_pool.append({
            "title": "数据规模可观",
            "emoji": "📈", "color": "#3498db",
            "desc": f"共{total_products}款爆品入选，覆盖多个细分类目，行业选品空间广阔，值得持续深耕。"
        })
    
    for fb in fallback_pool:
        if idx >= 4:
            break
        insights.append({
            "id": idx + 1,
            "title": fb["title"],
            "emoji": fb["emoji"],
            "color": fb["color"],
            "desc": fb["desc"],
        })
        idx += 1
    
    return insights

def build_multi_html(industries_data, version_key):
    """生成支持行业切换的完整HTML（version_key 用于隔离不同版本的DOM id）"""
    
    # 行业Tab
    industry_tabs = []
    industry_contents = []
    
    for idx, ind in enumerate(industries_data):
        is_first = (idx == 0)
        is_empty = (ind["total"] == 0)
        
        if is_first:
            tab_cls = 'active-industry-tab'
        else:
            tab_cls = ''
        
        industry_tabs.append(f'''    <div id="industry-tab-{version_key}-{idx}" onclick="switchIndustry('{version_key}', {idx})" style="--tab-color: {INDUSTRY_CONFIG[ind["name"]]["color_text"]}; --tab-color-deep: {INDUSTRY_CONFIG[ind["name"]]["color_deep"]};" class="industry-tab px-6 py-3 text-sm md:text-base font-black flex items-center gap-2 cursor-pointer select-none transition-all duration-300 {tab_cls}">
      <span>{ind["emoji"]}</span> {ind["name"]}
    </div>''')
        
        if is_empty:
            # 空状态
            content = build_empty_industry_content(ind, idx, is_first, version_key)
        else:
            content = build_industry_content(ind, idx, is_first, version_key)
        
        # 为每个行业注入CSS变量，使按钮颜色跟随行业主题
        c = INDUSTRY_CONFIG[ind["name"]]
        hex_color = c["color_text"].lstrip("#")
        r, g, b = int(hex_color[:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        css_vars = (
            f'--ind-color: {c["color_text"]}; '
            f'--ind-light: {c["color_light"]}; '
            f'--ind-border: {c["color_border"]}; '
            f'--ind-shadow: {c["shadow_color"]}; '
            f'--ind-shadow-hover: rgba({r},{g},{b},0.25); '
            f'--ind-link-bg: {c["color_bg"]}; '
            f'--ind-link-text: {c["color_dark"]}; '
            f'--ind-link-border: {c["color_border"]}; '
            f'--ind-quote-bg: rgba({r},{g},{b},0.06); '
            f'--ind-color-soft: rgba({r},{g},{b},0.4);'
        )
        content = f'<div class="industry-wrapper" style="{css_vars}">\n{content}\n</div>'
        industry_contents.append(content)
    
    all_tabs = "\n".join(industry_tabs)
    all_contents = "\n\n".join(industry_contents)
    
    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="0">
  <title>多行业爆品选品报告</title>
  <style>
    /* 离线可用：无需外部字体和 Tailwind CDN */
    * {{ box-sizing: border-box; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      margin: 0;
      padding: 20px 0;
      background-color: #f3f7f4;
      background-image: 
        radial-gradient(at 0% 0%, hsla(210,30%,96%,1) 0px, transparent 50%),
        radial-gradient(at 100% 100%, hsla(210,20%,92%,1) 0px, transparent 50%);
    }}
    .min-h-screen {{ min-height: 100vh; }}
    .text-slate-800 {{ color: #1e293b; }}
    .antialiased {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }}
    .pb-16 {{ padding-bottom: 4rem; }}
    .flex {{ display: flex; }}
    .items-end {{ align-items: flex-end; }}
    .items-center {{ align-items: center; }}
    .overflow-x-auto {{ overflow-x: auto; }}
    .gap-2 {{ gap: 0.5rem; }}
    .gap-1\.5 {{ gap: 0.375rem; }}
    .pl-2 {{ padding-left: 0.5rem; }}
    .px-5 {{ padding-left: 1.25rem; padding-right: 1.25rem; }}
    .px-6 {{ padding-left: 1.5rem; padding-right: 1.5rem; }}
    .py-3 {{ padding-top: 0.75rem; padding-bottom: 0.75rem; }}
    .text-xs {{ font-size: 0.75rem; }}
    .text-sm {{ font-size: 0.875rem; }}
    .md\\:text-base {{ font-size: 1rem; }}
    .md\\:text-sm {{ font-size: 0.875rem; }}
    .font-black {{ font-weight: 900; }}
    .cursor-pointer {{ cursor: pointer; }}
    .select-none {{ user-select: none; }}
    .transition-all {{ transition: all 0.3s; }}
    .duration-300 {{ transition-duration: 0.3s; }}
    .bg-white {{ background-color: #ffffff; }}
    .rounded-2xl {{ border-radius: 1rem; }}
    .rounded-tl-none {{ border-top-left-radius: 0; }}
    .p-5 {{ padding: 1.25rem; }}
    .md\\:p-6 {{ padding: 1.5rem; }}
    .shadow-xl {{ box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04); }}
    .relative {{ position: relative; }}
    .min-h-\\[400px\\] {{ min-height: 400px; }}
    .no-scrollbar {{
      -ms-overflow-style: none;
      scrollbar-width: none;
    }}
    .no-scrollbar::-webkit-scrollbar {{
      display: none;
    }}

    /* 行业级 Tab */
    .industry-tab {{
      position: relative;
      background: #ffffff;
      border-radius: 16px 16px 0 0;
      border: 2px solid rgba(0,0,0,0.06);
      border-bottom: none;
      white-space: nowrap;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      color: #94a3b8;
      font-size: 15px;
    }}

    .active-industry-tab {{
      background: var(--tab-color) !important;
      color: #ffffff !important;
      border-color: var(--tab-color) !important;
      box-shadow: none;
      z-index: 10;
      font-size: 16px;
    }}

    /* 类目分布卡片（可点击切换 = 兼具 Tab 功能） */
    .dist-card {{
      transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.2s ease, border-color 0.2s ease;
    }}
    .dist-card:hover {{
      transform: translateY(-2px);
    }}
    .dist-card.active-dist-card {{
      transform: translateY(-2px);
    }}

    .active-tab {{
      background: #1b3d22 !important;
      color: #ffffff !important;
      border-color: #1b3d22 !important;
      box-shadow: 0 -4px 10px rgba(27,61,34,0.08);
      z-index: 10;
    }}

    .no-scrollbar {{
      -ms-overflow-style: none;
      scrollbar-width: none;
    }}
    .no-scrollbar::-webkit-scrollbar {{
      display: none;
    }}

    .industry-content {{
      opacity: 0 !important;
      transform: translateY(10px) !important;
      display: none !important;
      transition: all 0.4s cubic-bezier(0.34, 1.56, 0.64, 1) !important;
    }}
    .industry-content.active {{
      display: block !important;
      opacity: 1 !important;
      transform: translateY(0) !important;
    }}

    .folder-content {{
      opacity: 0 !important;
      transform: translateY(10px) scale(0.99) !important;
      display: none !important;
      transition: all 0.35s cubic-bezier(0.34, 1.56, 0.64, 1) !important;
    }}
    .folder-content.folder-show,
    .folder-content.active {{
      display: block !important;
      opacity: 1 !important;
      transform: translateY(0) scale(1) !important;
    }}

    .text-forest {{ color: #1b3d22; }}
    .bg-forest {{ background-color: #1b3d22; }}
    .border-forest {{ border-color: #1b3d22; }}
    .selection-bg {{ }}
    ::selection {{ background-color: #dbeafe; color: #1e3a5f; }}

    /* ===== 移动端响应式 ===== */
    @media (max-width: 768px) {{
      body {{
        padding: 8px 0;
      }}

      /* 商品卡片：移动端占满宽度 */
      .product-card-mobile {{
        width: 100% !important;
        min-width: 0 !important;
      }}

      /* 顶部总览海报缩小内边距和字号 */
      .top-poster-mobile {{
        padding: 16px 14px 16px 14px !important;
      }}
      .top-poster-mobile h1 {{
        font-size: 14px !important;
        line-height: 1.25 !important;
        text-align: center !important;
      }}
      .top-poster-mobile h1 .poster-title-main {{
        font-size: 17px !important;
      }}
      .top-poster-mobile h1 .poster-title-sub {{
        font-size: 22px !important;
      }}
      .top-poster-mobile p {{
        font-size: 11px !important;
      }}

      /* 版本下拉框：移动端脱离 absolute 定位，移到标题上方居中显示 */
      .top-poster-mobile .version-selector {{
        position: static !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        margin: 0 auto 10px auto !important;
        padding: 0 !important;
        width: 100% !important;
      }}
      .top-poster-mobile .version-selector select {{
        min-width: 0 !important;
        width: auto !important;
        font-size: 11px !important;
        padding: 4px 26px 4px 12px !important;
        border-radius: 6px !important;
      }}

      /* 分布看板：移动端缩小 */
      .dist-board-mobile > div {{
        min-width: 72px !important;
        padding: 6px 6px !important;
      }}

      /* 行业Tab文字缩小 */
      .industry-tab {{
        font-size: 12px !important;
        padding: 8px 12px !important;
      }}
      .active-industry-tab {{
        font-size: 12px !important;
      }}

      /* 品类内容区域减小内边距 */
      .category-content-box {{
        padding: 12px !important;
      }}

      /* 卡片内部：图片和文字调整 */
      .product-card-mobile .card-image {{
        width: 92px !important;
        height: 92px !important;
      }}
      .product-card-mobile .card-noimage {{
        width: 92px !important;
        height: 92px !important;
      }}
      .product-card-mobile .product-title {{
        font-size: 12px !important;
        line-height: 1.35 !important;
        max-height: 36px !important;
        min-height: 36px !important;
      }}
      .product-card-mobile .badge {{
        font-size: 9px !important;
        padding: 2px 5px !important;
      }}

      /* 商家话术：移动端解除两行截断，完整展示 */
      .product-card-mobile .product-quote {{
        display: block !important;
        -webkit-line-clamp: 10 !important;
        -webkit-box-orient: unset !important;
        overflow: visible !important;
        max-height: none !important;
        font-size: 11px !important;
        line-height: 1.5 !important;
      }}

      /* 洞察卡片：单列 */
      .insights-grid {{
        grid-template-columns: 1fr !important;
      }}

      /* 行业内容容器减小padding */
      .industry-content-wrapper {{
        padding: 0 4px !important;
      }}
    }}

    @media (max-width: 480px) {{
      /* 更小屏幕进一步缩小 */
      .industry-tab {{
        font-size: 11px !important;
        padding: 6px 10px !important;
      }}
      .active-industry-tab {{
        font-size: 11px !important;
      }}
    }}
  </style>
</head>
<body class="min-h-screen text-slate-800 antialiased selection:bg-blue-100 selection:text-blue-900 pb-16">

  <!-- 顶部总览海报 -->
  <div style="max-width: 1200px; margin: 0 auto 20px auto; padding: 0 10px;" class="industry-content-wrapper">
    <div class="top-poster-mobile" style="background: linear-gradient(135deg, #1e293b 0%, #334155 100%); color: #ffffff; border-radius: 20px; padding: 35px 30px; text-align: center; box-shadow: 0 12px 40px rgba(30,41,59,0.15); position: relative; overflow: hidden;">
      <div style="position: absolute; top: -60px; right: -60px; width: 180px; height: 180px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
      <div style="position: absolute; bottom: -70px; left: -70px; width: 200px; height: 200px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
      
      <h1 style="margin: 0; font-size: 22px; font-weight: 700; letter-spacing: 1px; line-height: 1.4;">
        <span style="font-size: 22px;">商品消费业务部</span><br>
        <span style="font-size: 16px; font-weight: 500;">pdd爆品榜单</span>
      </h1>
      <div class="gold-line"></div>
      <p style="margin: 0; font-size: 13px; color: #94a3b8;">
        覆盖 {len([i for i in industries_data if i["total"] > 0])} 个行业 · 总计 <strong style="color: #fff; font-size: 15px;">{sum(i["total"] for i in industries_data)}</strong> 款爆品
      </p>
    </div>
  </div>

  <!-- 行业级 Tab 切换 -->
  <div style="max-width: 1200px; margin: 0 auto; padding: 0 10px;" class="industry-content-wrapper">
    <div class="flex items-end overflow-x-auto no-scrollbar gap-2 pl-2 select-none" id="industry-tabs-container">
{all_tabs}
    </div>
  </div>

  <!-- 行业内容区域 -->
  <div style="max-width: 1200px; margin: 0 auto;">
{all_contents}
  </div>
</body>
</html>'''
    return html


def build_product_card(item, cat_emoji, colors):
    """生成单个商品卡片，colors 为行业配色字典"""
    name = item.get("name", "")
    copy_text = item.get("copy", "")
    video_link = item.get("video_link", "#")
    link = item.get("link", "#")
    image_url = item.get("image_url", "")
    tags = item.get("tags", [])
    source = item.get("source", "")
    c_text = colors["color_text"]
    c_light = colors["color_light"]
    
    if image_url:
        # onerror 时把 img 隐藏，同时显示后面预置的 NO IMAGE 占位 div（高度自适应右侧总高度）
        # 图片支持点击放大：cursor: zoom-in + onclick 触发 showImageModal
        image_html = (
            f'<img src="{image_url}" alt="商品主图" class="card-image" style="width: 150px; height: 150px; flex-shrink: 0; border-radius: 8px; object-fit: cover; border: 1px solid #edebe5; cursor: zoom-in;" '
            f'onclick="showImageModal(this.src, event)" '
            f'onerror="this.onerror=null;this.style.display=\'none\';var n=this.nextElementSibling;if(n)n.style.display=\'flex\';">'
            f'<div class="card-noimage" style="width: 150px; height: 150px; flex-shrink: 0; border-radius: 8px; background: linear-gradient(135deg, var(--ind-light) 0%, var(--ind-border) 100%); display: none; flex-direction: column; align-items: center; justify-content: center; color: var(--ind-color);">'
            f'<span style="font-size: 32px;">{cat_emoji}</span>'
            f'<span style="font-size: 10px; font-weight: bold; margin-top: 6px; opacity: 0.7;">NO IMAGE</span>'
            f'</div>'
        )
    else:
        image_html = f'<div class="card-noimage" style="width: 150px; height: 150px; flex-shrink: 0; border-radius: 8px; background: linear-gradient(135deg, var(--ind-light) 0%, var(--ind-border) 100%); display: flex; flex-direction: column; align-items: center; justify-content: center; color: var(--ind-color);"><span style="font-size: 32px;">{cat_emoji}</span><span style="font-size: 10px; font-weight: bold; margin-top: 6px; opacity: 0.7;">NO IMAGE</span></div>'
    
    tags_html_parts = []
    for tag in tags:
        tags_html_parts.append(f'<span style="background-color: {c_light}; color: {c_text}; font-size: 10px; padding: 1px 6px; border-radius: 4px; margin-right: 4px; font-weight: 500;">{tag}</span>')
    tags_html = "".join(tags_html_parts)
    
    name_escaped = name.replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')
    copy_escaped = copy_text.replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')
    
    return f'''    <!-- 商品卡片 -->
    <div class="product-card-mobile product-card" style="width: calc(50% - 8px); min-width: 330px; background-color: #ffffff; border-radius: 12px; padding: 14px; box-sizing: border-box; box-shadow: 0 3px 10px rgba(0,0,0,0.015); border: 1px solid #eeebe3; display: flex; gap: 12px; align-items: stretch;">
      {image_html}
      <div style="display: flex; flex-direction: column; gap: 8px; flex: 1; min-width: 0; justify-content: space-between;">
        <div>
          <div style="display: flex; flex-wrap: wrap; gap: 2px; margin-bottom: 4px;">
            {tags_html}
          </div>
          <div class="product-title" style="font-size: 13px; line-height: 1.4; font-weight: 700; color: #2d2a26; max-height: 40px; min-height: 40px; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden;" title="{name_escaped}">
            {name}
          </div>
        </div>
        <div class="product-quote" style="background-color: var(--ind-quote-bg); border-left: 2px solid var(--ind-color-soft); padding: 6px 8px; border-radius: 0 6px 6px 0; font-size: 11px; color: #6b6b6b; font-style: italic; line-height: 1.5; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden;" title="{copy_escaped}">
          "{copy_text}"
        </div>
        <div style="display: flex; gap: 6px;">
          <a href="{video_link}" target="_blank" class="btn-visit btn-video"><span class="btn-emoji">🎬</span> 播放视频</a>
          <a href="{link}" target="_blank" class="btn-visit btn-link"><span class="btn-emoji">🔗</span> 直达链路</a>
        </div>
      </div>
    </div>'''


def build_industry_content(ind, ind_idx, is_first, version_key):
    """构建一个完整行业的内容区域（version_key 用于隔离不同版本的DOM id）"""
    total = ind["total"]
    categories = ind["categories"]
    keywords = ind["keywords"]
    insights = ind["insights"]
    platforms = ind.get("platforms", {})
    c = ind  # colors 就在 ind 里（INDUSTRY_CONFIG 的值）
    
    active_cls = 'active' if is_first else ''

    # 分布看板（每张卡片本身可点击切换品类 = 兼具 Tab 功能）
    dist_cards = []
    for cat_i, cat in enumerate(categories):
        is_first_cat = (cat_i == 0)
        # 激活态：行业主色描边 + 浅色背景；非激活：默认米色
        if is_first_cat:
            card_style = f'background-color: {c["color_bg"]}; border: 1.5px solid {c["color_deep"]};'
            num_color = c["color_deep"]
        else:
            card_style = 'background-color: #faf9f5; border: 1px solid #eeebe3;'
            num_color = c["color_text"]
        dist_cards.append(f'''      <div id="tab-header-{version_key}-{ind_idx}-{cat_i}" onclick="switchCategory('{version_key}',{ind_idx},{cat_i})" data-color-deep="{c["color_deep"]}" data-color-text="{c["color_text"]}" data-color-bg="{c["color_bg"]}" data-shadow="{c["shadow_color"]}" class="dist-card" style="flex: 1; min-width: 100px; {card_style} border-radius: 8px; padding: 8px 10px; text-align: center; cursor: pointer; transition: all 0.25s ease; box-shadow: {'0 2px 8px ' + c["shadow_color"] if is_first_cat else 'none'};" onmouseover="if(!this.classList.contains('active-dist-card')){{this.style.backgroundColor='#fff';this.style.borderColor='{c["color_text"]}';this.style.boxShadow='0 2px 8px {c["shadow_color"]}';}}" onmouseout="if(!this.classList.contains('active-dist-card')){{this.style.backgroundColor='#faf9f5';this.style.borderColor='#eeebe3';this.style.boxShadow='none';}}">
        <div style="display: flex; align-items: center; justify-content: center; gap: 4px; margin-bottom: 3px; white-space: nowrap; overflow: hidden;">
          <span style="font-size: 14px;">{cat["emoji"]}</span>
          <span style="font-size: 11px; color: #5c5a56; font-weight: 500; overflow: hidden; text-overflow: ellipsis;">{cat["name"]}</span>
        </div>
        <div style="font-size: 12px; color: {num_color}; font-weight: bold; white-space: nowrap;"><span class="dist-card-num" style="font-weight: bold;">{cat["count"]}</span><span style="font-size: 9px; font-weight: normal; color: #9c9995; margin: 0 3px;">款</span><span style="font-size: 9px; font-weight: normal; color: #9c9995;">{cat["percentage"]}</span></div>
      </div>''')
    dist_html = "\n".join(dist_cards)

    # 商品展示（每个品类直接平铺显示）
    tab_contents = []
    for cat_i, cat in enumerate(categories):
        # 商品卡片
        cards = []
        for item in cat.get("products", []):
            cards.append(build_product_card(item, cat["emoji"], c))
        cards_html = "\n".join(cards)

        # 品类之间的 margin-top 统一
        top_margin = '1px'
        show_cls = 'folder-show' if cat_i == 0 else ''
        tab_contents.append(f'''    <div id="tab-content-{version_key}-{ind_idx}-{cat_i}" class="folder-content {show_cls}">
  <div style="margin-top: {top_margin}; margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #eae7e0; padding-bottom: 8px;">
    <div style="font-size: 18px; font-weight: bold; color: {c["color_dark"]}; display: flex; align-items: center; gap: 8px;">
      <span style="font-size: 22px;">{cat["emoji"]}</span> {cat["name"]}
      <span style="background-color: #eae7e0; color: #5c5a56; font-size: 11px; padding: 1px 8px; border-radius: 10px; font-weight: normal; margin-left: 6px;">{cat["count"]} 款爆品</span>
    </div>
    <div style="font-size: 12px; color: #8c8985;">分行业精选</div>
  </div>
  <div style="display: flex; flex-wrap: wrap; gap: 15px;">
{cards_html}
  </div>
</div>''')

    contents_html = "\n\n".join(tab_contents)
    
    # 高频词
    kw_spans = []
    for kw in keywords:
        kw_spans.append(f'<span class="keyword-tag" style="background-color: {c["color_bg"]}; color: {c["color_text"]}; font-size: 12px; font-weight: bold; padding: 4px 10px; border-radius: 12px; border: 1px solid {c["color_border"]}; margin-bottom: 6px; white-space: nowrap;">{kw["word"]} ({kw["count"]}次)</span>')
    kw_html = " ".join(kw_spans)
    
    # 洞察卡片
    insight_cards = []
    for ins in insights:
        insight_cards.append(f'''        <div class="insight-card" style="display: flex; align-items: flex-start; gap: 12px;">
          <span class="insight-icon" style="background-color: {ins["color"]}15; color: {ins["color"]};">{ins["emoji"]}</span>
          <div style="flex: 1; min-width: 0;">
            <div style="font-weight: 700; color: {ins["color"]}; font-size: 13px; margin-bottom: 4px;">{ins["title"]}</div>
            <div style="font-size: 12px; color: #5c5a56; line-height: 1.6;">{ins["desc"]}</div>
          </div>
        </div>''')
    insights_html = "\n".join(insight_cards)
    
    # 平台分布
    platform_parts = []
    for plat, count in sorted(platforms.items(), key=lambda x: x[1], reverse=True):
        platform_parts.append(f'<span style="background-color: #eef2ff; color: #4f46e5; font-size: 11px; padding: 2px 8px; border-radius: 8px; font-weight: 500;">{plat}: {count}款</span>')
    platform_html = " ".join(platform_parts) if platform_parts else ""
    
    return f'''  <div id="industry-content-{version_key}-{ind_idx}" class="industry-content {active_cls} industry-content-wrapper" style="max-width: 1200px; margin: 0 auto; padding: 0 10px; box-sizing: border-box;">

    <!-- 爆品类目分布看板（每张卡片可点击切换品类，激活态为 Tab） -->
    <div style="background-color: #ffffff; border-radius: 14px; padding: 22px; margin-top: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.02); border: 1px solid #e9e7e0;">
      <h3 style="margin: 0 0 15px 0; font-size: 15px; color: {c["color_text"]}; font-weight: 700; display: flex; align-items: center; gap: 6px;">
        📊 爆品类目分布看板 <span style="font-size: 11px; color: #9c9995; font-weight: normal; margin-left: 6px;">（点击卡片切换品类 ↓）</span>
      </h3>
      <div class="dist-board-mobile" style="display: flex; flex-wrap: wrap; gap: 10px;">
{dist_html}
      </div>
    </div>

    <!-- 品类商品展示区 -->
    <div style="margin-top: 25px;">
      <div class="category-content-box bg-white rounded-2xl p-5 md:p-6 shadow-xl relative min-h-[400px]" style="border:2px solid {c["color_border"]};">
{contents_html}
      </div>

    </div>

    <!-- 爆品创意黄金法则与高频词 -->
    <div style="max-width: 1180px; margin: 30px auto 0 auto; padding: 0 10px; box-sizing: border-box;">
      <div style="background-color: #ffffff; border-radius: 14px; padding: 22px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.02); border: 1px solid #e9e7e0;">
        <h3 style="margin: 0 0 15px 0; font-size: 15px; color: {c["color_text"]}; font-weight: 700; display: flex; align-items: center; gap: 6px;">
          💡 爆品创意黄金法则与高频词
        </h3>
        
        <div style="margin-bottom: 18px;">
          <div style="font-size: 12px; color: #73706c; font-weight: 600; margin-bottom: 8px;">🏷️ 爆品标题与卖点高频核心词云：</div>
          <div style="display: flex; flex-wrap: wrap; gap: 8px;">
            {kw_html}
          </div>
        </div>
        
        <div style="border-top: 1px dashed #eeebe3; padding-top: 15px;">
          <div style="font-size: 12px; color: #73706c; font-weight: 600; margin-bottom: 10px;">🎯 爆款视频带货密码提炼：</div>
          <div class="insights-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px;">
{insights_html}
          </div>
        </div>
      </div>
    </div>

  </div>'''


def build_empty_industry_content(ind, ind_idx, is_first, version_key):
    """构建空行业（美护）占位内容"""
    active_cls = 'active' if is_first else ''
    c = ind  # colors
    return f'''  <div id="industry-content-{version_key}-{ind_idx}" class="industry-content {active_cls} industry-content-wrapper" style="max-width: 1200px; margin: 0 auto; padding: 0 10px; box-sizing: border-box;">

    <!-- 空状态占位 -->
    <div style="background-color: #ffffff; border-radius: 14px; padding: 60px 30px; text-align: center; margin-top: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.02); border: 2px dashed #e9e7e0;">
      <div style="font-size: 64px; margin-bottom: 20px;">{c["emoji"]}</div>
      <h3 style="margin: 0 0 10px 0; font-size: 18px; color: {c["color_text"]}; font-weight: 600;">
        {ind["name"]}行业数据即将上线
      </h3>
      <p style="margin: 0; font-size: 13px; color: #9c9995; line-height: 1.8;">
        该行业爆品数据正在采集中，完成后将自动填充<br>
        包含护肤品、彩妆、个护等细分类目
      </p>
    </div>

  </div>'''


def build_versioned_html(versions):
    """将所有数据源版本合并为带版本筛选器的最终HTML
    versions: [{"key", "label", "industries_data", "html"}, ...]，第一个为默认显示"""
    
    # 构建版本选择器选项 + 各版本内容块（剔除各自顶部海报框）
    version_options = []
    version_blocks = []
    for i, v in enumerate(versions):
        key = v["key"]
        label = v["label"]
        html_content = v["html"]
        
        industries_data = v["industries_data"]
        version_total = sum(ind["total"] for ind in industries_data)
        version_industry_count = len([ind for ind in industries_data if ind["total"] > 0])
        
        version_options.append(f'          <option value="{key}">{label}</option>')
        
        body_match = re.search(r'<body[^>]*>(.*)</body>', html_content, re.DOTALL)
        if body_match:
            body = body_match.group(1)
            # 剔除 build_multi_html 自带的顶部总览海报框（避免双框）
            body = re.sub(r'<!-- 顶部总览海报 -->.*?</div>\s*</div>', '', body, count=1, flags=re.DOTALL)
        else:
            body = html_content
        
        display = 'block' if i == 0 else 'none'
        version_blocks.append(f'    <!-- VERSION_START: {label} -->\n    <div id="version-{key}" data-label="{label}" data-total="{version_total}" data-industry-count="{version_industry_count}" class="version-block" style="display: {display}; width: 100%;">\n{body}\n    </div>\n    <!-- VERSION_END -->')
    
    version_select_html = '\n'.join(version_options)
    all_versions_html = '\n'.join(version_blocks)
    
    # 默认版本（第一个）的行业数据，用于顶部海报总计
    default_industries = versions[0]["industries_data"]
    
    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="0">
  <title>多行业爆品选品报告</title>
  <script src="https://cdn.busuanzi.cc/busuanzi/3.6.9/busuanzi.min.js" defer></script>
  <style>
    * {{ box-sizing: border-box; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      margin: 0;
      padding: 20px 0;
      background-color: #f3f7f4;
      background-image: 
        radial-gradient(at 0% 0%, hsla(210,30%,96%,1) 0px, transparent 50%),
        radial-gradient(at 100% 100%, hsla(210,20%,92%,1) 0px, transparent 50%);
    }}
    .min-h-screen {{ min-height: 100vh; }}
    .text-slate-800 {{ color: #1e293b; }}
    .antialiased {{ -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }}
    .pb-16 {{ padding-bottom: 4rem; }}
    .flex {{ display: flex; }}
    .items-end {{ align-items: flex-end; }}
    .items-center {{ align-items: center; }}
    .overflow-x-auto {{ overflow-x: auto; }}
    .gap-2 {{ gap: 0.5rem; }}
    .gap-1\\.5 {{ gap: 0.375rem; }}
    .pl-2 {{ padding-left: 0.5rem; }}
    .px-5 {{ padding-left: 1.25rem; padding-right: 1.25rem; }}
    .px-6 {{ padding-left: 1.5rem; padding-right: 1.5rem; }}
    .py-3 {{ padding-top: 0.75rem; padding-bottom: 0.75rem; }}
    .text-xs {{ font-size: 0.75rem; }}
    .text-sm {{ font-size: 0.875rem; }}
    .md\\:text-base {{ font-size: 1rem; }}
    .md\\:text-sm {{ font-size: 0.875rem; }}
    .font-black {{ font-weight: 900; }}
    .cursor-pointer {{ cursor: pointer; }}
    .select-none {{ user-select: none; }}
    .transition-all {{ transition: all 0.3s; }}
    .duration-300 {{ transition-duration: 0.3s; }}
    .bg-white {{ background-color: #ffffff; }}
    .rounded-2xl {{ border-radius: 1rem; }}
    .rounded-tl-none {{ border-top-left-radius: 0; }}
    .p-5 {{ padding: 1.25rem; }}
    .md\\:p-6 {{ padding: 1.5rem; }}
    .shadow-xl {{ box-shadow: 0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04); }}
    .relative {{ position: relative; }}
    .min-h-\\[400px\\] {{ min-height: 400px; }}
    .no-scrollbar {{
      -ms-overflow-style: none;
      scrollbar-width: none;
    }}
    .no-scrollbar::-webkit-scrollbar {{
      display: none;
    }}

    /* 版本筛选器 */
    .version-selector {{
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .version-selector select {{
      appearance: none;
      -webkit-appearance: none;
      background: #ffffff;
      border: 1.5px solid #e2e8f0;
      border-radius: 10px;
      padding: 8px 36px 8px 14px;
      font-size: 13px;
      font-weight: 600;
      color: #334155;
      cursor: pointer;
      background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");
      background-repeat: no-repeat;
      background-position: right 12px center;
      transition: border-color 0.2s, box-shadow 0.2s;
      outline: none;
      min-width: 160px;
    }}
    .version-selector select:hover {{
      border-color: #94a3b8;
    }}
    .version-selector select:focus {{
      border-color: #1e293b;
      box-shadow: 0 0 0 3px rgba(30,41,59,0.08);
    }}

    /* 行业级 Tab */
    .industry-tab {{
      position: relative;
      background: #ffffff;
      border-radius: 16px 16px 0 0;
      border: 2px solid rgba(0,0,0,0.06);
      border-bottom: none;
      white-space: nowrap;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      color: #94a3b8;
      font-size: 15px;
    }}

    .active-industry-tab {{
      background: var(--tab-color) !important;
      color: #ffffff !important;
      border-color: var(--tab-color) !important;
      box-shadow: none;
      z-index: 10;
      font-size: 16px;
    }}

    /* 类目分布卡片 */
    .dist-card {{
      transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.2s ease, border-color 0.2s ease;
    }}
    .dist-card:hover {{
      transform: translateY(-2px);
    }}
    .dist-card.active-dist-card {{
      transform: translateY(-2px);
    }}

    .active-tab {{
      background: #1b3d22 !important;
      color: #ffffff !important;
      border-color: #1b3d22 !important;
      box-shadow: 0 -4px 10px rgba(27,61,34,0.08);
      z-index: 10;
    }}

    .industry-content {{
      opacity: 0 !important;
      transform: translateY(10px) !important;
      display: none !important;
      transition: all 0.4s cubic-bezier(0.34, 1.56, 0.64, 1) !important;
    }}
    .industry-content.active {{
      display: block !important;
      opacity: 1 !important;
      transform: translateY(0) !important;
    }}

    .folder-content {{
      opacity: 0 !important;
      transform: translateY(10px) scale(0.99) !important;
      display: none !important;
      transition: all 0.35s cubic-bezier(0.34, 1.56, 0.64, 1) !important;
    }}
    .folder-content.folder-show,
    .folder-content.active {{
      display: block !important;
      opacity: 1 !important;
      transform: translateY(0) scale(1) !important;
    }}

    .text-forest {{ color: #1b3d22; }}
    .bg-forest {{ background-color: #1b3d22; }}
    .border-forest {{ border-color: #1b3d22; }}
    ::selection {{ background-color: #dbeafe; color: #1e3a5f; }}

    /* 商品卡片悬浮效果 */
    .product-card {{
      transition: transform 0.25s cubic-bezier(0.34, 1.56, 0.64, 1), box-shadow 0.25s ease;
      border: 1px solid transparent;
    }}
    .product-card:hover {{
      transform: translateY(-4px);
      box-shadow: 0 12px 28px rgba(0,0,0,0.1), 0 4px 10px rgba(0,0,0,0.06);
      border-color: rgba(0,0,0,0.06);
    }}

    /* 洞察卡片样式 */
    .insight-card {{
      background: linear-gradient(135deg, #ffffff 0%, #fafafa 100%);
      border: 1px solid #f0ebe0;
      border-radius: 14px;
      padding: 18px 20px;
      transition: transform 0.2s ease, box-shadow 0.2s ease;
    }}
    .insight-card:hover {{
      transform: translateY(-2px);
      box-shadow: 0 8px 24px rgba(0,0,0,0.06);
    }}
    .insight-icon {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 36px;
      height: 36px;
      border-radius: 10px;
      font-size: 18px;
      margin-right: 10px;
      flex-shrink: 0;
    }}

    /* 米色系分隔线微动效 */
    .gold-line {{
      width: 60px;
      height: 3px;
      background: linear-gradient(90deg, #c9b88f, #d4c39d, #c9b88f);
      background-size: 200% 100%;
      border-radius: 2px;
      margin: 18px auto 12px auto;
      animation: goldShimmer 3s ease-in-out infinite;
    }}
    @keyframes goldShimmer {{
      0%, 100% {{ background-position: 0% 50%; }}
      50% {{ background-position: 100% 50%; }}
    }}

    /* 按钮基础样式 - 浅色系，颜色跟随行业CSS变量 */
    .btn-visit {{
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 6px 12px;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.25s ease;
      text-decoration: none;
      flex: 1;
      text-align: center;
      justify-content: center;
    }}
    /* 播放视频按钮 - 跟随行业浅色 */
    .btn-video {{
      background-color: var(--ind-light);
      color: var(--ind-color);
      border: 1px solid var(--ind-border);
      box-shadow: 0 1px 2px var(--ind-shadow);
    }}
    .btn-video:hover {{
      background-color: var(--ind-color);
      color: #ffffff;
      border-color: var(--ind-color);
      transform: translateY(-1px);
      box-shadow: 0 4px 14px var(--ind-shadow-hover);
    }}
    .btn-video:hover .btn-emoji {{ filter: brightness(0) invert(1); }}
    .btn-emoji {{ transition: filter 0.2s ease; }}
    /* 直达链路按钮 - 行业为主题，稍重一点 */
    .btn-link {{
      background-color: var(--ind-link-bg);
      color: var(--ind-link-text);
      border: 1px solid var(--ind-link-border);
      box-shadow: 0 1px 2px var(--ind-shadow);
    }}
    .btn-link:hover {{
      background-color: var(--ind-link-text);
      color: #ffffff;
      border-color: var(--ind-link-text);
      transform: translateY(-1px);
      box-shadow: 0 4px 14px var(--ind-shadow-hover);
    }}
    .btn-link:hover .btn-emoji {{ filter: brightness(0) invert(1); }}

    /* 商品标题行高优化 */
    .product-title {{
      line-height: 1.5;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
    }}

    /* 品类分布卡片激活态增强 */
    .dist-card {{
      transition: transform 0.25s cubic-bezier(0.34, 1.56, 0.64, 1), box-shadow 0.25s ease, background-color 0.2s ease, border-color 0.2s ease;
    }}
    .dist-card:hover {{
      transform: translateY(-3px);
    }}
    .dist-card.active-dist-card {{
      transform: translateY(-3px);
    }}

    /* 关键字标签悬浮 */
    .keyword-tag {{
      transition: transform 0.15s ease, box-shadow 0.15s ease;
    }}
    .keyword-tag:hover {{
      transform: translateY(-1px);
      box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }}

    @media (max-width: 768px) {{
      body {{
        padding: 8px 0;
      }}
      .product-card-mobile {{
        width: 100% !important;
        min-width: 0 !important;
      }}
      .top-poster-mobile {{
        padding: 16px 14px !important;
      }}
      .top-poster-mobile h1 {{
        font-size: 14px !important;
        line-height: 1.25 !important;
        text-align: center !important;
      }}
      .top-poster-mobile h1 .poster-title-main {{
        font-size: 17px !important;
      }}
      .top-poster-mobile h1 .poster-title-sub {{
        font-size: 22px !important;
      }}
      .product-card-mobile .card-image {{
        width: 92px !important;
        height: 92px !important;
      }}
      .product-card-mobile .card-noimage {{
        width: 92px !important;
        height: 92px !important;
      }}
      .product-card-mobile .product-title {{
        font-size: 12px !important;
        line-height: 1.35 !important;
        max-height: 36px !important;
        min-height: 36px !important;
      }}
      .product-card-mobile .badge {{
        font-size: 9px !important;
        padding: 2px 5px !important;
      }}

      /* 商家话术：移动端解除两行截断，完整展示 */
      .product-card-mobile .product-quote {{
        display: block !important;
        -webkit-line-clamp: 10 !important;
        -webkit-box-orient: unset !important;
        overflow: visible !important;
        max-height: none !important;
        font-size: 11px !important;
        line-height: 1.5 !important;
      }}
      .dist-board-mobile {{
        gap: 6px !important;
      }}
      .industry-tab {{
        font-size: 11px !important;
        padding: 6px 10px !important;
      }}
      .active-industry-tab {{
        font-size: 11px !important;
      }}
      /* 版本下拉框：移动端脱离 absolute，移到标题上方居中显示 */
      .top-poster-mobile .version-selector {{
        position: static !important;
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        margin: 0 auto 10px auto !important;
        padding: 0 !important;
        width: 100% !important;
      }}
      .top-poster-mobile .version-selector select {{
        min-width: 0 !important;
        width: auto !important;
        font-size: 11px !important;
        padding: 4px 26px 4px 12px !important;
        border-radius: 6px !important;
      }}
    }}
  </style>
</head>
<body class="min-h-screen text-slate-800 antialiased pb-16">

  <!-- 顶部总览海报（含版本筛选器） -->
  <div style="max-width: 1200px; margin: 0 auto 20px auto; padding: 0 10px;">
    <div class="top-poster-mobile" style="background: #e8e2d0; color: #5a5347; border-radius: 20px; padding: 35px 30px; text-align: center; box-shadow: 0 6px 20px rgba(120,110,80,0.10); position: relative; overflow: hidden;">
      <div style="position: absolute; top: -60px; right: -60px; width: 180px; height: 180px; background: radial-gradient(circle, rgba(255,255,255,0.2) 0%, transparent 70%); border-radius: 50%;"></div>
      <div style="position: absolute; bottom: -70px; left: -70px; width: 200px; height: 200px; background: radial-gradient(circle, rgba(255,255,255,0.15) 0%, transparent 70%); border-radius: 50%;"></div>

      <!-- 版本筛选器（右上角） -->
      <div class="version-selector" style="position: absolute; top: 16px; right: 20px; z-index: 1;">
        <select id="version-select" onchange="switchVersion(this.value)" style="appearance: none; -webkit-appearance: none; background: #ffffff; border: 1px solid #c9b88f; border-radius: 8px; padding: 6px 32px 6px 12px; font-size: 12px; font-weight: 600; color: #5a5347; cursor: pointer; background-image: url(&quot;data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='none' stroke='%23c9b88f' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E&quot;); background-repeat: no-repeat; background-position: right 10px center; outline: none; min-width: 140px; transition: all 0.2s; box-shadow: 0 1px 3px rgba(120,110,80,0.12);">
{version_select_html}
        </select>
      </div>

      <h1 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: 1px; color: #3d3527; line-height: 1.4;">
        <span class="poster-title-sub" style="font-size: 34px; font-weight: 700; letter-spacing: 2px;">pdd爆品榜单</span>
      </h1>
      <div class="gold-line"></div>
      <p id="poster-summary" style="margin: 0; font-size: 13px; color: #6b6147;">
        覆盖 <span id="poster-industry-count">{len([i for i in default_industries if i["total"] > 0])}</span> 个行业 · 总计 <strong id="poster-total" style="color: #3d3527; font-size: 16px;">{sum(i["total"] for i in default_industries)}</strong> 款爆品
      </p>
    </div>
  </div>

  <!-- 所有版本的内容块 -->
{all_versions_html}

  <!-- 图片放大查看模态层 -->
  <div id="image-modal" onclick="closeImageModal()" style="display: none; position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background-color: rgba(0,0,0,0.85); z-index: 99999; cursor: zoom-out; align-items: center; justify-content: center; padding: 30px; box-sizing: border-box;">
    <img id="image-modal-img" src="" alt="放大图" style="max-width: 90vw; max-height: 90vh; object-fit: contain; border-radius: 8px; box-shadow: 0 10px 40px rgba(0,0,0,0.5); background-color: #fff;" />
    <div style="position: absolute; top: 20px; right: 30px; color: #fff; font-size: 14px; background: rgba(0,0,0,0.5); padding: 6px 14px; border-radius: 20px; pointer-events: none;">✕ 点击任意处关闭</div>
  </div>

  <script>
    // 版本切换
    function switchVersion(key) {{
      const blocks = document.querySelectorAll('.version-block');
      blocks.forEach(b => b.style.display = 'none');
      const target = document.getElementById('version-' + key);
      if (!target) return;
      target.style.display = 'block';

      // 顶部更新总计数字（覆盖行业数 + 总计爆品数）
      const totalEl = document.getElementById('poster-total');
      const countEl = document.getElementById('poster-industry-count');
      const totalVal = target.getAttribute('data-total');
      const countVal = target.getAttribute('data-industry-count');
      if (totalEl && totalVal) totalEl.textContent = totalVal;
      if (countEl && countVal) countEl.textContent = countVal;

      // 重置所有版本块内的 industry tab/content 激活态
      document.querySelectorAll('.industry-content').forEach(c => {{
        c.classList.remove('active');
        c.style.display = 'none';
      }});
      document.querySelectorAll('.industry-tab').forEach(t => t.classList.remove('active-industry-tab'));

      // 激活新版本第一个 industry tab + content
      const firstTab = document.getElementById('industry-tab-' + key + '-0');
      const firstContent = document.getElementById('industry-content-' + key + '-0');
      if (firstTab) firstTab.classList.add('active-industry-tab');
      if (firstContent) {{
        firstContent.classList.add('active');
        firstContent.style.display = 'block';
        // 同步激活该行业内的第一个 dist-card 和 folder-content
        switchCategory(key, 0, 0);
      }}
    }}

    // 行业切换（vk 为版本 key，idx 为行业索引）
    function switchIndustry(vk, idx) {{
      // 只重置当前版本块内的 industry-content，避免影响其他版本
      const versionBlock = document.getElementById('version-' + vk);
      if (versionBlock) {{
        versionBlock.querySelectorAll('.industry-content').forEach(c => {{
          c.classList.remove('active');
          c.style.display = 'none';
        }});
        versionBlock.querySelectorAll('.industry-tab').forEach(t => t.classList.remove('active-industry-tab'));
      }}

      const target = document.getElementById('industry-content-' + vk + '-' + idx);
      if (target) {{
        target.classList.add('active');
        target.style.display = 'block';
        // 切换行业时重置该行业的 dist-card 选中态为第一个
        switchCategory(vk, idx, 0);
      }}

      const activeTab = document.getElementById('industry-tab-' + vk + '-' + idx);
      if (activeTab) {{
        activeTab.classList.add('active-industry-tab');
        activeTab.scrollIntoView({{ behavior: 'smooth', block: 'nearest', inline: 'center' }});
      }}
    }}

    // 品类切换（vk 为版本 key，indPrefix 为行业索引，idx 为品类索引）
    function switchCategory(vk, indPrefix, idx) {{
      const container = document.getElementById('industry-content-' + vk + '-' + indPrefix);
      if (!container) return;

      const contents = container.querySelectorAll('.folder-content');
      contents.forEach(c => {{
        c.classList.remove('active', 'folder-show');
        c.style.display = 'none';
      }});

      const target = document.getElementById('tab-content-' + vk + '-' + indPrefix + '-' + idx);
      if (target) {{
        target.classList.add('folder-show');
        target.style.display = 'block';
      }}

      const distCards = container.querySelectorAll('.dist-card');
      distCards.forEach(card => {{
        card.classList.remove('active-dist-card');
        card.style.backgroundColor = '#faf9f5';
        card.style.borderColor = '#eeebe3';
        card.style.boxShadow = 'none';
        const numEl = card.querySelector('.dist-card-num');
        if (numEl) {{
          numEl.style.color = card.getAttribute('data-color-text') || '#1b3d22';
        }}
      }});

      // 3) 激活当前点击的"分布卡片"
      const activeCard = document.getElementById('tab-header-' + vk + '-' + indPrefix + '-' + idx);
      if (activeCard) {{
        const colorDeep = activeCard.getAttribute('data-color-deep') || '#1b3d22';
        const colorBg = activeCard.getAttribute('data-color-bg') || '#faf9f5';
        const shadow = activeCard.getAttribute('data-shadow') || 'rgba(27,61,34,0.08)';
        activeCard.classList.add('active-dist-card');
        activeCard.style.backgroundColor = colorBg;
        activeCard.style.borderColor = colorDeep;
        activeCard.style.borderWidth = '1.5px';
        activeCard.style.boxShadow = '0 2px 8px ' + shadow;
        const numEl = activeCard.querySelector('.dist-card-num');
        if (numEl) {{
          numEl.style.color = colorDeep;
        }}
        activeCard.scrollIntoView({{ behavior: 'smooth', block: 'nearest', inline: 'center' }});
      }}
    }}

    // 图片放大
    function showImageModal(src, event) {{
      event.stopPropagation();
      const modal = document.getElementById('image-modal');
      const modalImg = document.getElementById('image-modal-img');
      modal.style.display = 'flex';
      modalImg.src = src;
      document.body.style.overflow = 'hidden';
    }}

    function closeImageModal() {{
      const modal = document.getElementById('image-modal');
      modal.style.display = 'none';
      document.body.style.overflow = '';
    }}

    document.addEventListener('keydown', function(e) {{
      if (e.key === 'Escape') {{
        closeImageModal();
      }}
    }});
  </script>
  <!-- 访问量统计（不蒜子） -->
  <footer style="text-align: center; padding: 20px 16px 32px; color: #a89a7c; font-size: 12px; letter-spacing: 0.5px; line-height: 1.8;">
    <div style="display: inline-flex; gap: 20px; align-items: center; flex-wrap: wrap; justify-content: center; background: #faf8f2; border: 1px solid #eeebe3; border-radius: 12px; padding: 14px 26px; box-shadow: 0 2px 8px rgba(120,110,80,0.06);">
      <span>今日访问 <b id="busuanzi_today_pv" style="color: #6b5f45; font-weight: 700;">—</b> 次</span>
      <span style="color: #e0d4bc;">|</span>
      <span>今日访客 <b id="busuanzi_today_uv" style="color: #6b5f45; font-weight: 700;">—</b> 人</span>
      <span style="color: #e0d4bc;">|</span>
      <span>累计访问 <b id="busuanzi_site_pv" style="color: #6b5f45; font-weight: 700;">—</b> 次</span>
      <span style="color: #e0d4bc;">|</span>
      <span>累计访客 <b id="busuanzi_site_uv" style="color: #6b5f45; font-weight: 700;">—</b> 人</span>
    </div>
    <div style="margin-top: 8px; font-size: 11px; color: #c4b89d;">访问数据由不蒜子统计提供</div>
  </footer>
</body>
</html>'''
    return html


def main():
    print("📊 读取Excel数据...")
    
    # 遍历所有数据源，生成对应版本
    versions = []
    for source in DATA_SOURCES:
        print(f"  处理版本: {source['label']} [{source['key']}]")
        industries_data = load_industries_data(source)
        for ind in industries_data:
            print(f"    {ind['name']}: {ind['total']}款, {len(ind['categories'])}个类目")
        html = build_multi_html(industries_data, source["key"])
        versions.append({
            "key": source["key"],
            "label": source["label"],
            "industries_data": industries_data,
            "html": html,
        })
    
    output_dir = "/Users/krystalcao/Desktop/已完成"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "多行业爆品选品报告.html")
    
    print("🔨 生成HTML...")
    final_html = build_versioned_html(versions)
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(final_html)
    
    file_size = os.path.getsize(output_path)
    print(f"✅ 报告已生成: {output_path}")
    print(f"   文件大小: {file_size/1024:.1f} KB")
    print(f"   版本数: {len(versions)}")
    for v in versions:
        total = sum(i['total'] for i in v['industries_data'])
        print(f"   - {v['label']}: 共{total}款爆品")


if __name__ == "__main__":
    main()
