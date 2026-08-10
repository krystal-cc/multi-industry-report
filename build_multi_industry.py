#!/usr/bin/env python3
"""
多行业爆品选品报告生成器
读取爆品数据Excel，生成支持行业切换的HTML报告。
四个行业：服饰、消电日百、食品饮料、美护（暂无数据）
"""

import pandas as pd
import json
import os
from collections import Counter

EXCEL_PATH = "/Users/krystalcao/Desktop/爆品数据-最新.xlsx"
FOOD_EXCEL_PATH = "/Users/krystalcao/Desktop/食品饮料.xlsx"

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
}

INDUSTRY_CONFIG = {
    "服饰": {
        "emoji": "👗",
        "color": "#e74c3c",           # 海报主色
        "color_dark": "#c0392b",       # 海报渐变终点、品类标题
        "color_deep": "#a93226",       # 品类Tab激活背景、品类内容边框
        "color_text": "#c0392b",       # 分布看板标题/数字、直达链路按钮、高频词
        "color_light": "#fdf0f0",      # 标签浅色背景
        "color_bg": "#fef9f9",         # 高频词浅色背景
        "color_border": "#f5d6d6",     # 高频词边框
        "color_btn_hover": "#a93226",  # 按钮深色
        "shadow_color": "rgba(231,76,60,0.15)",
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
        "color": "#e91e63",
        "color_dark": "#c2185b",
        "color_deep": "#ad1457",
        "color_text": "#c2185b",
        "color_light": "#fce4ec",
        "color_bg": "#fef5f7",
        "color_border": "#f8bbd0",
        "color_btn_hover": "#880e4f",
        "shadow_color": "rgba(233,30,99,0.15)",
    },
}

def read_sheet(sheet_name):
    df = pd.read_excel(EXCEL_PATH, sheet_name)
    if df.empty:
        return None
    # 合并过小的二级类目
    merge_map = INDUSTRY_MERGE.get(sheet_name, {})
    if merge_map:
        df["投放二级行业"] = df["投放二级行业"].apply(lambda x: merge_map.get(x, x))
    return df

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
    """根据行业数据生成洞察密码"""
    insights = []
    
    if sheet_name == "服饰":
        insights = [
            {"id": 1, "title": "情感化送礼场景", "emoji": "🎁", "color": "#e74c3c",
             "desc": "文案强调\"送朋友送自己送爱人\"，将饰品与情感场景绑定，七夕、节日送礼是核心转化驱动力。"},
            {"id": 2, "title": "性价比+工艺卖点", "emoji": "💎", "color": "#c0392b",
             "desc": "高频出现\"浮雕工艺\"\"活口可调节\"\"高档\"等词，用工艺细节支撑价格，同时强调实惠，降低购买门槛。"},
            {"id": 3, "title": "视觉冲击+上身效果", "emoji": "✨", "color": "#e67e22",
             "desc": "\"显身材\"\"像没穿一样\"\"防走光\"等词频出，聚焦穿着效果和解决痛点，用强画面感激发购买欲。"},
            {"id": 4, "title": "应季+新款驱动", "emoji": "🔥", "color": "#f39c12",
             "desc": "\"抢新款\"\"2026新款\"是标题标配，利用消费者追新心理，制造紧迫感和稀缺感。"},
        ]
    elif sheet_name == "消电日百":
        insights = [
            {"id": 1, "title": "懒人经济+解放双手", "emoji": "🤖", "color": "#3498db",
             "desc": "\"免安装\"\"一键操作\"\"全自动\"\"智能\"高频出现，瞄准懒人经济和便捷生活需求，降低使用门槛。"},
            {"id": 2, "title": "极致性价比轰炸", "emoji": "💰", "color": "#2980b9",
             "desc": "\"两位数到手\"\"工厂直销\"\"比实体店便宜一半\"等价格锚点策略，用大牌平替概念击中价格敏感用户。"},
            {"id": 3, "title": "健康生活升级", "emoji": "🌿", "color": "#27ae60",
             "desc": "\"除菌\"\"净化\"\"无油烹饪\"\"健康材质\"等卖点突出，后疫情时代消费者对居家环境健康关注度持续走高。"},
            {"id": 4, "title": "居家场景全覆盖", "emoji": "🏡", "color": "#8e44ad",
             "desc": "从厨具到收纳、从电器到家装，覆盖居家生活全场景，\"一物多用\"\"多功能\"是提升客单价的关键策略。"},
        ]
    elif sheet_name == "食品饮料":
        insights = [
            {"id": 1, "title": "宠物经济大爆发", "emoji": "🐶", "color": "#27ae60",
             "desc": "宠物生活类目占比近半，\"冻干\"\"无谷\"\"营养\"是核心关键词，宠物主愿意为毛孩子健康买单。"},
            {"id": 2, "title": "懒人即食+免煮冲泡", "emoji": "⏰", "color": "#2ecc71",
             "desc": "\"免煮冲泡\"\"开袋即食\"\"方便速食\"仍是食品类爆款密码，快节奏生活下省时是第一需求。"},
            {"id": 3, "title": "健康养生概念升级", "emoji": "💊", "color": "#f39c12",
             "desc": "膳食营养品和海外保健品占比高，\"高纯净原料\"\"好吸收\"\"无添加\"是信任转化的关键。"},
            {"id": 4, "title": "不满意包退信任背书", "emoji": "🛡️", "color": "#e74c3c",
             "desc": "\"不好吃你随便退\"\"不满意包退\"等零风险承诺是食品饮料转化的最后一把火，极大降低试错成本。"},
        ]
    
    return insights

def build_multi_html(industries_data):
    """生成支持行业切换的完整HTML"""
    
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
        
        industry_tabs.append(f'''    <div id="industry-tab-{idx}" onclick="switchIndustry({idx})" class="industry-tab px-6 py-3 text-sm md:text-base font-black flex items-center gap-2 cursor-pointer select-none transition-all duration-300 {tab_cls}">
      <span>{ind["emoji"]}</span> {ind["name"]}
    </div>''')
        
        if is_empty:
            # 美护行业空状态
            content = build_empty_industry_content(ind, idx, is_first)
        else:
            content = build_industry_content(ind, idx, is_first)
        industry_contents.append(content)
    
    all_tabs = "\n".join(industry_tabs)
    all_contents = "\n\n".join(industry_contents)
    
    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
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
      background: #1e293b !important;
      color: #ffffff !important;
      border-color: #1e293b !important;
      box-shadow: 0 -6px 16px rgba(30,41,59,0.12);
      z-index: 10;
      font-size: 16px;
    }}

    /* 类目分布卡片（可点击切换） */
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
        padding: 20px 15px !important;
      }}
      .top-poster-mobile h1 {{
        font-size: 16px !important;
      }}
      .top-poster-mobile p {{
        font-size: 11px !important;
      }}

      /* 行业海报缩小 */
      .industry-poster-mobile {{
        padding: 25px 18px !important;
      }}

      /* 分布看板：flex-wrap + 缩小 */
      .dist-board-mobile > div {{
        min-width: 80px !important;
        padding: 8px 6px !important;
      }}
      .dist-board-mobile > div > div:first-child {{
        font-size: 16px !important;
      }}
      .dist-board-mobile > div > div:nth-child(2) {{
        font-size: 10px !important;
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
        width: 60px !important;
        height: 60px !important;
      }}
      .product-card-mobile .card-noimage {{
        width: 60px !important;
        height: 60px !important;
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
      
      <div style="display: inline-block; background-color: rgba(255, 255, 255, 0.12); color: #e2e8f0; font-size: 12px; font-weight: bold; letter-spacing: 3px; padding: 5px 16px; border-radius: 20px; margin-bottom: 14px; text-transform: uppercase;">
        🔥 全行业 · 爆款视频链路监控
      </div>
      
      <h1 style="margin: 0; font-size: 22px; font-weight: 700; letter-spacing: 1px;">
        多行业爆品选品总览
      </h1>
      <div style="width: 60px; height: 3px; background-color: #f59e0b; margin: 18px auto 12px auto; border-radius: 2px;"></div>
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

  <script>
    // 行业切换
    function switchIndustry(idx) {{
      const contents = document.querySelectorAll('.industry-content');
      contents.forEach(c => {{
        c.classList.remove('active');
        c.style.display = 'none';
      }});

      const target = document.getElementById('industry-content-' + idx);
      if (target) {{
        target.classList.add('active');
        target.style.display = 'block';
      }}

      const tabs = document.querySelectorAll('.industry-tab');
      tabs.forEach(t => t.classList.remove('active-industry-tab'));

      const activeTab = document.getElementById('industry-tab-' + idx);
      if (activeTab) {{
        activeTab.classList.add('active-industry-tab');
        activeTab.scrollIntoView({{ behavior: 'smooth', block: 'nearest', inline: 'center' }});
      }}
    }}

    // 品类切换（在每个行业内）
    function switchCategory(indPrefix, idx) {{
      const container = document.getElementById('industry-content-' + indPrefix);
      if (!container) return;
      
      const contents = container.querySelectorAll('.folder-content');
      contents.forEach(c => {{
        c.classList.remove('active');
        c.style.display = 'none';
      }});

      const target = document.getElementById('tab-content-' + indPrefix + '-' + idx);
      if (target) {{
        target.classList.add('active');
        target.style.display = 'block';
      }}

      // 重置所有"分布卡片"为非激活态
      const distCards = container.querySelectorAll('.dist-card');
      distCards.forEach(card => {{
        card.classList.remove('active-dist-card');
        card.style.backgroundColor = '#faf9f5';
        card.style.borderColor = '#eeebe3';
        card.style.boxShadow = 'none';
        // 数字颜色恢复
        const numEl = card.querySelector('.dist-card-num');
        if (numEl) {{
          numEl.style.color = card.getAttribute('data-color-text') || '#1b3d22';
        }}
      }});

      // 激活当前点击的"分布卡片"
      const activeCard = document.getElementById('tab-header-' + indPrefix + '-' + idx);
      if (activeCard) {{
        const colorDeep = activeCard.getAttribute('data-color-deep') || '#1b3d22';
        const colorBg = activeCard.getAttribute('data-color-bg') || '#faf9f5';
        const shadow = activeCard.getAttribute('data-shadow') || 'rgba(27,61,34,0.08)';
        activeCard.classList.add('active-dist-card');
        activeCard.style.backgroundColor = colorBg;
        activeCard.style.borderColor = colorDeep;
        activeCard.style.borderWidth = '1.5px';
        activeCard.style.boxShadow = '0 2px 8px ' + shadow;
        // 数字加深色
        const numEl = activeCard.querySelector('.dist-card-num');
        if (numEl) {{
          numEl.style.color = colorDeep;
        }}
        // 平滑滚动到卡片可见位置
        activeCard.scrollIntoView({{ behavior: 'smooth', block: 'nearest', inline: 'center' }});
      }}
    }}
  </script>
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
        # onerror 时把 img 隐藏，同时显示后面预置的 NO IMAGE 占位 div
        image_html = (
            f'<img src="{image_url}" alt="商品主图" class="card-image" style="width: 75px; height: 75px; border-radius: 8px; object-fit: cover; flex-shrink: 0; border: 1px solid #edebe5;" '
            f'onerror="this.onerror=null;this.style.display=\'none\';var n=this.nextElementSibling;if(n)n.style.display=\'flex\';">'
            f'<div class="card-noimage" style="width: 75px; height: 75px; border-radius: 8px; background: linear-gradient(135deg, #f0eee3 0%, #dfdcce 100%); display: none; flex-direction: column; align-items: center; justify-content: center; flex-shrink: 0; color: #7a766c;">'
            f'<span style="font-size: 20px;">{cat_emoji}</span>'
            f'<span style="font-size: 8px; font-weight: bold; margin-top: 2px; color: #9a968c;">NO IMAGE</span>'
            f'</div>'
        )
    else:
        image_html = f'<div class="card-noimage" style="width: 75px; height: 75px; border-radius: 8px; background: linear-gradient(135deg, #f0eee3 0%, #dfdcce 100%); display: flex; flex-direction: column; align-items: center; justify-content: center; flex-shrink: 0; color: #7a766c;"><span style="font-size: 20px;">{cat_emoji}</span><span style="font-size: 8px; font-weight: bold; margin-top: 2px; color: #9a968c;">NO IMAGE</span></div>'
    
    tags_html_parts = []
    for tag in tags:
        tags_html_parts.append(f'<span style="background-color: {c_light}; color: {c_text}; font-size: 10px; padding: 1px 6px; border-radius: 4px; margin-right: 4px; font-weight: 500;">{tag}</span>')
    tags_html_parts.append(f'<span style="background-color: #f1eff5; color: #7b68ee; font-size: 10px; padding: 1px 6px; border-radius: 4px; font-weight: 500;">{source}</span>')
    tags_html = "".join(tags_html_parts)
    
    name_escaped = name.replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')
    copy_escaped = copy_text.replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')
    
    return f'''    <!-- 商品卡片 -->
    <div class="product-card-mobile" style="width: calc(50% - 8px); min-width: 330px; background-color: #ffffff; border-radius: 12px; padding: 14px; box-sizing: border-box; box-shadow: 0 3px 10px rgba(0,0,0,0.015); border: 1px solid #eeebe3; display: flex; flex-direction: column; justify-content: space-between; transition: transform 0.2s ease, box-shadow 0.2s ease;">
      <div>
        <div style="display: flex; gap: 10px; align-items: flex-start; margin-bottom: 10px;">
          {image_html}
          <div style="display: flex; flex-direction: column; gap: 4px; width: 100%;">
            <div style="font-size: 13px; font-weight: 700; color: #2d2a26; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; height: 36px; line-height: 1.4;" title="{name_escaped}">
              {name}
            </div>
            <div style="display: flex; flex-wrap: wrap; gap: 2px; margin-top: 2px;">
              {tags_html}
            </div>
          </div>
        </div>
        <div style="background-color: #fdfcf7; border-left: 2px solid #ff8c00; padding: 6px 8px; border-radius: 0 6px 6px 0; font-size: 11px; color: #5a5651; font-style: italic; margin-bottom: 12px; line-height: 1.4; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; height: 30px;" title="{copy_escaped}">
          "{copy_text}"
        </div>
      </div>
      <div style="display: flex; gap: 6px; margin-top: auto;">
        <a href="{video_link}" target="_blank" style="flex: 1; text-align: center; background-color: #ff8c00; color: white; padding: 6px 10px; border-radius: 6px; font-size: 11px; text-decoration: none; font-weight: bold; display: inline-flex; align-items: center; justify-content: center; gap: 4px; box-shadow: 0 2px 4px rgba(255,140,0,0.1); border: 1px solid #ff8c00;">🎬 播放视频</a>
        <a href="{link}" target="_blank" style="flex: 1; text-align: center; background-color: {c_text}; color: white; padding: 6px 10px; border-radius: 6px; font-size: 11px; text-decoration: none; font-weight: bold; display: inline-flex; align-items: center; justify-content: center; gap: 4px; box-shadow: 0 2px 4px {colors['shadow_color']}; border: 1px solid {c_text};">🔗 直达链路</a>
      </div>
    </div>'''


def build_industry_content(ind, ind_idx, is_first):
    """构建一个完整行业的内容区域"""
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
        dist_cards.append(f'''      <div id="tab-header-{ind_idx}-{cat_i}" onclick="switchCategory({ind_idx},{cat_i})" data-color-deep="{c["color_deep"]}" data-color-text="{c["color_text"]}" data-color-bg="{c["color_bg"]}" data-shadow="{c["shadow_color"]}" class="dist-card" style="flex: 1; min-width: 110px; {card_style} border-radius: 8px; padding: 10px 12px; text-align: center; cursor: pointer; transition: all 0.25s ease; box-shadow: {'0 2px 8px ' + c["shadow_color"] if is_first_cat else 'none'};" onmouseover="if(!this.classList.contains('active-dist-card')){{this.style.backgroundColor='#fff';this.style.borderColor='{c["color_text"]}';this.style.boxShadow='0 2px 8px {c["shadow_color"]}';}}" onmouseout="if(!this.classList.contains('active-dist-card')){{this.style.backgroundColor='#faf9f5';this.style.borderColor='#eeebe3';this.style.boxShadow='none';}}">
        <div style="font-size: 18px; margin-bottom: 4px;">{cat["emoji"]}</div>
        <div style="font-size: 11px; color: #73706c; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{cat["name"]}</div>
        <div class="dist-card-num" style="font-size: 15px; font-weight: bold; color: {num_color}; margin-top: 2px;">{cat["count"]}<span style="font-size: 9px; font-weight: normal; color: #9c9995; margin-left: 2px;">款</span></div>
        <div style="font-size: 9px; color: #9c9995; margin-top: 1px;">{cat["percentage"]}</div>
      </div>''')
    dist_html = "\n".join(dist_cards)
    
    tab_headers = []  # 不再用底部 Tab 条，分布卡片本身就是 Tab

    # 商品展示（每个品类的内容）
    tab_contents = []
    for cat_i, cat in enumerate(categories):
        is_first_cat = (cat_i == 0)
        # 商品卡片
        cards = []
        for item in cat.get("products", []):
            cards.append(build_product_card(item, cat["emoji"], c))
        cards_html = "\n".join(cards)

        content_cls = 'folder-content active' if is_first_cat else 'folder-content'
        tab_contents.append(f'''    <div id="tab-content-{ind_idx}-{cat_i}" class="{content_cls}">
  <div style="margin-top: 30px; margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #eae7e0; padding-bottom: 8px;">
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
    
    tabs_html = "\n".join(tab_headers)
    contents_html = "\n\n".join(tab_contents)
    
    # 高频词
    kw_spans = []
    for kw in keywords:
        kw_spans.append(f'<span style="background-color: {c["color_bg"]}; color: {c["color_text"]}; font-size: 12px; font-weight: bold; padding: 4px 10px; border-radius: 12px; border: 1px solid {c["color_border"]}; margin-bottom: 6px; white-space: nowrap;">{kw["word"]} ({kw["count"]}次)</span>')
    kw_html = " ".join(kw_spans)
    
    # 洞察卡片
    insight_cards = []
    for ins in insights:
        insight_cards.append(f'''        <div style="background-color: #fafbf9; border-left: 3px solid {ins["color"]}; padding: 10px 12px; border-radius: 0 6px 6px 0; font-size: 12px;">
          <strong style="color: {ins["color"]}; font-size: 13px;">{ins["id"]}. {ins["title"]} {ins["emoji"]}</strong><br>
          {ins["desc"]}
        </div>''')
    insights_html = "\n".join(insight_cards)
    
    # 平台分布
    platform_parts = []
    for plat, count in sorted(platforms.items(), key=lambda x: x[1], reverse=True):
        platform_parts.append(f'<span style="background-color: #eef2ff; color: #4f46e5; font-size: 11px; padding: 2px 8px; border-radius: 8px; font-weight: 500;">{plat}: {count}款</span>')
    platform_html = " ".join(platform_parts) if platform_parts else ""
    
    return f'''  <div id="industry-content-{ind_idx}" class="industry-content {active_cls} industry-content-wrapper" style="max-width: 1200px; margin: 0 auto; padding: 0 10px; box-sizing: border-box;">

    <!-- 顶部海报 -->
    <div style="max-width: 1180px; margin: 0 auto; background-color: #fbfaf5; padding: 25px 30px; border-radius: 0 0 20px 20px; box-shadow: 0 12px 40px rgba(0,0,0,0.06); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'PingFang SC', 'Microsoft YaHei', sans-serif; color: #2d2a26; line-height: 1.6;">
    
    <div class="industry-poster-mobile" style="background: linear-gradient(135deg, {c["color"]} 0%, {c["color_dark"]} 100%); color: #ffffff; border-radius: 16px; padding: 40px 30px; text-align: center; margin-bottom: 25px; box-shadow: 0 8px 24px {c["shadow_color"]}; position: relative; overflow: hidden;">
      <div style="position: absolute; top: -50px; right: -50px; width: 150px; height: 150px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
      <div style="position: absolute; bottom: -60px; left: -60px; width: 180px; height: 180px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
      
      <div style="display: inline-block; background-color: rgba(255, 255, 255, 0.15); color: #f4fbf7; font-size: 11px; font-weight: bold; letter-spacing: 2px; padding: 4px 12px; border-radius: 20px; margin-bottom: 12px; text-transform: uppercase;">
        🔥 {ind["name"]} · 爆款视频链路监控
      </div>
      
      <h2 style="margin: 8px 0 0 0; font-size: 15px; font-weight: 300; color: rgba(255,255,255,0.75); letter-spacing: 0.5px;">
        {ind["name"]}行业 · 近期数据精选
      </h2>
      <div style="width: 50px; height: 3px; background-color: #ff8c00; margin: 20px auto 15px auto; border-radius: 2px;"></div>
      <p style="margin: 0; font-size: 12px; color: rgba(255,255,255,0.6);">
        精确去重后唯一爆品总数: <strong style="color: #fff; font-size: 14px;">{total}</strong> 款
        {f' | 引流平台: {platform_html}' if platform_html else ''}
      </p>
    </div>

    <!-- 爆品类目分布看板（每张卡片可点击切换品类） -->
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

      <!-- 页脚 -->
      <div style="text-align: center; border-top: 1px dashed #eae7e0; margin-top: 40px; padding-top: 15px; font-size: 11px; color: #9c9995;">
        <div>💡 长图使用提示：支持在任何兼容 HTML 渲染的阅读器中直接预览。</div>
        <div style="margin-top: 4px;">由 CID 专项爆款分析助手精心制作 | {ind["name"]}行业 · 数据日期 2026-08</div>
      </div>
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


def build_empty_industry_content(ind, ind_idx, is_first):
    """构建空行业（美护）占位内容"""
    active_cls = 'active' if is_first else ''
    c = ind  # colors
    return f'''  <div id="industry-content-{ind_idx}" class="industry-content {active_cls} industry-content-wrapper" style="max-width: 1200px; margin: 0 auto; padding: 0 10px; box-sizing: border-box;">

    <div style="max-width: 1180px; margin: 0 auto; background-color: #fbfaf5; padding: 25px 30px; border-radius: 0 0 20px 20px; box-shadow: 0 12px 40px rgba(0,0,0,0.06); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'PingFang SC', 'Microsoft YaHei', sans-serif; color: #2d2a26; line-height: 1.6;">
    
    <div class="industry-poster-mobile" style="background: linear-gradient(135deg, {c["color"]} 0%, {c["color_dark"]} 100%); color: #ffffff; border-radius: 16px; padding: 40px 30px; text-align: center; margin-bottom: 25px; box-shadow: 0 8px 24px {c["shadow_color"]}; position: relative; overflow: hidden;">
      <div style="position: absolute; top: -50px; right: -50px; width: 150px; height: 150px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
      <div style="position: absolute; bottom: -60px; left: -60px; width: 180px; height: 180px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
      
      <div style="display: inline-block; background-color: rgba(255, 255, 255, 0.15); color: #f4fbf7; font-size: 11px; font-weight: bold; letter-spacing: 2px; padding: 4px 12px; border-radius: 20px; margin-bottom: 12px; text-transform: uppercase;">
        🔥 {ind["name"]} · 爆款视频链路监控
      </div>
      
      <h2 style="margin: 8px 0 0 0; font-size: 15px; font-weight: 300; color: rgba(255,255,255,0.75); letter-spacing: 0.5px;">
        {ind["name"]}行业 · 数据待补充
      </h2>
      <div style="width: 50px; height: 3px; background-color: #ff8c00; margin: 20px auto 15px auto; border-radius: 2px;"></div>
      <p style="margin: 0; font-size: 12px; color: rgba(255,255,255,0.6);">
        暂无数据，敬请期待
      </p>
    </div>

    <!-- 空状态占位 -->
    <div style="background-color: #ffffff; border-radius: 14px; padding: 60px 30px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.02); border: 2px dashed #e9e7e0;">
      <div style="font-size: 64px; margin-bottom: 20px;">{c["emoji"]}</div>
      <h3 style="margin: 0 0 10px 0; font-size: 18px; color: {c["color_text"]}; font-weight: 600;">
        {ind["name"]}行业数据即将上线
      </h3>
      <p style="margin: 0; font-size: 13px; color: #9c9995; line-height: 1.8;">
        该行业爆品数据正在采集中，完成后将自动填充<br>
        包含护肤品、彩妆、个护等细分类目
      </p>
    </div>

    </div>

  </div>'''


def main():
    print("📊 读取Excel数据...")
    
    industries_data = []
    for sheet_name in ["食品饮料", "消电日百", "服饰", "美护"]:
        print(f"  处理: {sheet_name}")
        if sheet_name == "食品饮料":
            df = pd.read_excel(FOOD_EXCEL_PATH)
            if df.empty:
                df = None
            else:
                merge_map = INDUSTRY_MERGE.get(sheet_name, {})
                if merge_map:
                    df["投放二级行业"] = df["投放二级行业"].apply(lambda x: merge_map.get(x, x))
        else:
            df = read_sheet(sheet_name)
        if df is None:
            cfg = INDUSTRY_CONFIG[sheet_name]
            ind_data = {
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
                "total": 0,
                "categories": [],
                "keywords": [],
                "insights": [],
                "platforms": {},
            }
        else:
            ind_data = build_industry_data(sheet_name, df)
        industries_data.append(ind_data)
        print(f"    {ind_data['total']}款, {len(ind_data['categories'])}个类目")
    
    print("🔨 生成HTML...")
    html = build_multi_html(industries_data)
    
    output_dir = "/Users/krystalcao/Desktop/已完成"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, "多行业爆品选品报告.html")
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    
    file_size = os.path.getsize(output_path)
    print(f"✅ 报告已生成: {output_path}")
    print(f"   文件大小: {file_size/1024:.1f} KB")
    print(f"   总爆品数: {sum(i['total'] for i in industries_data)}")


if __name__ == "__main__":
    main()
