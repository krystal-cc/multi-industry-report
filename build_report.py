#!/usr/bin/env python3
"""
爆品选品报告生成器
用法：将数据填入 data.json 中，运行 python3 build_report.py 即可生成 HTML 报告。
"""

import json
import os

def load_template():
    """加载模板"""
    with open("template.html", "r", encoding="utf-8") as f:
        return f.read()

def build_badge_html(tags, source):
    """生成标签 HTML"""
    parts = []
    for tag in tags:
        parts.append(f'<span style="background-color: #eaf2eb; color: #426b50; font-size: 10px; padding: 1px 6px; border-radius: 4px; margin-right: 4px; font-weight: 500;">{tag}</span>')
    parts.append(f'<span style="background-color: #f1eff5; color: #7b68ee; font-size: 10px; padding: 1px 6px; border-radius: 4px; font-weight: 500;">数据源:{source}</span>')
    return "".join(parts)

def build_product_card(item, category_emoji):
    """生成单个商品卡片 HTML"""
    # 图片区域
    if item.get("image_url"):
        image_html = f'<img src="{item["image_url"]}" alt="商品主图" style="width: 75px; height: 75px; border-radius: 8px; object-fit: cover; flex-shrink: 0; border: 1px solid #edebe5;">'
    else:
        image_html = f'''<div style="width: 75px; height: 75px; border-radius: 8px; background: linear-gradient(135deg, #f0eee3 0%, #dfdcce 100%); display: flex; flex-direction: column; align-items: center; justify-content: center; flex-shrink: 0; color: #7a766c;">
              <span style="font-size: 20px;">{category_emoji}</span>
              <span style="font-size: 8px; font-weight: bold; margin-top: 2px; color: #9a968c;">NO IMAGE</span>
            </div>'''

    tags_html = build_badge_html(item.get("tags", []), item.get("source", ""))
    
    return f'''    <!-- 商品卡片 -->
    <div style="width: calc(50% - 8px); min-width: 330px; background-color: #ffffff; border-radius: 12px; padding: 14px; box-sizing: border-box; box-shadow: 0 3px 10px rgba(0,0,0,0.015); border: 1px solid #eeebe3; display: flex; flex-direction: column; justify-content: space-between; transition: transform 0.2s ease, box-shadow 0.2s ease;">
      <div>
        <!-- 卡片头部: 图片 + 名称 + 标签 -->
        <div style="display: flex; gap: 10px; align-items: flex-start; margin-bottom: 10px;">
          {image_html}
          <div style="display: flex; flex-direction: column; gap: 4px; width: 100%;">
            <div style="font-size: 13px; font-weight: 700; color: #2d2a26; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; height: 36px; line-height: 1.4;" title="{title_escaped}">
              {title}
            </div>
            <div style="display: flex; flex-wrap: wrap; gap: 2px; margin-top: 2px;">
              {tags_html}
            </div>
          </div>
        </div>
        
        <!-- 创意广告语 -->
        <div style="background-color: #fdfcf7; border-left: 2px solid #ff8c00; padding: 6px 8px; border-radius: 0 6px 6px 0; font-size: 11px; color: #5a5651; font-style: italic; margin-bottom: 12px; line-height: 1.4; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; height: 30px;" title="{copy_escaped}">
          "{copy}"
        </div>
      </div>
      
      <!-- 底部并排按钮 -->
      <div style="display: flex; gap: 6px; margin-top: auto;">
        <a href="{video_link}" target="_blank" style="flex: 1; text-align: center; background-color: #ff8c00; color: white; padding: 6px 10px; border-radius: 6px; font-size: 11px; text-decoration: none; font-weight: bold; display: inline-flex; align-items: center; justify-content: center; gap: 4px; box-shadow: 0 2px 4px rgba(255,140,0,0.1); border: 1px solid #ff8c00;">🎬 播放视频</a>
        <a href="{link}" target="_blank" style="flex: 1; text-align: center; background-color: #426b50; color: white; padding: 6px 10px; border-radius: 6px; font-size: 11px; text-decoration: none; font-weight: bold; display: inline-flex; align-items: center; justify-content: center; gap: 4px; box-shadow: 0 2px 4px rgba(66,107,80,0.1); border: 1px solid #426b50;">🔗 直达链路</a>
      </div>
    </div>'''.format(
        title=item.get("name", ""),
        title_escaped=item.get("name", "").replace('"', '&quot;'),
        copy=item.get("copy", ""),
        copy_escaped=item.get("copy", "").replace('"', '&quot;'),
        video_link=item.get("video_link", "#"),
        link=item.get("link", "#"),
        image_html=image_html,
        tags_html=tags_html
    )

def build_distribution_card(cat):
    """生成分布看板卡片"""
    return f'''      <div style="flex: 1; min-width: 110px; background-color: #faf9f5; border-radius: 8px; padding: 10px 12px; border: 1px solid #eeebe3; text-align: center;">
        <div style="font-size: 18px; margin-bottom: 4px;">{cat["emoji"]}</div>
        <div style="font-size: 11px; color: #73706c; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{cat["name"]}</div>
        <div style="font-size: 15px; font-weight: bold; color: #426b50; margin-top: 2px;">{cat["count"]}<span style="font-size: 9px; font-weight: normal; color: #9c9995; margin-left: 2px;">款</span></div>
        <div style="font-size: 9px; color: #9c9995; margin-top: 1px;">{cat["percentage"]}</div>
      </div>'''

def build_tab_header(cat, idx, is_first):
    """生成 Tab 头部"""
    if is_first:
        cls = 'active-tab text-forest'
    else:
        cls = 'text-slate-400 bg-transparent'
    return f'''    <div id="tab-header-{idx}" onclick="switchCategory({idx})" class="folder-tab px-5 py-3 text-xs md:text-sm font-black flex items-center gap-1.5 cursor-pointer select-none transition-all duration-300 {cls}">
      <span>{cat["emoji"]}</span> {cat["name"]}
    </div>'''

def build_tab_content(cat, idx, is_first):
    """生成 Tab 内容区域"""
    if is_first:
        cls = 'folder-content active'
    else:
        cls = 'folder-content'
    
    cards_html = ""
    for item in cat.get("products", []):
        cards_html += build_product_card(item, cat["emoji"]) + "\n"
    
    return f'''    <div id="tab-content-{idx}" class="{cls}">
      
  <div style="margin-top: 30px; margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #eae7e0; padding-bottom: 8px;">
    <div style="font-size: 18px; font-weight: bold; color: #224332; display: flex; align-items: center; gap: 8px;">
      <span style="font-size: 22px;">{cat["emoji"]}</span> {cat["name"]}
      <span style="background-color: #eae7e0; color: #5c5a56; font-size: 11px; padding: 1px 8px; border-radius: 10px; font-weight: normal; margin-left: 6px;">{cat["count"]} 款爆品</span>
    </div>
    <div style="font-size: 12px; color: #8c8985;">分行业精选</div>
  </div>
  
  <div style="display: flex; flex-wrap: wrap; gap: 15px;">
{cards_html}
  </div>'''

def build_insight_card(insight):
    """生成带货密码卡片"""
    return f'''        <div style="background-color: #fafbf9; border-left: 3px solid {insight["color"]}; padding: 10px 12px; border-radius: 0 6px 6px 0; font-size: 12px;">
          <strong style="color: {insight["color"]}; font-size: 13px;">{insight["id"]}. {insight["title"]} {insight["emoji"]}</strong><br>
          {insight["desc"]}
        </div>'''

def build_keyword_span(kw):
    """生成关键词 span"""
    return f'<span style="background-color: #f2f7f3; color: #426b50; font-size: 12px; font-weight: bold; padding: 4px 10px; border-radius: 12px; border: 1px solid #dbe6e0; margin-bottom: 6px; white-space: nowrap;">{kw["word"]} ({kw["count"]}次)</span>'

def generate(data):
    """根据数据生成完整 HTML"""
    
    # 1. 分布看板
    dist_cards = "\n".join(build_distribution_card(c) for c in data["categories"])
    
    # 2. Tab 头部
    tab_headers = "\n".join(build_tab_header(c, i, i == 0) for i, c in enumerate(data["categories"]))
    
    # 3. Tab 内容
    tab_contents = "\n\n".join(build_tab_content(c, i, i == 0) for i, c in enumerate(data["categories"]))
    
    # 4. 高频词
    keywords = " ".join(build_keyword_span(k) for k in data.get("keywords", []))
    
    # 5. 带货密码
    insights = "\n".join(build_insight_card(ins) for ins in data.get("insights", []))
    
    # 组装
    html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{data["title"]}</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Noto+Sans+SC:wght@300;400;500;700;900&display=swap');
    
    body {{
      font-family: 'Outfit', 'Noto Sans SC', -apple-system, BlinkMacSystemFont, sans-serif;
      margin: 0;
      padding: 20px 0;
      background-color: #f3f7f4;
      background-image: 
        radial-gradient(at 0% 0%, hsla(143,32%,94%,1) 0px, transparent 50%),
        radial-gradient(at 100% 100%, hsla(150,30%,90%,1) 0px, transparent 50%);
    }}

    .folder-tab {{
      position: relative;
      background: #ffffff;
      border-radius: 12px 12px 0 0;
      border: 1px solid rgba(27,61,34,0.1);
      border-bottom: none;
      white-space: nowrap;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
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

    .folder-content {{
      opacity: 0;
      transform: translateY(10px) scale(0.99);
      display: none;
      transition: all 0.35s cubic-bezier(0.34, 1.56, 0.64, 1);
    }}
    .folder-content.active {{
      display: block;
      opacity: 1;
      transform: translateY(0) scale(1);
    }}

    .text-forest {{ color: #1b3d22; }}
    .bg-forest {{ background-color: #1b3d22; }}
    .border-forest {{ border-color: #1b3d22; }}
  </style>
</head>
<body class="min-h-screen text-slate-800 antialiased selection:bg-emerald-100 selection:text-emerald-900 pb-16">

  <div style="max-width: 850px; margin: 0 auto; background-color: #fbfaf5; padding: 25px 30px; border-radius: 20px; box-shadow: 0 12px 40px rgba(0,0,0,0.06); font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'PingFang SC', 'Microsoft YaHei', sans-serif; color: #2d2a26; line-height: 1.6;">

  <!-- 顶部视觉海报 -->
  <div style="background: linear-gradient(135deg, #426b50 0%, #224332 100%); color: #ffffff; border-radius: 16px; padding: 40px 30px; text-align: center; margin-bottom: 25px; box-shadow: 0 8px 24px rgba(40, 75, 55, 0.15); position: relative; overflow: hidden;">
    <div style="position: absolute; top: -50px; right: -50px; width: 150px; height: 150px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
    <div style="position: absolute; bottom: -60px; left: -60px; width: 180px; height: 180px; background: rgba(255,255,255,0.03); border-radius: 50%;"></div>
    
    <div style="display: inline-block; background-color: rgba(255, 255, 255, 0.15); color: #f4fbf7; font-size: 11px; font-weight: bold; letter-spacing: 2px; padding: 4px 12px; border-radius: 20px; margin-bottom: 12px; text-transform: uppercase;">
      🔥 {data["badge"]}
    </div>
    
    <h2 style="margin: 8px 0 0 0; font-size: 15px; font-weight: 300; color: #cbd9ce; letter-spacing: 0.5px;">
      {data["subtitle"]}
    </h2>
    <div style="width: 50px; height: 3px; background-color: #ff8c00; margin: 20px auto 15px auto; border-radius: 2px;"></div>
    <p style="margin: 0; font-size: 12px; color: #a4bea8;">
      {data["description"]}
    </p>
  </div>

  <!-- 数据大屏面板 -->
  <div style="background-color: #ffffff; border-radius: 14px; padding: 22px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.02); border: 1px solid #e9e7e0;">
    <h3 style="margin: 0 0 15px 0; font-size: 15px; color: #426b50; font-weight: 700; display: flex; align-items: center; gap: 6px;">
      📊 {data["dashboard_title"]}
    </h3>
    <div style="display: flex; flex-wrap: wrap; gap: 10px;">
{dist_cards}
    </div>
  </div>

  <!-- 品类 Tab 切换区域 -->
  <div style="max-width: 850px; margin: 30px auto 0 auto; padding: 0 10px; box-sizing: border-box;">
    
    <div class="flex items-end overflow-x-auto no-scrollbar gap-1.5 pl-2 select-none" id="category-tabs-container">
{tab_headers}
    </div>

    <div class="bg-white border-2 border-emerald-900/10 rounded-2xl rounded-tl-none p-5 md:p-6 shadow-xl relative min-h-[400px]">
      
{tab_contents}

      <!-- 页脚说明 -->
      <div style="text-align: center; border-top: 1px dashed #eae7e0; margin-top: 40px; padding-top: 15px; font-size: 11px; color: #9c9995;">
        <div>💡 长图使用提示：支持在任何兼容 HTML 渲染的 Markdown 阅读器（如 Notion、VS Code、Typora）中直接预览。</div>
        <div style="margin-top: 4px;">{data["footer"]}</div>
      </div>

    </div>
  </div>

  <!-- 爆品创意黄金法则与高频词 -->
  <div style="max-width: 850px; margin: 30px auto 0 auto; padding: 0 10px; box-sizing: border-box;">
    <div style="background-color: #ffffff; border-radius: 14px; padding: 22px; margin-bottom: 25px; box-shadow: 0 4px 15px rgba(0,0,0,0.02); border: 1px solid #e9e7e0;">
      <h3 style="margin: 0 0 15px 0; font-size: 15px; color: #426b50; font-weight: 700; display: flex; align-items: center; gap: 6px;">
        💡 {data["insights_title"]}
      </h3>
      
      <div style="margin-bottom: 18px;">
        <div style="font-size: 12px; color: #73706c; font-weight: 600; margin-bottom: 8px;">🏷️ {data["keywords_title"]}：</div>
        <div style="display: flex; flex-wrap: wrap; gap: 8px;">
          {keywords}
        </div>
      </div>
      
      <div style="border-top: 1px dashed #eeebe3; padding-top: 15px;">
        <div style="font-size: 12px; color: #73706c; font-weight: 600; margin-bottom: 10px;">🎯 {data["secrets_title"]}：</div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 12px;">
{insights}
        </div>
      </div>
    </div>
  </div>

  <script>
    function switchCategory(idx) {{
      const contents = document.querySelectorAll('.folder-content');
      contents.forEach(content => content.classList.remove('active'));

      const targetContent = document.getElementById('tab-content-' + idx);
      if (targetContent) {{
        targetContent.classList.add('active');
      }}

      const tabs = document.querySelectorAll('.folder-tab');
      tabs.forEach(tab => {{
        tab.classList.remove('active-tab', 'text-forest');
        tab.classList.add('text-slate-400', 'bg-transparent');
      }});

      const activeTab = document.getElementById('tab-header-' + idx);
      if (activeTab) {{
        activeTab.classList.remove('text-slate-400', 'bg-transparent');
        activeTab.classList.add('active-tab', 'text-forest');
        
        activeTab.scrollIntoView({{
          behavior: 'smooth',
          block: 'nearest',
          inline: 'center'
        }});
      }}
    }}
  </script>
</body>
</html>'''
    return html


def main():
    # 加载数据
    with open("data.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    
    # 生成 HTML
    html = generate(data)
    
    # 确定输出文件名
    output_file = data.get("output_file", "report.html")
    output_dir = data.get("output_dir", "")
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, output_file)
    else:
        output_path = output_file
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    
    print(f"✅ 报告已生成: {output_path}")


if __name__ == "__main__":
    main()
