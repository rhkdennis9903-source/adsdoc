import streamlit as st
import pandas as pd
from docx import Document
import io
import datetime

# --- 設定頁面配置 ---
st.set_page_config(page_title="Meta 廣告投放配置生成器", layout="wide")

# --- 預設資料 (基於您提供的 CSV 檔案提取，作為備用) ---
DEFAULT_TAGS = [
    "Facebook Page Admins (粉專管理員)", "Small business (小型企業)", "Entrepreneurship (創業)",
    "商業 / 生產力軟體和應用程式", "Digital marketing (數位行銷)", "Management (管理)",
    "企業顧問 (Business consultant)", "專業商業技能 (Professional business skills)",
    "Web design (網頁設計)", "Business travelers (商務旅客)", "Coworking", "Software"
]

# --- 核心函式：解析 Word 檔中的廣告 ID (增強版) ---
def parse_ad_ids_from_docx(uploaded_file):
    """
    解析 Word 文件中的【廣告組合 ID】。
    邏輯：
    1. 尋找關鍵字 "【廣告組合 ID】"
    2. 策略 A：檢查同一行是否有內容 (支援去除逗號、冒號等符號)
    3. 策略 B：若同一行是空的，則抓取下一行的內容
    """
    ad_ids = []
    try:
        doc = Document(uploaded_file)
        paragraphs = doc.paragraphs
        
        for i, para in enumerate(paragraphs):
            text = para.text.strip()
            
            # 偵測關鍵字
            if "【廣告組合 ID】" in text:
                # --- 策略 1: 嘗試在同一行抓取 ---
                parts = text.split("【廣告組合 ID】")
                
                # 預設候選字串為空
                candidate = ""
                
                if len(parts) > 1:
                    # 清洗同一行的內容 (去除全形/半形 逗號與冒號)
                    candidate = parts[1].strip().lstrip(",").lstrip(":").lstrip("，").lstrip("：").strip()
                
                if candidate:
                    # 如果同一行有東西，就直接使用
                    ad_ids.append(candidate)
                else:
                    # --- 策略 2: 如果同一行是空的，嘗試抓下一行 ---
                    if i + 1 < len(paragraphs):
                        next_para_text = paragraphs[i + 1].text.strip()
                        if next_para_text:
                            # 同樣進行清洗，以防萬一
                            clean_next = next_para_text.lstrip(",").lstrip(":").lstrip("，").lstrip("：").strip()
                            if clean_next:
                                ad_ids.append(clean_next)
                                
    except Exception as e:
        st.error(f"解析 Word 檔時發生錯誤: {e}")
    
    # 移除重複值並排序，讓選單整齊
    return sorted(list(set(ad_ids)))

# --- 輔助函式：解析 CSV/Excel 受眾標籤 ---
def parse_tags_from_csv(uploaded_file):
    tags = []
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # 尋找 "受眾標籤 (Tag)" 欄位
        if "受眾標籤 (Tag)" in df.columns:
            tags = df["受眾標籤 (Tag)"].dropna().unique().tolist()
        else:
            st.warning("上傳的檔案中找不到 '受眾標籤 (Tag)' 欄位，請檢查檔案格式。")
    except Exception as e:
        st.error(f"解析受眾標籤檔時發生錯誤: {e}")
    return tags

# --- 初始化 Session State ---
if 'ad_sets' not in st.session_state:
    st.session_state.ad_sets = []

# --- 側邊欄：資料來源與設定 ---
with st.sidebar:
    st.header("📂 資料來源設定")
    st.markdown("上傳檔案以自動載入選項，若不上傳則使用預設值。")
    
    # 上傳 Word (抓取廣告 ID)
    docx_file = st.file_uploader("上傳上刊文件 (Word)", type=['docx'])
    available_ad_ids = []
    if docx_file:
        available_ad_ids = parse_ad_ids_from_docx(docx_file)
        st.success(f"已載入 {len(available_ad_ids)} 個廣告 ID")
        if available_ad_ids:
            with st.expander("查看已抓取到的 ID"):
                st.write(available_ad_ids)
    else:
        st.info("尚未上傳上刊文件，將無法選擇特定的廣告 ID。")

    # 上傳 Excel/CSV (抓取受眾標籤)
    tag_file = st.file_uploader("上傳受眾標籤 (CSV/Excel)", type=['csv', 'xlsx'])
    available_tags = DEFAULT_TAGS.copy()
    if tag_file:
        uploaded_tags = parse_tags_from_csv(tag_file)
        if uploaded_tags:
            available_tags = uploaded_tags
            st.success(f"已更新受眾標籤庫 ({len(available_tags)} 筆)")

# --- 主畫面 ---
st.title("🚀 Meta 廣告投放配置生成工具")
st.markdown("---")

# 第一部分：行銷活動 (Campaign)
st.header("1. 行銷活動 (Campaign)")
col1, col2 = st.columns(2)

with col1:
    campaign_name = st.text_input("行銷活動名稱", placeholder="例如：2025_Q1_品牌推廣_轉換")
    campaign_objective = st.selectbox("行銷活動目標", 
        ["銷售 (Sales)", "潛在客戶 (Leads)", "流量 (Traffic)", "互動 (Engagement)", "知名度 (Awareness)", "應用程式推廣 (App Promotion)"]
    )
    is_cbo = st.toggle("開啟行銷活動預算最佳化 (Advantage+ Campaign Budget)", value=True)

with col2:
    budget_type = st.selectbox("預算類型", ["單日預算", "總預算"])
    budget_amount = st.number_input("預算金額 (TWD)", min_value=0, value=1000, step=100)
    start_date = st.date_input("開始日期", datetime.date.today())

st.markdown("---")

# 第二部分：廣告組合 (Ad Sets) 管理
st.header("2. 廣告組合 (Ad Sets)")

# 新增廣告組合按鈕
if st.button("➕ 新增一個廣告組合"):
    st.session_state.ad_sets.append({
        "id": len(st.session_state.ad_sets) + 1,
        "name": "",
        "goal": "網站",
        "tags": [],
        "custom_audience": "",
        "age_min": 25,
        "age_max": 55,
        "locations": ["台灣"],
        "gender": "所有性別",
        "advantage_audience": True,
        "placements_type": "Advantage+ (自動版位)",
        "excluded_placements": [],
        "selected_ads": []
    })

# 顯示所有廣告組合
ad_sets_to_remove = []

for index, ad_set in enumerate(st.session_state.ad_sets):
    with st.expander(f"廣告組合 #{index + 1}: {ad_set['name'] if ad_set['name'] else '未命名'}", expanded=True):
        
        # 刪除按鈕
        if st.button("🗑️ 刪除此組合", key=f"del_{index}"):
            ad_sets_to_remove.append(index)
        
        c1, c2 = st.columns(2)
        
        # 組合基本設定
        with c1:
            ad_set['name'] = st.text_input("廣告組合名稱", value=ad_set['name'], key=f"name_{index}")
            ad_set['goal'] = st.selectbox("轉換目標位置", ["網站", "應用程式", "訊息", "通話"], key=f"goal_{index}")
        
        with c2:
            ad_set['custom_audience'] = st.text_area("自訂受眾 / 類似受眾 (請手動輸入)", value=ad_set['custom_audience'], placeholder="例如：官網訪客 30天, 購買者 1% 類似受眾", key=f"ca_{index}")

        st.subheader("受眾設定")
        ac1, ac2 = st.columns(2)
        with ac1:
            # 興趣標籤 (Multiselect + 自行輸入)
            ad_set['tags'] = st.multiselect(
                "興趣標籤 (Tag) - 可多選", 
                options=available_tags,
                default=list(set(ad_set['tags']) & set(available_tags)), # 防止預設值不在選項中報錯
                key=f"tags_{index}"
            )
            # 補充手動輸入標籤
            extra_tags = st.text_input("自行輸入其他標籤 (以逗號分隔)", key=f"extra_tags_{index}")
            if extra_tags:
                current_extras = [t.strip() for t in extra_tags.split(",") if t.strip()]
                ad_set['manual_tags'] = current_extras
            else:
                ad_set['manual_tags'] = []

            ad_set['advantage_audience'] = st.checkbox("開啟 Advantage+ 受眾 (高效速成受眾)", value=ad_set['advantage_audience'], key=f"aa_{index}")

        with ac2:
            lc1, lc2 = st.columns(2)
            with lc1:
                ad_set['age_min'] = st.number_input("年齡下限", 13, 65, ad_set['age_min'], key=f"amin_{index}")
            with lc2:
                ad_set['age_max'] = st.number_input("年齡上限 (65+)", 13, 65, ad_set['age_max'], key=f"amax_{index}")
            
            ad_set['gender'] = st.radio("性別", ["所有性別", "男性", "女性"], horizontal=True, key=f"gen_{index}")
            
            # 地點
            loc_str = st.text_input("地理位置", value=",".join(ad_set['locations']), key=f"loc_{index}")
            ad_set['locations'] = [l.strip() for l in loc_str.split(",")]

        st.subheader("版位設定")
        pc1, pc2 = st.columns(2)
        with pc1:
            placement_mode = st.radio("版位模式", ["Advantage+ (自動版位)", "手動指定版位"], key=f"pmode_{index}")
            ad_set['placements_type'] = placement_mode
        
        with pc2:
            if placement_mode == "手動指定版位":
                ad_set['excluded_placements'] = st.multiselect("排除版位 / 指定版位說明", 
                    ["Facebook 動態牆", "Instagram 動態牆", "Stories (限時動態)", "Reels", "Audience Network", "Right Column"],
                    key=f"excl_{index}"
                )
                ad_set['placement_note'] = st.text_input("其他版位備註", key=f"pnote_{index}")
            else:
                ad_set['excluded_placements'] = []
                ad_set['placement_note'] = ""

        st.subheader("3. 廣告素材選擇")
        if available_ad_ids:
            # 這裡使用解析出來的 ID 讓使用者選擇
            ad_set['selected_ads'] = st.multiselect(
                "選擇要包含在此組合的廣告 ID (來自上刊文件)",
                options=available_ad_ids,
                default=list(set(ad_set['selected_ads']) & set(available_ad_ids)),
                key=f"ads_{index}"
            )
        else:
            st.warning("請先在左側上傳 Word 上刊文件以選擇廣告 ID")
            # 允許手動輸入 ID 作為備案
            manual_ad_ids = st.text_area("或手動輸入廣告 ID (一行一個)", key=f"man_ads_{index}")
            if manual_ad_ids:
                ad_set['manual_ads'] = [line.strip() for line in manual_ad_ids.split("\n") if line.strip()]
            else:
                ad_set['manual_ads'] = []

# 移除被標記刪除的廣告組合
for i in sorted(ad_sets_to_remove, reverse=True):
    del st.session_state.ad_sets[i]

st.markdown("---")

# --- 生成報告 ---
st.header("📄 產出投放配置說明")

if st.button("生成 Markdown 文件", type="primary"):
    # 建構文件內容
    report = f"# Meta 廣告投放配置指令\n\n"
    report += f"**生成時間**: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"
    
    report += f"## 1. 行銷活動 (Campaign)\n"
    report += f"- **名稱**: {campaign_name}\n"
    report += f"- **目標**: {campaign_objective}\n"
    report += f"- **預算**: {budget_type} NT$ {budget_amount}\n"
    report += f"- **預算最佳化 (CBO)**: {'開啟' if is_cbo else '關閉'}\n"
    report += f"- **走期開始**: {start_date}\n\n"
    
    report += f"## 2. 廣告組合與素材配置\n"
    
    if not st.session_state.ad_sets:
        report += "> 尚未建立任何廣告組合。\n"
    
    for i, ad_set in enumerate(st.session_state.ad_sets):
        report += f"### 廣告組合 {i+1}: {ad_set['name']}\n"
        report += f"**A. 目標設定**\n"
        report += f"- 轉換位置: {ad_set['goal']}\n\n"
        
        report += f"**B. 受眾設定**\n"
        report += f"- **地點**: {', '.join(ad_set['locations'])}\n"
        report += f"- **年齡**: {ad_set['age_min']} - {ad_set['age_max']}\n"
        report += f"- **性別**: {ad_set['gender']}\n"
        report += f"- **高效速成受眾 (Advantage+)**: {'開啟' if ad_set['advantage_audience'] else '關閉'}\n"
        
        all_tags = ad_set['tags'] + ad_set.get('manual_tags', [])
        tags_str = ", ".join(all_tags) if all_tags else "無"
        report += f"- **興趣標籤**: {tags_str}\n"
        
        ca_str = ad_set['custom_audience'] if ad_set['custom_audience'] else "無"
        report += f"- **自訂/類似受眾**: {ca_str}\n\n"
        
        report += f"**C. 版位設定**\n"
        report += f"- **模式**: {ad_set['placements_type']}\n"
        if ad_set['placements_type'] == "手動指定版位":
             report += f"- **排除/指定**: {', '.join(ad_set['excluded_placements'])}\n"
             if ad_set['placement_note']:
                 report += f"- **備註**: {ad_set['placement_note']}\n"
        
        report += f"\n**D. 廣告素材 (Ads)**\n"
        report += f"請使用以下上刊文件中的 ID 進行設定：\n"
        
        # 合併選單選擇的 ID 與 手動輸入的 ID
        final_ads = ad_set.get('selected_ads', []) + ad_set.get('manual_ads', [])
        
        if final_ads:
            for ad_id in final_ads:
                report += f"- ` {ad_id} `\n"
        else:
            report += "- (未指定廣告 ID)\n"
        
        report += "\n---\n"

    st.text_area("複製下方內容提供給廣告投手：", value=report, height=600)
    
    # 提供下載按鈕
    st.download_button(
        label="下載 .txt 檔案",
        data=report,
        file_name=f"Meta廣告配置_{datetime.date.today()}.txt",
        mime="text/plain"
    )
