import streamlit as st
import pandas as pd
from docx import Document
import io
import datetime

# --- 設定頁面配置 ---
st.set_page_config(page_title="Meta 廣告全策略配置工具", layout="wide")

# --- 預設資料 ---
DEFAULT_TAGS = [
    "Facebook Page Admins (粉專管理員)", "Small business (小型企業)", "Entrepreneurship (創業)",
    "商業 / 生產力軟體和應用程式", "Digital marketing (數位行銷)", "Management (管理)",
    "企業顧問 (Business consultant)", "專業商業技能 (Professional business skills)",
    "Web design (網頁設計)", "Business travelers (商務旅客)", "Coworking", "Software"
]

# --- 核心函式：解析 Word 檔 (包含表格與段落) ---
def parse_ad_ids_from_docx(uploaded_file):
    """全面解析 Word 文件，包含「一般段落」與「表格內容」"""
    ad_ids = []
    try:
        doc = Document(uploaded_file)
        all_paragraphs = []
        all_paragraphs.extend(doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    all_paragraphs.extend(cell.paragraphs)
        
        for i, para in enumerate(all_paragraphs):
            text = para.text.strip()
            if "【廣告組合 ID】" in text:
                parts = text.split("【廣告組合 ID】")
                candidate = ""
                if len(parts) > 1:
                    candidate = parts[1].strip().lstrip(",").lstrip(":").lstrip("，").lstrip("：").strip()
                
                if candidate:
                    ad_ids.append(candidate)
                else:
                    if i + 1 < len(all_paragraphs):
                        next_text = all_paragraphs[i + 1].text.strip()
                        if next_text:
                            clean_next = next_text.lstrip(",").lstrip(":").lstrip("，").lstrip("：").strip()
                            if clean_next:
                                ad_ids.append(clean_next)
    except Exception as e:
        st.error(f"解析 Word 檔時發生錯誤: {e}")
    return sorted(list(set(ad_ids)))

# --- 輔助函式：解析 CSV/Excel ---
def parse_tags_from_csv(uploaded_file):
    tags = []
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        if "受眾標籤 (Tag)" in df.columns:
            tags = df["受眾標籤 (Tag)"].dropna().unique().tolist()
    except Exception as e:
        st.error(f"解析受眾標籤檔時發生錯誤: {e}")
    return tags

# --- 初始化 Session State ---
# 結構改為：campaigns = [ { settings..., ad_sets: [ ... ] }, ... ]
if 'campaigns' not in st.session_state:
    st.session_state.campaigns = []

# --- 側邊欄 ---
with st.sidebar:
    st.header("📂 資料來源庫")
    docx_file = st.file_uploader("上傳上刊文件 (Word)", type=['docx'])
    available_ad_ids = []
    if docx_file:
        available_ad_ids = parse_ad_ids_from_docx(docx_file)
        if available_ad_ids:
            st.success(f"已抓取 {len(available_ad_ids)} 個廣告 ID")
            with st.expander("查看 ID 列表"):
                st.write(available_ad_ids)
    
    tag_file = st.file_uploader("上傳受眾標籤 (CSV/Excel)", type=['csv', 'xlsx'])
    available_tags = DEFAULT_TAGS.copy()
    if tag_file:
        uploaded_tags = parse_tags_from_csv(tag_file)
        if uploaded_tags:
            available_tags = uploaded_tags
            st.success(f"已載入 {len(available_tags)} 個標籤")

# --- 主畫面 ---
st.title("🚀 Meta 廣告全策略配置工具")
st.info("💡 支援 **多行銷活動 (Multi-Campaign)** 架構。您可以新增多個活動，並在每個活動下獨立管理受眾。")

# --- 全局控制 ---
if st.button("➕ 建立一個新的行銷活動 (Campaign)", type="primary"):
    st.session_state.campaigns.append({
        "name": "",
        "objective": "銷售 (Sales)",
        "is_cbo": True,
        "budget_type": "單日預算",
        "budget_amount": 1000,
        "start_date": datetime.date.today(),
        "ad_sets": [] # 每個 Campaign 獨立的 Ad Sets 列表
    })

st.markdown("---")

# --- 迴圈渲染每個 Campaign ---
campaigns_to_remove = []

for c_idx, campaign in enumerate(st.session_state.campaigns):
    # 使用 Container 區隔每個 Campaign
    with st.container():
        st.markdown(f"## 📢 行銷活動 #{c_idx + 1}")
        
        # Campaign 設定區塊
        with st.expander(f"設定活動內容：{campaign['name'] if campaign['name'] else '(未命名)'}", expanded=True):
            # 刪除 Campaign 按鈕
            col_del_camp, _ = st.columns([1, 6])
            if col_del_camp.button(f"🗑️ 刪除整個活動 #{c_idx + 1}", key=f"del_cmp_{c_idx}"):
                campaigns_to_remove.append(c_idx)

            col1, col2 = st.columns(2)
            with col1:
                campaign['name'] = st.text_input("行銷活動名稱", value=campaign['name'], placeholder="例如：2025_TOF_流量開發", key=f"cmp_name_{c_idx}")
                campaign['objective'] = st.selectbox("行銷活動目標", ["銷售 (Sales)", "潛在客戶 (Leads)", "流量 (Traffic)", "互動 (Engagement)", "知名度 (Awareness)", "應用程式推廣 (App Promotion)"], key=f"cmp_obj_{c_idx}")
                campaign['is_cbo'] = st.toggle("開啟 CBO 預算最佳化", value=campaign['is_cbo'], key=f"cmp_cbo_{c_idx}")
            with col2:
                campaign['budget_type'] = st.selectbox("預算類型", ["單日預算", "總預算"], key=f"cmp_btype_{c_idx}")
                campaign['budget_amount'] = st.number_input("預算金額 (TWD)", min_value=0, value=campaign['budget_amount'], step=100, key=f"cmp_bamt_{c_idx}")
                campaign['start_date'] = st.date_input("開始日期", value=campaign['start_date'], key=f"cmp_date_{c_idx}")

        # --- 該 Campaign 下的 Ad Sets 管理 ---
        st.markdown(f"**👇 設定活動 #{c_idx + 1} 的廣告組合 (受眾)**")
        
        if st.button(f"➕ 新增廣告組合 (至活動 #{c_idx + 1})", key=f"add_adset_{c_idx}"):
            campaign['ad_sets'].append({
                "name": "", "goal": "網站", "tags": [], "manual_tags": [],
                "custom_audience": "", "age_min": 25, "age_max": 55,
                "locations": ["台灣"], "gender": "所有性別",
                "advantage_audience": True, "placements_type": "Advantage+ (自動版位)",
                "excluded_placements": [], "placement_note": "", "selected_ads": [], "manual_ads": []
            })

        adsets_to_remove = []
        # 顯示該 Campaign 下的所有 Ad Sets
        for a_idx, ad_set in enumerate(campaign['ad_sets']):
            # 使用嵌套的 Expander 或不同背景色塊
            st.info(f"廣告組合 {c_idx+1}-{a_idx+1}: {ad_set['name'] if ad_set['name'] else '設定中...'}")
            
            # Ad Set 內容
            c1, c2 = st.columns(2)
            
            with c1:
                ad_set['name'] = st.text_input("組合名稱", value=ad_set['name'], key=f"c{c_idx}_a{a_idx}_name")
                ad_set['goal'] = st.selectbox("轉換位置", ["網站", "應用程式", "訊息", "通話"], key=f"c{c_idx}_a{a_idx}_goal")
                
                # 受眾
                ad_set['tags'] = st.multiselect("興趣標籤", options=available_tags, default=list(set(ad_set['tags']) & set(available_tags)), key=f"c{c_idx}_a{a_idx}_tags")
                extra = st.text_input("手動輸入標籤", key=f"c{c_idx}_a{a_idx}_extra")
                ad_set['manual_tags'] = [t.strip() for t in extra.split(",") if t.strip()]
                
                ad_set['advantage_audience'] = st.checkbox("開啟 Advantage+ 受眾", value=ad_set['advantage_audience'], key=f"c{c_idx}_a{a_idx}_aa")

            with c2:
                # 刪除 Ad Set 按鈕
                if st.button("🗑️", key=f"del_adset_c{c_idx}_a{a_idx}", help="刪除此廣告組合"):
                    adsets_to_remove.append(a_idx)
                
                ad_set['custom_audience'] = st.text_area("自訂受眾", value=ad_set['custom_audience'], height=68, key=f"c{c_idx}_a{a_idx}_ca")
                
                col_age1, col_age2 = st.columns(2)
                ad_set['age_min'] = col_age1.number_input("最小年齡", 13, 65, ad_set['age_min'], key=f"c{c_idx}_a{a_idx}_amin")
                ad_set['age_max'] = col_age2.number_input("最大年齡", 13, 65, ad_set['age_max'], key=f"c{c_idx}_a{a_idx}_amax")
                ad_set['gender'] = st.radio("性別", ["所有性別", "男性", "女性"], horizontal=True, key=f"c{c_idx}_a{a_idx}_gen")
                
                # 素材選擇 (關鍵功能)
                if available_ad_ids:
                    ad_set['selected_ads'] = st.multiselect("選擇素材 ID", options=available_ad_ids, default=list(set(ad_set['selected_ads']) & set(available_ad_ids)), key=f"c{c_idx}_a{a_idx}_ads")
                else:
                    man_ads = st.text_area("手動輸入素材 ID", key=f"c{c_idx}_a{a_idx}_mads")
                    ad_set['manual_ads'] = [x.strip() for x in man_ads.split("\n") if x.strip()]

            # 版位設定 (選填)
            with st.expander("版位設定 (進階)", expanded=False):
                ad_set['placements_type'] = st.radio("版位模式", ["Advantage+ (自動)", "手動指定"], key=f"c{c_idx}_a{a_idx}_pmode")
                if ad_set['placements_type'] == "手動指定":
                    ad_set['excluded_placements'] = st.multiselect("排除/指定版位", ["FB動態", "IG動態", "Stories", "Reels"], key=f"c{c_idx}_a{a_idx}_excl")
        
        # 執行 Ad Set 刪除
        for i in sorted(adsets_to_remove, reverse=True):
            del campaign['ad_sets'][i]
            
    st.markdown("---")

# 執行 Campaign 刪除
for i in sorted(campaigns_to_remove, reverse=True):
    del st.session_state.campaigns[i]

# --- 輸出報告 ---
if st.button("📝 生成全策略 Brief 文件", type="primary"):
    report = f"# Meta 廣告全策略 Brief\n"
    report += f"**生成時間**: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
    report += f"**總計**: {len(st.session_state.campaigns)} 個行銷活動\n\n"
    report += "---"
    
    if not st.session_state.campaigns:
        report += "\n\n(尚未建立任何行銷活動)"

    for c_idx, campaign in enumerate(st.session_state.campaigns):
        report += f"\n\n## 📢 行銷活動 {c_idx+1}: {campaign['name']}\n"
        report += f"> **目標**: {campaign['objective']} | **預算**: {campaign['budget_type']} ${campaign['budget_amount']} | **CBO**: {'✅' if campaign['is_cbo'] else '❌'}\n\n"
        
        if not campaign['ad_sets']:
            report += "   *(此活動下尚未設定廣告組合)*\n"

        for a_idx, ad_set in enumerate(campaign['ad_sets']):
            report += f"### 🔹 組合 {c_idx+1}-{a_idx+1}: {ad_set['name']}\n"
            
            # 受眾區塊
            report += f"**【受眾設定】**\n"
            report += f"- **人口**: {ad_set['age_min']}-{ad_set['age_max']} 歲 / {ad_set['gender']} / {', '.join(ad_set['locations'])}\n"
            
            tags = ad_set['tags'] + ad_set['manual_tags']
            tags_str = ", ".join(tags) if tags else "無"
            report += f"- **興趣標籤**: {tags_str}\n"
            
            if ad_set['custom_audience']:
                report += f"- **自訂受眾**: {ad_set['custom_audience']}\n"
            report += f"- **Advantage+ 受眾**: {'開啟' if ad_set['advantage_audience'] else '關閉'}\n"

            # 版位區塊
            report += f"**【版位設定】**\n"
            report += f"- {ad_set['placements_type']}"
            if ad_set['excluded_placements']:
                report += f" (指定: {', '.join(ad_set['excluded_placements'])})"
            report += "\n"

            # 素材區塊
            report += f"**【素材配置】**\n"
            final_ads = ad_set['selected_ads'] + ad_set['manual_ads']
            if final_ads:
                for ad_id in final_ads:
                    report += f"- ` {ad_id} `\n"
            else:
                report += "- (尚未指定素材)\n"
            
            report += "\n"
        report += "---\n"

    st.text_area("Markdown 輸出預覽", value=report, height=600)
    st.download_button("下載完整 Brief (.txt)", data=report, file_name=f"Meta_Full_Strategy_{datetime.date.today()}.txt")
