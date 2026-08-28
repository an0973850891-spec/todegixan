import streamlit as st
import pandas as pd
import os
from datetime import datetime

# 設定網頁標題與版面
st.set_page_config(page_title="土地增值稅即時試算系統", page_icon="🏡", layout="wide")

# --- 🎨 自定義 CSS 美化與「列印專屬過濾」 ---
st.markdown("""
    <style>
    h1 {
        color: #1E3A8A;
        font-weight: 700;
    }
    h2, h3 {
        color: #C2410C;
    }
    label, div[data-baseweb="select"] span {
        font-size: 17px !important;
        font-weight: 600 !important;
        color: #1F2937 !important;
    }
    .stSelectbox div[data-baseweb="select"] *, .stTextInput input, .stNumberInput input {
        font-size: 16px !important;
    }
    .stForm {
        background-color: #FFF7ED !important;
        padding: 25px;
        border-radius: 12px;
        border: 1px solid #FFEDD5;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    /* 🖨️ 核心列印過濾規則：列印時只保留計算總結與明細區塊，隱藏所有輸入與操作按鈕 */
    @media print {
        header, footer, nav, 
        .stForm, 
        button, 
        .streamlit-expanderHeader,
        div[data-testid="stSidebar"] {
            display: none !important;
        }
        body {
            background-color: white !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

st.title("🏡 土地增值稅多筆土地累加與合併試算系統")
st.markdown("##### 請依序選擇縣市與鄉鎮市區（自動連動），輸入地段地號與明細後點擊「加入清單」，即可進行多筆累加計算！")

# --- 1. 自動取得當前即時年月 ---
now = datetime.now()
current_roc_year = now.year - 1911
current_month = now.month

# --- 2. 智慧判斷路徑：本地用 D:\PY\，雲端用相對路徑 ---
if os.path.exists(r"D:\PY\cpispleym.xls"):
    excel_path = r"D:\PY\cpispleym.xls"
else:
    excel_path = "cpispleym.xls"

@st.cache_data
def load_local_cpi_excel(path):
    if not os.path.exists(path):
        alt_path = path + "x"
        if os.path.exists(alt_path):
            path = alt_path
        else:
            return None, f"找不到檔案！系統無法在指定路徑找到檔案：{path}。請確認 cpispleym.xls 是否已上傳至 GitHub 或放在 D:\\PY 資料夾中。"
    try:
        df = pd.read_excel(path, header=None)
        return df, None
    except Exception as e:
        return None, f"讀取 Excel 發生錯誤: {e}"

df_cpi, error_msg = load_local_cpi_excel(excel_path)

def get_cpi_from_local_matrix(df, target_year, target_month):
    if df is None:
        return 100.0
    try:
        month_col_idx = None
        for col in range(df.shape[1]):
            cell_val = df.iloc[3, col]
            try:
                if int(cell_val) == target_month:
                    month_col_idx = col
                    break
            except:
                if str(cell_val).strip() in [str(target_month), f"{target_month}月", f"0{target_month}月"]:
                    month_col_idx = col
                    break
        if month_col_idx is None:
            month_col_idx = target_month
            
        target_row_idx = None
        for row in range(4, df.shape[0]):
            year_val = df.iloc[row, 0]
            try:
                if int(year_val) == target_year:
                    target_row_idx = row
                    break
            except:
                continue
                
        if target_row_idx is not None and month_col_idx is not None:
            raw_val = float(df.iloc[target_row_idx, month_col_idx])
            if raw_val > 0:
                return max(raw_val, 100.0)
    except:
        pass
    return 100.0

if error_msg:
    st.error(error_msg)

# --- 3. 初始化 Session State ---
if "calculated" not in st.session_state:
    st.session_state.calculated = False
if "results_data" not in st.session_state:
    st.session_state.results_data = None

if "land_list" not in st.session_state:
    st.session_state.land_list = []

county_district_map = {
    "基隆市": ["仁愛區", "信義區", "中正區", "中山區", "安樂區", "暖暖區", "七堵區"],
    "臺北市": ["中正區", "大同區", "中山區", "松山區", "大安區", "萬華區", "信義區", "士林區", "北投區", "內湖區", "南港區", "文山區"],
    "新北市": ["板橋區", "三重區", "中和區", "永和區", "新莊區", "新店區", "土城區", "蘆洲區", "汐止區", "樹林區", "淡水區", "鶯歌區", "三峽區", "瑞芳區", "五股區", "泰山區", "林口區", "深坑區", "石碇區", "坪林區", "三芝區", "石門區", "八里區", "平溪區", "雙溪區", "貢寮區", "金山區", "萬里區", "烏來區"],
    "桃園市": ["桃園區", "中壢區", "大溪區", "楊梅區", "蘆竹區", "大園區", "龜山區", "八德區", "龍潭區", "平鎮區", "新屋區", "觀音區", "復興區"],
    "新竹市": ["東區", "北區", "香山區"],
    "新竹縣": ["竹北市", "竹東鎮", "新埔鎮", "關西鎮", "湖口鄉", "新豐鄉", "峨眉鄉", "寶山鄉", "北埔鄉", "橫山鄉", "尖石鄉", "五峰鄉", "芎林鄉"],
    "苗栗縣": ["苗栗市", "頭份市", "竹南鎮", "後龍鎮", "通霄鎮", "苑裡鎮", "卓蘭鎮", "造橋鄉", "西湖鄉", "頭屋鄉", "公館鄉", "銅鑼鄉", "三義鄉", "大湖鄉", "獅潭鄉", "三灣鄉", "南庄鄉", "泰安鄉"],
    "臺中市": ["中區", "東區", "南區", "西區", "北區", "西屯區", "南屯區", "北屯區", "豐原區", "東勢區", "大甲區", "清水區", "沙鹿區", "梧棲區", "后里區", "神岡區", "大雅區", "潭子區", "新社區", "石岡區", "外埔區", "大安區", "烏日區", "大肚區", "龍井區", "霧峰區", "太平區", "大里區", "和平區"],
    "彰化縣": ["彰化市", "員林市", "和美鎮", "鹿港鎮", "溪湖鎮", "二林鎮", "北斗鎮", "田中鎮", "大城鄉", "芳苑鄉", "二水鄉", "田尾鄉", "埤頭鄉", "溪州鄉", "社頭鄉", "永靖鄉", "埔心鄉", "埔鹽鄉", "大村鄉", "芬園鄉", "花壇鄉", "秀水鄉"],
    "南投縣": ["南投市", "埔里鎮", "草屯鎮", "竹山鎮", "集集鎮", "名間鄉", "鹿谷鄉", "中寮鄉", "魚池鄉", "國姓鄉", "水里鄉", "信義鄉", "仁愛鄉"],
    "雲林縣": ["斗六市", "斗南鎮", "虎尾鎮", "西螺鎮", "土庫鎮", "北港鎮", "古坑鄉", "大埤鄉", "莿桐鄉", "林內鄉", "褒忠鄉", "臺西鄉", "崙背鄉", "麥寮鄉", "二崙鄉", "水林鄉", "口湖鄉", "四湖鄉", "元長鄉", "東勢鄉"],
    "嘉義市": ["東區", "西區"],
    "嘉義縣": ["太保市", "朴子市", "布袋鎮", "大林鎮", "民雄鄉", "溪口鄉", "新港鄉", "六腳鄉", "東石鄉", "義竹鄉", "鹿草鄉", "水上鄉", "中埔鄉", "竹崎鄉", "梅山鄉", "番路鄉", "大埔鄉", "阿里山鄉"],
    "臺南市": ["永康區", "歸仁區", "新化區", "左鎮區", "玉井區", "楠西區", "南化區", "仁德區", "關廟區", "龍崎區", "官田區", "麻豆區", "佳里區", "西港區", "七股區", "將軍區", "學甲區", "北門區", "新營區", "後壁區", "白河區", "東山區", "六甲區", "下營區", "柳營區", "鹽水區", "善化區", "大內區", "安定區", "山上區", "新市區", "中西區", "東區", "南區", "北區", "安平區", "安南區"],
    "高雄市": ["楠梓區", "左營區", "鼓山區", "三民區", "鹽埕區", "前金區", "新興區", "苓雅區", "前鎮區", "旗津區", "小港區", "鳳山區", "林園區", "大寮區", "大樹區", "大社區", "仁武區", "鳥松區", "岡山區", "橋頭區", "燕巢區", "田寮區", "阿蓮區", "路竹區", "湖內區", "茄安區", "永安區", "彌陀區", "梓官區", "旗山區", "美濃區", "六龜區", "甲仙區", "杉林區", "內門區", "茂林區", "桃源區", "那瑪夏區"],
    "屏東縣": ["屏東市", "潮州鎮", "東港鎮", "恆春鎮", "萬丹鄉", "長治鄉", "麟洛鄉", "九如鄉", "里港鄉", "鹽埔鄉", "高樹鄉", "萬巒鄉", "內埔鄉", "竹田鄉", "新埤鄉", "枋寮鄉", "新園鄉", "崁頂鄉", "林邊鄉", "南州鄉", "佳冬鄉", "琉球鄉", "車城鄉", "滿州鄉", "枋山鄉", "三地門鄉", "霧臺鄉", "瑪家鄉", "泰武鄉", "來義鄉", "春日鄉", "獅子鄉", "牡丹鄉"],
    "宜蘭縣": ["宜蘭市", "羅東鎮", "蘇澳鎮", "頭城鎮", "礁溪鄉", "壯圍鄉", "員山鄉", "冬山鄉", "五結鄉", "三星鄉", "大同鄉", "南澳鄉"],
    "花蓮縣": ["花蓮市", "鳳林鎮", "玉里鎮", "新城鄉", "吉安鄉", "壽豐鄉", "光復鄉", "豐濱鄉", "瑞穗鄉", "富里鄉", "秀林鄉", "萬榮鄉", "卓溪鄉"],
    "臺東縣": ["臺東市", "成功鎮", "關山鎮", "長濱鄉", "池上鄉", "東河鄉"],
    "澎湖縣": ["馬公市", "湖西鄉", "白沙鄉", "西嶼鄉", "望安鄉", "七美鄉"],
    "金門縣": ["金城鎮", "金湖鎮", "金沙鎮", "金寧鄉", "烈嶼鄉", "烏坵鄉"],
    "連江縣": ["南竿鄉", "北竿鄉", "莒光鄉", "東引鄉"]
}

county_options = list(county_district_map.keys())

# --- 4. 獨立在 Form 外面的縣市選擇 ---
st.markdown("### 📌 新增土地資料")

col_c1, col_c2 = st.columns(2)
with col_c1:
    selected_county = st.selectbox("縣市", options=county_options, key="selected_county_widget")

district_options = county_district_map.get(selected_county, ["其他"])

with col_c2:
    selected_district = st.selectbox("鄉鎮市區", options=district_options, key="selected_district_widget")

# --- 5. 表單內部：地段、地號與其他明細 ---
with st.form("add_land_form", clear_on_submit=False):
    col3, col4 = st.columns(2)
    with col3:
        section_name = st.text_input("地段名稱", placeholder="例如：茄安段")
    with col4:
        land_number = st.text_input("地號", placeholder="例如：497")

    st.markdown("---")
    
    col5, col6, col7, col8 = st.columns(4)
    with col5:
        year_options = list(range(48, current_roc_year + 1))
        last_year = st.selectbox("前次移轉年份 (民國)", options=year_options, index=len(year_options)-15 if len(year_options)>=15 else 0)
    with col6:
        last_month = st.selectbox("前次移轉月份", options=list(range(1, 13)), index=7)
    with col7:
        curr_val_str = st.text_input("公告現值 (元/m²)", value="0.0")
    with col8:
        last_val_str = st.text_input("原地價 (元/m²)", value="0.0")

    col9, col10, col11, col12 = st.columns(4)
    with col9:
        area_str = st.text_input("面積 (m²)", value="0.0")
    with col10:
        num_str = st.text_input("持分分子", value="1.0")
    with col11:
        den_str = st.text_input("持分分母", value="1.0")
    with col12:
        offset_str = st.text_input("抵繳地價稅額 (元)", value="0.0")

    add_submitted = st.form_submit_button("➕ 加入清單", use_container_width=True)

if add_submitted:
    try:
        curr_val = float(curr_val_str)
        last_val = float(last_val_str)
        area = float(area_str)
        num = float(num_str)
        den = float(den_str)
        offset = float(offset_str)
    except ValueError:
        curr_val, area = 0.0, 0.0

    if curr_val == 0 or area == 0:
        st.warning("⚠️ 公告現值與面積不能為 0 或空白！")
    else:
        new_land = {
            "縣市": selected_county,
            "鄉鎮市區": selected_district,
            "地段名稱": section_name.strip(),
            "地號": land_number.strip(),
            "前次年(民國)": last_year,
            "前次月": last_month,
            "公告現值(元/m²)": curr_val,
            "原地價(元/m²)": last_val,
            "面積(m²)": area,
            "持分分子": num,
            "持分分母": den,
            "抵繳地價稅(元)": offset
        }
        st.session_state.land_list.append(new_land)
        st.success(f"✅ 已成功加入：{selected_county}{selected_district} {section_name}{land_number}號")

# --- 6. 顯示目前已加入的土地清單與操作按鈕 ---
if st.session_state.land_list:
    st.markdown("---")
    st.subheader("📋 目前已加入的土地清單")
    df_current_lands = pd.DataFrame(st.session_state.land_list)
    st.dataframe(df_current_lands, use_container_width=True)

    col_btn1, col_btn2, _ = st.columns([1, 1, 4])
    with col_btn1:
        calc_pressed = st.button("🔥 開始計算與合併", use_container_width=True, type="primary")
    with col_btn2:
        clear_pressed = st.button("🧹 清除全部資料", use_container_width=True)

    if clear_pressed:
        st.session_state.land_list = []
        st.session_state.calculated = False
        st.session_state.results_data = None
        st.rerun()

    if calc_pressed:
        raw_results = []
        for idx, row in df_current_lands.iterrows():
            county = row["縣市"]
            district = row["鄉鎮市區"]
            section = row["地段名稱"]
            land_no = row["地號"]
            
            if not section and not land_no:
                land_name = f"{county}{district} 土地 {idx+1}"
            else:
                land_name = f"{county}{district} {section} {land_no}號".strip()

            last_y = int(row["前次年(民國)"])
            last_m = int(row["前次月"])
            curr_y = current_roc_year
            
            curr_val_sqm = float(row["公告現值(元/m²)"])
            last_val_sqm = float(row["原地價(元/m²)"])
            area_val = float(row["面積(m²)"])
            num_val = float(row["持分分子"])
            den_val = float(row["持分分母"])
            offset_val = float(row["抵繳地價稅(元)"])

            row_cpi = get_cpi_from_local_matrix(df_cpi, last_y, last_m)
            start_months = last_y * 12 + last_m
            end_months = curr_y * 12 + current_month
            holding_years = max(0, (end_months - start_months) // 12)

            share_ratio = (num_val / den_val) if den_val > 0 else 1.0
            land_curr_total = curr_val_sqm * area_val * share_ratio
            land_last_total = last_val_sqm * (row_cpi / 100.0) * area_val * share_ratio

            appreciation = land_curr_total - land_last_total
            if appreciation < 0:
                appreciation = 0

            b = land_last_total
            a = appreciation
            ratio = (a / b) if b > 0 else 0

            # 自用住宅稅額 (10%)
            self_cha = a * 0.10
            self_tax = max(0, self_cha - offset_val)

            # 一般用地稅額 (三級累進與長期持有減徵)
            gen_cha = 0
            if holding_years > 40:
                if ratio <= 1: gen_cha = a * 0.20
                elif ratio <= 2: gen_cha = (a * 0.26) - (b * 0.06)
                else: gen_cha = (a * 0.32) - (b * 0.18)
            elif holding_years > 30:
                if ratio <= 1: gen_cha = a * 0.20
                elif ratio <= 2: gen_cha = (a * 0.27) - (b * 0.07)
                else: gen_cha = (a * 0.34) - (b * 0.21)
            elif holding_years > 20:
                if ratio <= 1: gen_cha = a * 0.20
                elif ratio <= 2: gen_cha = (a * 0.28) - (b * 0.08)
                else: gen_cha = (a * 0.36) - (b * 0.24)
            else:
                if ratio <= 1: gen_cha = a * 0.20
                elif ratio <= 2: gen_cha = (a * 0.30) - (b * 0.10)
                else: gen_cha = (a * 0.40) - (b * 0.30)

            gen_tax = max(0, gen_cha - offset_val)

            raw_results.append({
                "土地名稱 / 地號": land_name,
                "持有年限": holding_years,
                "套用CPI": row_cpi,
                "申報現值總額": land_curr_total,
                "漲價總數額(a)": a,
                "一般用地稅額": gen_tax,
                "自用住宅稅額": self_tax
            })

        df_res = pd.DataFrame(raw_results)
        df_grouped = df_res.groupby("土地名稱 / 地號").agg({
            "申報現值總額": "sum",
            "漲價總數額(a)": "sum",
            "一般用地稅額": "sum",
            "自用住宅稅額": "sum",
            "持有年限": "max",
            "套用CPI": "max"
        }).reset_index()

        st.session_state.calculated = True
        st.session_state.results_data = df_grouped

# --- 7. 呈現累加與合併計算結果總表與列印匯出功能 ---
if st.session_state.calculated and st.session_state.results_data is not None:
    df_grouped = st.session_state.results_data

    total_gen_tax = df_grouped["一般用地稅額"].sum()
    total_self_tax = df_grouped["自用住宅稅額"].sum()
    total_appreciation = df_grouped["漲價總數額(a)"].sum()
    total_current_val = df_grouped["申報現值總額"].sum()

    st.markdown("---")
    
    head_col1, head_col2 = st.columns([3, 1])
    with head_col1:
        st.subheader("📊 多筆土地試算總結與同地號合併明細")
    with head_col2:
        csv_data = df_grouped.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 下載明細報表 (CSV)",
            data=csv_data,
            file_name=f"土地增值稅試算報表_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    sum_col1, sum_col2, sum_col3 = st.columns(3)
    with sum_col1:
        st.metric(label="所有土地申報現值總額", value=f"{round(total_current_val):,} 元")
    with sum_col2:
        st.metric(label="所有土地漲價總數額合計", value=f"{round(total_appreciation):,} 元")
    with sum_col3:
        st.metric(label="合併後土地筆數", value=f"共計 {len(df_grouped)} 筆")

    st.markdown("---")
    
    res_col1, res_col2 = st.columns(2)
    with res_col1:
        st.markdown("### 🏢 全部土地 - 一般用地總稅額")
        st.markdown(f"#### 總應納稅額：**`{round(total_gen_tax):,} 元`**")
    with res_col2:
        st.markdown("### 🏡 全部土地 - 自用住宅總稅額")
        st.markdown(f"#### 總應納稅額：**`{round(total_self_tax):,} 元`**")

    st.markdown("---")
    
    st.subheader("📍 各筆土地（同地號已自動合併）明細列表")
    for _, row in df_grouped.iterrows():
        with st.container():
            st.markdown(f"#### 🔹 土地名稱/地號：**{row['土地名稱 / 地號']}** (已自動合併同地號)")
            card_col1, card_col2, card_col3, card_col4 = st.columns(4)
            with card_col1:
                st.caption("⏳ 持有年限 / CPI")
                st.write(f"**{int(row['持有年限'])} 年** (CPI: {row['套用CPI']}%)")
            with card_col2:
                st.caption("💰 合併後漲價總數額")
                st.write(f"**{round(row['漲價總數額(a)']):,} 元**")
            with card_col3:
                st.caption("🏢 一般用地稅額")
                st.write(f"**{round(row['一般用地稅額']):,} 元**")
            with card_col4:
                st.caption("🏡 自用住宅稅額")
                st.write(f"**{round(row['自用住宅稅額']):,} 元**")
            st.markdown("<hr style='margin: 10px 0; border-top: 1px dashed #ccc;'>", unsafe_allow_html=True)

    with st.expander("📝 點此展開完整合併後數據對照表"):
        df_display = df_grouped.copy()
        df_display["申報現值總額"] = df_display["申報現值總額"].round()
        df_display["漲價總數額(a)"] = df_display["漲價總數額(a)"].round()
        df_display["一般用地稅額"] = df_display["一般用地稅額"].round()
        df_display["自用住宅稅額"] = df_display["自用住宅稅額"].round()
        st.dataframe(df_display, use_container_width=True)

    diff_total = total_gen_tax - total_self_tax
    if diff_total > 0:
        st.success(f"💡 **整體節稅提示：** 若全部土地皆符合自用住宅資格，採用自用住宅總計可比一般用地**節省約 {round(diff_total):,} 元**！")

    # --- 8. 列印輸出指引 ---
    st.markdown("---")
    st.info("💡 **列印說明：** 當您按下鍵盤快捷鍵 **`Ctrl + P`**（Mac 為 `Cmd + P`）時，系統會**自動隱藏所有輸入與操作按鈕**，只保留上方的加總數據與下方的土地明細列表供您列印或儲存為 PDF！")
