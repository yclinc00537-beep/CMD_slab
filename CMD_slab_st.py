import streamlit as st
import os
import math
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ==========================================
# 1. 法規材質資料庫 (原始單位: kg/cm2)
# ==========================================
MATERIAL_DB = {
    "木材-A級 (柳安等)": {"E": 110000, "Fb": 135, "Fv": 12},
    "木材-B級 (松、杉)": {"E": 90000, "Fb": 110, "Fv": 9},
    "木材-C級 (常用模板)": {"E": 70000, "Fb": 90, "Fv": 7},
    "鋼材-SS400 (一般鋼)": {"E": 2100000, "Fb": 1470, "Fv": 960},
    "鋼材-SS490 (高強鋼)": {"E": 2100000, "Fb": 1770, "Fv": 1150},
    "鋁合金 (支撐樑)": {"E": 700000, "Fb": 1050, "Fv": 600}
}

# ==========================================
# 2. 核心力學計算函數
# ==========================================
def beam_check(w, span, b_cm, h_cm, mat, member_type):
    b_m, h_m = b_cm / 100.0, h_cm / 100.0
    E_kpa = mat["E"] * 98.0665
    Fb_long_kpa = mat["Fb"] * 98.0665
    Fb_allow_kpa = Fb_long_kpa * 1.33  # 短期施工修正
    
    Z = b_m * (h_m ** 2) / 6.0
    I = b_m * (h_m ** 3) / 12.0
    
    if member_type == "sheet":
        M = w * (span ** 2) / 10.0
        delta_m = (w * (span ** 4)) / (128.0 * E_kpa * I)
    else:
        M = w * (span ** 2) / 10.0
        delta_m = (w * (span ** 4)) / (150.0 * E_kpa * I)
        
    stress_kpa = M / Z
    delta_mm = delta_m * 1000.0
    limit_mm = (span / 360.0) * 1000.0
    
    return {
        "w": w, "M": M, "stress": stress_kpa, "delta": delta_mm, "limit": limit_mm,
        "Fb_allow": Fb_allow_kpa,
        "ratio_bending": stress_kpa / Fb_allow_kpa,
        "ratio_deflection": delta_mm / limit_mm if limit_mm > 0 else 0
    }

def run_single_logic(data, slab_t):
    gamma_kn = data["load"]["concrete_unit_weight_kn"]
    live_kn = data["load"]["live_load_kn"]
    dead_kn = (slab_t * gamma_kn) + 0.5
    q_kn = 1.3 * dead_kn + live_kn
    
    res = {"q": q_kn, "dead": dead_kn, "slab_t": slab_t}
    for p in ["sheet", "joist", "waler"]:
        m_type = "sheet" if p == "sheet" else "beam"
        mat = MATERIAL_DB.get(data[p]["material"])
        w_line = q_kn * data[p]["spacing"]
        res[p] = beam_check(w_line, data[p]["span"], data[p]["b"], data[p]["h"], mat, m_type)
        
    P_post = q_kn * data["post"]["sx"] * data["post"]["sy"]
    res["post"] = {
        "P": P_post,
        "ratio": P_post / data["post"]["capacity_kn"]
    }
    return res

def set_cell_background(cell, fill_hex):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

# ==========================================
# 3. Streamlit 網頁 UI 介面設計
# ==========================================
st.set_page_config(page_title="樓板模板結構應力計算與優化系統", layout="wide")

# 標題區
st.markdown("<h1 style='text-align: center; color: #1a5276;'>🧱 樓板模板結構應力計算與優化系統</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #566573;'>Streamlit 商業網頁平臺版 / 多材與多階板厚深度智慧檢核</p>", unsafe_allow_html=True)
st.write("---")

# 建立左右兩大區塊（左側輸入欄 vs. 右側看板結果）
col_input, col_result = st.columns([1, 1])

with col_input:
    st.subheader("⚙️ 二、工程環境參數")
    h_input = st.text_input("澆置設計總高度清單 (H) [m]", "3.5, 4.2")
    t_input = st.text_input("自訂檢核板厚位置 (以半形逗號隔開) [m]", "0.15, 0.20")
    
    concrete_w = st.number_input("混凝土單位重 (yc) [kN/m³]", value=24.0, step=0.1)
    live_l = st.number_input("施工活荷載 [kN/m²]", value=2.5, step=0.1)
    
    try:
        hs = [float(x.strip()) for x in h_input.split(",")]
        ts = [float(x.strip()) for x in t_input.split(",")]
    except Exception:
        st.error("⚠️ 請檢查高度或板厚輸入格式（必須以半形逗號隔開數字）")
        hs, ts = [], []

    st.subheader("📐 三、模板結構間距與跨距配置")
    
    with st.expander("面板參數 (Sheet)", expanded=True):
        sheet_mat = st.selectbox("面板材料選擇", list(MATERIAL_DB.keys()), index=2)
        sheet_space = st.number_input("面板設計跨距 (Panel Span) [m]", value=0.30, step=0.05, key="s_space")
        sheet_span = st.number_input("面板支撐跨距 [m]", value=0.30, step=0.05, key="s_span")
        sheet_b = st.number_input("面板寬度 b [cm]", value=30.0, step=1.0)
        sheet_h = st.number_input("面板厚度 h [cm]", value=1.8, step=0.1)
    
    with st.expander("閣柵參數 (Joist)", expanded=True):
        joist_mat = st.selectbox("閣柵材料選擇", list(MATERIAL_DB.keys()), index=2)
        joist_space = st.number_input("Joist (立桿) 佈置間距 [m]", value=0.30, step=0.05)
        joist_span = st.number_input("Joist (立桿) 支撐跨距 [m]", value=0.60, step=0.05)
        joist_b = st.number_input("閣柵寬度 b [cm]", value=5.0, step=0.5)
        joist_h = st.number_input("閣柵高度 h [cm]", value=10.0, step=0.5)
    
    with st.expander("橫楞參數 (Waler)", expanded=True):
        waler_mat = st.selectbox("橫楞材料選擇", list(MATERIAL_DB.keys()), index=3)
        waler_space = st.number_input("橫楞佈置間距 [m]", value=0.60, step=0.05)
        waler_span = st.number_input("橫楞支撐跨距 [m]", value=0.90, step=0.05)
        waler_b = st.number_input("橫楞寬度 b [cm]", value=10.0, step=0.5)
        waler_h = st.number_input("橫楞高度 h [cm]", value=10.0, step=0.5)
    
    with st.expander("支撐鋼管參數 (Post)", expanded=True):
        post_sx = st.number_input("縱向間距 Sx [m]", value=0.90, step=0.05)
        post_sy = st.number_input("橫向間距 Sy [m]", value=0.90, step=0.05)
        post_capacity = st.number_input("單支容許承載力 [kN]", value=10.0, step=1.0) # 👈 預設 10.0 KN 完美過關

    # 打包數據
    data = {
        "load": {"concrete_unit_weight_kn": concrete_w, "live_load_kn": live_l},
        "sheet": {"material": sheet_mat, "spacing": sheet_space, "span": sheet_span, "b": sheet_b, "h": sheet_h},
        "joist": {"material": joist_mat, "spacing": joist_space, "span": joist_span, "b": joist_b, "h": joist_h},
        "waler": {"material": waler_mat, "spacing": waler_space, "span": waler_span, "b": waler_b, "h": waler_h},
        "post": {"sx": post_sx, "sy": post_sy, "capacity_kn": post_capacity}
    }

with col_result:
    st.subheader("📊 各自訂板厚全構件應力安全判定看板")
    
    if len(hs) != len(ts) or len(hs) == 0:
        st.warning("⚠️ 請確認上方『高度清單』與『板厚清單』的數量必須完全一致。")
    else:
        # 計算所有階層的結果
        all_results = []
        max_ratio = 0.0
        
        for t in ts:
            res = run_single_logic(data, t)
            all_results.append(res)
            # 尋找全案最大控制應力利用率
            for k in ["sheet", "joist", "waler"]:
                max_ratio = max(max_ratio, res[k]["ratio_bending"], res[k]["ratio_deflection"])
            max_ratio = max(max_ratio, res["post"]["ratio"])
        
        # 儀表看板
        k_col1, k_col2 = st.columns(2)
        with k_col1:
            st.metric(label="全案最大控制利用率 (應力/撓度/軸力)", value=f"{max_ratio:.2f}")
        with k_col2:
            if max_ratio <= 1.0:
                st.markdown("<h2 style='color: #27ae60; margin-top: 15px;'>✅ 安全判定：通過 (OK)</h2>", unsafe_allow_html=True)
            else:
                st.markdown("<h2 style='color: #c0392b; margin-top: 15px;'>⚠️ 安全判定：不通過 (NG)</h2>", unsafe_allow_html=True)
                
        st.write("---")
        
        # 顯示各板厚的詳細數據卡片
        for i, t in enumerate(ts):
            res = all_results[i]
            with st.expander(f"📋 檢核斷面 - 板厚: {t:.2f}m (對應層高: {hs[i]}m)", expanded=True):
                st.write(f"**設計組合總荷載 q:** `{res['q']:.2f}` kN/m²")
                
                # 📢 使用完美校正後的單一 f""" 多行字串渲染 Markdown 表格
                st.markdown(f"""

| 構件名稱 | 抗彎應力利用率 | 撓度變形利用率 | 狀態 |
| :--- | :---: | :---: | :---: |
| **面板 (Sheet)** | {res['sheet']['ratio_bending']*100:.1f}% | {res['sheet']['ratio_deflection']*100:.1f}% | {'🟢 OK' if res['sheet']['ratio_bending']<=1 and res['sheet']['ratio_deflection']<=1 else '🔴 NG'} |
| **閣柵 (Joist)** | {res['joist']['ratio_bending']*100:.1f}% | {res['joist']['ratio_deflection']*100:.1f}% | {'🟢 OK' if res['joist']['ratio_bending']<=1 and res['joist']['ratio_deflection']<=1 else '🔴 NG'} |
| **橫楞 (Waler)** | {res['waler']['ratio_bending']*100:.1f}% | {res['waler']['ratio_deflection']*100:.1f}% | {'🟢 OK' if res['waler']['ratio_bending']<=1 and res['waler']['ratio_deflection']<=1 else '🔴 NG'} |
""")
                p_ratio = res['post']['ratio']
                st.write(f"**支撐鋼管軸力 P:** `{res['post']['P']:.2f}` kN / 容許值: `{data['post']['capacity_kn']:.2f}` kN (利用率: `{p_ratio*100:.1f}%` -> {'🟢 OK' if p_ratio<=1 else '🔴 NG'})")

        # Word 報告點擊一鍵直接儲存按鈕
        st.write("---")
        if st.button("📝 一鍵秒殺匯出專業 Word 結構計算書", use_container_width=True):
            try:
                doc = Document()
                style = doc.styles['Normal']
                style.font.name = 'Microsoft JhengHei'
                style.font.size = Pt(11)
                
                title = doc.add_heading('樓板模板及支撐結構詳細計算書', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                p_meta = doc.add_paragraph()
                p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_meta.add_run(f"產製日期：{datetime.now().strftime('%Y年%m月%d日')}\n規範依據：建築物結構設計規範 / 施工安全標準")

                for floor_h, slab_t, res in zip(hs, ts, all_results):
                    doc.add_heading(f'一、 樓層斷面控制 (層高: {floor_h}m, 板厚: {slab_t}m)', level=1)
                    
                    p_load = doc.add_paragraph()
                    p_load.add_run("1. 施工設計荷載計算：\n").bold = True
                    p_load.add_run(f"   - 混凝土自重：{data['load']['concrete_unit_weight_kn']:.2f} kN/m³ × {slab_t} m = {float(slab_t)*data['load']['concrete_unit_weight_kn']:.2f} kN/m²\n")
                    p_load.add_run(f"   - 模板與施工自重假設：0.50 kN/m²\n")
                    p_load.add_run(f"   - 施工活荷載：{data['load']['live_load_kn']:.2f} kN/m²\n")
                    p_load.add_run(f"   - 組合設計總荷載 q = 1.3 × 恆載 + 活載 = {res['q']:.2f} kN/m²").bold = True
                    
                    doc.add_heading('2. 結構構件力學檢核表', level=2)
                    
                    table = doc.add_table(rows=1, cols=5)
                    table.style = 'Light Shading Accent 1'
                    hdr_cells = table.rows[0].cells
                    headers = ['構件名稱', '設計應力 / 變形', '容許值 / 上限', '比值 (Ratio)', '判定結果']
                    
                    for idx, h_text in enumerate(headers):
                        hdr_cells[idx].text = h_text
                        set_cell_background(hdr_cells[idx], "1A5276")
                        hdr_cells[idx].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
                        hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                    
                    members_map = [("面板 (Sheet)", "sheet"), ("閣柵 (Joist)", "joist"), ("橫楞 (Waler)", "waler")]
                    for label, key in members_map:
                        r = res[key]
                        cur_mat_name = data[key]["material"]
                        size_txt = f"厚度 = {data[key]['h']} cm" if key == "sheet" else f"尺寸 = {data[key]['b']} x {data[key]['h']} cm"
                        
                        row_b = table.add_row().cells
                        row_b[0].text = f"{label}\n({cur_mat_name})\n抗彎強度"
                        row_b[1].text = f"{r['stress']:.1f} kPa"
                        row_b[2].text = f"{r['Fb_allow']:.1f} kPa"
                        row_b[3].text = f"{r['ratio_bending']*100:.1f}%"
                        
                        row_d = table.add_row().cells
                        row_d[0].text = f"{label}\n({cur_mat_name})\n撓度變形"
                        row_d[1].text = f"{r['delta']:.3f} mm"
                        row_d[2].text = f"{r['limit']:.2f} mm"
                        row_d[3].text = f"{r['ratio_deflection']*100:.1f}%"
                        
                        for row_cells, ratio in [(row_b, r['ratio_bending']), (row_d, r['ratio_deflection'])]:
                            if ratio <= 1.0:
                                p_run = row_cells[4].paragraphs[0].add_run("通過 (OK)")
                                p_run.font.color.rgb = RGBColor(39, 174, 96)
                            else:
                                p_run = row_cells[4].paragraphs[0].add_run("不通過 (NG)")
                                p_run.font.color.rgb = RGBColor(192, 57, 43)
                                for cell in row_cells:
                                    set_cell_background(cell, "FADBD8")
                            p_run.font.bold = True
                    
                    doc.add_heading('3. 支撐鋼管軸力檢核', level=2)
                    p_post = doc.add_paragraph()
                    p_post.add_run(f"   - 支撐配置範圍：{data['post']['sx']:.2f} m × {data['post']['sy']:.2f} m\n")
                    p_post.add_run(f"   - 支撐軸力設計值 P = {res['post']['P']:.2f} kN / 容許值 = {data['post']['capacity_kn']:.2f} kN\n")
                    
                    ratio_p = res['post']['ratio']
                    if ratio_p <= 1.0:
                        f_run = p_post.add_run(f"   - 判定結果：安全通過 (利用率 {ratio_p*100:.1f}%)")
                        f_run.font.color.rgb = RGBColor(39, 174, 96)
                    else:
                        f_run = p_post.add_run(f"   - 判定結果：超限不通過 (利用率 {ratio_p*100:.1f}%)")
                        f_run.font.color.rgb = RGBColor(192, 57, 43)
                    f_run.bold = True
                    doc.add_page_break()

                # 固定自動存檔路徑
                fixed_dir = r"D:\PythonTraining\slab_formwork_app\CMD_slab_app"
                if not os.path.exists(fixed_dir):
                    os.makedirs(fixed_dir)
                
                time_str = datetime.now().strftime('%Y%m%d_%H%M%S')
                file_name = f"樓板模板結構計算書_{time_str}.docx"
                save_path = os.path.join(fixed_dir, file_name)
                
                doc.save(save_path)
                st.success(f"🎉 專業結構計算書已成功自動儲存至：{save_path}")
            except Exception as e:
                st.error(f"❌ 匯出 Word 過程中發生錯誤：{str(e)}")