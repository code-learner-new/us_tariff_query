import streamlit as st
import sqlite3
import re
from openpyxl import Workbook
from io import BytesIO
import os
import sys

# ===================== 配置 =====================
DB = r"./tariff.db"
ORIGIN_COUNTRY_LIST = ["China","Vietnam","Malaysia","Mexico","India","Indonesia","Thailand"]

# 数据库不存在则自动运行init_db重建（解决Streamlit容器重启丢失db）
if not os.path.exists(DB):
    st.warning("数据库不存在，正在重建税则数据库，请稍候...")
    import init_db
    init_db.main()

# ===================== DB工具函数（原样复用） =====================
def get_db_conn():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def clean_input_hs(s):
    if not s:
        return ""
    return "".join([c for c in str(s) if c.isdigit()])

def get_8digit_tariff(conn, full_clean_hs):
    """逐级截断：10→8→6，优先最长匹配父节点general_rate"""
    candidates = []
    if len(full_clean_hs)>=10:
        candidates.append(full_clean_hs[:10])
    if len(full_clean_hs)>=8:
        candidates.append(full_clean_hs[:8])
    if len(full_clean_hs)>=6:
        candidates.append(full_clean_hs[:6])
    for seg in candidates:
        row = conn.execute('''
            SELECT general_rate FROM hts
            WHERE REPLACE(hs_code, '.', '') = ?
        ''', (seg,)).fetchone()
        if row and row["general_rate"] is not None and row["general_rate"].strip()!="":
            raw_rate = row["general_rate"].strip()
            rate_strip = raw_rate.upper()
            if rate_strip == "FREE":
                return "0.00%"
            return raw_rate
    return "N/A"

def get_section232_info(conn, product_hs_clean):
    rows_232 = conn.execute('''
        SELECT hs_code, add_tariff_rate, type FROM tariff_232
    ''').fetchall()
    for r in rows_232:
        list_hs_raw = r["hs_code"]
        list_hs_clean = clean_input_hs(list_hs_raw)
        if not list_hs_clean:
            continue
        if product_hs_clean.startswith(list_hs_clean):
            rate = r["add_tariff_rate"]
            typ = r["type"]
            return f"{rate} | {typ}"
    return "-"

def parse_rate_number(rate_str: str):
    if not rate_str or rate_str == "-" or rate_str == "N/A":
        return None
    s = str(rate_str).strip().upper()
    if "FREE" in s:                    # ← 去掉 or "0.00%" in s
        return 0.0
    match = re.search(r"(\d+(?:\.\d+)?)", s)   # ← 正则收紧，避免匹配到游离小数点
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None

def parse_forced_labor_rate(text, target_country, is_excl_flag, lang="zh"):
    if is_excl_flag == 1:
        if lang == "en":
            return "0 (Exempted)"
        else:
            return "0（豁免清单内）"
    if not text or text.strip() == "":
        if lang == "en":
            return "Not in list"
        else:
            return "不在清单"
    pattern = re.findall(r'"([\w]+)\s*,\s*([\d\.]+)"', text)
    rate_val = None
    for country,rate in pattern:
        if country.strip().upper() == target_country.strip().upper():
            rate_val = f"{rate}%"
            break
    if rate_val is None:
        if lang == "en":
            return "Not in list"
        else:
            return "不在清单"
    return rate_val

def calc_composite_tariff(base: str, t301: str, fl_text: str, sec232_full: str, metal_flag: str) -> dict:
    expr_parts = []
    num_list = []
    use_232 = False
    b_num = parse_rate_number(base)
    t301_num = parse_rate_number(t301)
    sec232_rate = None
    sec232_type = ""
    if sec232_full != "-":
        sec232_rate, sec232_type = sec232_full.split(" | ",1)
    cond1 = False
    if metal_flag == "是" and sec232_full != "-":
        t = sec232_type.upper()
        if "STEEL" in t or "COPPER" in t or "ALUMINUM" in t or "AUTOMOBILE" in t:
            cond1 = True
    if cond1:
        use_232 = True
        expr_parts.append(base)
        if b_num is not None:
            num_list.append(b_num)
        if t301 != "-":
            expr_parts.append(t301)
            if t301_num is not None:
                num_list.append(t301_num)
        expr_parts.append(sec232_rate)
        s232_num = parse_rate_number(sec232_rate)
        if s232_num is not None:
            num_list.append(s232_num)
    else:
        expr_parts.append(base)
        if b_num is not None:
            num_list.append(b_num)
        if t301 != "-":
            expr_parts.append(t301)
            if t301_num is not None:
                num_list.append(t301_num)
        fl_skip_set = {"不在清单", "Not in list", "0（豁免清单内）", "0 (Exempted)"}
        fl_effective = fl_text not in fl_skip_set
        if fl_effective:
            expr_parts.append(fl_text)
            fl_num = parse_rate_number(fl_text)
            if fl_num is not None:
                num_list.append(fl_num)
    expr_str = " + ".join(expr_parts)
    if len(num_list) >0:
        total_val = sum(num_list)
        total_str = f"{total_val:.2f}%"
    else:
        total_str = "(无法计算)"
    return {"expr": expr_str, "total": total_str, "use_232": use_232}

# ===================== Excel导出函数（修改为返回二进制字节流） =====================
def generate_excel(rows, lang="zh") -> bytes:
    wb = Workbook()
    ws = wb.active
    if lang == "en":
        headers = ["HS CODE","Description","Base Tariff","Section301(China)","Section232","Forced-Labor Tariff","Composite Formula","Total"]
    else:
        headers = ["HS CODE","商品描述","基础关税","对华301加征","232加征","强迫劳动加征","综合关税公式","综合关税合计"]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.get("hs_code",""),
            r.get("product_desc",""),
            r.get("mfn_rate",""),
            r.get("tariff301_CN",""),
            r.get("sec232",""),
            r.get("tariff301_forcedlabor",""),
            r.get("composite_expr",""),
            r.get("composite_total","")
        ])
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()

# ===================== Streamlit页面主体，替换Flask index路由 =====================
st.set_page_config(page_title="US关税查询工具", layout="wide")
st.title("美国综合关税查询工具")

# 语言切换
lang = st.radio("Language / 语言", ["zh","en"], horizontal=True)

# UI文本
ui_text = {
    "zh":{
        "origin_label":"选择原产国",
        "hs_input_label":"输入HS编码（支持单行，多行批量仅支持10位）",
        "metal_label":"是否包含钢铁/铝/铜/汽车零部件",
        "warn_metal":"请选择【钢铁铝及衍生品/汽车零部件】，此项为必填",
        "warn_input":"请输入HS编码，单行支持6/8/10位；多行批量仅支持10位HS编码。",
        "warn_no_result":"输入HS编码未在HTS数据库找到，请核对编码。",
        "btn_search":"开始查询",
        "btn_export":"导出Excel"
    },
    "en":{
        "origin_label":"Select Origin Country",
        "hs_input_label":"Input HS code(s). Single search support 6/8/10-digit; batch multiple lines only support 10-digit HS.",
        "metal_label":"Contains Steel / Aluminum / Copper / Automobile parts?",
        "warn_metal":"Please select whether contains steel‑aluminum‑copper-autoparts (Yes / No)",
        "warn_input":"Please input HS code(s). Single search support 6/8/10‑digit; batch multiple lines only support 10‑digit HS.",
        "warn_no_result":"HS code not found in HTS database, please check your input.",
        "btn_search":"Search",
        "btn_export":"Export Excel"
    }
}[lang]

# 表单控件
sel_origin = st.selectbox(ui_text["origin_label"], ORIGIN_COUNTRY_LIST, index=0)
search_hs = st.text_area(ui_text["hs_input_label"], height=120)
metal_flag = st.radio(ui_text["metal_label"], ["是","否"], horizontal=True)

warn_msg = ""
result_data = st.session_state.get("query_result", [])

if st.button(ui_text["btn_search"]):
    if not metal_flag:
        warn_msg = ui_text["warn_metal"]
    else:
        raw_lines = search_hs.splitlines()
        hs_batch_raw = []
        for line in raw_lines:
            clean_hs = clean_input_hs(line)
            if clean_hs:
                hs_batch_raw.append(clean_hs)
        if len(hs_batch_raw) == 0:
            warn_msg = ui_text["warn_input"]
        else:
            conn = get_db_conn()
            is_batch_mode = len(hs_batch_raw) > 1
            result_data = []
            for search_hs_item in hs_batch_raw:
                if is_batch_mode:
                    if len(search_hs_item) != 10:
                        continue
                    rows_hts = conn.execute('''
                        SELECT * FROM hts
                        WHERE REPLACE(REPLACE(hs_code, '.',''),'-','') = ?
                    ''', (search_hs_item,)).fetchall()
                else:
                    if len(search_hs_item) >= 8:
                        query_prefix = search_hs_item[:8] + "%"
                    else:
                        query_prefix = search_hs_item + "%"
                    rows_hts = conn.execute('''
                        SELECT * FROM hts
                        WHERE REPLACE(REPLACE(hs_code, '.',''),'-','') LIKE ?
                        ORDER BY hs_code
                    ''', (query_prefix,)).fetchall()
                if len(rows_hts) ==0:
                    continue
                for hts_row in rows_hts:
                    hs = hts_row["hs_code"]
                    desc = hts_row["product_desc"]
                    hs_clean = hs.replace(".", "").replace("-","")
                    base_rate = get_8digit_tariff(conn, hs_clean)
                    rate_301 = None
                    if sel_origin == "China":
                        r301 = conn.execute('''
                            SELECT add_tariff_rate
                            FROM tariff_301
                            WHERE ? LIKE REPLACE(REPLACE(hs_code, '.',''),'-','') || '%'
                        ''', (hs_clean,)).fetchone()
                        if r301:
                            rate_301 = r301["add_tariff_rate"]
                    sec232_text = get_section232_info(conn, hs_clean)
                    fl_rate_text_raw = ""
                    fl_excl = 0
                    fl_row = conn.execute('''
                        SELECT country_rate_text,is_exclusion
                        FROM forcedlabor_2607
                        WHERE ? LIKE REPLACE(REPLACE(hs_code, '.',''),'-','') || '%'
                    ''',(hs_clean,)).fetchone()
                    if fl_row:
                        fl_rate_text_raw = fl_row["country_rate_text"]
                        fl_excl = fl_row["is_exclusion"]
                    fl_display_text = parse_forced_labor_rate(fl_rate_text_raw, sel_origin, fl_excl, lang)
                    composite_result = calc_composite_tariff(
                        base=base_rate,
                        t301=rate_301 if rate_301 else "-",
                        fl_text=fl_display_text,
                        sec232_full=sec232_text,
                        metal_flag=metal_flag
                    )
                    if composite_result["use_232"]:
                        fl_display_text = "0"
                    if is_batch_mode:
                        disp_mfn = base_rate
                        disp_301 = rate_301 if rate_301 else "-"
                        disp_232 = sec232_text
                        disp_fl = fl_display_text
                        disp_expr = composite_result["expr"]
                        disp_total = composite_result["total"]
                    else:
                        is_8digit = (len(hs_clean) == 8)
                        if is_8digit:
                            disp_mfn = ""
                            disp_301 = ""
                            disp_232 = ""
                            disp_fl = ""
                            disp_expr = ""
                            disp_total = ""
                        else:
                            disp_mfn = base_rate
                            disp_301 = rate_301 if rate_301 else "-"
                            disp_232 = sec232_text
                            disp_fl = fl_display_text
                            disp_expr = composite_result["expr"]
                            disp_total = composite_result["total"]
                    item = {
                        "hs_code": hs,
                        "product_desc": desc,
                        "mfn_rate": disp_mfn,
                        "tariff301_CN": disp_301,
                        "sec232": disp_232,
                        "tariff301_forcedlabor": disp_fl,
                        "composite_expr": disp_expr,
                        "composite_total": disp_total
                    }
                    result_data.append(item)
            conn.close()
            if len(result_data) == 0:
                warn_msg = ui_text["warn_no_result"]
    # 存入st.session_state，替代Flask session
    st.session_state["query_result"] = result_data

# 提示信息
if warn_msg:
    st.error(warn_msg)

# 展示查询结果表格
result_data = st.session_state.get("query_result", [])
if result_data:
    st.subheader("查询结果 / Query Result")
    st.dataframe(result_data, use_container_width=True)
    excel_bytes = generate_excel(result_data, lang)
    st.download_button(
        label=ui_text["btn_export"],
        data=excel_bytes,
        file_name="US_Tariff_Query.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.info("⚠️本工具仅供公开参考，报关请以美国海关官方数据为准。")
