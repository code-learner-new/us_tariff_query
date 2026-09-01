from flask import Flask, render_template, request, make_response, session
import sqlite3
import re
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)
app.secret_key = "us_tariff_2026_secret_001"
DB = r"./tariff.db"

ORIGIN_COUNTRY_LIST = ["China","Vietnam","Malaysia","Mexico","India","Indonesia","Thailand"]


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
    if "FREE" in s or "0.00%" in s:
        return 0.0
    match = re.search(r"([\d\.]+)", s)
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


@app.route("/", methods=["GET","POST"])
def index():
    select_origin = "China"
    input_text = ""
    result_data = []
    warn_msg = ""
    lang = request.args.get("lang", "zh")
    if request.method == "POST":
        lang = request.form.get("lang_hidden", "zh")

    metal_flag = ""

    if request.method == "POST":
        select_origin = request.form.get("sel_origin","China").strip()
        input_text = request.form.get("search_hs","")
        metal_flag = request.form.get("metal_flag","").strip()

        if not metal_flag:
            if lang == "en":
                warn_msg = "Please select whether contains steel-aluminum-copper (Yes / No)"
            else:
                warn_msg = "请选择【钢铁铝及衍生品或汽车零部件】，此项为必填"
        else:
            raw_lines = input_text.splitlines()
            hs_batch_raw = []
            for line in raw_lines:
                clean_hs = clean_input_hs(line)
                if clean_hs:
                    hs_batch_raw.append(clean_hs)

            if len(hs_batch_raw) == 0:
                if lang == "en":
                    warn_msg = "Please input HS code(s). Single search support 6/8/10-digit; batch multiple lines only support 10-digit HS."
                else:
                    warn_msg = "请输入HS编码，单行支持6/8/10位；多行批量仅支持10位HS编码。"
            else:
                conn = get_db_conn()
                is_batch_mode = len(hs_batch_raw) > 1

                for search_hs in hs_batch_raw:
                    if is_batch_mode:
                        if len(search_hs) != 10:
                            continue
                        rows_hts = conn.execute('''
                            SELECT * FROM hts
                            WHERE REPLACE(REPLACE(hs_code, '.',''),'-','') = ?
                        ''', (search_hs,)).fetchall()
                    else:
                        if len(search_hs) >= 8:
                            query_prefix = search_hs[:8] + "%"
                        else:
                            query_prefix = search_hs + "%"
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
                        if select_origin == "China":
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

                        fl_display_text = parse_forced_labor_rate(fl_rate_text_raw, select_origin, fl_excl, lang)

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
                            "tariff301": disp_301,
                            "sec232": disp_232,
                            "fl_text": disp_fl,
                            "composite_expr": disp_expr,
                            "composite_total": disp_total
                        }
                        result_data.append(item)
                conn.close()
                if len(result_data) == 0:
                    if lang=="en":
                        warn_msg = "HS code not found in HTS database, please check your input."
                    else:
                        warn_msg = "输入HS编码未在HTS数据库找到，请核对编码。"

    session["query_result"] = result_data
    session["ui_lang"] = lang
    session["metal_flag_session"] = metal_flag

    return render_template("index.html",
                           origin_list=ORIGIN_COUNTRY_LIST,
                           sel_origin=select_origin,
                           search_hs=input_text,
                           metal_flag=metal_flag,
                           result=result_data,
                           warn_msg=warn_msg,
                           lang=lang)

@app.route("/export_excel", methods=["POST"])
def export_excel():
    json_data = request.get_json()
    lang = json_data.get("lang", session.get("ui_lang", "zh"))
    rows = session.get("query_result", [])
    if not rows:
        return make_response("no data", 400)

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
            r.get("tariff301",""),
            r.get("sec232",""),
            r.get("fl_text",""),
            r.get("composite_expr",""),
            r.get("composite_total","")
        ])

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    resp = make_response(output.read())
    # =========修复：全部使用普通ASCII减号，禁止特殊连字符=========
    filename = "US_Tariff_Query.xlsx"
    resp.headers["Content-Disposition"] = "attachment; filename=" + filename
    resp.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return resp

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)