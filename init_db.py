import sqlite3
import csv
import os
import re   # ← 必须导入，用于税率清洗

DB_FILE = r"./tariff.db"

def clean_hs(s):
    """清洗HS编码：剔除小数点，只保留数字字符"""
    if s is None:
        return ""
    raw = str(s)
    return "".join([c for c in raw if c.isdigit()])

def safe_int(val, default=0):
    """安全转int，空/空格/非数字返回默认0，防止CSV空值报错跳过整行"""
    if val is None:
        return default
    s = str(val).strip()
    if s == "":
        return default
    try:
        return int(s)
    except ValueError:
        return default

def clean_rate_text(val):
    """清洗税率字段：只保留数字、小数点、百分号、减号，消除不可见脏字符"""
    if val is None:
        return "-"
    raw = str(val)
    cleaned = re.sub(r"[^0-9.%\-]", "", raw)
    return cleaned if cleaned != "" else "-"

def init_database():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    # =========建表=========
    cur.execute('''DROP TABLE IF EXISTS hts''')
    cur.execute('''
    CREATE TABLE hts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hs_code TEXT,
        product_desc TEXT,
        unit TEXT,
        general_rate TEXT,
        applicable_origin TEXT
    )
    ''')
    cur.execute('''DROP TABLE IF EXISTS tariff_301''')
    cur.execute('''
    CREATE TABLE tariff_301 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hs_code TEXT,
        add_tariff_rate TEXT,
        applicable_origin TEXT,
        is_exclusion INTEGER DEFAULT 0,
        list_id TEXT,
        effective_date TEXT
    )
    ''')
    cur.execute('''DROP TABLE IF EXISTS forcedlabor_2607''')
    cur.execute('''
    CREATE TABLE forcedlabor_2607 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hs_code TEXT,
        country_rate_text TEXT,
        applicable_origin TEXT,
        is_control INTEGER DEFAULT 0,
        is_exclusion INTEGER DEFAULT 0
    )
    ''')
    cur.execute('''DROP TABLE IF EXISTS tariff_232''')
    cur.execute('''
    CREATE TABLE tariff_232 (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        hs_code TEXT,
        add_tariff_rate TEXT,
        type TEXT,
        applicable_origin TEXT
    )
    ''')
    log_messages = []

    # =========1、导入hts.csv=========
    csv_hts = r"./data/hts.csv"
    success_hts = 0
    skip_hts = 0
    try:
        with open(csv_hts,"r",encoding="utf-8-sig",newline="") as f:
            reader = csv.DictReader(f)
            for idx,row in enumerate(reader):
                raw_hs = row.get("hs_code","")
                hs = clean_hs(raw_hs)
                if not hs:
                    skip_hts +=1
                    continue
                cur.execute('''
                INSERT INTO hts(hs_code,product_desc,unit,general_rate,applicable_origin)
                VALUES(?,?,?,?,?)
                ''',(
                    hs,
                    row.get("product_desc",""),
                    row.get("unit",""),
                    clean_rate_text(row.get("general_rate")),
                    row.get("applicable_origin","ALL")
                ))
                success_hts +=1
    except FileNotFoundError:
        log_messages.append("⚠️ hts.csv 文件缺失！")
    except Exception as e:
        log_messages.append(f"⚠️ hts.csv读取异常:{str(e)}")
    log_messages.append(f"hts导入完成：成功{success_hts}，跳过{skip_hts}")

    # =========2、导入301_china.csv=========
    csv_301 = r"./data/301_china.csv"
    success_301 = 0
    skip_301 = 0
    try:
        with open(csv_301,"r",encoding="utf-8-sig",newline="") as f:
            reader = csv.DictReader(f)
            for idx,row in enumerate(reader):
                raw_hs = row.get("hs_code","")
                hs = clean_hs(raw_hs)
                if not hs:
                    skip_301 +=1
                    continue
                # 【关键修复】301税率入库前清洗，消除csv脏字符（换行/零宽空格/前置0）
                clean_rate = clean_rate_text(row.get("add_tariff_rate"))
                cur.execute('''
                INSERT INTO tariff_301(hs_code,add_tariff_rate,applicable_origin,is_exclusion,list_id,effective_date)
                VALUES(?,?,?,?,?,?)
                ''',(
                    hs,
                    clean_rate,
                    row.get("applicable_origin","China"),
                    safe_int(row.get("is_exclusion")),
                    row.get("list_id",""),
                    row.get("effective_date","")
                ))
                success_301 +=1
    except FileNotFoundError:
        log_messages.append("⚠️未找到301_china.csv，跳过301数据导入")
    except Exception as e:
        log_messages.append(f"⚠️301_china.csv异常:{str(e)}")
    log_messages.append(f"301导入完成：成功{success_301}，跳过{skip_301}")

    # =========3、导入 Forcelabor_2607.csv=========
    csv_fl = r"./data/Forcelabor_2607.csv"
    success_fl = 0
    skip_fl = 0
    try:
        with open(csv_fl,"r",encoding="utf-8-sig",newline="") as f:
            reader = csv.DictReader(f)
            for idx,row in enumerate(reader):
                raw_hs = row.get("hs_code","")
                hs = clean_hs(raw_hs)
                if not hs:
                    skip_fl +=1
                    continue
                cur.execute('''
                INSERT INTO forcedlabor_2607(hs_code,country_rate_text,applicable_origin,is_control,is_exclusion)
                VALUES(?,?,?,?,?)
                ''',(
                    hs,
                    row.get("country_rate_text",""),
                    "China",
                    safe_int(row.get("is_control")),
                    safe_int(row.get("is_exclusion"))
                ))
                success_fl +=1
    except FileNotFoundError:
        log_messages.append("⚠️未找到Forcelabor_2607.csv，请核对data目录文件名大小写")
    except Exception as e:
        log_messages.append(f"⚠️Forcelabor_2607.csv异常:{str(e)}")
    log_messages.append(f"强迫劳动导入完成：成功{success_fl}，跳过{skip_fl}")

    # =========4、导入 section232.csv=========
    csv_232 = r"./data/section232.csv"
    success_232 = 0
    skip_232 = 0
    try:
        with open(csv_232,"r",encoding="utf-8-sig",newline="") as f:
            reader = csv.DictReader(f)
            for idx,row in enumerate(reader):
                raw_hs = row.get("hs_code","")
                hs = clean_hs(raw_hs)
                if not hs:
                    skip_232 +=1
                    continue
                cur.execute('''
                INSERT INTO tariff_232(hs_code, add_tariff_rate, type, applicable_origin)
                VALUES(?,?,?,?)
                ''',(
                    hs,
                    clean_rate_text(row.get("add_tariff_rate")),
                    row.get("Type",""),
                    row.get("applicable_origin","ALL")
                ))
                success_232 += 1
    except FileNotFoundError:
        log_messages.append("⚠️未找到section232.csv，跳过232数据导入")
    except Exception as e:
        log_messages.append(f"⚠️section232.csv异常:{str(e)}")
    log_messages.append(f"Section232导入完成：成功{success_232}，跳过{skip_232}")

    conn.commit()
    conn.close()
    return log_messages

def main():
    """供streamlit app.py调用的入口函数"""
    logs = init_database()
    for msg in logs:
        print(msg)

if __name__ == "__main__":
    # 本地运行直接执行
    main()
