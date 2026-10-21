# -*- coding: utf-8 -*-
"""
Chạy NỘI BỘ (trong mạng Tima): lấy số liệu từ SQL Server (CHỈ SELECT) và ghi ra file CSV nhỏ trong data/
để app Streamlit online chỉ đọc CSV, không cần kết nối SQL Server.

  python export_data.py                       # 12 tháng gần nhất đã chốt
  python export_data.py --from 2026-01 --to 2026-09
  python export_data.py --push                # xong tự git add/commit/push thư mục data/

Kết nối: file sql_config.json cạnh file này (KHÔNG đưa lên GitHub) hoặc biến môi trường
TIMA_SQL_SERVER / TIMA_SQL_DATABASE / TIMA_SQL_USER / TIMA_SQL_PASSWORD.

Logic giống file SQL "Lấy dữ liệu các khoản B00 CIMB không xử lý được trong tháng":
  - Tập B00 tháng N: DB_Monthly (BUSINESS_DATE = cuối tháng N) JOIN tblDanhMucLoanID (Lender 75564 = CIMB)
                     JOIN tblSynthesisReport (Fordate = cuối tháng N), DaydueFirst < ngày 1 tháng N+1, BUCKET_BOM = 'B00'
  - POS_BOM  = SUM(POS_BOM) * 1e6 của tập B00;  HĐ roll = tập B00 có DPD_CURRENT > 0
  - Roll_EOM = SUM(POS_CURRENT) * 1e6 của các HĐ roll
"""
import argparse
import calendar
import json
import os
import re
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
DATA_DIR = ROOT / "data"
LENDER_ID = 75564  # CIMB
NAME_COL_OVERRIDE = None  # vd "loans.[CUSTOMER_NAME]"; None = tự dò
NAME_EXACT = ["customer_name", "customername", "fullname", "full_name", "tenkh", "ten_kh", "tenkhachhang",
              "ten_khach_hang", "clientname", "borrowername", "customer_full_name"]
JOINS = """
FROM DB_Monthly AS loans
INNER JOIN THNRE.[MecashNew].[dbo].[tblDanhMucLoanID] AS loan_details ON loans.LOAN_ID = loan_details.LoanID
INNER JOIN THNRE.[Report_THN].[dbo].[tblSynthesisReport] AS synthesis ON loans.LOAN_ID = synthesis.LoanID
"""


# ───────────────────────────── Kết nối (chỉ SELECT) ─────────────────────────────
def conn_cfg():
    f = ROOT / "sql_config.json"
    c = json.loads(f.read_text(encoding="utf-8")) if f.exists() else {}
    get = lambda k, env: os.environ.get(env) or c.get(k, "")  # noqa
    cfg = (get("server", "TIMA_SQL_SERVER"), get("database", "TIMA_SQL_DATABASE"),
           get("user", "TIMA_SQL_USER"), get("password", "TIMA_SQL_PASSWORD"))
    if not (cfg[0] and cfg[2] and cfg[3]):
        sys.exit("Thiếu thông tin kết nối: tạo sql_config.json (server, database, user, password) hoặc đặt biến môi trường.")
    return cfg


def _connect(server, database, user, password):
    err = None
    try:
        import pyodbc

        drivers = [d for d in pyodbc.drivers() if "SQL Server" in d]
        pref = ["ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server",
                "ODBC Driver 13 for SQL Server", "SQL Server Native Client 11.0", "SQL Server"]
        drv = next((p for p in pref if p in drivers), drivers[0] if drivers else None)
        if drv is None:
            raise RuntimeError("Không có ODBC driver SQL Server")
        cs = (f"DRIVER={{{drv}}};SERVER={server};UID={user};PWD={{{password.replace('}', '}}')}}};"
              "Encrypt=no;TrustServerCertificate=yes;")
        if database:
            cs += f"DATABASE={database};"
        return pyodbc.connect(cs, timeout=30)
    except Exception as e:  # noqa
        err = e
    try:
        import pymssql

        return pymssql.connect(server=server, user=user, password=password,
                               database=database or None, login_timeout=30)
    except Exception as e:  # noqa
        raise RuntimeError(f"Không kết nối được SQL Server. pyodbc: {err} | pymssql: {e}")


def run_select(sql):
    if not sql.lstrip().upper().startswith("SELECT"):
        raise ValueError("Chỉ cho phép SELECT")
    conn = _connect(*conn_cfg())
    try:
        cur = conn.cursor()
        cur.execute(sql)
        cols = [c[0] for c in cur.description]
        return pd.DataFrame([tuple(r) for r in cur.fetchall()], columns=cols)
    finally:
        conn.close()


_NAME_CACHE = []


def find_name_col():
    """Tìm cột Tên KH (CUSTOMER_NAME...) trong các bảng của SQL gốc; None nếu không có."""
    if _NAME_CACHE:
        return _NAME_CACHE[0]
    found, probed = NAME_COL_OVERRIDE, False
    if not found:
        sources = [("loans", "INFORMATION_SCHEMA.COLUMNS", "DB_Monthly"),
                   ("loan_details", "THNRE.[MecashNew].INFORMATION_SCHEMA.COLUMNS", "tblDanhMucLoanID"),
                   ("synthesis", "THNRE.[Report_THN].INFORMATION_SCHEMA.COLUMNS", "tblSynthesisReport")]
        for alias, view, tbl in sources:
            try:
                cols = list(run_select(f"SELECT COLUMN_NAME FROM {view} WHERE TABLE_NAME = '{tbl}'")["COLUMN_NAME"])
            except Exception:  # noqa
                continue
            probed = probed or bool(cols)
            for pref in NAME_EXACT:
                hit = next((c for c in cols if str(c).lower() == pref), None)
                if hit and not found:
                    found = f"{alias}.[{hit}]"
        if not found and not probed:
            found = "loans.[CUSTOMER_NAME]"
    _NAME_CACHE.append(found)
    return found


# ───────────────────────────── Truy vấn 1 tháng ─────────────────────────────
def export_month(y, m):
    """Trả về (dòng tổng hợp, danh sách HĐ roll) của tháng y/m."""
    last = date(y, m, calendar.monthrange(y, m)[1])
    ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
    d, n = last.strftime("%Y%m%d"), date(ny, nm, 1).strftime("%Y-%m-%d")

    def make_detail_q(nc):
        name_sel = f"{nc} AS [Tên KH]," if nc else ""
        return f"""SELECT
    loans.BUSINESS_DATE, loans.MA_TC, loans.LOAN_ID, {name_sel} loans.Fromdate, synthesis.Nextdate,
    loans.DPD_CURRENT, loans.POS_BOM, loans.POS_CURRENT,
    CASE
        WHEN loans.PRODUCT_NAME_CL LIKE N'%XM%' THEN N'Xe máy'
        WHEN loans.PRODUCT_NAME_CL LIKE N'%Ô tô%' THEN N'Ô tô'
        ELSE N'Khác'
    END AS [Sản phẩm]
{JOINS}
WHERE loans.BUSINESS_DATE = '{d}'
    AND synthesis.DaydueFirst < '{n}'
    AND loan_details.LenderIDDisbursement = {LENDER_ID}
    AND loans.DPD_CURRENT > 0
    AND loans.BUCKET_BOM = 'B00'
    AND synthesis.Fordate = '{d}'"""

    q_pos = f"""SELECT SUM(loans.POS_BOM) * 1e6 AS pos_bom, COUNT(*) AS n
{JOINS}
WHERE loans.BUSINESS_DATE = '{d}'
  AND synthesis.fordate = '{d}'
  AND synthesis.DaydueFirst < '{n}'
  AND loan_details.LenderIDDisbursement = {LENDER_ID}
  AND loans.BUCKET_BOM = 'B00'"""

    name_col = find_name_col()
    try:
        detail = run_select(make_detail_q(name_col))
    except Exception:  # noqa  (cột tên lỗi -> chạy lại không có Tên KH)
        if not name_col:
            raise
        detail = run_select(make_detail_q(None))
    pos = run_select(q_pos)
    for c in ["DPD_CURRENT", "POS_BOM", "POS_CURRENT"]:
        detail[c] = pd.to_numeric(detail[c], errors="coerce")
    pos_bom = float(pd.to_numeric(pos["pos_bom"], errors="coerce").fillna(0).iloc[0]) if len(pos) else 0.0
    n_b00 = int(pd.to_numeric(pos["n"], errors="coerce").fillna(0).iloc[0]) if len(pos) else 0
    row = dict(ym=y * 100 + m, pos=pos_bom, roll=float(detail["POS_CURRENT"].sum() * 1e6),
               n_b00=n_b00, n_roll=len(detail))
    return row, detail


def month_range(a, b):
    out, (y, m) = [], a
    while (y, m) <= b:
        out.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def parse_ym(s):
    y, m = s.split("-")
    return int(y), int(m)


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--from", dest="a", help="tháng bắt đầu, dạng YYYY-MM")
    ap.add_argument("--to", dest="b", help="tháng kết thúc, dạng YYYY-MM (mặc định: tháng trước)")
    ap.add_argument("--push", action="store_true", help="git add/commit/push thư mục data/")
    args = ap.parse_args(argv)

    t = date.today()
    end = parse_ym(args.b) if args.b else ((t.year, t.month - 1) if t.month > 1 else (t.year - 1, 12))
    start = parse_ym(args.a) if args.a else (end[0] - 1, end[1] + 1) if end[1] < 12 else (end[0], 1)
    months = month_range(start, end)

    DATA_DIR.mkdir(exist_ok=True)
    rows, ok = [], 0
    for y, m in months:
        try:
            row, det = export_month(y, m)
        except Exception as e:  # noqa
            print(f"[LỖI] {m:02d}/{y}: {e}")
            continue
        if row["pos"] <= 0:
            print(f"[BỎ QUA] {m:02d}/{y}: chưa có dữ liệu chốt")
            continue
        det.to_csv(DATA_DIR / f"roll_{row['ym']}.csv", index=False, encoding="utf-8-sig")
        rows.append(row)
        ok += 1
        print(f"[OK] {m:02d}/{y}: POS_BOM={row['pos']:,.0f} | Roll_EOM={row['roll']:,.0f} | "
              f"{row['n_roll']}/{row['n_b00']} HĐ")
    if not ok:
        sys.exit("Không xuất được tháng nào.")

    f = DATA_DIR / "monthly_summary.csv"
    old = pd.read_csv(f) if f.exists() else pd.DataFrame(columns=["ym", "pos", "roll", "n_b00", "n_roll"])
    new = pd.DataFrame(rows)
    out = pd.concat([old[~old.ym.isin(new.ym)], new]).sort_values("ym").reset_index(drop=True)
    out.to_csv(f, index=False)
    (DATA_DIR / "last_update.txt").write_text(datetime.now().strftime("%d/%m/%Y %H:%M"), encoding="utf-8")
    print(f"Đã ghi {ok} tháng vào {DATA_DIR}")

    if args.push:
        msg = f"Cập nhật dữ liệu {datetime.now():%d/%m/%Y %H:%M}"
        for cmd in (["git", "add", "data"], ["git", "commit", "-m", msg], ["git", "push"]):
            r = subprocess.run(cmd, cwd=ROOT)
            if r.returncode and cmd[1] != "commit":  # commit rỗng (không có thay đổi) thì bỏ qua
                sys.exit(f"Lệnh {' '.join(cmd)} lỗi")


if __name__ == "__main__":
    main()
