# -*- coding: utf-8 -*-
# BÁO CÁO KẾT QUẢ CHĂM SÓC KHÁCH HÀNG TIMA-CIMB  (Streamlit)
# Chạy:  streamlit run app.py
# Nguồn: file CSV trong data/ (export_data.py lấy từ SQL Server, chỉ SELECT)
# Logic: giống file SQL "Lấy dữ liệu các khoản B00 CIMB không xử lý được trong tháng"
#   - Tập B00 tháng N : DB_Monthly (BUSINESS_DATE = cuối tháng N) JOIN tblDanhMucLoanID (Lender = 75564, CIMB)
#                       JOIN tblSynthesisReport (Fordate = cuối tháng N), DaydueFirst < ngày 1 tháng N+1, BUCKET_BOM = 'B00'
#   - POS_BOM    : SUM(POS_BOM) * 1e6 của tập B00 trên
#   - HĐ Roll    : tập B00 trên có DPD_CURRENT > 0 (chưa xử lý được trong tháng)
#   - Roll_EOM   : SUM(POS_CURRENT) * 1e6 của các HĐ Roll (dư nợ gốc cuối kỳ)
#   - % Roll EOM : Roll_EOM / POS_BOM
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Báo cáo CSKH TIMA-CIMB", page_icon="📊", layout="wide")

# ── Nguồn dữ liệu: file CSV trong thư mục data/ (do export_data.py chạy nội bộ tạo ra) ──
DATA_DIR = Path(__file__).parent / "data"
STATUS_COL = "Tình trạng chăm sóc cuối cùng"  # để trống, nhập tay trong file Excel tải về
AUTHOR = "Dungpv-QC.DVKH.TIMA"

# Bảng màu theo báo cáo mẫu
GOLD, BROWN, BLUE, PURPLE = "#B8860B", "#7A4A00", "#2A5CA8", "#7E63B0"
GOLD_LIGHT = "#E3CF9A"
INK, INK2, GRID = "#3D3220", "#7A6A4A", "#EDE6D3"

# ── (Tùy chọn) khóa mật khẩu: khai báo APP_PASSWORD trong Secrets của Streamlit Cloud ──
try:
    _app_pw = str(st.secrets["APP_PASSWORD"]) if "APP_PASSWORD" in st.secrets else ""
except Exception:  # noqa
    _app_pw = ""
if _app_pw:
    if st.text_input("Mật khẩu truy cập", type="password") != _app_pw:
        st.stop()

DETAIL_COLS = ["BUSINESS_DATE", "MA_TC", "LOAN_ID", "Tên KH", "Fromdate", "Nextdate", "DPD_CURRENT",
               "POS_BOM", "POS_CURRENT", "Sản phẩm"]


def shift_month(y, m, k):
    t = y * 12 + (m - 1) + k
    return t // 12, t % 12 + 1


@st.cache_data(ttl=300, show_spinner=False)
def load_summary():
    """Bảng tổng hợp theo tháng: ym, pos (VND), roll (VND), n_b00, n_roll."""
    f = DATA_DIR / "monthly_summary.csv"
    if not f.exists():
        return pd.DataFrame(columns=["ym", "pos", "roll", "n_b00", "n_roll"])
    t = pd.read_csv(f)
    return t.sort_values("ym").reset_index(drop=True)


@st.cache_data(ttl=300, show_spinner=False)
def load_detail(ym):
    """Danh sách HĐ roll của tháng ym (yyyymm)."""
    f = DATA_DIR / f"roll_{ym}.csv"
    if not f.exists():
        return pd.DataFrame(columns=DETAIL_COLS)
    d = pd.read_csv(f, dtype={"MA_TC": str}, encoding="utf-8-sig")
    for c in ["DPD_CURRENT", "POS_BOM", "POS_CURRENT"]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    return d


@st.cache_data(ttl=300, show_spinner=False)
def last_update():
    f = DATA_DIR / "last_update.txt"
    return f.read_text(encoding="utf-8").strip() if f.exists() else ""


def load_trend(summary, y, m, n_months):
    yms = [yy * 100 + mm for yy, mm in (shift_month(y, m, k) for k in range(-(n_months - 1), 1))]
    t = summary[summary.ym.isin(yms)].copy()
    t["pct"] = t.roll / t.pos.where(t.pos > 0)
    t["pct_n"] = t.n_roll / t.n_b00.where(t.n_b00 > 0)
    t["label"] = t.ym.map(lambda v: f"T{v % 100:02d}/{str(v // 100)[2:]}")
    return t[t.pos > 0].reset_index(drop=True)


# ───────────────────────────── Định dạng ─────────────────────────────
def vn(x, d=0):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "—"
    return f"{x:,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def ty(x):
    return vn(x / 1e9, 2) + " tỷ"


def present(detail):
    """Định dạng giống Sheet1 trong file Excel mẫu (cột tình trạng chăm sóc để trống)."""
    def dt(s):
        return pd.to_datetime(s, errors="coerce").dt.strftime("%Y-%m-%d")

    out = pd.DataFrame({
        "Date": pd.to_datetime(detail["BUSINESS_DATE"], errors="coerce").dt.date,
        "Mã hợp đồng (Partner user ID)": detail["MA_TC"].astype(str),
        "LoanID": detail["LOAN_ID"],
        **({"Tên KH": detail["Tên KH"].fillna("").astype(str).str.strip()} if "Tên KH" in detail.columns else {}),
        "Ngày giải ngân": dt(detail["Fromdate"]),
        "Ngày đến hạn": dt(detail["Nextdate"]),
        "Số ngày trễ hạn": detail["DPD_CURRENT"],
        "POS_BOM": detail["POS_BOM"],
        "POS_EOM": detail["POS_CURRENT"],
        "Product": detail["Sản phẩm"],
        "Roll_EOM": 1,
        "Roll_POS": detail["POS_CURRENT"],
        STATUS_COL: "",
    })
    return out.sort_values("POS_BOM", ascending=False).reset_index(drop=True)


# ───────────────────────────── CSS ─────────────────────────────
st.markdown(f"""
<style>
.stApp {{ background:#f1f4f4; }}
.block-container {{ max-width:1400px; padding-top:1.2rem; padding-bottom:3rem; }}
header[data-testid="stHeader"] {{ background:transparent; }}
#MainMenu, footer, .stDeployButton, [data-testid="stToolbar"], [data-testid="stDecoration"],
[data-testid="stStatusWidget"], [data-testid="stElementToolbar"], [data-testid="stDownloadButton"],
[data-testid="stToolbarActions"], [data-testid="stAppDeployButton"] {{ display:none !important; visibility:hidden !important; }}
section[data-testid="stSidebar"] {{ background:#fffdf7; }}
.sb-t {{ font-weight:800; font-size:13px; letter-spacing:1px; color:#6b5a3a; margin:6px 0 4px; }}
.hdr {{ position:relative; background:linear-gradient(100deg,#7c6840 0%,#a58d58 55%,#bda568 100%); border-radius:12px;
  padding:22px 30px 26px; color:#fff; display:flex; justify-content:space-between; align-items:flex-start;
  border-bottom:3px solid #d9a93a; margin-bottom:14px; }}
.hdr .eb {{ color:#6fe3d0; font-size:11px; letter-spacing:2.5px; font-weight:700; }}
.hdr h1 {{ color:#fff; font-size:27px; font-weight:800; margin:4px 0 6px; padding:0; letter-spacing:.3px; }}
.hdr .sub {{ font-size:13px; opacity:.88; }}
.hdr .pill {{ background:rgba(255,255,255,.18); border:1px solid rgba(255,255,255,.35); border-radius:20px;
  padding:5px 14px; font-size:12px; white-space:nowrap; }}
.hdr .by {{ position:absolute; right:16px; bottom:6px; font-size:10.5px; opacity:.85; letter-spacing:.3px; }}
.sec {{ display:flex; align-items:center; gap:12px; padding:10px 4px 2px; }}
.sec .bd {{ background:#8a6b2f; color:#fff; font-weight:800; width:26px; height:26px; border-radius:6px;
  display:flex; align-items:center; justify-content:center; font-size:13px; }}
.sec .t {{ font-weight:800; font-size:16px; color:{INK}; }}
.sec .s {{ font-size:11px; color:{INK2}; }}
.kpis {{ display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin:8px 0 10px; }}
@media (max-width:1000px) {{ .kpis {{ grid-template-columns:repeat(3,1fr); }} }}
@media (max-width:640px) {{ .kpis {{ grid-template-columns:repeat(2,1fr); }} }}
.kc {{ background:#fdf6ea; border:1px solid #f1dcc0; border-radius:10px; padding:12px 14px 12px; }}
.kc .ti {{ text-align:center; font-size:11.5px; font-weight:800; letter-spacing:.6px; color:{INK}; margin-bottom:8px; }}
.kc .r {{ display:flex; justify-content:space-between; font-size:12px; padding:2px 0; color:{INK2}; }}
.kc .r b {{ color:{BROWN}; font-size:15px; }}
.kc .r.big b {{ font-size:19px; }}
.kc .pl {{ white-space:nowrap; gap:6px; margin-top:8px; border-radius:7px; padding:7px 10px; display:flex;
  justify-content:space-between; font-size:11.5px; font-weight:700; }}
.kc .pl b {{ font-size:14px; }}
.kc .pl.good {{ background:#d9f5ea; color:#0f7d57; }}
.kc .pl.bad {{ background:#fde8ea; color:#c5303a; }}
.kc .pl.neu {{ background:#f4ead2; color:#7a5c14; }}
[class*="st-key-card_"] {{ background:#fffdf7; border:1px solid #efe3c4 !important; border-radius:10px;
  border-top:3px solid {GOLD} !important; padding:4px 8px 0; }}
.st-key-card_blue {{ border-top-color:{BLUE} !important; }}
.st-key-card_purple {{ border-top-color:{PURPLE} !important; }}
.ct {{ font-weight:800; font-size:14px; color:{INK}; letter-spacing:.2px; padding:4px 0 0; }}
.cs {{ font-size:11px; color:{INK2}; }}
table.rpt {{ border-collapse:collapse; width:100%; font-size:14px; background:#fff; }}
table.rpt th {{ background:{GOLD}; color:#fff; padding:8px 12px; border:1px solid #e6d7a8; }}
table.rpt td {{ padding:8px 12px; border:1px solid #efe3c4; text-align:right; }}
table.rpt td:first-child {{ text-align:center; font-weight:700; }}
table.rpt td.hl {{ font-weight:800; color:{BROWN}; }}
.note {{ font-size:12px; color:{INK2}; line-height:1.7; margin-top:8px; }}
</style>
""", unsafe_allow_html=True)


# ───────────────────────────── Thành phần dùng chung ─────────────────────────────
def sec_head(num, title, sub=""):
    st.markdown(f'<div class="sec"><div class="bd">{num}</div><div><div class="t">{title}</div>'
                f'<div class="s">{sub}</div></div></div>', unsafe_allow_html=True)


def chart_title(t, s=""):
    st.markdown(f'<div class="ct">{t}</div>' + (f'<div class="cs">{s}</div>' if s else ""), unsafe_allow_html=True)


def kpi_card(title, cur, prev, fmt, kind, neutral=False):
    """kind: 'rel' (% thay đổi) | 'pp' (điểm %) | 'abs'. Chỉ số càng thấp càng tốt (trừ khi neutral)."""
    ok = prev is not None and cur is not None and not pd.isna(prev) and not pd.isna(cur)
    if ok:
        diff = cur - prev
        if kind == "rel":
            dtxt = f"{diff / prev * 100:+.1f}%".replace(".", ",") if prev else "—"
        elif kind == "pp":
            dtxt = f"{diff * 100:+.2f} đ%".replace(".", ",")
        else:
            dtxt = f"{diff:+,.0f}".replace(",", ".")
        arrow = "▲" if diff > 0 else ("▼" if diff < 0 else "■")
        tone = "neu" if (diff == 0 or neutral) else ("bad" if diff > 0 else "good")
        pill = f'<div class="pl {tone}"><span>Chênh lệch</span><b>{arrow} {dtxt}</b></div>'
    else:
        pill = '<div class="pl neu"><span>Chênh lệch</span><b>—</b></div>'
    return (f'<div class="kc"><div class="ti">{title}</div>'
            f'<div class="r big"><span>Kỳ này</span><b>{fmt(cur)}</b></div>'
            f'<div class="r"><span>Tháng trước</span><b style="font-size:13px">{fmt(prev) if prev is not None else "—"}</b></div>'
            f'{pill}</div>')


def style(fig, h=300, legend=False):
    fig.update_layout(
        height=h, margin=dict(l=8, r=8, t=8, b=8), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Segoe UI, Inter, sans-serif", size=12, color=INK), separators=",.",
        hovermode="x unified", showlegend=legend,
        hoverlabel=dict(bgcolor="#fffdf7", font_color=INK, bordercolor="#e6d7a8"))
    fig.update_xaxes(showgrid=False, linecolor=GRID, tickangle=0, tickfont=dict(color=INK2, size=10.5))
    fig.update_yaxes(gridcolor=GRID, zeroline=False, tickfont=dict(color=INK2))
    return fig


def show(fig, key):
    st.plotly_chart(fig, width="stretch", key=key, config={"displayModeBar": False})


def bar_month(mt, col, scale, name, unit, key, color, light):
    y = mt[col] / scale
    cols = [light] * (len(mt) - 1) + [color]
    txt = [""] * (len(mt) - 1) + [vn(y.iloc[-1], 2 if scale > 1 else 0)]
    fig = go.Figure(go.Bar(x=mt["label"], y=y, marker_color=cols, marker_cornerradius=4, text=txt,
                           textposition="outside", cliponaxis=False, name=name,
                           hovertemplate=f"{name}: %{{y:,.2f}} {unit}<extra></extra>"))
    fig.update_yaxes(title=unit)
    show(style(fig), key)


# ───────────────────────────── Bộ lọc (sidebar, hàng dọc) ─────────────────────────────
summary = load_summary()
avail = sorted(summary.ym[summary.pos > 0].astype(int).tolist()) if len(summary) else []
if not avail:
    st.error("Chưa có dữ liệu. Chạy export_data.py trong mạng nội bộ rồi đẩy thư mục data/ lên GitHub.")
    st.stop()
years = sorted({v // 100 for v in avail}, reverse=True)

with st.sidebar:
    st.markdown('<div class="sb-t">BỘ LỌC</div>', unsafe_allow_html=True)
    sel_y = st.selectbox("Năm", years)
    months = sorted([v % 100 for v in avail if v // 100 == sel_y], reverse=True)
    sel_m = st.selectbox("Tháng", months, format_func=lambda m: f"Tháng {m}")
    n_months = st.selectbox("Số tháng so sánh", [3, 6, 9, 12], index=1)
    if st.button("🔄 Tải lại dữ liệu"):
        st.cache_data.clear()
        st.rerun()

ym = sel_y * 100 + sel_m
detail_raw = load_detail(ym)
mt = load_trend(summary, sel_y, sel_m, n_months)
py_, pm_ = shift_month(sel_y, sel_m, -1)
rc = mt[mt.ym == ym]
rp = mt[mt.ym == py_ * 100 + pm_]
if rc.empty:
    st.info(f"Chưa có dữ liệu B00 CIMB cho tháng {sel_m:02d}/{sel_y}.")
    st.stop()
cur = rc.iloc[0].to_dict()
prev = rp.iloc[0].to_dict() if not rp.empty else None
det = present(detail_raw)


def pv(k):
    return prev[k] if prev else None


st.markdown(f"""<div class="hdr"><div><div class="eb">KHỐI DỊCH VỤ KHÁCH HÀNG · TIMA</div>
<h1>BÁO CÁO KẾT QUẢ CHĂM SÓC KHÁCH HÀNG TIMA-CIMB</h1>
<div class="sub">Dư nợ gốc B00 · Roll EOM · % Roll · Danh sách KH roll · So sánh nhiều tháng</div></div>
<div class="pill">Tháng {sel_m:02d}/{sel_y} · {vn(len(det))} HĐ roll{(" · cập nhật " + last_update()) if last_update() else ""}</div>
<div class="by">{AUTHOR}</div></div>""", unsafe_allow_html=True)

# ───────────────────────────── 1. Tổng quan ─────────────────────────────
sec_head(1, "TỔNG QUAN", f"CIMB · B00 chốt BOM tháng {sel_m:02d}/{sel_y}")
pct_f = lambda v: vn(v * 100, 2) + "%"  # noqa
cards = [
    kpi_card("DƯ NỢ GỐC POS BOM", cur["pos"], pv("pos"), ty, "rel", neutral=True),
    kpi_card("ROLL EOM", cur["roll"], pv("roll"), ty, "rel"),
    kpi_card("% ROLL EOM", cur["pct"], pv("pct"), pct_f, "pp"),
    kpi_card("SỐ HĐ ROLL", cur["n_roll"], pv("n_roll"), vn, "abs"),
    kpi_card("% ROLL THEO SỐ HĐ", cur["pct_n"], pv("pct_n"), pct_f, "pp"),
]
st.markdown('<div class="kpis">' + "".join(cards) + "</div>", unsafe_allow_html=True)
st.markdown(f"""<table class="rpt"><tr><th>Agency</th><th>POS_BOM</th><th>Roll_EOM</th><th>%_Roll_EOM</th></tr>
<tr><td>CIMB</td><td>{vn(cur['pos'])}</td><td>{vn(cur['roll'])}</td><td class="hl">{pct_f(cur['pct'])}</td></tr></table>
<div class="note"><b>Trong đó:</b><br>
- <b>POS BOM</b>: Dư nợ gốc các HĐ có ngày thanh toán trong tháng {sel_m:02d} chốt tại ngày đầu tháng {sel_m:02d}<br>
- <b>Roll EOM</b>: Dư nợ gốc các HĐ không thanh toán trong tháng {sel_m:02d}, bị chuyển nợ quá hạn<br>
- <b>% Roll EOM</b>: Tỉ lệ nợ gốc bị chuyển nhóm quá hạn từ B00 lên B01 chốt tại ngày cuối cùng của tháng {sel_m:02d}</div>""",
            unsafe_allow_html=True)

# ───────────────────────────── 2. So sánh nhiều tháng ─────────────────────────────
sec_head(2, "SO SÁNH NHIỀU THÁNG", f"{n_months} tháng gần nhất đến {sel_m:02d}/{sel_y} · cột đậm = tháng đang xem")
c1, c2, c3 = st.columns(3)
with c1, st.container(key="card_gold1"):
    chart_title("% ROLL EOM", "Roll EOM / POS BOM")
    fig = go.Figure(go.Scatter(x=mt["label"], y=mt["pct"] * 100, mode="lines+markers", name="% Roll EOM",
                               line=dict(color=GOLD, width=2.5),
                               marker=dict(size=8, color=GOLD, line=dict(color="#fffdf7", width=2)),
                               hovertemplate="% Roll EOM: %{y:.2f}%<extra></extra>"))
    fig.update_yaxes(ticksuffix="%", rangemode="tozero")
    show(style(fig), "c_pct")
with c2, st.container(key="card_gold2"):
    chart_title("DƯ NỢ GỐC POS BOM", "Đơn vị: tỷ đồng")
    bar_month(mt, "pos", 1e9, "POS BOM", "tỷ đồng", "c_pos", GOLD, GOLD_LIGHT)
with c3, st.container(key="card_purple1"):
    chart_title("ROLL EOM", "Đơn vị: tỷ đồng")
    bar_month(mt, "roll", 1e9, "Roll EOM", "tỷ đồng", "c_roll", PURPLE, "#CDBFE3")

with st.expander("Bảng số liệu các tháng"):
    t = mt[["label", "pos", "roll", "pct", "n_b00", "n_roll", "pct_n"]].copy()
    t.columns = ["Tháng", "POS_BOM (VND)", "Roll_EOM (VND)", "% Roll EOM", "Số HĐ B00", "Số HĐ roll", "% Roll theo HĐ"]
    t["% Roll EOM"] *= 100
    t["% Roll theo HĐ"] *= 100
    for c in ["POS_BOM (VND)", "Roll_EOM (VND)", "Số HĐ B00", "Số HĐ roll"]:
        t[c] = t[c].map(vn)
    for c in ["% Roll EOM", "% Roll theo HĐ"]:
        t[c] = t[c].map(lambda x: vn(x, 2) + "%")
    st.dataframe(t, hide_index=True, width="stretch")

# ───────────────────────────── 3. Danh sách KH roll ─────────────────────────────
sec_head(3, "DANH SÁCH KH ROLL TRONG THÁNG",
         "HĐ B00 đầu tháng, có due trong tháng nhưng chưa xử lý được (còn trễ hạn cuối tháng)")
if det.empty:
    st.success("Không có HĐ nào roll trong tháng này.")
else:
    q = st.text_input("Tìm mã HĐ / LoanID / tên KH", "", placeholder="Nhập mã hợp đồng hoặc LoanID")
    v = det
    if q.strip():
        v = det[det["Mã hợp đồng (Partner user ID)"].str.contains(q.strip(), case=False, na=False)
                | det["LoanID"].astype(str).str.contains(q.strip(), na=False)
                | (det["Tên KH"].str.contains(q.strip(), case=False, na=False) if "Tên KH" in det.columns else False)]
    st.caption(f"{vn(len(v))} HĐ · dư nợ gốc roll {ty(v['Roll_POS'].sum() * 1e6)} · POS tính theo triệu đồng")
    st.dataframe(v, hide_index=True, width="stretch", height=420,
                 column_order=[c for c in v.columns if c != STATUS_COL],
                 column_config={"POS_BOM": st.column_config.NumberColumn(format="%.3f"),
                                "POS_EOM": st.column_config.NumberColumn(format="%.3f"),
                                "Roll_POS": st.column_config.NumberColumn(format="%.3f"),
                                "LoanID": st.column_config.NumberColumn(format="%d")})
