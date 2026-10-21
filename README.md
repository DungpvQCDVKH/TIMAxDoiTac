# Báo cáo CSKH TIMA-CIMB

**Luồng:** máy nội bộ (có SQL Server) chạy `export_data.py` → CSV nhỏ trong `data/` → push GitHub → Streamlit Cloud chỉ đọc CSV.

1. Nội bộ: `pip install -r requirements-export.txt`, tạo `sql_config.json` (đã có, **không push**).
2. Cập nhật số liệu: `python export_data.py --push` (hoặc không `--push` rồi tự commit/push thư mục `data/`).
   Tháng cụ thể: `python export_data.py --from 2026-01 --to 2026-09`.
3. Streamlit Cloud: repo **private**, main file `app.py`, requirements.txt (nhẹ, không có driver SQL).
   Nên khai báo `APP_PASSWORD = "..."` trong Secrets để khóa app (dữ liệu có tên KH).
4. Chạy thử local: `streamlit run app.py`.

Lên lịch tự động: Windows Task Scheduler chạy `python export_data.py --push` mỗi ngày/tháng.
