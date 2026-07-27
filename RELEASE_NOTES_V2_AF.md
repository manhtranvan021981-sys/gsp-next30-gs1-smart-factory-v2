# GS1 Smart Factory V2 — Dòng hàng mẹ AF

## Phạm vi thay đổi

- Bộ lọc **Dòng hàng mẹ (AF)** dùng duy nhất cột AF làm khóa.
- Hiển thị cố định đủ 19 nhóm theo thứ tự `01–19`.
- Chuẩn hóa AF bằng cách bỏ khoảng trắng thừa và chuyển thành chữ hoa trước
  khi so khớp chính xác.
- Không suy luận mảng từ mã/tên vật tư, máy, công đoạn hoặc dòng hàng con.

## Nhóm kiểm soát dữ liệu

- `00_Chưa khai báo dòng hàng mẹ`: AF trống.
- `98_Xung đột AF theo LTT/phiếu`: cùng LTT hoặc phiếu thống kê có nhiều AF.
- `99_AF chưa ánh xạ`: AF có giá trị nhưng không thuộc danh mục 19 mã.

Các nhóm này vẫn được tính trong **Tất cả dòng hàng** và được đánh dấu Data
đỏ; không gộp vào `17_Nhóm hàng khác`.

## Các luồng đã đồng bộ

- Bộ xử lý Excel tạo dữ liệu nén theo tháng.
- Luồng đọc dữ liệu live CSV/JSONP trong dashboard.
- Bộ lọc chính, OEE/capa và Pareto downtime.
- Lịch máy hiện tại/tuần tiếp theo.
- Data Quality, bảng LTT, máy, thợ chính và action list.
- Cache GitHub Actions V2 tách khỏi dữ liệu phân nhóm V1.

## Quy tắc không thay đổi

- Nhà máy nhận dữ liệu: cột E phải bằng `GS1`.
- Mapping KPI cố định: `S/U/AR/AT/BG/BI/CH`.
- Công thức KPI, quy tắc khử trùng LTT/phiếu và cấu hình nguồn không đổi.
- Dashboard GS6 không thuộc phạm vi thay đổi.
- Giữ bản V1 để rollback.

## Kiểm thử phát hành

Chạy:

```bash
python3 scripts/test_af_contract.py --out .qa/af-contract-data
python3 scripts/verify_build.py \
  --config factory-config.json \
  --data .qa/af-contract-data
```

Bộ test hợp đồng bao phủ đủ 19 mã hợp lệ, AF trống, AF ngoài danh mục, xung
đột theo LTT và xung đột theo phiếu thống kê.
