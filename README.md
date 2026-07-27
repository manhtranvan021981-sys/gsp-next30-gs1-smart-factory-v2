# GSP NEXT30 – GS1 Smart Factory V2

Gói phát hành độc lập cho **GS1 – Nhà máy Goldsun Hà Nội**. Dashboard dùng
chung lõi tính toán với GS6 nhưng có cấu hình nguồn, schema kiểm tra và
namespace riêng.

## Nguồn đã cấu hình

- File ID: `1c3_CSdnh9sxALEt-6FKyrd-4-J1K_Yhl`
- Tệp: `P3_Tong_Hop_LTT_2507-HN.xlsx`
- Sheet: `P3.Tổng hợp lệnh thao tác`
- Header: dòng 9; vùng đọc: `A9:CT`
- Điều kiện nhận dòng: cột E bằng `GS1`
- Bộ lọc `Dòng hàng mẹ (AF)` dùng duy nhất cột AF, chuẩn hóa khoảng trắng và
  chữ hoa rồi so khớp chính xác với danh mục 19 mã `HOC…TCKT`.
- Không suy luận dòng hàng từ mã vật tư, tên vật tư, máy, công đoạn hoặc
  dòng hàng con.
- Dữ liệu bất thường được tách riêng:
  `00_Chưa khai báo dòng hàng mẹ`,
  `98_Xung đột AF theo LTT/phiếu`,
  `99_AF chưa ánh xạ`.

## Nguyên tắc an toàn

- Mapping số liệu giữ cố định theo vị trí cột `S/U/AR/AT/BG/BI/CH`.
- Manifest phải có `plant = GS1`; dữ liệu nhà máy khác bị chặn trước phát hành.
- Nguồn lỗi thì workflow dừng; GitHub Pages tiếp tục giữ bản hợp lệ gần nhất.
- Dashboard mặc định tải tháng mới nhất; chỉ tải toàn bộ tháng khi người dùng
  chọn `Tất cả các tháng`.
- Data Quality kiểm tra tỷ lệ ánh xạ AF, AF trống, mã ngoài danh mục và xung
  đột AF theo LTT/phiếu thống kê.
- Không upload Excel vào repository.

## Tệp cấu hình riêng

- `factory-config.json`: bộ xử lý dữ liệu và workflow.
- `factory-config.js`: giao diện và nhận diện nhà máy.

Giữ bản V1 làm rollback; V2 dùng cache dữ liệu riêng để không khôi phục nhầm
gói phân loại 6 value stream cũ.
