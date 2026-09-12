# Trang tính số tiền thuế cần nộp — Thiết kế

Ngày: 2026-09-12

## Mục tiêu

Một trang web tĩnh (chạy trên GitHub Pages) giúp người thu thuế thôn tra cứu
thuế sử dụng đất phi nông nghiệp của một hộ từ file Excel sổ bộ thuế, rồi nhập
số nhân khẩu để tính tổng các khoản phải nộp.

## Phạm vi

Có:
- Đọc file Excel (.xlsx/.xls) do người dùng chọn trên trình duyệt.
- Tra cứu theo tên hoặc số CCCD.
- Tính 4 khoản và tổng cộng.
- Cho phép sửa đơn giá và số tháng mặc định, lưu trên trình duyệt.

Không có:
- Server, đăng nhập, lưu lịch sử tra cứu, in phiếu, xuất Excel.
- File Excel không được đưa lên repo.

## Kiến trúc

Một file `index.html` duy nhất chứa HTML, CSS, JavaScript. Thư viện duy nhất:
SheetJS (`xlsx`) tải từ CDN cdnjs, phiên bản ghim cố định. Không có bước build.

Ba đơn vị logic trong JavaScript, tách thành các hàm thuần để dễ kiểm thử:

1. `parseWorkbook(arrayBuffer) -> { rows, missingColumns }`
   Đọc sheet đầu tiên, tìm dòng tiêu đề, ánh xạ cột theo tên, trả về mảng bản
   ghi chuẩn hoá `{ mst, ten, ngaySinh, cccd, dienTich, thon, thuePhaiNop }`.
2. `searchRows(rows, query) -> rows[]`
   Tìm theo tên hoặc CCCD, không phân biệt hoa thường và dấu, khớp một phần.
3. `computeTotals(input, rates) -> { lines, total }`
   Tính 4 dòng và tổng từ số thuế đất, số người từng mục, số tháng, đơn giá.

Phần giao diện chỉ gọi ba hàm trên và vẽ kết quả.

## Đọc file Excel

- Dòng tiêu đề: dòng đầu tiên có ô chứa "MST" và ô chứa "Tên" (so khớp không
  dấu, không phân biệt hoa thường). Tiêu đề có thể trải hai dòng (dòng dưới
  chỉ có số thứ tự cột) nên bỏ qua dòng ngay dưới tiêu đề nếu toàn số nguyên nhỏ.
- Ánh xạ cột theo từ khoá trong tên tiêu đề (đã bỏ dấu, thường hoá):
  - `mst`: chứa "mst"
  - `ten`: chứa "ten nnt" hoặc "ten"
  - `ngaySinh`: chứa "ngay sinh"
  - `cccd`: chứa "cccd"
  - `dienTich`: chứa "dien tich"
  - `thon`: chứa "thon"
  - `thuePhaiNop`: chứa "tong so thue phai nop" (ưu tiên), sau đó "phai nop"
- Cột bắt buộc: `ten`, `cccd`, `thuePhaiNop`. Thiếu cột nào thì báo lỗi liệt kê
  tên cột thiếu và không nạp dữ liệu.
- Bỏ qua dòng không có tên.
- `thuePhaiNop`: chuyển chuỗi có dấu chấm/phẩy phân cách nghìn thành số. Ô trống
  hoặc không phải số thì lưu `null`.
- CCCD đọc dạng chuỗi để giữ số 0 đầu.

## Tra cứu

- Ô nhập duy nhất. Chuẩn hoá chuỗi tìm: bỏ dấu, thường hoá, gộp khoảng trắng.
- Khớp nếu tên chuẩn hoá chứa chuỗi tìm, hoặc CCCD chứa chuỗi tìm.
- Hiện tối đa 50 kết quả dạng danh sách: tên, ngày sinh, CCCD, thôn, thuế phải
  nộp. Bấm một dòng để chọn.
- Chuỗi tìm rỗng thì không hiện kết quả.
- Không có kết quả thì hiện "Không tìm thấy trong file". Người dùng vẫn có thể
  bấm "Tính không có thuế đất" để tiếp tục với thuế đất = 0 đ.

## Tính tiền

| Khoản | Công thức | Người dùng nhập |
|---|---|---|
| Thuế sử dụng đất phi nông nghiệp | lấy từ cột "Tổng số thuế phải nộp" | không |
| Nghĩa trang nhân dân | đơnGiáNT × sốNgười | số người |
| Quỹ phòng chống thiên tai | đơnGiáPCTT × sốNgười | số người |
| Bảo vệ môi trường | đơnGiáBVMT × sốNgười × sốTháng | số người, số tháng |
| Tổng cộng | tổng 4 dòng | |

- Đơn giá mặc định: nghĩa trang 15.000; phòng chống thiên tai 10.000; bảo vệ
  môi trường 15.000 đ/người/tháng; số tháng mặc định 6.
- Số người mặc định 0, số nguyên không âm. Số tháng số nguyên không âm.
- Nếu hộ được chọn có `thuePhaiNop` là `null`, hiện ghi chú "Không tìm thấy
  trong file" trên dòng thuế đất và lấy 0 đ.
- Tổng cập nhật ngay khi thay đổi bất kỳ ô nhập.
- Định dạng tiền: phân cách nghìn bằng dấu chấm, hậu tố " đ" (ví dụ
  `1.234.000 đ`).

## Cài đặt và lưu trữ trên trình duyệt

- Khối "Đơn giá" thu gọn được, gồm 3 đơn giá và số tháng mặc định. Lưu vào
  `localStorage` khoá `thue-settings` ngay khi thay đổi; đọc lại khi mở trang.
  Có nút "Khôi phục mặc định".
- Không lưu dữ liệu Excel vào trình duyệt; mỗi lần mở trang phải chọn file lại.
- Mọi truy cập `localStorage` bọc try/catch; thiếu hoặc hỏng thì dùng mặc định.

## Xử lý lỗi

- File không đọc được: "Không đọc được file. Hãy chọn file .xlsx hoặc .xls."
- Không tìm thấy dòng tiêu đề: "Không tìm thấy dòng tiêu đề (cần có cột MST và
  Tên NNT)."
- Thiếu cột bắt buộc: "Thiếu cột: ...".
- Sau khi nạp thành công hiện "Đã nạp N hộ".

## Giao diện

- Một cột, tối đa khoảng 760px, dùng được trên điện thoại.
- Thứ tự từ trên xuống: tiêu đề trang, chọn file + trạng thái, ô tìm kiếm +
  danh sách kết quả, thông tin hộ đã chọn, bảng tính với ô nhập ngay trong
  bảng, dòng tổng nổi bật, khối đơn giá thu gọn ở cuối.
- Tiếng Việt toàn bộ.

## Kiểm thử

- Tạo file `test/mau.xlsx` khoảng 10 dòng lấy từ ảnh sổ bộ (tiêu đề giống bảng
  in) để kiểm tra thủ công. File này được commit vì chỉ chứa dữ liệu mẫu.
- File `test/test.html` chạy các hàm thuần bằng assert đơn giản trong trình
  duyệt: tìm dòng tiêu đề, ánh xạ cột, chuyển số có dấu chấm, tìm không dấu,
  tìm theo CCCD, tính tổng, xử lý `null`.
- Mở `index.html` trong trình duyệt của app, nạp `test/mau.xlsx`, tra một tên
  và kiểm tra tổng.

## Triển khai

- Khởi tạo git tại thư mục dự án. `.gitignore` loại ảnh (`*.JPG`, `*.HEIC`,
  `*.jpg`) và file Excel thật (`*.xlsx`, `*.xls`) trừ `test/mau.xlsx`.
- Đẩy lên GitHub, bật Pages từ nhánh `main`, thư mục gốc. Truy cập
  `https://<user>.github.io/<repo>/`.
