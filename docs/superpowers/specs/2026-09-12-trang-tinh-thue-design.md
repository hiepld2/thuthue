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

Hai file: `thue.js` chứa các hàm thuần (chạy được trong Node và trình duyệt) và `index.html` chứa giao diện, nạp SheetJS và `thue.js`. Thư viện duy nhất:
SheetJS (`xlsx`) tải từ CDN cdnjs, phiên bản ghim cố định. Không có bước build.

Ba đơn vị logic trong JavaScript, tách thành các hàm thuần để dễ kiểm thử:

1. `parseSheetRows(matrix) -> { ok, rows, missingColumns, error }`
   `matrix` là mảng hai chiều lấy từ SheetJS `sheet_to_json(sheet, { header: 1 })`.
   Tìm dòng tiêu đề, ánh xạ cột theo tên, trả về mảng bản ghi chuẩn hoá
   `{ mst, ten, ngaySinh, cccd, dienTich, thon, stt, tieuMuc, thuePhaiNop, daNop }`.
2. `searchRows(rows, query, limit) -> rows[]` và `searchInfo(rows, query, limit) -> { rows, total }`
   Tìm theo tên hoặc CCCD, không phân biệt hoa thường và dấu, khớp một phần;
   `searchInfo` trả thêm tổng số dòng khớp thật trước khi cắt theo `limit`.
3. `computeTotals(input, rates) -> { lines, total }`
   Tính 4 dòng và tổng từ số thuế đất, số người từng mục, số tháng, đơn giá.

Phần giao diện chỉ gọi ba hàm trên và vẽ kết quả.

## Đọc file Excel

- Dòng tiêu đề: dòng đầu tiên có ô chứa "MST" và ô chứa "Tên" (so khớp không
  dấu, không phân biệt hoa thường). Tiêu đề có thể trải hai dòng (dòng dưới
  chỉ có số thứ tự cột) nên bỏ qua dòng ngay dưới tiêu đề nếu toàn số nguyên nhỏ.
- Ánh xạ cột theo từ khoá trong tên tiêu đề (đã bỏ dấu, thường hoá):
  - `mst`: chứa "mst"
  - `ten`: chứa "ten nnt", "ten nguoi nop thue", "ho ten" hoặc "ho va ten"
  - `ngaySinh`: chứa "ngay sinh"
  - `cccd`: chứa "cccd"
  - `dienTich`: chứa "dien tich"
  - `thon`: chứa "thon"
  - `thuePhaiNop`: chứa "tong so thue phai nop" (ưu tiên), sau đó "phai nop"
  - `stt` (tuỳ chọn): chứa "stt" hoặc "so thu tu"
  - `tieuMuc` (tuỳ chọn): chứa "tieu muc"
  - `daNop` (tuỳ chọn): chứa "so tien da nop" hoặc "da nop"
- Cột bắt buộc: `ten`, `cccd`, `thuePhaiNop`. Không dùng từ khoá `ten` đơn lẻ để
  tránh bắt nhầm cột khác có chữ "tên"; chấp nhận `ten nnt`, `ten nguoi nop thue`,
  `ho ten`, `ho va ten`. Thiếu cột nào thì báo lỗi liệt kê
  tên cột thiếu và không nạp dữ liệu.
- Bỏ qua dòng không có tên.
- `thuePhaiNop`: chuyển chuỗi có dấu chấm/phẩy phân cách nghìn thành số. Ô trống
  hoặc không phải số thì lưu `null`.
- CCCD đọc dạng chuỗi để giữ số 0 đầu.

## Tra cứu

- Ô nhập duy nhất. Chuẩn hoá chuỗi tìm: bỏ dấu, thường hoá, gộp khoảng trắng.
- Khớp nếu tên chuẩn hoá chứa chuỗi tìm, hoặc CCCD chứa chuỗi tìm.
- Luôn hiện danh sách kết quả (tối đa 50), kể cả khi chỉ có một. Mỗi dòng gồm
  ô tích, tên, ngày sinh, CCCD, thôn, diện tích, thuế phải nộp để phân biệt các
  hộ trùng tên. Không tự chọn thay người dùng.
- Một hộ có thể có nhiều dòng trong Excel (nhiều thửa đất hoặc khoản thu, ví
  dụ dòng thuế và dòng tiền chậm nộp — cột "Tiểu mục" khác nhau — cùng
  MST/tên/CCCD). Bấm một dòng sẽ tích dòng đó và tự tích các dòng khác cùng
  tên chuẩn hoá: khớp theo CCCD khi CCCD khác rỗng và khác "0"; nếu CCCD rỗng
  hoặc "0" thì khớp theo MST (khi MST khác rỗng). Người dùng có thể bỏ tích
  từng dòng. Nút "Tính cho N dòng đã chọn" chuyển sang bước 3.
- Kết quả tìm kiếm giới hạn tối đa 50 dòng hiển thị; nếu tổng số dòng khớp còn
  nhiều hơn 50, hiện thêm thông báo mời gõ thêm để thu hẹp.
- Bước 3 liệt kê các dòng đã chọn; thuế đất = tổng "Tổng số thuế phải nộp" của
  các dòng đó. Dòng nào trống số thuế thì tính 0 đ và ghi chú; nếu mọi dòng đều
  trống thì ghi "Không tìm thấy trong file".
- Chuỗi tìm rỗng thì không hiện kết quả.
- Không có kết quả thì hiện "Không tìm thấy trong file". Người dùng vẫn có thể
  bấm "Tính với thuế đất 0 đ" để tiếp tục với thuế đất = 0 đ.

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
- Sau khi nạp thành công hiện "Đã nạp N dòng".

## Giao diện

Thiết kế ưu tiên màn hình điện thoại (mobile-first), một cột, tối đa khoảng
760px trên máy tính. Thứ tự từ trên xuống:

1. **Tiêu đề trang.**
2. **Khối "Đơn giá"** (cấu hình đầu tiên, ít thay đổi) dạng `<details>` thu gọn;
   dòng tóm tắt hiện đơn giá đang dùng, ví dụ "Nghĩa trang 15.000 đ/người ·
   Thiên tai 10.000 đ/người · Môi trường 15.000 đ/người/tháng · 6 tháng". Mở
   ra có 4 ô nhập và nút "Khôi phục mặc định". Lưu `localStorage`.
3. **Bước 1: Chọn file Excel** + dòng trạng thái.
4. **Bước 2: Tìm hộ**: ô tìm, danh sách kết quả có ô tích (tên, STT, tiểu mục,
   ngày sinh, CCCD, thôn, diện tích, đã nộp, thuế), thông báo cắt 50, hộp
   "Không tìm thấy trong file" + nút "Tính với thuế đất 0 đ".
5. **Bước 3: Nhập số người** — hiện ngay bên dưới khi đã tích ít nhất một dòng
   (hoặc bấm "Tính với thuế đất 0 đ"). Gồm: tên hộ và số dòng đã chọn; ba ô số
   người (nghĩa trang, phòng chống thiên tai, bảo vệ môi trường); ô "Số tháng"
   đặt ngay cạnh ô môi trường với nhãn "chỉ áp dụng cho bảo vệ môi trường",
   mặc định theo đơn giá; nút lớn "Tính tiền".
6. **Hoá đơn** — chỉ hiện sau khi bấm "Tính tiền": tên hộ; các dòng sổ bộ đã
   chọn (STT, tiểu mục, đã nộp, thuế); bốn khoản, mỗi khoản một hàng gồm tên,
   cách tính và số tiền căn phải; mỗi khoản có một màu nền nhạt riêng để dễ
   phân biệt (thuế đất xanh dương nhạt, nghĩa trang tím nhạt, thiên tai cam
   nhạt, môi trường xanh lá nhạt; chữ vẫn tối, tương phản cao); hàng "TỔNG
   CỘNG" nền trắng, chữ đậm; nút "Tính cho hộ khác" quay về Bước 2 và xoá lựa
   chọn. Khi tích dòng ở Bước 2, Bước 3 chỉ hiện ra bên dưới, không tự cuộn
   hay chuyển focus để người dùng tiếp tục tích/bỏ tích. Bấm "Tính tiền" lần nữa thì vẽ lại
   hoá đơn theo số liệu mới; không tự tính khi gõ.

Quy tắc gộp hộ (đối xứng): hai dòng cùng hộ khi tên chuẩn hoá bằng nhau và
(CCCD hợp lệ bằng nhau, hoặc MST khác rỗng bằng nhau). CCCD hợp lệ là khác rỗng
và khác "0".

Yêu cầu chung: tiếng Việt toàn bộ; cỡ chữ cơ bản 18px; ô nhập và nút cao tối
thiểu 48px, trên điện thoại nút "Tính tiền" và "Chọn file Excel" chiếm hết
chiều ngang; `inputmode="numeric"` cho ô số; không dùng `<table>` cho hoá đơn
(dùng hàng flex để không tràn ngang); không tràn ngang ở bề rộng 360px; màu
tương phản cao; bước sau chỉ hiện khi bước trước xong.

## Kiểm thử

- Tạo file `test/mau.xlsx` khoảng 10 dòng lấy từ ảnh sổ bộ (tiêu đề giống bảng
  in) để kiểm tra thủ công. File này được commit vì chỉ chứa dữ liệu mẫu.
- File `test/test.js` chạy bằng `node test/test.js` với `node:assert`: tìm dòng tiêu đề, ánh xạ cột, chuyển số có dấu chấm, tìm không dấu,
  tìm theo CCCD, tính tổng, xử lý `null`.
- Mở `index.html` trong trình duyệt của app, nạp `test/mau.xlsx`, tra một tên
  và kiểm tra tổng.

## Triển khai

- Khởi tạo git tại thư mục dự án. `.gitignore` loại ảnh (`*.JPG`, `*.HEIC`,
  `*.jpg`) và file Excel thật (`*.xlsx`, `*.xls`) trừ `test/mau.xlsx`.
- Đẩy lên GitHub, bật Pages từ nhánh `main`, thư mục gốc. Truy cập
  `https://<user>.github.io/<repo>/`.
