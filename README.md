# Tính tiền thuế cần nộp

Trang web tĩnh tra cứu thuế sử dụng đất phi nông nghiệp từ file Excel sổ bộ
thuế và tính tổng cùng ba khoản thu theo nhân khẩu: nghĩa trang nhân dân, quỹ
phòng chống thiên tai, bảo vệ môi trường.

## Cách dùng

0. Mở trang (địa chỉ GitHub Pages hoặc mở `index.html`). Kiểm tra khối **Đơn giá**
   ở đầu trang, chỉ sửa khi đơn giá thực tế thay đổi (được ghi nhớ trên trình duyệt).
1. Bước 1: bấm **Chọn file Excel**, chọn file sổ bộ thuế (.xlsx/.xls).
   File chỉ được đọc trên máy bạn, không gửi đi đâu. Trang sẽ nhớ file này trên
   trình duyệt: lần sau mở lại trang, dữ liệu tự nạp lại mà không cần chọn file.
   Bấm **Chọn file khác** để nạp file khác, hoặc **Xoá dữ liệu đã lưu** để xoá
   dữ liệu đang được nhớ trên trình duyệt. Khi đã có dữ liệu, khối này tự thu gọn
   (giống khối Đơn giá); bấm vào tiêu đề để mở lại.
2. Bước 2: gõ tên hoặc số CCCD, bấm vào hộ cần tính.
   Một hộ có thể có nhiều dòng (nhiều thửa đất hoặc khoản thu, ví dụ tiền chậm nộp); bấm một dòng sẽ tự tích các dòng cùng hộ, bạn có thể bỏ tích dòng không cần. Bước 3 hiện ngay bên dưới khi có ít nhất một dòng được tích.
3. Bước 3: nhập số người cho từng khoản (nghĩa trang, phòng chống thiên tai,
   bảo vệ môi trường); số tháng chỉ áp dụng cho khoản bảo vệ môi trường. Bấm
   **Tính tiền**.
4. Xem hoá đơn ở cuối trang: các dòng sổ bộ đã chọn, từng khoản và tổng cộng.
   Bấm **Tính cho hộ khác** để quay lại Bước 2 và tính cho hộ tiếp theo.

## Yêu cầu file Excel

Sheet đầu tiên phải có dòng tiêu đề chứa các cột: `MST`, `Tên NNT`, `Số CCCD`,
`Tổng số thuế phải nộp`. Các cột khác (Ngày sinh, Diện tích, Thôn) là tuỳ chọn.
Ngoài ra còn nhận thêm các cột tuỳ chọn `STT`, `Tiểu mục`, `Số tiền đã nộp` —
hiện kèm trong danh sách để phân biệt các dòng của cùng một hộ (ví dụ dòng
thuế và dòng tiền chậm nộp). Số tiền được coi là số nguyên đồng (dấu chấm/phẩy
chỉ là phân cách nghìn). Xem ví dụ tại `test/mau.xlsx`.

## Đưa lên GitHub Pages

1. Tạo repo trên GitHub, đẩy nhánh `main`.
2. Vào **Settings → Pages**, mục *Build and deployment* chọn *Deploy from a
   branch*, nhánh `main`, thư mục `/ (root)`, bấm Save.
3. Sau khoảng một phút, truy cập `https://hiepld2.github.io/thuthue/`.

## Phát triển

- Chạy test: `node test/test.js`
- Sinh lại file mẫu: `python3 test/tao_mau.py`
- Không đưa file Excel thật lên repo (`.gitignore` đã chặn `*.xlsx` trừ `test/mau.xlsx`).
