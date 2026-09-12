# Tính tiền thuế cần nộp

Trang web tĩnh tra cứu thuế sử dụng đất phi nông nghiệp từ file Excel sổ bộ
thuế và tính tổng cùng ba khoản thu theo nhân khẩu: nghĩa trang nhân dân, quỹ
phòng chống thiên tai, bảo vệ môi trường.

## Cách dùng

1. Mở trang (địa chỉ GitHub Pages hoặc mở `index.html`).
2. Bước 1: bấm **Chọn file Excel**, chọn file sổ bộ thuế (.xlsx/.xls).
   File chỉ được đọc trên máy bạn, không gửi đi đâu.
3. Bước 2: gõ tên hoặc số CCCD, bấm vào hộ cần tính.
   Một hộ có thể có nhiều dòng (nhiều thửa đất hoặc khoản thu, ví dụ tiền chậm nộp); bấm một dòng sẽ tự tích các dòng cùng hộ, bạn có thể bỏ tích rồi bấm **Tính cho N dòng đã chọn**.
4. Bước 3: nhập số người từng khoản, xem tổng tiền.

Đơn giá có thể sửa ở khối **Đơn giá** cuối trang, được ghi nhớ trên trình duyệt.

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
