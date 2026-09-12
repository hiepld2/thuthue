# test/tao_mau.py — sinh test/mau.xlsx (dữ liệu mẫu lấy từ ảnh sổ bộ)
from pathlib import Path
from openpyxl import Workbook

HEADER = ['STT', 'MST', 'Tên NNT', 'Ngày sinh', 'Số CCCD', 'Diện tích', 'Đoạn đường', 'Giá đất',
          'Mã phi nông nghiệp', 'Thôn/Tổ', 'Tiểu mục', 'Số thuế kỳ trước chuyển sang',
          'Số phát sinh trong kỳ', 'Số truy thu năm trước', 'Thuế miễn giảm',
          'Tổng số thuế phải nộp', 'Số tiền đã nộp', 'Chuyển kỳ sau', 'Ghi chú']

DUONG = 'Các trục đường có mặt cắt dưới 2,5m (xã Đồng Tảo)'
THON = 'Thôn Thống Nhất (Xã Đồng Tảo cũ)'

# (STT, MST, Tên, Ngày sinh, CCCD, Diện tích, Giá đất, Phát sinh, Tổng phải nộp)
DATA = [
    (177, '8220707543', 'LÊ VĂN AN',          '04/06/1936', '145665829',    200, 900000,  54000,  54000),
    (178, '8220707543', 'LÊ VĂN AN',          '04/06/1936', '145665829',    100, 900000,  27000,  27000),
    (343, '8220708963', 'NGUYỄN VĂN AN',      '18/04/1962', '011618238',    325, 2400000, 234000, 468000),
    (25,  '8023404726', 'LÊ THỊ VÂN ANH',     '17/09/1984', '001184013031', 120, 2400000, 4282,   4282),
    (133, '8220707166', 'GIANG LÊ ANH',       '01/01/1956', '145344429',    200, 900000,  54000,  54000),
    (369, '8220709131', 'GIANG LÊ TUẤN ANH',  '23/10/1997', '033097008492', 120, 2400000, 86400,  172800),
    (365, '8220709075', 'LÊ VĂN BÀN',         '23/03/1949', '033049004943', 200, 2400000, 144000, 144000),
    (465, '8905401838', 'CHU THỊ BẰNG',       '26/11/1967', '033167009147', 75,  5500000, 123750, 123750),
    (438, '8608470444', 'NGUYỄN CHIẾN BINH',  '25/04/1959', '030509004203', 60,  1800000, 32400,  32400),
    (256, '8220708226', 'GIANG LÊ BÌNH',      '01/08/1956', '033056006676', 50,  900000,  13500,  13500),
    (122, '8220707053', 'LÊ VĂN BÌNH',        '01/01/1970', '0',            50,  900000,  None,   None),
]

wb = Workbook()
ws = wb.active
ws.title = 'Sổ bộ'
ws.append(['SỔ BỘ THUẾ SỬ DỤNG ĐẤT PHI NÔNG NGHIỆP NĂM 2026'])
ws.append([])
ws.append(HEADER)
ws.append(list(range(1, len(HEADER) + 1)))
for stt, mst, ten, ns, cccd, dt, gia, ps, tong in DATA:
    ws.append([stt, mst, ten, ns, cccd, dt, DUONG, gia, '1019031264', THON, 1601,
               None, ps, None, None, tong, None, None, None])
# Cột CCCD (E) và MST (B) để dạng text để giữ số 0 đầu
for r in range(5, 5 + len(DATA)):
    ws.cell(row=r, column=2).number_format = '@'
    ws.cell(row=r, column=5).number_format = '@'

out = Path(__file__).with_name('mau.xlsx')
wb.save(out)
print('Đã ghi', out)
