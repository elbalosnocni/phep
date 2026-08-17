# Đồng bộ phép năm Excel -> Google Sheets -> GitHub

## Cấu trúc nguồn thực tế

File:

X:\DS + PN + TP - 2014\Npn2023.xlsb

Mật khẩu mở file:

2000

Hai sheet cần lấy dữ liệu:

- 2026
- 2026 PL

Hai sheet được xem là 2 xưởng.

## Các cột

- A = HỌ VÀ TÊN
- Y = Paid
- Z = Left 2025
- AA = Left 2026
- AD = Gross

## Cách tính

### Đã sử dụng

Cột Y được SUM theo HỌ VÀ TÊN trên toàn bộ dòng tháng của cả 2 sheet.

Ví dụ:

NGUYỄN VĂN A:
Jan Y = 1
Feb Y = 0
Mar Y = 2
Apr Y = 1

Paid = 4

### Phép năm cũ

Lấy giá trị số đầu tiên không trống của cột Z theo tên.

### Phép năm hiện tại

Lấy giá trị số đầu tiên không trống của cột AA theo tên.

### Phép còn lại

Lấy giá trị số đầu tiên không trống của cột AD (Gross) theo tên.

KHÔNG SUM cột AD, vì trong file mẫu AD/Gross chỉ xuất hiện ở dòng tổng của nhân viên.

## Lưu ý

Hệ thống đang gom theo HỌ VÀ TÊN.

Nếu 2 nhân viên khác nhau có cùng họ tên, họ sẽ bị gộp thành một người. Khi đó nên bổ sung Mã nhân viên vào nguồn dữ liệu để khóa chính xác.

## Cài đặt

1. Mở một workbook macro-enabled dùng cho HR.
2. Alt + F11 -> Insert -> Module.
3. Import file AnnualLeaveSync.bas.
4. Mở Google Sheet mới.
5. Extensions -> Apps Script.
6. Dán GoogleAppsScript.gs.
7. Chạy setupSheet() một lần.
8. Deploy -> New deployment -> Web app.
9. Execute as: Me.
10. Who has access: Anyone with the link.
11. Copy Web App URL.
12. Dán URL vào WEB_APP_URL trong VBA.
13. Dán cùng URL vào API_URL trong index.html.
14. Upload index.html lên GitHub Pages.

## Tự động mỗi ngày

Sau khi chạy thủ công thành công, có thể dùng Windows Task Scheduler:

- Mở Excel
- Mở workbook macro-enabled chứa macro
- Chạy SyncAnnualLeaveToGoogle
- Đóng Excel

Nên đặt khoảng 06:00 mỗi ngày.

## Quan trọng

Không đưa mật khẩu Excel 2000 lên GitHub.

Mật khẩu chỉ nằm trong macro trên máy HR.

Google Sheets chỉ lưu dữ liệu cần thiết để tra cứu.
