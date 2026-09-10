ANNUAL LEAVE - AUTO DAILY v3

1. index.html: hien thi toi da 3 so thap phan, khong them so 0.
   7 -> 7
   7.5 -> 7.5
   8.167 -> 8.167

2. Dat AnnualLeaveSync.xlsm tai:
   X:\DS + PN + TP - 2014\AnnualLeaveSync.xlsm

3. Macro trong workbook phai co:
   Public Sub SyncAnnualLeaveToGoogle

4. Chay Test_RunAnnualLeaveSync.bat de test.

5. Sau khi test thanh cong, tao Windows Task Scheduler chay RunAnnualLeaveSync.vbs moi ngay, de xuat 06:00.

6. LUU Y: neu X: la network drive, Task Scheduler co the khong thay X:. Khi do dung UNC path thay cho X: trong VBS. Mo Properties cua o X: de xem Network path.

7. Trinh tu:
   Npn2023.xlsb -> VBA -> Google Sheets -> GitHub Pages.


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

ANNUAL LEAVE - FIX DSCNV

1. VBA source:
   \\192.168.0.253\vn hr\DS + PN + TP - 2014\DSCNV-23.xlsb
   Sheet: DSCNV

2. DSCNV columns:
   H  = Citizen ID / CCCD
   AF = Department / Phòng ban
   AG = Section / Bộ phận

3. Matching:
   Match by employee name after normalizing spaces, case and Vietnamese accents.
   The EmployeeName displayed/sent to Google is preserved exactly from Npn2023,
   so it is NOT converted to lowercase.

6. IMPORTANT:
   The name column in DSCNV is assumed to be column A, same as Npn2023.
   If the employee name in DSCNV is in another column, change COL_NAME accordingly
   or create a separate COL_EMPLOYEE_NAME_MASTER constant.

4. DSCNV NAME COLUMN:
   Họ tên trong DSCNV-23.xlsb nằm ở cột B. Code đã dùng riêng COL_MASTER_NAME = 2.

7. CCCD:
   The VBA reads .Text first and formats numeric values as 12 digits.
   If Excel has already lost a leading zero and has no formatting information,
   the missing zero cannot be known with certainty; best is to keep H as Text
   or custom format 000000000000.

6. GAS:
   LeaveData becomes A:V:
   A EmployeeName
   B PaidLeave
   C Left2025
   D Left2026
   E RemainingLeave
   F Factories
   G Department
   H Section
   I CitizenID
   J UpdatedAt
   K:V Jan-Dec

7. Định dạng Google Sheet:
   - I = CitizenID: Plain text (@), để giữ số 0 đầu.
   - J = UpdatedAt: dd/MM/yyyy HH:mm:ss.
   - K:V = Jan-Dec dạng số 0.000.
