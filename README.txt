ANNUAL LEAVE - FIX DSCNV

1. VBA source:
   \\192.168.0.253\vn hr\DS + PN + TP - 2014\DSCNV-23.xlsb
   Password: DSCNV
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
