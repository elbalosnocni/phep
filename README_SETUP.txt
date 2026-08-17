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
