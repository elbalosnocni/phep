ANNUAL LEAVE LOOKUP v4
=======================

Mục tiêu:
- Công nhân tra cứu bằng CCCD 12 số, kể cả số 0 ở đầu.
- Dữ liệu phép vẫn lấy từ Npn2023.xlsb, sheet 2026 và 2026 PL.
- Thông tin Phòng ban/Bộ phận/CCCD lấy từ DSCNV-2023.xlsb, sheet DSCNV, dò theo tên.
- LeaveData được mở rộng từ A:S thành A:V.

Cột DSCNV:
- H = CCCD
- 31 = AE = Phòng ban
- 32 = AF = Bộ phận

Nguồn:
\\192.168.0.253\vn hr\DS + PN + TP - 2014\Npn2023.xlsb
\\192.168.0.253\vn hr\DS + PN + TP - 2014\DSCNV-2023.xlsb [DSCNV]

Cài đặt:
1. Thay toàn bộ code Apps Script bằng AnnualLeave.gs.
2. Deploy Web App, lấy URL /exec. Nếu URL thay đổi, cập nhật API_URL trong index.html và WEB_APP_URL trong AnnualLeaveSync.bas.
3. Thay module VBA AnnualLeaveSync.bas.
4. Run macro SyncAnnualLeaveToGoogle.
5. Upload index.html lên GitHub Pages/Cloudflare Pages.
6. RunAnnualLeaveSync.vbs vẫn gọi workbook AnnualLeaveSync.xlsm như hiện tại.

LƯU Ý CCCD:
- VBA đọc c.Text trước để giữ định dạng hiển thị 12 số.
- Nếu ô là số, VBA dùng format 000000000000 để khôi phục số 0 đầu.
- Khuyến nghị cột H trong DSCNV được đặt định dạng Text hoặc 000000000000.
- Nếu Excel đã lưu CCCD dạng số và đã làm mất số 0 đầu, code chỉ có thể khôi phục đúng khi biết CCCD luôn đủ 12 số.

Kiến trúc:
DSCNV-2023.xlsb --(dò tên)--> CCCD + Phòng ban + Bộ phận
Npn2023.xlsb  --(tên)---------> Phép + Xưởng + tháng
                                  |
                                  v
                             Google Sheet
                               LeaveData
                                  |
                                  v
                         Google Apps Script
                                  |
                                  v
                         phepnam.pages.dev
