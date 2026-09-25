Sub CopySpecificColumnsToGoogleSheets()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim folderPath As String, fileName As String
    Dim passwordExcel As String
    Dim lastRow As Long, i As Long
    Dim http As Object, url As String, payload As String
    Dim jsonRows As String
    
    ' === CAU HINH THONG TIN TAI DAY ===
    ' 1. BAN PHAI PHAI DIEN DIEN CHINH XAC LINK WEB APP GOOGLE APPS SCRIPT VAO DAY (KHONG DE GOOGLE.COM)
    url = "https://script.google.com/macros/s/AKfycbxtfVxsHIdtX6Z5nSBbGSM52aD7oTiIYc8Yt_jv1o8OihVVE2VURsKJJmDjAzLxH7X7/exec"
    passwordExcel = "DSCNV"
    folderPath = "\\192.168.0.253\vn hr\DS + PN + TP - 2014\"
    
    ' CAU HINH GOOGLE SHEET TACH RIENG TAI DAY:
    Dim googleSpreadsheetId As String
    Dim googleSheetName As String
    googleSpreadsheetId = "1caxiuh1jyzZi8rkz1EqDkzwcr_gf5sIk4XzCh-XBjjQ"
    googleSheetName = "DSCNV"
    ' ==================================
    
    ' Tìm file co ten bat dau bang DSCNV va duoi .xlsb
    fileName = Dir(folderPath & "DSCNV-*.xlsb")
    
    If fileName = "" Then
        MsgBox "Khong tim thay file nao bat dau bang 'DSCNV' trong thu muc!", vbCritical, "Loi"
        Exit Sub
    End If
    
    ' === KICH HOAT CHE DO AN THONG BAO VA TAT HOI LINK ===
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False ' Tat tat ca thong bao canh bao cua Excel
    Application.AskToUpdateLinks = False ' Tat thong bao bat cap nhat lien ket ngoai
    
    ' Mo file nguon: Bo qua cap nhat link (UpdateLinks:=0), Dien mat khau, Khong thong bao thieu link
    On Error Resume Next
    Set wbSource = Workbooks.Open( _
        fileName:=folderPath & fileName, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        password:=passwordExcel, _
        WriteResPassword:=passwordExcel, _
        IgnoreReadOnlyRecommended:=True)
    On Error GoTo 0
    
    If wbSource Is Nothing Then
        MsgBox "Khong the mo duoc file. Vui long kiem tra lai thu muc hoac mat khau!", vbCritical, "Loi"
        GoTo CleanExit
    End If
    
    ' Tro vao sheet "DSCNV"
    On Error Resume Next
    Set wsSource = wbSource.Sheets("DSCNV")
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Khong tim thay sheet 'DSCNV' trong file " & fileName, vbCritical, "Loi"
        wbSource.Close SaveChanges:=False
        GoTo CleanExit
    End If
    
    ' Khoi tao HTTP Request
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    lastRow = wsSource.Cells(wsSource.Rows.count, "A").End(xlUp).row
    
    ' Duyet qua tung dong de loc du lieu, xoa dau tieng Viet va gom vao JSON
    jsonRows = ""
    Dim colH_Value As String
    Dim validRowCount As Long
    validRowCount = 0
    
    For i = 8 To lastRow
        colH_Value = SafeCell(wsSource.Cells(i, 8))
        
        ' Kiem tra neu cot H co chua so can cuoc (khong bi trong)
        If Trim(colH_Value) <> "" Then
            validRowCount = validRowCount + 1
            
            Dim colB As String, colD As String, colE As String, colF As String, colG As String
            Dim colAF As String, colAG As String, colAH As String, colAM As String
            Dim colAO As String, colAP As String, colAQ As String, colAR As String, colAT As String, colAW As String
            
            colB = SafeCell(wsSource.Cells(i, 2))
            colD = SafeCell(wsSource.Cells(i, 4))
            colE = SafeCell(wsSource.Cells(i, 5))
            colF = SafeCell(wsSource.Cells(i, 6))
            colG = SafeCell(wsSource.Cells(i, 7))
            colH_Value = (colH_Value)
            colAF = SafeCell(wsSource.Cells(i, 32))
            colAG = SafeCell(wsSource.Cells(i, 33))
            colAH = SafeCell(wsSource.Cells(i, 34))
            colAM = SafeCell(wsSource.Cells(i, 38))
            colAO = SafeCell(wsSource.Cells(i, 41))
            colAP = SafeCell(wsSource.Cells(i, 42))
            colAQ = SafeCell(wsSource.Cells(i, 43))
            colAR = SafeCell(wsSource.Cells(i, 44))
            colAT = SafeCell(wsSource.Cells(i, 46))
            colAW = SafeCell(wsSource.Cells(i, 49))
            
            Dim rowJson As String
            rowJson = "[""" & colB & """,""" & colD & """,""" & colE & """,""" & colF & """,""" & colG & """,""" & colH_Value & """,""" & colAF & """,""" & colAG & """,""" & colAH & """,""" & colAM & """,""" & colAO & """,""" & colAP & """,""" & colAQ & """,""" & colAR & """,""" & colAT & """,""" & colAW & """]"
            
            If jsonRows = "" Then
                jsonRows = rowJson
            Else
                jsonRows = jsonRows & "," & rowJson
            End If
        End If
    Next i
    
    ' Dong file nguon ngay sau khi doc xong
    wbSource.Close SaveChanges:=False
    
    ' Kiem tra neu khong co dong nao hop le thi dung
    If validRowCount = 0 Then
        MsgBox "Khong tim thay dong nao co chua so can cuoc tai cot H!", vbExclamation, "Thong bao"
        GoTo CleanExit
    End If
    
    ' Tao Payload JSON
    payload = "{""spreadsheet_id"":""" & googleSpreadsheetId & """,""sheet_name"":""" & googleSheetName & """,""values"":[" & jsonRows & "]}"

    
    ' Gui Request len Google Sheets
    On Error Resume Next
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send payload
    
    If Err.Number <> 0 Then
        MsgBox "Loi ket noi Internet: " & Err.Description, vbCritical, "That Bai"
        GoTo CleanExit
    End If
    On Error GoTo 0
    
    ' Kiem tra phan hoi tu Google
    If http.Status = 200 Then
        Dim responseText As String
        responseText = http.responseText
        
        ' Kiem tra xem phan hoi có phai loi he thong dang HTML không hoac có chua chu thành công không
        If InStr(responseText, "status") > 0 And InStr(responseText, "success") > 0 Then
            MsgBox "Da loc va sao chep thanh cong " & validRowCount & " dong len Google Sheets!", vbInformation, "Hoan Tat"
        Else
            ' Neu phan hoi tra ve chuoi HTML loi hoac loi JSON tu script
            MsgBox "May chu Google tra ve loi hoac chua cau hinh dung hàm doGet:" & vbCrLf & _
                   Left(responseText, 300), vbExclamation, "Loi Ghi Du Lieu"
        End If
    Else
        MsgBox "Loi ket noi den may chu Google: " & http.Status & " - " & http.statusText, vbExclamation, "Loi Dong Bo"
    End If


CleanExit:
    ' KHOI PHUC LAI CAC CAU HINH HE THONG CUA EXCEL
    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
End Sub

Function SafeCell(rng As Range) As String
    If IsError(rng.Value) Then
        SafeCell = ""
    ElseIf IsEmpty(rng.Value) Then
        SafeCell = ""
    Else
        Dim txt As String
        txt = CStr(rng.Value)
        
        ' 1. X? lý d?u g?ch chéo ngu?c (ph?i thay th? d?u tiên d? tránh l?i ch?ng chéo)
        txt = Replace(txt, "\", "\\")
        
        ' 2. X? lý d?u ngo?c kép h?p l? cho JSON
        txt = Replace(txt, """", "\""")
        
        ' 3. X? lý ký t? xu?ng dòng (Alt + Enter trong Excel)
        txt = Replace(txt, vbCrLf, "\n")
        txt = Replace(txt, vbCr, "\n")
        txt = Replace(txt, vbLf, "\n")
        
        ' 4. X? lý ký t? Tab (n?u có)
        txt = Replace(txt, vbTab, "\t")
        
        SafeCell = txt
    End If
End Function



