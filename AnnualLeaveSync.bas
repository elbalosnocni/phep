Option Explicit

' ============================================================
' ANNUAL LEAVE SYNC - FINAL + MONTHLY LEAVE
' ============================================================

Private Const SOURCE_FILE As String = "\\192.168.0.253\vn hr\DS + PN + TP - 2014\Npn2023.xlsb"
Private Const SOURCE_PASSWORD As String = "2000"
Private Const EMPLOYEE_SOURCE_FILE As String = "\\192.168.0.253\vn hr\DS + PN + TP - 2014\DSCNV-23.xlsb"
Private Const EMPLOYEE_SOURCE_PASSWORD As String = "DSCNV"
Private Const EMPLOYEE_SHEET As String = "DSCNV"
Private Const WEB_APP_URL As String = "https://script.google.com/macros/s/AKfycbxUSCx1x8scN0Xq-3ec-KcDop9bb-AZy8iH9TJyDWgPhmMYr14smVac2MB0GD2L1TRkQw/exec"
Private Const FIRST_DATA_ROW As Long = 3

' Source columns
Private Const COL_NAME As Long = 1        ' A
Private Const COL_MONTH As Long = 3       ' C
Private Const COL_PAID As Long = 25       ' Y
Private Const COL_LEFT_2025 As Long = 26  ' Z
Private Const COL_LEFT_2026 As Long = 27  ' AA
Private Const COL_GROSS As Long = 30      ' AD
Private Const COL_CITIZEN_ID As Long = 8  ' H
Private Const COL_DEPARTMENT As Long = 31 ' AE
Private Const COL_SECTION As Long = 32    ' AF
Private Const COL_MASTER_NAME As Long = 2 ' B - Họ tên trong DSCNV

' ============================================================
' MAIN SYNC
' ============================================================

Public Sub SyncAnnualLeaveToGoogle()

    Dim wb As Workbook
    Dim openedByMacro As Boolean

    Dim yearText As String
    Dim sheet1 As String
    Dim sheet2 As String

    Dim dict As Object
    Dim masterDict As Object
    Dim ws As Worksheet

    Dim payload As String
    Dim updated As String

    Dim responseText As String
    Dim httpStatus As Long
    Dim employeeCount As Long

    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Application.StatusBar = "Annual Leave Sync: starting..."

    If Len(Trim$(WEB_APP_URL)) = 0 Then
        Err.Raise vbObjectError + 100, , "WEB_APP_URL is empty."
    End If

    If InStr(1, WEB_APP_URL, "PASTE_", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 101, , "WEB_APP_URL has not been configured."
    End If

    yearText = CStr(Year(Date))

    sheet1 = yearText
    sheet2 = yearText & " PL"

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ' --------------------------------------------------------
    ' OPEN SOURCE FILE
    ' --------------------------------------------------------

    Set wb = GetOpenWorkbookByFullName(SOURCE_FILE)

    If wb Is Nothing Then
        If Dir(SOURCE_FILE) = "" Then
            Err.Raise vbObjectError + 102, , "Cannot find source file:" & vbCrLf & SOURCE_FILE
        End If

        Set wb = Workbooks.Open( _
            Filename:=SOURCE_FILE, _
            Password:=SOURCE_PASSWORD, _
            ReadOnly:=True, _
            UpdateLinks:=False, _
            Notify:=False, _
            AddToMru:=False)

        openedByMacro = True
    End If

    ' --------------------------------------------------------
    ' READ SHEET 2026
    ' --------------------------------------------------------

    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(sheet1)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 103, , "Cannot find sheet: " & sheet1
    End If

    Application.StatusBar = "Reading " & sheet1 & " -> XUONG BANH..."
    ReadFactorySheet ws, dict, "XUONG BANH"

    ' --------------------------------------------------------
    ' READ SHEET 2026 PL
    ' --------------------------------------------------------

    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(sheet2)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 104, , "Cannot find sheet: " & sheet2
    End If

    Application.StatusBar = "Reading " & sheet2 & " -> XUONG IN..."
    ReadFactorySheet ws, dict, "XUONG IN"

    employeeCount = dict.Count

    If employeeCount = 0 Then
        Err.Raise vbObjectError + 105, , "No employee data found in " & sheet1 & " and " & sheet2
    End If

    ' --------------------------------------------------------
    ' READ EMPLOYEE MASTER: DSCNV-23.xlsb / DSCNV
    ' B = EmployeeName, H = Citizen ID, AE = Department, AF = Section
    ' Match by normalized employee name, but keep original display name.
    ' --------------------------------------------------------
    Set masterDict = CreateObject("Scripting.Dictionary")
    masterDict.CompareMode = vbTextCompare

    Application.StatusBar = "Reading employee master DSCNV..." 
    ReadEmployeeMaster masterDict
    MergeEmployeeMaster dict, masterDict

    ' --------------------------------------------------------
    ' UPDATE TIME & BUILD PAYLOAD
    ' --------------------------------------------------------

    updated = Format(Now, "dd/mm/yyyy HH:nn:ss")
    payload = BuildPayload(dict, updated, SOURCE_FILE, sheet1, sheet2)

    ' --------------------------------------------------------
    ' SEND GOOGLE
    ' --------------------------------------------------------

    Application.StatusBar = "Sending data to Google Sheets..."
    PostJson WEB_APP_URL, payload, responseText, httpStatus

    ' --------------------------------------------------------
    ' CLOSE SOURCE
    ' --------------------------------------------------------

    If openedByMacro Then
        wb.Close SaveChanges:=False
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox "DONG BO THANH CONG" & vbCrLf & vbCrLf & _
           "Sheets: " & sheet1 & " + " & sheet2 & vbCrLf & _
           "Nhan vien: " & employeeCount & vbCrLf & _
           "Xuong: Banh / In" & vbCrLf & _
           "Da dong bo phep theo thang: Jan-Dec" & vbCrLf & _
           "Da doi chieu DSCNV: CitizenID + Department + Section" & vbCrLf & _
           "So lieu: 3 chu so thap phan" & vbCrLf & _
           "Cap nhat (VN): " & updated & vbCrLf & _
           "HTTP: " & httpStatus, _
           vbInformation, "Annual Leave Sync"

    Exit Sub

EH:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next

    If openedByMacro Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox "DONG BO THAT BAI" & vbCrLf & vbCrLf & _
           "Error " & errNumber & vbCrLf & _
           errDescription & vbCrLf & vbCrLf & _
           IIf(Len(responseText) > 0, "Google response:" & vbCrLf & responseText, ""), _
           vbCritical, "Annual Leave Sync"

End Sub

' ============================================================
' TEST GOOGLE
' ============================================================

Public Sub TestGoogleConnection()

    Dim responseText As String
    Dim httpStatus As Long

    On Error GoTo EH

    GetGoogle WEB_APP_URL, responseText, httpStatus

    Application.StatusBar = False

    MsgBox "KET NOI GOOGLE THANH CONG" & vbCrLf & vbCrLf & _
           "HTTP: " & httpStatus & vbCrLf & vbCrLf & responseText, _
           vbInformation, "Google Connection Test"

    Exit Sub

EH:
    Application.StatusBar = False

    MsgBox "KET NOI GOOGLE THAT BAI" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & vbCrLf & Err.Description, _
           vbCritical, "Google Connection Test"

End Sub

' ============================================================
' READ FACTORY SHEET
' ============================================================

Private Sub ReadFactorySheet( _
    ByVal ws As Worksheet, _
    ByVal dict As Object, _
    ByVal factoryName As String)

    Dim lastRow As Long
    Dim r As Long

    Dim employeeName As String
    Dim employeeKey As String
    Dim monthName As String

    Dim item As Variant
    Dim mIndex As Long
    Dim paidValue As Double

    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row

    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow

        employeeName = NormalizeEmployeeName(CStr(ws.Cells(r, COL_NAME).Value))
        employeeKey = NormalizeEmployeeKey(employeeName)

        If Len(employeeKey) > 0 Then

            If Not dict.Exists(employeeKey) Then
                dict.Add employeeKey, Array( _
                    employeeName, 0#, Empty, Empty, Empty, factoryName, _
                    0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, _
                    "", "", "")
            Else
                item = dict(employeeKey)

                ' Giữ nguyên họ tên đúng như file Npn2023, không lowercase.
                ' Nếu cùng nhân viên xuất hiện ở 2 xưởng thì ghép xưởng.
                If InStr(1, CStr(item(5)), factoryName, vbTextCompare) = 0 Then
                    item(5) = CStr(item(5)) & " + " & factoryName
                End If

                dict(employeeKey) = item
            End If

            item = dict(employeeKey)

            ' Paid leave total
            If IsNumericCell(ws.Cells(r, COL_PAID).Value) Then
                paidValue = CDbl(ws.Cells(r, COL_PAID).Value)
                item(1) = Round(CDbl(item(1)) + paidValue, 3)
            End If

            ' Monthly leave
            monthName = CleanMonthName(CStr(ws.Cells(r, COL_MONTH).Value))
            mIndex = MonthIndex(monthName)

            If mIndex > 0 Then
                If IsNumericCell(ws.Cells(r, COL_PAID).Value) Then
                    paidValue = CDbl(ws.Cells(r, COL_PAID).Value)
                    item(5 + mIndex) = Round(CDbl(item(5 + mIndex)) + paidValue, 3)
                End If
            End If

            ' Left 2025
            If IsEmptyValue(item(2)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2025).Value) Then
                    item(2) = Round(CDbl(ws.Cells(r, COL_LEFT_2025).Value), 3)
                End If
            End If

            ' Left 2026
            If IsEmptyValue(item(3)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2026).Value) Then
                    item(3) = Round(CDbl(ws.Cells(r, COL_LEFT_2026).Value), 3)
                End If
            End If

            ' Remaining leave
            If IsEmptyValue(item(4)) Then
                If IsNumericCell(ws.Cells(r, COL_GROSS).Value) Then
                    item(4) = Round(CDbl(ws.Cells(r, COL_GROSS).Value), 3)
                End If
            End If

            dict(employeeKey) = item

        End If

    Next r

End Sub

' ============================================================
' READ EMPLOYEE MASTER
' ============================================================

Private Sub ReadEmployeeMaster(ByVal masterDict As Object)

    Dim wbEmp As Workbook
    Dim wsEmp As Worksheet
    Dim openedByMacro As Boolean

    Dim lastRow As Long
    Dim r As Long

    Dim displayName As String
    Dim employeeKey As String
    Dim citizenId As String
    Dim department As String
    Dim sectionName As String

    On Error GoTo EH

    Set wbEmp = GetOpenWorkbookByFullName(EMPLOYEE_SOURCE_FILE)

    If wbEmp Is Nothing Then
        If Dir(EMPLOYEE_SOURCE_FILE) = "" Then
            Err.Raise vbObjectError + 500, , _
                "Cannot find employee master file:" & vbCrLf & EMPLOYEE_SOURCE_FILE
        End If

        Set wbEmp = Workbooks.Open( _
            Filename:=EMPLOYEE_SOURCE_FILE, _
            Password:=EMPLOYEE_SOURCE_PASSWORD, _
            ReadOnly:=True, _
            UpdateLinks:=False, _
            Notify:=False, _
            AddToMru:=False)

        openedByMacro = True
    End If

    Set wsEmp = Nothing

    On Error Resume Next
    Set wsEmp = wbEmp.Worksheets(EMPLOYEE_SHEET)
    On Error GoTo EH

    If wsEmp Is Nothing Then
        Err.Raise vbObjectError + 501, , _
            "Cannot find sheet '" & EMPLOYEE_SHEET & "' in:" & vbCrLf & EMPLOYEE_SOURCE_FILE
    End If

    lastRow = wsEmp.Cells(wsEmp.Rows.Count, COL_MASTER_NAME).End(xlUp).Row

    For r = FIRST_DATA_ROW To lastRow

        displayName = NormalizeEmployeeName(CStr(wsEmp.Cells(r, COL_MASTER_NAME).Value))
        employeeKey = NormalizeEmployeeKey(displayName)

        If Len(employeeKey) > 0 Then

            citizenId = ReadCitizenId(wsEmp.Cells(r, COL_CITIZEN_ID))
            department = CleanText(CStr(wsEmp.Cells(r, COL_DEPARTMENT).Value))
            sectionName = CleanText(CStr(wsEmp.Cells(r, COL_SECTION).Value))

            ' Array: CitizenID, Department, Section
            masterDict(employeeKey) = Array(citizenId, department, sectionName)

        End If

    Next r

    If openedByMacro Then
        wbEmp.Close SaveChanges:=False
    End If

    Exit Sub

EH:
    On Error Resume Next
    If openedByMacro Then
        If Not wbEmp Is Nothing Then wbEmp.Close SaveChanges:=False
    End If

    Err.Raise Err.Number, , Err.Description

End Sub

' ============================================================
' MERGE MASTER DATA INTO LEAVE DATA
' ============================================================

Private Sub MergeEmployeeMaster(ByVal dict As Object, ByVal masterDict As Object)

    Dim key As Variant
    Dim item As Variant
    Dim master As Variant

    Dim matched As Long
    Dim unmatched As Long

    For Each key In dict.Keys

        item = dict(key)

        If masterDict.Exists(CStr(key)) Then

            master = masterDict(CStr(key))

            ' item(18) = CitizenID
            ' item(19) = Department
            ' item(20) = Section
            item(18) = CStr(master(0))
            item(19) = CStr(master(1))
            item(20) = CStr(master(2))

            matched = matched + 1

        Else

            unmatched = unmatched + 1

        End If

        dict(key) = item

    Next key

    Application.StatusBar = _
        "Matched employee master: " & matched & _
        " | Not matched: " & unmatched

End Sub

' ============================================================
' HELPERS
' ============================================================

Private Function NormalizeEmployeeKey(ByVal s As String) As String
    ' Chỉ dùng để dò tên.
    ' Không dùng giá trị này để hiển thị EmployeeName.
    s = CleanText(s)
    s = LCase$(s)

    ' Loại một số dấu tiếng Việt để tăng khả năng match giữa 2 file.
    s = Replace(s, "à", "a"): s = Replace(s, "á", "a")
    s = Replace(s, "ạ", "a"): s = Replace(s, "ả", "a")
    s = Replace(s, "ã", "a"): s = Replace(s, "ă", "a")
    s = Replace(s, "ằ", "a"): s = Replace(s, "ắ", "a")
    s = Replace(s, "ặ", "a"): s = Replace(s, "ẳ", "a")
    s = Replace(s, "ẵ", "a"): s = Replace(s, "â", "a")
    s = Replace(s, "ầ", "a"): s = Replace(s, "ấ", "a")
    s = Replace(s, "ậ", "a"): s = Replace(s, "ẩ", "a")
    s = Replace(s, "ẫ", "a")

    s = Replace(s, "è", "e"): s = Replace(s, "é", "e")
    s = Replace(s, "ẹ", "e"): s = Replace(s, "ẻ", "e")
    s = Replace(s, "ẽ", "e"): s = Replace(s, "ê", "e")
    s = Replace(s, "ề", "e"): s = Replace(s, "ế", "e")
    s = Replace(s, "ệ", "e"): s = Replace(s, "ể", "e")
    s = Replace(s, "ễ", "e")

    s = Replace(s, "ì", "i"): s = Replace(s, "í", "i")
    s = Replace(s, "ị", "i"): s = Replace(s, "ỉ", "i")
    s = Replace(s, "ĩ", "i")

    s = Replace(s, "ò", "o"): s = Replace(s, "ó", "o")
    s = Replace(s, "ọ", "o"): s = Replace(s, "ỏ", "o")
    s = Replace(s, "õ", "o"): s = Replace(s, "ô", "o")
    s = Replace(s, "ồ", "o"): s = Replace(s, "ố", "o")
    s = Replace(s, "ộ", "o"): s = Replace(s, "ổ", "o")
    s = Replace(s, "ỗ", "o"): s = Replace(s, "ơ", "o")
    s = Replace(s, "ờ", "o"): s = Replace(s, "ớ", "o")
    s = Replace(s, "ợ", "o"): s = Replace(s, "ở", "o")
    s = Replace(s, "ỡ", "o")

    s = Replace(s, "ù", "u"): s = Replace(s, "ú", "u")
    s = Replace(s, "ụ", "u"): s = Replace(s, "ủ", "u")
    s = Replace(s, "ũ", "u"): s = Replace(s, "ư", "u")
    s = Replace(s, "ừ", "u"): s = Replace(s, "ứ", "u")
    s = Replace(s, "ự", "u"): s = Replace(s, "ử", "u")
    s = Replace(s, "ữ", "u")

    s = Replace(s, "ỳ", "y"): s = Replace(s, "ý", "y")
    s = Replace(s, "ỵ", "y"): s = Replace(s, "ỷ", "y")
    s = Replace(s, "ỹ", "y")

    s = Replace(s, "đ", "d")

    NormalizeEmployeeKey = s
End Function

Private Function CleanText(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CleanText = Trim$(s)
End Function

Private Function ReadCitizenId(ByVal cell As Range) As String
    Dim s As String
    Dim v As Variant

    ' .Text ưu tiên để giữ số 0 đầu nếu Excel đang hiển thị 12 chữ số.
    On Error Resume Next
    s = Trim$(CStr(cell.Text))
    On Error GoTo 0

    s = Replace(s, " ", "")

    If Len(s) = 12 And IsAllDigits(s) Then
        ReadCitizenId = s
        Exit Function
    End If

    v = cell.Value2

    If IsNumeric(v) Then
        ' CCCD luôn 12 số.
        ReadCitizenId = Format$(CDbl(v), "000000000000")
        Exit Function
    End If

    s = Trim$(CStr(v))
    s = Replace(s, " ", "")

    If Len(s) > 0 And IsAllDigits(s) Then
        ReadCitizenId = Right$("000000000000" & s, 12)
    Else
        ReadCitizenId = ""
    End If
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    IsAllDigits = True
End Function

' ============================================================
' HELPER FUNCTIONS FOR STRING & MONTH
' ============================================================

Private Function CleanMonthName(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    CleanMonthName = LCase$(Trim$(s))
End Function

Private Function MonthIndex(ByVal monthName As String) As Long
    Select Case LCase$(Trim$(monthName))
        Case "jan", "january": MonthIndex = 1
        Case "feb", "february": MonthIndex = 2
        Case "mar", "march": MonthIndex = 3
        Case "apr", "april": MonthIndex = 4
        Case "may": MonthIndex = 5
        Case "jun", "june": MonthIndex = 6
        Case "jul", "july": MonthIndex = 7
        Case "aug", "august": MonthIndex = 8
        Case "sep", "sept", "september": MonthIndex = 9
        Case "oct", "october": MonthIndex = 10
        Case "nov", "november": MonthIndex = 11
        Case "dec", "december": MonthIndex = 12
        Case Else: MonthIndex = 0
    End Select
End Function

Private Function NormalizeEmployeeName(ByVal s As String) As String
    ' Dùng cho tên HIỂN THỊ: giữ nguyên hoa/thường như file nguồn.
    NormalizeEmployeeName = CleanText(s)
End Function

' ============================================================
' BUILD JSON PAYLOAD
' ============================================================

Private Function BuildPayload( _
    ByVal dict As Object, _
    ByVal updatedAt As String, _
    ByVal source As String, _
    ByVal sheet1 As String, _
    ByVal sheet2 As String) As String

    Dim key As Variant
    Dim item As Variant
    Dim employees As String

    employees = "["

    For Each key In dict.Keys
        item = dict(key)

        If Len(employees) > 1 Then employees = employees & ","

        employees = employees & _
            "{""employeeName"":""" & JsonEscape(CStr(item(0))) & """," & _
            """paidLeave"":" & NumberOrNull(item(1)) & "," & _
            """left2025"":" & NumberOrNull(item(2)) & "," & _
            """left2026"":" & NumberOrNull(item(3)) & "," & _
            """remainingLeave"":" & NumberOrNull(item(4)) & "," & _
            """factories"":""" & JsonEscape(CStr(item(5))) & """," & _
            """department"":""" & JsonEscape(CStr(item(19))) & """," & _
            """section"":""" & JsonEscape(CStr(item(20))) & """," & _
            """citizenId"":""" & JsonEscape(CStr(item(18))) & """," & _
            """jan"":" & NumberOrNull(item(6)) & "," & _
            """feb"":" & NumberOrNull(item(7)) & "," & _
            """mar"":" & NumberOrNull(item(8)) & "," & _
            """apr"":" & NumberOrNull(item(9)) & "," & _
            """may"":" & NumberOrNull(item(10)) & "," & _
            """jun"":" & NumberOrNull(item(11)) & "," & _
            """jul"":" & NumberOrNull(item(12)) & "," & _
            """aug"":" & NumberOrNull(item(13)) & "," & _
            """sep"":" & NumberOrNull(item(14)) & "," & _
            """oct"":" & NumberOrNull(item(15)) & "," & _
            """nov"":" & NumberOrNull(item(16)) & "," & _
            """dec"":" & NumberOrNull(item(17)) & "}"
    Next key

    employees = employees & "]"

    BuildPayload = _
        "{""updatedAt"":""" & JsonEscape(updatedAt) & """," & _
        """source"":""" & JsonEscape(source) & """," & _
        """sheets"":""" & JsonEscape(sheet1 & " + " & sheet2) & """," & _
        """employees"":" & employees & "}"

End Function

' ============================================================
' OTHER HELPER FUNCTIONS
' ============================================================

Private Function NumberOrNull(ByVal v As Variant) As String
    Dim n As Double
    Dim s As String

    If IsEmptyValue(v) Or Not IsNumeric(v) Then
        NumberOrNull = "null"
        Exit Function
    End If

    n = Round(CDbl(v), 3)
    s = Format$(n, "0.000")
    s = Replace(s, Application.International(xlDecimalSeparator), ".")
    NumberOrNull = s
End Function

Private Function IsEmptyValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        IsEmptyValue = True
    ElseIf VarType(v) = vbString Then
        IsEmptyValue = (Trim$(CStr(v)) = "")
    Else
        IsEmptyValue = False
    End If
End Function

Private Function IsNumericCell(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If Trim$(CStr(v)) = "" Or Trim$(CStr(v)) = "-" Then Exit Function
    IsNumericCell = IsNumeric(v)
End Function

Private Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    JsonEscape = s
End Function

Private Function GetOpenWorkbookByFullName(ByVal fullName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.fullName, fullName, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

' ============================================================
' HTTP REQUESTS (UPDATED FOR GOOGLE APPS SCRIPT REDIRECTS)
' ============================================================

Private Sub PostJson( _
    ByVal url As String, _
    ByVal body As String, _
    ByRef responseText As String, _
    ByRef httpStatus As Long)

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.Option(6) = True ' WinHttpRequestOption_EnableRedirects = True
    http.SetTimeouts 30000, 30000, 30000, 120000
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "User-Agent", "AnnualLeaveSync/2.0"

    http.Send body

    httpStatus = http.Status
    responseText = http.responseText

    If httpStatus < 200 Or httpStatus >= 400 Then
        Err.Raise vbObjectError + 300, , _
            "Google Apps Script returned HTTP " & httpStatus & "." & vbCrLf & vbCrLf & responseText
    End If
End Sub

Private Sub GetGoogle( _
    ByVal url As String, _
    ByRef responseText As String, _
    ByRef httpStatus As Long)

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False
    http.Option(6) = True ' WinHttpRequestOption_EnableRedirects = True
    http.SetTimeouts 30000, 30000, 30000, 30000
    http.SetRequestHeader "User-Agent", "AnnualLeaveSync/2.0"

    http.Send

    httpStatus = http.Status
    responseText = http.responseText

    If httpStatus < 200 Or httpStatus >= 400 Then
        Err.Raise vbObjectError + 400, , "HTTP " & httpStatus & "." & vbCrLf & responseText
    End If
End Sub

