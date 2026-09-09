Option Explicit

' ============================================================
' ANNUAL LEAVE SYNC - v4
' ============================================================
' Nguon 1: phep nam
Private Const LEAVE_SOURCE_FILE As String = "\\192.168.0.253\vn hr\DS + PN + TP - 2014\Npn2023.xlsb"
Private Const LEAVE_SOURCE_PASSWORD As String = "2000"

' Nguon 2: thong tin nhan vien
Private Const EMPLOYEE_SOURCE_FILE As String = "\\192.168.0.253\vn hr\DS + PN + TP - 2014\DSCNV-2023.xlsb"
Private Const EMPLOYEE_SOURCE_PASSWORD As String = ""

Private Const EMPLOYEE_SHEET As String = "DSCNV"

Private Const WEB_APP_URL As String = "https://script.google.com/macros/s/AKfycbxUSCx1x8scN0Xq-3ec-KcDop9bb-AZy8iH9TJyDWgPhmMYr14smVac2MB0GD2L1TRkQw/exec"
Private Const FIRST_DATA_ROW As Long = 3

' Npn2023 columns
Private Const COL_NAME As Long = 1        ' A
Private Const COL_MONTH As Long = 3       ' C
Private Const COL_PAID As Long = 25       ' Y
Private Const COL_LEFT_2025 As Long = 26  ' Z
Private Const COL_LEFT_2026 As Long = 27  ' AA
Private Const COL_GROSS As Long = 30      ' AD

' DSCNV-2023 columns
Private Const COL_CITIZEN_ID As Long = 8  ' H
Private Const COL_DEPARTMENT As Long = 31 ' AE
Private Const COL_SECTION As Long = 32    ' AF

' ============================================================
' MAIN
' ============================================================
Public Sub SyncAnnualLeaveToGoogle()

    Dim wbLeave As Workbook, wbEmp As Workbook
    Dim openedLeave As Boolean, openedEmp As Boolean
    Dim yearText As String, sheet1 As String, sheet2 As String
    Dim dict As Object, empDict As Object
    Dim ws As Worksheet
    Dim payload As String, updated As String
    Dim responseText As String, httpStatus As Long
    Dim employeeCount As Long
    Dim errNumber As Long, errDescription As String

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.StatusBar = "Annual Leave Sync: starting..."

    If Len(Trim$(WEB_APP_URL)) = 0 Then
        Err.Raise vbObjectError + 100, , "WEB_APP_URL is empty."
    End If

    yearText = CStr(Year(Date))
    sheet1 = yearText
    sheet2 = yearText & " PL"

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set empDict = CreateObject("Scripting.Dictionary")
    empDict.CompareMode = vbTextCompare

    ' --------------------------------------------------------
    ' OPEN Npn2023.xlsb
    ' --------------------------------------------------------
    Set wbLeave = GetOpenWorkbookByFullName(LEAVE_SOURCE_FILE)

    If wbLeave Is Nothing Then
        If Dir(LEAVE_SOURCE_FILE) = "" Then
            Err.Raise vbObjectError + 102, , "Cannot find leave source file:" & vbCrLf & LEAVE_SOURCE_FILE
        End If

        Set wbLeave = Workbooks.Open( _
            Filename:=LEAVE_SOURCE_FILE, _
            Password:=LEAVE_SOURCE_PASSWORD, _
            ReadOnly:=True, _
            UpdateLinks:=False, _
            Notify:=False, _
            AddToMru:=False)

        openedLeave = True
    End If

    ' --------------------------------------------------------
    ' READ 2026 -> XUONG BANH
    ' --------------------------------------------------------
    Set ws = Nothing
    On Error Resume Next
    Set ws = wbLeave.Worksheets(sheet1)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 103, , "Cannot find sheet: " & sheet1
    End If

    Application.StatusBar = "Reading " & sheet1 & " -> XUONG BANH..."
    ReadFactorySheet ws, dict, "XUONG BANH"

    ' --------------------------------------------------------
    ' READ 2026 PL -> XUONG IN
    ' --------------------------------------------------------
    Set ws = Nothing
    On Error Resume Next
    Set ws = wbLeave.Worksheets(sheet2)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 104, , "Cannot find sheet: " & sheet2
    End If

    Application.StatusBar = "Reading " & sheet2 & " -> XUONG IN..."
    ReadFactorySheet ws, dict, "XUONG IN"

    If dict.Count = 0 Then
        Err.Raise vbObjectError + 105, , "No employee leave data found."
    End If

    ' --------------------------------------------------------
    ' OPEN DSCNV-2023.xlsb / DSCNV
    ' --------------------------------------------------------
    Application.StatusBar = "Reading DSCNV-2023.xlsb -> department / section / citizen ID..."

    Set wbEmp = GetOpenWorkbookByFullName(EMPLOYEE_SOURCE_FILE)

    If wbEmp Is Nothing Then
        If Dir(EMPLOYEE_SOURCE_FILE) = "" Then
            Err.Raise vbObjectError + 106, , "Cannot find employee source file:" & vbCrLf & EMPLOYEE_SOURCE_FILE
        End If

        Set wbEmp = Workbooks.Open( _
            Filename:=EMPLOYEE_SOURCE_FILE, _
            Password:=EMPLOYEE_SOURCE_PASSWORD, _
            ReadOnly:=True, _
            UpdateLinks:=False, _
            Notify:=False, _
            AddToMru:=False)

        openedEmp = True
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wbEmp.Worksheets(EMPLOYEE_SHEET)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 107, , "Cannot find sheet: " & EMPLOYEE_SHEET
    End If

    ReadEmployeeMaster ws, empDict

    ' --------------------------------------------------------
    ' MERGE MASTER INFO BY NAME
    ' --------------------------------------------------------
    MergeEmployeeMaster dict, empDict

    employeeCount = dict.Count
    updated = Format(Now, "dd/mm/yyyy HH:nn:ss")
    payload = BuildPayload(dict, updated, LEAVE_SOURCE_FILE, sheet1, sheet2)

    ' --------------------------------------------------------
    ' SEND GOOGLE
    ' --------------------------------------------------------
    Application.StatusBar = "Sending data to Google Sheets..."
    PostJson WEB_APP_URL, payload, responseText, httpStatus

    If openedLeave Then wbLeave.Close SaveChanges:=False
    If openedEmp Then wbEmp.Close SaveChanges:=False

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox "DONG BO THANH CONG" & vbCrLf & vbCrLf & _
           "Sheets phep: " & sheet1 & " + " & sheet2 & vbCrLf & _
           "Nhan vien: " & employeeCount & vbCrLf & _
           "Xuong: Banh / In" & vbCrLf & _
           "Master: " & EMPLOYEE_SOURCE_FILE & " [" & EMPLOYEE_SHEET & "]" & vbCrLf & _
           "CCCD: cot H" & vbCrLf & _
           "Phong ban: cot 31" & vbCrLf & _
           "Bo phan: cot 32" & vbCrLf & _
           "Da dong bo phep theo thang: Jan-Dec" & vbCrLf & _
           "HTTP: " & httpStatus, _
           vbInformation, "Annual Leave Sync"

    Exit Sub

EH:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If openedLeave Then
        If Not wbLeave Is Nothing Then wbLeave.Close SaveChanges:=False
    End If
    If openedEmp Then
        If Not wbEmp Is Nothing Then wbEmp.Close SaveChanges:=False
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

Public Sub TestGoogleConnection()
    Dim responseText As String
    Dim httpStatus As Long

    On Error GoTo EH
    GetGoogle WEB_APP_URL, responseText, httpStatus

    MsgBox "KET NOI GOOGLE THANH CONG" & vbCrLf & vbCrLf & _
           "HTTP: " & httpStatus & vbCrLf & vbCrLf & responseText, _
           vbInformation, "Google Connection Test"
    Exit Sub

EH:
    MsgBox "KET NOI GOOGLE THAT BAI" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & vbCrLf & Err.Description, _
           vbCritical, "Google Connection Test"
End Sub

' ============================================================
' READ LEAVE SHEETS
' ============================================================
Private Sub ReadFactorySheet(ByVal ws As Worksheet, ByVal dict As Object, ByVal factoryName As String)

    Dim lastRow As Long, r As Long
    Dim employeeName As String, monthName As String
    Dim item As Variant, mIndex As Long, paidValue As Double

    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow

        employeeName = NormalizeEmployeeName(CStr(ws.Cells(r, COL_NAME).Value))

        If Len(employeeName) > 0 Then

            If Not dict.Exists(employeeName) Then
                dict.Add employeeName, Array( _
                    employeeName, 0#, Empty, Empty, Empty, factoryName, _
                    "", "", "", "", _
                    0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
            Else
                item = dict(employeeName)
                If InStr(1, CStr(item(5)), factoryName, vbTextCompare) = 0 Then
                    If Len(Trim$(CStr(item(5)))) > 0 Then
                        item(5) = CStr(item(5)) & " + " & factoryName
                    Else
                        item(5) = factoryName
                    End If
                End If
                dict(employeeName) = item
            End If

            item = dict(employeeName)

            If IsNumericCell(ws.Cells(r, COL_PAID).Value) Then
                paidValue = CDbl(ws.Cells(r, COL_PAID).Value)
                item(1) = Round(CDbl(item(1)) + paidValue, 3)
            End If

            monthName = CleanMonthName(CStr(ws.Cells(r, COL_MONTH).Value))
            mIndex = MonthIndex(monthName)

            If mIndex > 0 Then
                If IsNumericCell(ws.Cells(r, COL_PAID).Value) Then
                    paidValue = CDbl(ws.Cells(r, COL_PAID).Value)
                    item(9 + mIndex) = Round(CDbl(item(9 + mIndex)) + paidValue, 3)
                End If
            End If

            If IsEmptyValue(item(2)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2025).Value) Then
                    item(2) = Round(CDbl(ws.Cells(r, COL_LEFT_2025).Value), 3)
                End If
            End If

            If IsEmptyValue(item(3)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2026).Value) Then
                    item(3) = Round(CDbl(ws.Cells(r, COL_LEFT_2026).Value), 3)
                End If
            End If

            If IsEmptyValue(item(4)) Then
                If IsNumericCell(ws.Cells(r, COL_GROSS).Value) Then
                    item(4) = Round(CDbl(ws.Cells(r, COL_GROSS).Value), 3)
                End If
            End If

            dict(employeeName) = item
        End If
    Next r
End Sub

' ============================================================
' READ DSCNV-2023.xlsb
' empDict(name) = Array(citizenId, department, section)
' ============================================================
Private Sub ReadEmployeeMaster(ByVal ws As Worksheet, ByVal empDict As Object)

    Dim lastRow As Long, r As Long
    Dim employeeName As String
    Dim citizenId As String, department As String, section As String

    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow

        employeeName = NormalizeEmployeeName(CStr(ws.Cells(r, COL_NAME).Value))

        If Len(employeeName) > 0 Then
            citizenId = ReadCitizenId(ws.Cells(r, COL_CITIZEN_ID))
            department = CleanText(CStr(ws.Cells(r, COL_DEPARTMENT).Value))
            section = CleanText(CStr(ws.Cells(r, COL_SECTION).Value))

            ' If duplicate name exists, keep the first complete record.
            If Not empDict.Exists(employeeName) Then
                empDict.Add employeeName, Array(citizenId, department, section)
            Else
                Dim oldItem As Variant
                oldItem = empDict(employeeName)

                If Len(CStr(oldItem(0))) = 0 And Len(citizenId) > 0 Then oldItem(0) = citizenId
                If Len(CStr(oldItem(1))) = 0 And Len(department) > 0 Then oldItem(1) = department
                If Len(CStr(oldItem(2))) = 0 And Len(section) > 0 Then oldItem(2) = section

                empDict(employeeName) = oldItem
            End If
        End If
    Next r
End Sub

Private Sub MergeEmployeeMaster(ByVal leaveDict As Object, ByVal empDict As Object)

    Dim key As Variant, item As Variant, master As Variant

    For Each key In leaveDict.Keys
        item = leaveDict(key)

        If empDict.Exists(key) Then
            master = empDict(key)
            item(6) = CStr(master(1)) ' Department
            item(7) = CStr(master(2)) ' Section
            item(8) = CStr(master(0)) ' Citizen ID
        End If

        leaveDict(key) = item
    Next key
End Sub

' ============================================================
' STRING / MONTH HELPERS
' ============================================================
Private Function CleanText(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = Trim$(s)
End Function

Private Function CleanMonthName(ByVal s As String) As String
    CleanMonthName = LCase$(Trim$(CleanText(s)))
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
    NormalizeEmployeeName = LCase$(CleanText(s))
End Function

' ============================================================
' CITIZEN ID - PRESERVE LEADING ZERO
' ============================================================
Private Function ReadCitizenId(ByVal c As Range) As String

    Dim s As String
    Dim v As Variant

    On Error GoTo SafeExit

    ' Prefer displayed text. This preserves custom number formats such as 000000000000.
    s = Trim$(CStr(c.Text))
    s = DigitsOnly(s)

    If Len(s) = 12 Then
        ReadCitizenId = s
        Exit Function
    End If

    v = c.Value2

    If IsNumeric(v) Then
        ' CCCD has 12 digits. If Excel stored it as numeric, restore leading zeroes.
        s = Format$(CDbl(v), "000000000000")
        s = DigitsOnly(s)
        If Len(s) = 12 Then
            ReadCitizenId = s
            Exit Function
        End If
    Else
        s = DigitsOnly(CStr(v))
        If Len(s) < 12 And Len(s) > 0 Then s = String$(12 - Len(s), "0") & s
        If Len(s) = 12 Then
            ReadCitizenId = s
            Exit Function
        End If
    End If

SafeExit:
    If Len(ReadCitizenId) = 0 Then
        ReadCitizenId = DigitsOnly(CStr(c.Value2))
        If Len(ReadCitizenId) < 12 And Len(ReadCitizenId) > 0 Then
            ReadCitizenId = String$(12 - Len(ReadCitizenId), "0") & ReadCitizenId
        End If
    End If
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, result As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then result = result & ch
    Next i

    DigitsOnly = result
End Function

' ============================================================
' BUILD JSON
' ============================================================
Private Function BuildPayload(ByVal dict As Object, ByVal updatedAt As String, _
                              ByVal source As String, ByVal sheet1 As String, _
                              ByVal sheet2 As String) As String

    Dim key As Variant, item As Variant, employees As String
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
            """department"":""" & JsonEscape(CStr(item(6))) & """," & _
            """section"":""" & JsonEscape(CStr(item(7))) & """," & _
            """citizenId"":""" & JsonEscape(CStr(item(8))) & """," & _
            """jan"":" & NumberOrNull(item(10)) & "," & _
            """feb"":" & NumberOrNull(item(11)) & "," & _
            """mar"":" & NumberOrNull(item(12)) & "," & _
            """apr"":" & NumberOrNull(item(13)) & "," & _
            """may"":" & NumberOrNull(item(14)) & "," & _
            """jun"":" & NumberOrNull(item(15)) & "," & _
            """jul"":" & NumberOrNull(item(16)) & "," & _
            """aug"":" & NumberOrNull(item(17)) & "," & _
            """sep"":" & NumberOrNull(item(18)) & "," & _
            """oct"":" & NumberOrNull(item(19)) & "," & _
            """nov"":" & NumberOrNull(item(20)) & "," & _
            """dec"":" & NumberOrNull(item(21)) & "}"
    Next key

    employees = employees & "]"

    BuildPayload = _
        "{""updatedAt"":""" & JsonEscape(updatedAt) & """," & _
        """source"":""" & JsonEscape(source) & """," & _
        """sheets"":""" & JsonEscape(sheet1 & " + " & sheet2) & """," & _
        """employees"":" & employees & "}"
End Function

Private Function NumberOrNull(ByVal v As Variant) As String
    Dim n As Double, s As String

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
        If StrComp(wb.FullName, fullName, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

' ============================================================
' HTTP
' ============================================================
Private Sub PostJson(ByVal url As String, ByVal body As String, _
                     ByRef responseText As String, ByRef httpStatus As Long)

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False
    http.Option(6) = True
    http.SetTimeouts 30000, 30000, 30000, 120000
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.SetRequestHeader "User-Agent", "AnnualLeaveSync/4.0"

    http.Send body

    httpStatus = http.Status
    responseText = http.ResponseText

    If httpStatus < 200 Or httpStatus >= 400 Then
        Err.Raise vbObjectError + 300, , _
            "Google Apps Script returned HTTP " & httpStatus & "." & vbCrLf & vbCrLf & responseText
    End If
End Sub

Private Sub GetGoogle(ByVal url As String, ByRef responseText As String, ByRef httpStatus As Long)

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False
    http.Option(6) = True
    http.SetTimeouts 30000, 30000, 30000, 30000
    http.SetRequestHeader "User-Agent", "AnnualLeaveSync/4.0"

    http.Send

    httpStatus = http.Status
    responseText = http.ResponseText

    If httpStatus < 200 Or httpStatus >= 400 Then
        Err.Raise vbObjectError + 400, , "HTTP " & httpStatus & "." & vbCrLf & responseText
    End If
End Sub
