Option Explicit

' ============================================================
' ANNUAL LEAVE SYNC - FINAL
' 2 FACTORIES:
'   Sheet 2026    -> XUONG BANH
'   Sheet 2026 PL -> XUONG IN
'
' SOURCE:
'   X:\DS + PN + TP - 2014\Npn2023.xlsb
' PASSWORD:
'   2000
'
' COLUMNS:
'   A  = HO VA TEN
'   Y  = Paid          -> SUM all monthly rows
'   Z  = Left 2025     -> first numeric value
'   AA = Left 2026     -> first numeric value
'   AD = Gross         -> first numeric value = remaining leave
'
' All numbers are rounded/sent with 3 decimals.
' UpdatedAt is Vietnam local time in dd/mm/yyyy HH:nn:ss format.
' ============================================================

Private Const SOURCE_FILE As String = "X:\DS + PN + TP - 2014\Npn2023.xlsb"
Private Const SOURCE_PASSWORD As String = "2000"

Private Const WEB_APP_URL As String = _
    "https://script.google.com/macros/s/AKfycbxUSCx1x8scN0Xq-3ec-KcDop9bb-AZy8iH9TJyDWgPhmMYr14smVac2MB0GD2L1TRkQw/exec"

Private Const FIRST_DATA_ROW As Long = 4

Private Const COL_NAME As Long = 1       ' A
Private Const COL_PAID As Long = 25      ' Y
Private Const COL_LEFT_2025 As Long = 26 ' Z
Private Const COL_LEFT_2026 As Long = 27 ' AA
Private Const COL_GROSS As Long = 30     ' AD

Public Sub SyncAnnualLeaveToGoogle()

    Dim wb As Workbook
    Dim openedByMacro As Boolean
    Dim yearText As String
    Dim sheet1 As String, sheet2 As String
    Dim dict As Object
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

    Set wb = GetOpenWorkbookByFullName(SOURCE_FILE)

    If wb Is Nothing Then

        If Dir(SOURCE_FILE) = "" Then
            Err.Raise vbObjectError + 102, , _
                "Cannot find source file:" & vbCrLf & SOURCE_FILE
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

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(sheet1)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 103, , _
            "Cannot find sheet: " & sheet1
    End If

    Application.StatusBar = "Reading " & sheet1 & " -> XUONG BANH..."
    ReadFactorySheet ws, dict, "XUONG BANH"

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(sheet2)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 104, , _
            "Cannot find sheet: " & sheet2
    End If

    Application.StatusBar = "Reading " & sheet2 & " -> XUONG IN..."
    ReadFactorySheet ws, dict, "XUONG IN"

    employeeCount = dict.Count

    If employeeCount = 0 Then
        Err.Raise vbObjectError + 105, , _
            "No employee data found in " & sheet1 & " and " & sheet2
    End If

    ' Vietnam local time on the HR PC
    updated = Format(Now, "dd/mm/yyyy HH:nn:ss")

    payload = BuildPayload(dict, updated, SOURCE_FILE, sheet1, sheet2)

    Application.StatusBar = "Sending data to Google Sheets..."
    PostJson WEB_APP_URL, payload, responseText, httpStatus

    If openedByMacro Then
        wb.Close SaveChanges:=False
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    MsgBox _
        "DONG BO THANH CONG" & vbCrLf & vbCrLf & _
        "Sheets: " & sheet1 & " + " & sheet2 & vbCrLf & _
        "Nhan vien: " & employeeCount & vbCrLf & _
        "Xuong: Banh / In" & vbCrLf & _
        "So lieu: 3 chu so thap phan" & vbCrLf & _
        "Cap nhat (VN): " & updated & vbCrLf & _
        "HTTP: " & httpStatus, _
        vbInformation, _
        "Annual Leave Sync"

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

    MsgBox _
        "DONG BO THAT BAI" & vbCrLf & vbCrLf & _
        "Error " & errNumber & vbCrLf & _
        errDescription & vbCrLf & vbCrLf & _
        IIf(Len(responseText) > 0, _
            "Google response:" & vbCrLf & responseText, _
            ""), _
        vbCritical, _
        "Annual Leave Sync"

End Sub

Public Sub TestGoogleConnection()

    Dim responseText As String
    Dim httpStatus As Long

    On Error GoTo EH

    GetGoogle WEB_APP_URL, responseText, httpStatus

    Application.StatusBar = False

    MsgBox _
        "KET NOI GOOGLE THANH CONG" & vbCrLf & vbCrLf & _
        "HTTP: " & httpStatus & vbCrLf & vbCrLf & _
        responseText, _
        vbInformation, _
        "Google Connection Test"

    Exit Sub

EH:
    Application.StatusBar = False

    MsgBox _
        "KET NOI GOOGLE THAT BAI" & vbCrLf & vbCrLf & _
        "Error " & Err.Number & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "Google Connection Test"

End Sub

Private Sub ReadFactorySheet( _
    ByVal ws As Worksheet, _
    ByVal dict As Object, _
    ByVal factoryName As String)

    Dim lastRow As Long
    Dim r As Long
    Dim employeeName As String
    Dim item As Variant

    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row

    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow

        employeeName = NormalizeEmployeeName(CStr(ws.Cells(r, COL_NAME).Value))

        If Len(employeeName) > 0 Then

            If Not dict.Exists(employeeName) Then

                ' 0 Name
                ' 1 PaidLeave (SUM Y)
                ' 2 Left2025 (Z)
                ' 3 Left2026 (AA)
                ' 4 RemainingLeave (AD)
                ' 5 Factories
                dict.Add employeeName, Array( _
                    employeeName, _
                    0#, _
                    Empty, _
                    Empty, _
                    Empty, _
                    factoryName)

            Else

                item = dict(employeeName)

                If InStr(1, CStr(item(5)), factoryName, vbTextCompare) = 0 Then
                    item(5) = CStr(item(5)) & " + " & factoryName
                End If

                dict(employeeName) = item

            End If

            item = dict(employeeName)

            ' Y = Paid -> SUM every monthly row
            If IsNumericCell(ws.Cells(r, COL_PAID).Value) Then
                item(1) = Round( _
                    CDbl(item(1)) + CDbl(ws.Cells(r, COL_PAID).Value), _
                    3)
            End If

            ' Z = Left 2025 -> first numeric value
            If IsEmptyValue(item(2)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2025).Value) Then
                    item(2) = Round(CDbl(ws.Cells(r, COL_LEFT_2025).Value), 3)
                End If
            End If

            ' AA = Left 2026 -> first numeric value
            If IsEmptyValue(item(3)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2026).Value) Then
                    item(3) = Round(CDbl(ws.Cells(r, COL_LEFT_2026).Value), 3)
                End If
            End If

            ' AD = Gross / Remaining Leave -> first numeric value, no SUM
            If IsEmptyValue(item(4)) Then
                If IsNumericCell(ws.Cells(r, COL_GROSS).Value) Then
                    item(4) = Round(CDbl(ws.Cells(r, COL_GROSS).Value), 3)
                End If
            End If

            dict(employeeName) = item

        End If

    Next r

End Sub

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
            """factories"":""" & JsonEscape(CStr(item(5))) & """}"

    Next key

    employees = employees & "]"

    BuildPayload = _
        "{""updatedAt"":""" & JsonEscape(updatedAt) & """," & _
        """source"":""" & JsonEscape(source) & """," & _
        """sheets"":""" & JsonEscape(sheet1 & " + " & sheet2) & """," & _
        """employees"":" & employees & "}"

End Function

Private Function NumberOrNull(ByVal v As Variant) As String

    Dim n As Double
    Dim s As String

    If IsEmptyValue(v) Then
        NumberOrNull = "null"
        Exit Function
    End If

    If Not IsNumeric(v) Then
        NumberOrNull = "null"
        Exit Function
    End If

    n = Round(CDbl(v), 3)

    s = Format$(n, "0.000")

    s = Replace( _
        s, _
        Application.International(xlDecimalSeparator), _
        ".")

    NumberOrNull = s

End Function

Private Function IsEmptyValue(ByVal v As Variant) As Boolean

    If IsEmpty(v) Then
        IsEmptyValue = True
    ElseIf IsNull(v) Then
        IsEmptyValue = True
    ElseIf VarType(v) = vbString Then
        IsEmptyValue = (Trim$(CStr(v)) = "")
    Else
        IsEmptyValue = False
    End If

End Function

Private Function IsNumericCell(ByVal v As Variant) As Boolean

    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        IsNumericCell = False
        Exit Function
    End If

    If Trim$(CStr(v)) = "" Or Trim$(CStr(v)) = "-" Then
        IsNumericCell = False
        Exit Function
    End If

    IsNumericCell = IsNumeric(v)

End Function

Private Function NormalizeEmployeeName(ByVal s As String) As String

    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeEmployeeName = Trim$(s)

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

Private Sub PostJson( _
    ByVal url As String, _
    ByVal body As String, _
    ByRef responseText As String, _
    ByRef httpStatus As Long)

    Dim http As Object

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False

    http.SetTimeouts 30000, 30000, 30000, 120000

    http.SetRequestHeader _
        "Content-Type", _
        "application/json; charset=utf-8"

    http.SetRequestHeader _
        "User-Agent", _
        "AnnualLeaveSync/1.0"

    http.Send body

    httpStatus = http.Status
    responseText = http.responseText

    If httpStatus < 200 Or httpStatus >= 300 Then
        Err.Raise vbObjectError + 300, , _
            "Google Apps Script returned HTTP " & httpStatus & "." & _
            vbCrLf & vbCrLf & responseText
    End If

End Sub

Private Sub GetGoogle( _
    ByVal url As String, _
    ByRef responseText As String, _
    ByRef httpStatus As Long)

    Dim http As Object

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False

    http.SetTimeouts 30000, 30000, 30000, 30000

    http.SetRequestHeader _
        "User-Agent", _
        "AnnualLeaveSync/1.0"

    http.Send

    httpStatus = http.Status
    responseText = http.responseText

    If httpStatus < 200 Or httpStatus >= 300 Then
        Err.Raise vbObjectError + 400, , _
            "HTTP " & httpStatus & "." & vbCrLf & responseText
    End If

End Sub
