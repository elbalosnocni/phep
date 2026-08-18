Option Explicit

' ============================================================
' ANNUAL LEAVE SYNC - FINAL + MONTHLY LEAVE
'
' SOURCE:
'   X:\DS + PN + TP - 2014\Npn2023.xlsb
'
' PASSWORD:
'   2000
'
' SHEETS:
'   2026    -> XUONG BANH
'   2026 PL -> XUONG IN
'
' COLUMNS:
'   A  = HO VA TEN
'   C  = MONTH (Jan, Feb, Mar ... Dec)
'   Y  = Paid Leave
'   Z  = Left 2025
'   AA = Left 2026
'   AD = Gross / Remaining Leave
'
' LOGIC:
'   Y  -> SUM all monthly rows = PaidLeave
'   Z  -> first numeric value = Left2025
'   AA -> first numeric value = Left2026
'   AD -> first numeric value = RemainingLeave
'
' MONTHLY:
'   C = Jan + Y -> Jan
'   C = Feb + Y -> Feb
'   ...
'   C = Dec + Y -> Dec
'
' Google Sheet:
'   A EmployeeName
'   B PaidLeave
'   C Left2025
'   D Left2026
'   E RemainingLeave
'   F Factories
'   G UpdatedAt
'   H Jan
'   I Feb
'   J Mar
'   K Apr
'   L May
'   M Jun
'   N Jul
'   O Aug
'   P Sep
'   Q Oct
'   R Nov
'   S Dec
'
' All numbers rounded to 3 decimals.
' ============================================================

Private Const SOURCE_FILE As String = _
    "X:\DS + PN + TP - 2014\Npn2023.xlsb"

Private Const SOURCE_PASSWORD As String = "2000"

Private Const WEB_APP_URL As String = _
    "https://script.google.com/macros/s/AKfycbxUSCx1x8scN0Xq-3ec-KcDop9bb-AZy8iH9TJyDWgPhmMYr14smVac2MB0GD2L1TRkQw/exec"

Private Const FIRST_DATA_ROW As Long = 4

' Source columns
Private Const COL_NAME As Long = 1       ' A
Private Const COL_MONTH As Long = 3      ' C
Private Const COL_PAID As Long = 25      ' Y
Private Const COL_LEFT_2025 As Long = 26 ' Z
Private Const COL_LEFT_2026 As Long = 27 ' AA
Private Const COL_GROSS As Long = 30     ' AD

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
        Err.Raise vbObjectError + 100, , _
            "WEB_APP_URL is empty."
    End If

    If InStr(1, WEB_APP_URL, "PASTE_", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 101, , _
            "WEB_APP_URL has not been configured."
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

    ' --------------------------------------------------------
    ' READ SHEET 2026
    ' --------------------------------------------------------

    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(sheet1)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 103, , _
            "Cannot find sheet: " & sheet1
    End If

    Application.StatusBar = _
        "Reading " & sheet1 & " -> XUONG BANH..."

    ReadFactorySheet ws, dict, "XUONG BANH"

    ' --------------------------------------------------------
    ' READ SHEET 2026 PL
    ' --------------------------------------------------------

    Set ws = Nothing

    On Error Resume Next
    Set ws = wb.Worksheets(sheet2)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 104, , _
            "Cannot find sheet: " & sheet2
    End If

    Application.StatusBar = _
        "Reading " & sheet2 & " -> XUONG IN..."

    ReadFactorySheet ws, dict, "XUONG IN"

    employeeCount = dict.Count

    If employeeCount = 0 Then
        Err.Raise vbObjectError + 105, , _
            "No employee data found in " & sheet1 & _
            " and " & sheet2
    End If

    ' --------------------------------------------------------
    ' UPDATE TIME
    ' --------------------------------------------------------

    updated = Format(Now, "dd/mm/yyyy HH:nn:ss")

    ' --------------------------------------------------------
    ' BUILD JSON
    ' --------------------------------------------------------

    payload = BuildPayload( _
        dict, _
        updated, _
        SOURCE_FILE, _
        sheet1, _
        sheet2)

    ' --------------------------------------------------------
    ' SEND GOOGLE
    ' --------------------------------------------------------

    Application.StatusBar = _
        "Sending data to Google Sheets..."

    PostJson _
        WEB_APP_URL, _
        payload, _
        responseText, _
        httpStatus

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

    MsgBox _
        "DONG BO THANH CONG" & vbCrLf & vbCrLf & _
        "Sheets: " & sheet1 & " + " & sheet2 & vbCrLf & _
        "Nhan vien: " & employeeCount & vbCrLf & _
        "Xuong: Banh / In" & vbCrLf & _
        "Da dong bo phep theo thang: Jan-Dec" & vbCrLf & _
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
        If Not wb Is Nothing Then
            wb.Close SaveChanges:=False
        End If
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

' ============================================================
' TEST GOOGLE
' ============================================================

Public Sub TestGoogleConnection()

    Dim responseText As String
    Dim httpStatus As Long

    On Error GoTo EH

    GetGoogle _
        WEB_APP_URL, _
        responseText, _
        httpStatus

    Application.StatusBar = False

    MsgBox _
        "KET NOI GOOGLE THANH CONG" & vbCrLf & vbCrLf & _
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
    Dim monthName As String

    Dim item As Variant
    Dim monthIndex As Long

    Dim paidValue As Double

    lastRow = ws.Cells( _
        ws.Rows.Count, _
        COL_NAME).End(xlUp).Row

    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow

        employeeName = NormalizeEmployeeName( _
            CStr(ws.Cells(r, COL_NAME).Value))

        If Len(employeeName) > 0 Then

            ' ------------------------------------------------
            ' CREATE EMPLOYEE
            '
            ' item:
            ' 0  Name
            ' 1  PaidLeave total
            ' 2  Left2025
            ' 3  Left2026
            ' 4  RemainingLeave
            ' 5  Factories
            ' 6  Jan
            ' 7  Feb
            ' 8  Mar
            ' 9  Apr
            ' 10 May
            ' 11 Jun
            ' 12 Jul
            ' 13 Aug
            ' 14 Sep
            ' 15 Oct
            ' 16 Nov
            ' 17 Dec
            ' ------------------------------------------------

            If Not dict.Exists(employeeName) Then

                dict.Add employeeName, Array( _
                    employeeName, _
                    0#, _
                    Empty, _
                    Empty, _
                    Empty, _
                    factoryName, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#, _
                    0#)

            Else

                item = dict(employeeName)

                If InStr( _
                    1, _
                    CStr(item(5)), _
                    factoryName, _
                    vbTextCompare) = 0 Then

                    item(5) = CStr(item(5)) & _
                              " + " & factoryName
                End If

                dict(employeeName) = item

            End If

            item = dict(employeeName)

            ' ------------------------------------------------
            ' Y = PAID LEAVE
            ' Keep existing calculation:
            ' SUM all monthly rows
            ' ------------------------------------------------

            If IsNumericCell( _
                ws.Cells(r, COL_PAID).Value) Then

                paidValue = CDbl( _
                    ws.Cells(r, COL_PAID).Value)

                item(1) = Round( _
                    CDbl(item(1)) + paidValue, _
                    3)

            End If

            ' ------------------------------------------------
            ' MONTHLY LEAVE
            ' C = Jan / Feb / Mar...
            ' Y = Paid Leave
            ' ------------------------------------------------

            monthName = Trim$( _
                CStr(ws.Cells(r, COL_MONTH).Value))

            monthIndex = MonthIndex(monthName)

            If monthIndex > 0 Then

                If IsNumericCell( _
                    ws.Cells(r, COL_PAID).Value) Then

                    paidValue = CDbl( _
                        ws.Cells(r, COL_PAID).Value)

                    item(5 + monthIndex) = Round( _
                        CDbl(item(5 + monthIndex)) + _
                        paidValue, _
                        3)

                End If

            End If

            ' ------------------------------------------------
            ' Z = LEFT 2025
            ' first numeric value
            ' ------------------------------------------------

            If IsEmptyValue(item(2)) Then

                If IsNumericCell( _
                    ws.Cells(r, COL_LEFT_2025).Value) Then

                    item(2) = Round( _
                        CDbl(ws.Cells( _
                            r, COL_LEFT_2025).Value), _
                        3)

                End If

            End If

            ' ------------------------------------------------
            ' AA = LEFT 2026
            ' first numeric value
            ' ------------------------------------------------

            If IsEmptyValue(item(3)) Then

                If IsNumericCell( _
                    ws.Cells(r, COL_LEFT_2026).Value) Then

                    item(3) = Round( _
                        CDbl(ws.Cells( _
                            r, COL_LEFT_2026).Value), _
                        3)

                End If

            End If

            ' ------------------------------------------------
            ' AD = REMAINING LEAVE
            ' first numeric value
            ' ------------------------------------------------

            If IsEmptyValue(item(4)) Then

                If IsNumericCell( _
                    ws.Cells(r, COL_GROSS).Value) Then

                    item(4) = Round( _
                        CDbl(ws.Cells( _
                            r, COL_GROSS).Value), _
                        3)

                End If

            End If

            dict(employeeName) = item

        End If

    Next r

End Sub

' ============================================================
' MONTH INDEX
' ============================================================

Private Function MonthIndex( _
    ByVal monthName As String) As Long

    Select Case LCase$(Trim$(monthName))

        Case "jan", "january"
            MonthIndex = 1

        Case "feb", "february"
            MonthIndex = 2

        Case "mar", "march"
            MonthIndex = 3

        Case "apr", "april"
            MonthIndex = 4

        Case "may"
            MonthIndex = 5

        Case "jun", "june"
            MonthIndex = 6

        Case "jul", "july"
            MonthIndex = 7

        Case "aug", "august"
            MonthIndex = 8

        Case "sep", "sept", "september"
            MonthIndex = 9

        Case "oct", "october"
            MonthIndex = 10

        Case "nov", "november"
            MonthIndex = 11

        Case "dec", "december"
            MonthIndex = 12

        Case Else
            MonthIndex = 0

    End Select

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

        If Len(employees) > 1 Then
            employees = employees & ","
        End If

        employees = employees & _
            "{""employeeName"":""" & _
            JsonEscape(CStr(item(0))) & """," & _

            """paidLeave"":" & _
            NumberOrNull(item(1)) & "," & _

            """left2025"":" & _
            NumberOrNull(item(2)) & "," & _

            """left2026"":" & _
            NumberOrNull(item(3)) & "," & _

            """remainingLeave"":" & _
            NumberOrNull(item(4)) & "," & _

            """factories"":""" & _
            JsonEscape(CStr(item(5))) & """," & _

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
        "{""updatedAt"":""" & _
        JsonEscape(updatedAt) & """," & _

        """source"":""" & _
        JsonEscape(source) & """," & _

        """sheets"":""" & _
        JsonEscape(sheet1 & " + " & sheet2) & """," & _

        """employees"":" & employees & "}"

End Function

' ============================================================
' NUMBER -> JSON
' ============================================================

Private Function NumberOrNull( _
    ByVal v As Variant) As String

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

' ============================================================
' EMPTY VALUE
' ============================================================

Private Function IsEmptyValue( _
    ByVal v As Variant) As Boolean

    If IsEmpty(v) Then

        IsEmptyValue = True

    ElseIf IsNull(v) Then

        IsEmptyValue = True

    ElseIf VarType(v) = vbString Then

        IsEmptyValue = _
            (Trim$(CStr(v)) = "")

    Else

        IsEmptyValue = False

    End If

End Function

' ============================================================
' NUMERIC CELL
' ============================================================

Private Function IsNumericCell( _
    ByVal v As Variant) As Boolean

    If IsError(v) Or _
       IsEmpty(v) Or _
       IsNull(v) Then

        IsNumericCell = False
        Exit Function

    End If

    If Trim$(CStr(v)) = "" Or _
       Trim$(CStr(v)) = "-" Then

        IsNumericCell = False
        Exit Function

    End If

    IsNumericCell = IsNumeric(v)

End Function

' ============================================================
' NORMALIZE EMPLOYEE NAME
' ============================================================

Private Function NormalizeEmployeeName( _
    ByVal s As String) As String

    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeEmployeeName = Trim$(s)

End Function

' ============================================================
' JSON ESCAPE
' ============================================================

Private Function JsonEscape( _
    ByVal s As String) As String

    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")

    JsonEscape = s

End Function

' ============================================================
' FIND OPEN WORKBOOK
' ============================================================

Private Function GetOpenWorkbookByFullName( _
    ByVal fullName As String) As Workbook

    Dim wb As Workbook

    For Each wb In Application.Workbooks

        If StrComp( _
            wb.FullName, _
            fullName, _
            vbTextCompare) = 0 Then

            Set GetOpenWorkbookByFullName = wb
            Exit Function

        End If

    Next wb

End Function

' ============================================================
' POST JSON
' ============================================================

Private Sub PostJson( _
    ByVal url As String, _
    ByVal body As String, _
    ByRef responseText As String, _
    ByRef httpStatus As Long)

    Dim http As Object

    Set http = _
        CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "POST", url, False

    http.SetTimeouts _
        30000, _
        30000, _
        30000, _
        120000

    http.SetRequestHeader _
        "Content-Type", _
        "application/json; charset=utf-8"

    http.SetRequestHeader _
        "User-Agent", _
        "AnnualLeaveSync/2.0"

    http.Send body

    httpStatus = http.Status
    responseText = http.responseText

    If httpStatus < 200 Or _
       httpStatus >= 300 Then

        Err.Raise _
            vbObjectError + 300, , _
            "Google Apps Script returned HTTP " & _
            httpStatus & "." & _
            vbCrLf & vbCrLf & responseText

    End If

End Sub

' ============================================================
' GET GOOGLE
' ============================================================

Private Sub GetGoogle( _
    ByVal url As String, _
    ByRef responseText As String, _
    ByRef httpStatus As Long)

    Dim http As Object

    Set http = _
        CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False

    http.SetTimeouts _
        30000, _
        30000, _
        30000, _
        30000

    http.SetRequestHeader _
        "User-Agent", _
        "AnnualLeaveSync/2.0"

    http.Send

    httpStatus = http.Status
    responseText = http.responseText

    If httpStatus < 200 Or _
       httpStatus >= 300 Then

        Err.Raise _
            vbObjectError + 400, , _
            "HTTP " & httpStatus & "." & _
            vbCrLf & responseText

    End If

End Sub
