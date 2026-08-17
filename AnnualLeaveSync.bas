Option Explicit

' ============================================================
' ANNUAL LEAVE SYNC - 2 FACTORIES
'
' Source workbook:
' X:\DS + PN + TP - 2014\Npn2023.xlsb
'
' Password:
' 2000
'
' CURRENT YEAR SHEETS:
'   2026
'   2026 PL
'
' DATA COLUMNS IN BOTH SHEETS:
'   A  = HỌ VÀ TÊN
'   Y  = Paid
'   Z  = Left 2025
'   AA = Left 2026
'   AD = Gross (phép còn lại)
'
' RULE:
'   - Y: SUM all monthly rows by employee name across BOTH factories.
'   - Z: take the first non-blank numeric value for that employee.
'   - AA: take the first non-blank numeric value for that employee.
'   - AD: take the first non-blank numeric value for that employee.
'
' IMPORTANT:
' Put this code in a macro-enabled workbook that the HR computer
' can open (for example PERSONAL.XLSB).
' ============================================================

Private Const SOURCE_FILE As String = "X:\DS + PN + TP - 2014\Npn2023.xlsb"
Private Const SOURCE_PASSWORD As String = "2000"

' Paste your deployed Google Apps Script Web App URL here.
Private Const WEB_APP_URL As String = "https://script.google.com/macros/s/AKfycbxUSCx1x8scN0Xq-3ec-KcDop9bb-AZy8iH9TJyDWgPhmMYr14smVac2MB0GD2L1TRkQw/exec"

Private Const FIRST_DATA_ROW As Long = 4

' Fixed columns from the source workbook.
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

    On Error GoTo EH

    If WEB_APP_URL = "" Or InStr(1, WEB_APP_URL, "PASTE_", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 100, , _
            "Please paste the Google Apps Script Web App URL into WEB_APP_URL."
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    yearText = CStr(Year(Date))
    sheet1 = yearText
    sheet2 = yearText & " PL"

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set wb = GetOpenWorkbookByFullName(SOURCE_FILE)

    If wb Is Nothing Then
        If Dir(SOURCE_FILE) = "" Then
            Err.Raise vbObjectError + 101, , _
                "Cannot find source file: " & SOURCE_FILE
        End If

        Set wb = Workbooks.Open( _
            Filename:=SOURCE_FILE, _
            Password:=SOURCE_PASSWORD, _
            ReadOnly:=True, _
            UpdateLinks:=False)

        openedByMacro = True
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(sheet1)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 102, , _
            "Cannot find sheet: " & sheet1
    End If

    ReadFactorySheet ws, dict, sheet1

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(sheet2)
    On Error GoTo EH

    If ws Is Nothing Then
        Err.Raise vbObjectError + 103, , _
            "Cannot find sheet: " & sheet2
    End If

    ReadFactorySheet ws, dict, sheet2

    If dict.Count = 0 Then
        Err.Raise vbObjectError + 104, , _
            "No employee data found in sheets " & sheet1 & " and " & sheet2
    End If

    updated = Format(Now, "yyyy-mm-dd HH:nn:ss")
    payload = BuildPayload(dict, updated, SOURCE_FILE, sheet1, sheet2)

    PostJson WEB_APP_URL, payload

    If openedByMacro Then wb.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Đồng bộ phép năm thành công." & vbCrLf & vbCrLf & _
           "Sheets: " & sheet1 & " + " & sheet2 & vbCrLf & _
           "Nhân viên: " & dict.Count & vbCrLf & _
           "Cập nhật: " & updated, vbInformation

    Exit Sub

EH:
    On Error Resume Next
    If openedByMacro Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Đồng bộ thất bại:" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub ReadFactorySheet(ByVal ws As Worksheet, ByVal dict As Object, ByVal factoryName As String)

    Dim lastRow As Long
    Dim r As Long
    Dim employeeName As String
    Dim item As Variant

    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row

    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow

        employeeName = Trim$(CStr(ws.Cells(r, COL_NAME).Value))

        If Len(employeeName) > 0 Then

            If Not dict.Exists(employeeName) Then
                ' Array:
                ' 0 Name
                ' 1 Paid (Y) SUM
                ' 2 Left2025 (Z) first numeric
                ' 3 Left2026 (AA) first numeric
                ' 4 Gross (AD) first numeric
                ' 5 Source factories
                dict.Add employeeName, Array( _
                    employeeName, _
                    0#, _
                    Empty, _
                    Empty, _
                    Empty, _
                    factoryName)
            Else
                item = dict(employeeName)

                If InStr(1, item(5), factoryName, vbTextCompare) = 0 Then
                    item(5) = item(5) & " + " & factoryName
                End If

                dict(employeeName) = item
            End If

            item = dict(employeeName)

            ' Y = Paid: SUM all monthly rows.
            If IsNumericCell(ws.Cells(r, COL_PAID).Value) Then
                item(1) = CDbl(item(1)) + CDbl(ws.Cells(r, COL_PAID).Value)
            End If

            ' Z = Left 2025: first non-blank numeric value.
            If IsEmptyValue(item(2)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2025).Value) Then
                    item(2) = CDbl(ws.Cells(r, COL_LEFT_2025).Value)
                End If
            End If

            ' AA = Left 2026: first non-blank numeric value.
            If IsEmptyValue(item(3)) Then
                If IsNumericCell(ws.Cells(r, COL_LEFT_2026).Value) Then
                    item(3) = CDbl(ws.Cells(r, COL_LEFT_2026).Value)
                End If
            End If

            ' AD = Gross / remaining leave:
            ' take first non-blank numeric value, do NOT SUM.
            If IsEmptyValue(item(4)) Then
                If IsNumericCell(ws.Cells(r, COL_GROSS).Value) Then
                    item(4) = CDbl(ws.Cells(r, COL_GROSS).Value)
                End If
            End If

            dict(employeeName) = item
        End If
    Next r

End Sub

Private Function BuildPayload(ByVal dict As Object, _
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

    If IsEmptyValue(v) Then
        NumberOrNull = "null"
    ElseIf IsNumeric(v) Then
        NumberOrNull = Replace(CStr(CDbl(v)), ",", ".")
    Else
        NumberOrNull = "null"
    End If

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

Private Sub PostJson(ByVal url As String, ByVal body As String)

    Dim http As Object

    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send body

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 105, , _
            "Google Apps Script returned HTTP " & http.Status & _
            ": " & http.responseText
    End If

End Sub
