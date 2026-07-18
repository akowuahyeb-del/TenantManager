Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' TENANT MANAGER - button logic
' Paste this whole module into a STANDARD module
' (Alt+F11 -> Insert -> Module -> paste)
' ============================================================

Const FIRST_DATA_ROW As Long = 4
Const LAST_DATA_ROW As Long = 53

Private Function TenantsSheet() As Worksheet
    Set TenantsSheet = ThisWorkbook.Worksheets("TENANTS")
End Function

Private Function FormSheet() As Worksheet
    Set FormSheet = ThisWorkbook.Worksheets("TENANT FORM")
End Function

' Find the row number in TENANTS matching a given Tenant ID, 0 if not found
Private Function FindTenantRow(tenantID As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = TenantsSheet()
    FindTenantRow = 0
    If Trim(tenantID) = "" Then Exit Function
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If Trim(CStr(ws.Cells(r, "A").Value)) = Trim(tenantID) Then
            FindTenantRow = r
            Exit Function
        End If
    Next r
End Function

' First empty row in TENANTS (based on column A)
Private Function FirstEmptyRow() As Long
    Dim ws As Worksheet, r As Long
    Set ws = TenantsSheet()
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If Trim(CStr(ws.Cells(r, "A").Value)) = "" Then
            FirstEmptyRow = r
            Exit Function
        End If
    Next r
    FirstEmptyRow = 0 ' full
End Function

Private Function NextTenantID() As String
    Dim ws As Worksheet, r As Long, maxNum As Long, n As Long, idVal As String
    Set ws = TenantsSheet()
    maxNum = 0
    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        idVal = Trim(CStr(ws.Cells(r, "A").Value))
        If Left(idVal, 2) = "T-" And IsNumeric(Mid(idVal, 3)) Then
            n = CLng(Mid(idVal, 3))
            If n > maxNum Then maxNum = n
        End If
    Next r
    NextTenantID = "T-" & Format(maxNum + 1, "000")
End Function




' Load a tenant record from TENANTS into the form
Public Sub LoadTenantToForm(ByVal tenantID As String)

    Dim f As Worksheet, t As Worksheet
    Dim row As Long

    Set f = FormSheet()
    Set t = TenantsSheet()

    row = FindTenantRow(tenantID)

    If row = 0 Then
        MsgBox "Tenant not found.", vbExclamation
        Exit Sub
    End If

    f.Range("C5").Value = t.Cells(row, "A").Value
    f.Range("C7").Value = t.Cells(row, "B").Value
    f.Range("C9").Value = t.Cells(row, "C").Value
    f.Range("C11").Value = t.Cells(row, "D").Value
    f.Range("C13").Value = t.Cells(row, "E").Value
    f.Range("C15").Value = t.Cells(row, "F").Value
    f.Range("C17").Value = t.Cells(row, "G").Value
    f.Range("C19").Value = t.Cells(row, "H").Value

    f.Range("F5").Value = t.Cells(row, "I").Value
    f.Range("F7").Value = t.Cells(row, "J").Value
    f.Range("F9").Value = t.Cells(row, "L").Value
    f.Range("F11").Value = t.Cells(row, "M").Value
    f.Range("F13").Value = t.Cells(row, "N").Value
    f.Range("F15").Value = t.Cells(row, "O").Value
    f.Range("F17").Value = t.Cells(row, "P").Value

    f.Range("I5").Value = t.Cells(row, "S").Value
    f.Range("I7").Value = t.Cells(row, "T").Value
    f.Range("I9").Value = t.Cells(row, "U").Value

    f.Range("F19").Value = t.Cells(row, "V").Value

End Sub



' Write the form's input fields into a TENANTS row
Private Sub WriteFormToRow(row As Long, tenantID As String)
    Dim f As Worksheet, t As Worksheet
    Set f = FormSheet()
    Set t = TenantsSheet()

    t.Cells(row, "A").Value = tenantID
    t.Cells(row, "B").Value = f.Range("C7").Value    ' Full Name / Company
    t.Cells(row, "C").Value = f.Range("C9").Value    ' Phone
    t.Cells(row, "D").Value = f.Range("C11").Value   ' Email
    t.Cells(row, "E").Value = f.Range("C13").Value   ' ID / Reg No.
    t.Cells(row, "F").Value = f.Range("C15").Value   ' Unit Number
    t.Cells(row, "G").Value = f.Range("C17").Value   ' Unit Type
    t.Cells(row, "H").Value = f.Range("C19").Value   ' Floor / Block
    t.Cells(row, "I").Value = f.Range("F5").Value    ' Contract Start
    t.Cells(row, "J").Value = f.Range("F7").Value    ' Contract End
    t.Cells(row, "L").Value = f.Range("F9").Value    ' Monthly Rent
    t.Cells(row, "M").Value = f.Range("F11").Value   ' Currency
    t.Cells(row, "N").Value = f.Range("F13").Value   ' Payment Frequency
    t.Cells(row, "O").Value = f.Range("F15").Value   ' Deposit Paid
    t.Cells(row, "P").Value = f.Range("F17").Value   ' Payment Method
    t.Cells(row, "S").Value = f.Range("I5").Value    ' Emergency Contact
    t.Cells(row, "T").Value = f.Range("I7").Value    ' Relationship
    t.Cells(row, "U").Value = f.Range("I9").Value    ' Emergency Phone
    t.Cells(row, "V").Value = f.Range("F19").Value   ' Residential Address

    ' Re-apply the K / Q / R formulas for this row (Term, Days to Expiry, Status)
    t.Cells(row, "K").Formula = "=IFERROR(DATEDIF(I" & row & ",J" & row & ",""M""),"""")"
    t.Cells(row, "Q").Formula = "=IFERROR(IF(J" & row & "="""","""",J" & row & "-TODAY()),"""")"
    t.Cells(row, "R").Formula = "=IFERROR(IF(B" & row & "="""",""Vacant""," & _
        "IF(Q" & row & "<0,""Expired""," & _
        "IF(Q" & row & "<=30,""Expiring Soon""," & _
        "IF(Q" & row & "<=90,""Active - Review"",""Active"")))),""Vacant"")"
End Sub

Public Sub AddTenant()
    Dim f As Worksheet, targetRow As Long, tenantID As String

    Set f = FormSheet()
    tenantID = Trim(f.Range("C5").Value)

    If tenantID = "" Then
        tenantID = NextTenantID()
    ElseIf FindTenantRow(tenantID) > 0 Then
        MsgBox "Tenant ID " & tenantID & " already exists. Use Update instead, or clear the Tenant ID field to auto-generate a new one.", vbExclamation
        Exit Sub
    End If

    If Trim(f.Range("C7").Value) = "" Then
        MsgBox "Please enter the tenant's full name / company.", vbExclamation
        Exit Sub
    End If

    targetRow = FirstEmptyRow()
    If targetRow = 0 Then
        MsgBox "The TENANTS sheet is full (max " & (LAST_DATA_ROW - FIRST_DATA_ROW + 1) & " rows). Extend the sheet before adding more.", vbCritical
        Exit Sub
    End If

    WriteFormToRow targetRow, tenantID
    f.Range("C5").Value = tenantID
    RefreshGrid
    MsgBox "Tenant " & tenantID & " added.", vbInformation
End Sub

Public Sub UpdateTenant()
    Dim f As Worksheet, row As Long, tenantID As String
    Set f = FormSheet()
    tenantID = Trim(f.Range("C5").Value)
    row = FindTenantRow(tenantID)
    

    If row = 0 Then
        MsgBox "No tenant found with ID '" & tenantID & "'. Type an existing Tenant ID into the Tenant ID field, or use Add for a new tenant.", vbExclamation
        Exit Sub
    End If

    WriteFormToRow row, tenantID
    RefreshGrid
    MsgBox "Tenant " & tenantID & " updated.", vbInformation
End Sub

Public Sub DeleteTenant()
    Dim f As Worksheet, t As Worksheet, row As Long, tenantID As String
    Set f = FormSheet()
    Set t = TenantsSheet()
    tenantID = Trim(f.Range("C5").Value)
    row = FindTenantRow(tenantID)

    If row = 0 Then
        MsgBox "No tenant found with ID '" & tenantID & "'.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Delete tenant " & tenantID & "? This cannot be undone.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    Dim cols As Variant, i As Long
    cols = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V")
    For i = LBound(cols) To UBound(cols)
        t.Cells(row, cols(i)).ClearContents
    Next i

    ResetForm
    RefreshGrid
    MsgBox "Tenant " & tenantID & " deleted.", vbInformation
End Sub

Public Sub ResetForm()
    Dim f As Worksheet
    Set f = FormSheet()
    Dim cells As Variant, i As Long
    cells = Array("C5", "C7", "C9", "C11", "C13", "C15", "C17", "C19", _
                  "F5", "F7", "F9", "F11", "F13", "F15", "F17", "F19", _
                  "I5", "I7", "I9", "C25")
    For i = LBound(cells) To UBound(cells)
        f.Range(cells(i)).ClearContents
    Next i
    f.Range("G25").Value = "All"
    RestoreGridFormulas
End Sub

Public Sub CloseForm()
    ResetForm
    ThisWorkbook.Worksheets("DASHBOARD").Activate
End Sub

Public Sub UploadDocument()
    Dim f As Worksheet, filePath As Variant
    Set f = FormSheet()
    filePath = Application.GetOpenFilename( _
        FileFilter:="Documents (*.pdf;*.jpg;*.jpeg;*.png),*.pdf;*.jpg;*.jpeg;*.png", _
        Title:="Select tenant ID photo or lease scan")
    If filePath = False Then Exit Sub
    f.Range("H12").Value = filePath
    MsgBox "File path saved. It will be stored against this tenant's PicturePath when you click Add.", vbInformation
End Sub

' Put the original INDEX-based formulas back into the 6-row preview grid
Public Sub RestoreGridFormulas()
    Dim f As Worksheet, i As Long, row As Long
    Set f = FormSheet()
    Dim headerRow As Long: headerRow = 33
    For i = 0 To 5
        row = headerRow + 1 + i
        f.Cells(row, "A").Formula = "=IFERROR(INDEX(TENANTS!A$4:A$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "B").Formula = "=IFERROR(INDEX(TENANTS!B$4:B$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "C").Formula = "=IFERROR(INDEX(TENANTS!F$4:F$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "D").Formula = "=IFERROR(INDEX(TENANTS!L$4:L$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "E").Formula = "=IFERROR(INDEX(TENANTS!M$4:M$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "F").Formula = "=IFERROR(INDEX(TENANTS!J$4:J$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "G").Formula = "=IFERROR(INDEX(TENANTS!R$4:R$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "H").Formula = "=IFERROR(INDEX(TENANTS!C$4:C$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "I").Formula = "=IFERROR(INDEX(TENANTS!S$4:S$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "J").Formula = "=IFERROR(INDEX(TENANTS!U$4:U$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "K").Formula = "=IFERROR(INDEX(TENANTS!V$4:V$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.Cells(row, "L").Formula = "=IFERROR(INDEX(TENANTS!P$4:P$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
    Next i
End Sub

' Filters TENANTS by the Search box text + Status dropdown, writes matches
' as literal values into the 6-row preview grid (overwrites formulas until Reset)
Public Sub SearchTenants()
    Dim f As Worksheet, t As Worksheet
    Dim searchText As String, statusFilter As String
    Dim r As Long, matchCount As Long, headerRow As Long, outRow As Long
    Dim nameVal As String, idVal As String, unitVal As String, statusVal As String

    Set f = FormSheet()
    Set t = TenantsSheet()
        ' Direct Tenant ID lookup
    If Trim(f.Range("C25").Value) <> "" Then

        Dim foundRow As Long

        foundRow = FindTenantRow(Trim(f.Range("C25").Value))

        If foundRow > 0 Then
            LoadTenantToForm Trim(f.Range("C25").Value)
            MsgBox "Tenant loaded into form.", vbInformation
            Exit Sub
        End If

    End If
    searchText = LCase(Trim(f.Range("C25").Value))
    statusFilter = Trim(f.Range("G25").Value)
    If statusFilter = "" Then statusFilter = "All"
    headerRow = 33
    matchCount = 0

    ' clear the grid first
    Dim clearRow As Long, c As Long
    For clearRow = headerRow + 1 To headerRow + 6
        For c = 1 To 12
            f.Cells(clearRow, c).ClearContents
        Next c
    Next clearRow

    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If matchCount >= 6 Then Exit For
        idVal = CStr(t.Cells(r, "A").Value)
        If Trim(idVal) = "" Then GoTo ContinueLoop

        nameVal = CStr(t.Cells(r, "B").Value)
        unitVal = CStr(t.Cells(r, "F").Value)
        statusVal = CStr(t.Cells(r, "R").Value)

        Dim textMatch As Boolean, statusMatch As Boolean
        textMatch = (searchText = "") Or _
                    (InStr(1, LCase(nameVal), searchText) > 0) Or _
                    (InStr(1, LCase(idVal), searchText) > 0) Or _
                    (InStr(1, LCase(unitVal), searchText) > 0)
        statusMatch = (statusFilter = "All") Or (statusFilter = statusVal)

        If textMatch And statusMatch Then
            outRow = headerRow + 1 + matchCount
            f.Cells(outRow, "A").Value = idVal
            f.Cells(outRow, "B").Value = nameVal
            f.Cells(outRow, "C").Value = unitVal
            f.Cells(outRow, "D").Value = t.Cells(r, "L").Value
            f.Cells(outRow, "E").Value = t.Cells(r, "M").Value
            f.Cells(outRow, "F").Value = t.Cells(r, "J").Value
            f.Cells(outRow, "G").Value = statusVal
            f.Cells(outRow, "H").Value = t.Cells(r, "C").Value
            f.Cells(outRow, "I").Value = t.Cells(r, "S").Value
            f.Cells(outRow, "J").Value = t.Cells(r, "U").Value
            f.Cells(outRow, "K").Value = t.Cells(r, "V").Value
            f.Cells(outRow, "L").Value = t.Cells(r, "P").Value
            matchCount = matchCount + 1
        End If
ContinueLoop:
    Next r

    If matchCount = 0 Then
    MsgBox "No tenants matched your search.", vbInformation
Else
    MsgBox matchCount & " tenant(s) found.", vbInformation
End If
End Sub

Public Sub RefreshGrid()
    ' Only restore formulas if the grid is currently formula-driven (i.e. not mid-search)
    RestoreGridFormulas
End Sub
