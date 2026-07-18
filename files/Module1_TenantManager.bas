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
        If Trim(CStr(ws.cells(r, "A").Value)) = Trim(tenantID) Then
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
        If Trim(CStr(ws.cells(r, "A").Value)) = "" Then
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
        idVal = Trim(CStr(ws.cells(r, "A").Value))
        If Left(idVal, 2) = "T-" And IsNumeric(Mid(idVal, 3)) Then
            n = CLng(Mid(idVal, 3))
            If n > maxNum Then maxNum = n
        End If
    Next r
    NextTenantID = "T-" & Format(maxNum + 1, "000")
End Function

Public Sub LoadTenantToForm(ByVal tenantID As String)

    Dim f As Worksheet
    Dim t As Worksheet
    Dim rw As Long

    Set f = FormSheet()
    Set t = TenantsSheet()

    rw = FindTenantRow(tenantID)

    If rw = 0 Then
        MsgBox "Tenant not found.", vbExclamation
        Exit Sub
    End If

    f.Range("C5").Value = t.cells(rw, "A").Value
    f.Range("C7").Value = t.cells(rw, "B").Value
    f.Range("C9").Value = t.cells(rw, "C").Value
    f.Range("C11").Value = t.cells(rw, "D").Value
    f.Range("C13").Value = t.cells(rw, "E").Value
    f.Range("C15").Value = t.cells(rw, "F").Value
    f.Range("C17").Value = t.cells(rw, "G").Value
    f.Range("C19").Value = t.cells(rw, "H").Value

    f.Range("F5").Value = t.cells(rw, "I").Value
    f.Range("F7").Value = t.cells(rw, "J").Value
    f.Range("F9").Value = t.cells(rw, "L").Value
    f.Range("F11").Value = t.cells(rw, "M").Value
    f.Range("F13").Value = t.cells(rw, "N").Value
    f.Range("F15").Value = t.cells(rw, "O").Value
    f.Range("F17").Value = t.cells(rw, "P").Value

    f.Range("I5").Value = t.cells(rw, "S").Value
    f.Range("I7").Value = t.cells(rw, "T").Value
    f.Range("I9").Value = t.cells(rw, "U").Value

    f.Range("F19").Value = t.cells(rw, "V").Value
    f.Range("H12").Value = t.Cells(rw, "W").Value


End Sub




' Write the form's input fields into a TENANTS row
Private Sub WriteFormToRow(row As Long, tenantID As String)
    Dim f As Worksheet, t As Worksheet
    Set f = FormSheet()
    Set t = TenantsSheet()

    t.cells(row, "A").Value = tenantID
    t.cells(row, "B").Value = f.Range("C7").Value    ' Full Name / Company
    t.cells(row, "C").Value = f.Range("C9").Value    ' Phone
    t.cells(row, "D").Value = f.Range("C11").Value   ' Email
    t.cells(row, "E").Value = f.Range("C13").Value   ' ID / Reg No.
    t.cells(row, "F").Value = f.Range("C15").Value   ' Unit Number
    t.cells(row, "G").Value = f.Range("C17").Value   ' Unit Type
    t.cells(row, "H").Value = f.Range("C19").Value   ' Floor / Block
    t.cells(row, "I").Value = f.Range("F5").Value    ' Contract Start
    t.cells(row, "J").Value = f.Range("F7").Value    ' Contract End
    t.cells(row, "L").Value = f.Range("F9").Value    ' Monthly Rent
    t.cells(row, "M").Value = f.Range("F11").Value   ' Currency
    t.cells(row, "N").Value = f.Range("F13").Value   ' Payment Frequency
    t.cells(row, "O").Value = f.Range("F15").Value   ' Deposit Paid
    t.cells(row, "P").Value = f.Range("F17").Value   ' Payment Method
    t.cells(row, "S").Value = f.Range("I5").Value    ' Emergency Contact
    t.cells(row, "T").Value = f.Range("I7").Value    ' Relationship
    t.cells(row, "U").Value = f.Range("I9").Value    ' Emergency Phone
    t.cells(row, "V").Value = f.Range("F19").Value   ' Residential Address
    t.Cells(row, "W").Value = f.Range("H12").Value


    ' Re-apply the K / Q / R formulas for this row (Term, Days to Expiry, Status)
    t.cells(row, "K").Formula = "=IFERROR(DATEDIF(I" & row & ",J" & row & ",""M""),"""")"
    t.cells(row, "Q").Formula = "=IFERROR(IF(J" & row & "="""","""",J" & row & "-TODAY()),"""")"
    t.cells(row, "R").Formula = "=IFERROR(IF(B" & row & "="""",""Vacant""," & _
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
    cols = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V","W")
    For i = LBound(cols) To UBound(cols)
        t.cells(row, cols(i)).ClearContents
    Next i

    ResetForm
    RefreshGrid
    MsgBox "Tenant " & tenantID & " deleted.", vbInformation
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
        f.cells(row, "A").Formula = "=IFERROR(INDEX(TENANTS!A$4:A$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "B").Formula = "=IFERROR(INDEX(TENANTS!B$4:B$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "C").Formula = "=IFERROR(INDEX(TENANTS!F$4:F$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "D").Formula = "=IFERROR(INDEX(TENANTS!L$4:L$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "E").Formula = "=IFERROR(INDEX(TENANTS!M$4:M$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "F").Formula = "=IFERROR(INDEX(TENANTS!J$4:J$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "G").Formula = "=IFERROR(INDEX(TENANTS!R$4:R$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "H").Formula = "=IFERROR(INDEX(TENANTS!C$4:C$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "I").Formula = "=IFERROR(INDEX(TENANTS!S$4:S$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "J").Formula = "=IFERROR(INDEX(TENANTS!U$4:U$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "K").Formula = "=IFERROR(INDEX(TENANTS!V$4:V$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
        f.cells(row, "L").Formula = "=IFERROR(INDEX(TENANTS!P$4:P$" & LAST_DATA_ROW & "," & (i + 1) & "),"""")"
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
    Dim foundRow As Long

foundRow = FindTenantRow(Trim(f.Range("C25").Value))

If foundRow > 0 Then
    LoadTenantToForm Trim(f.Range("C25").Value)
    Exit Sub
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
            f.cells(clearRow, c).ClearContents
        Next c
    Next clearRow

    For r = FIRST_DATA_ROW To LAST_DATA_ROW
        If matchCount >= 6 Then Exit For
        idVal = CStr(t.cells(r, "A").Value)
        If Trim(idVal) = "" Then GoTo ContinueLoop

        nameVal = CStr(t.cells(r, "B").Value)
        unitVal = CStr(t.cells(r, "F").Value)
        statusVal = CStr(t.cells(r, "R").Value)

        Dim textMatch As Boolean, statusMatch As Boolean
        textMatch = (searchText = "") Or _
                    (InStr(1, LCase(nameVal), searchText) > 0) Or _
                    (InStr(1, LCase(idVal), searchText) > 0) Or _
                    (InStr(1, LCase(unitVal), searchText) > 0)
        statusMatch = (statusFilter = "All") Or (statusFilter = statusVal)

        If textMatch And statusMatch Then
            outRow = headerRow + 1 + matchCount
            f.cells(outRow, "A").Value = idVal
            f.cells(outRow, "B").Value = nameVal
            f.cells(outRow, "C").Value = unitVal
            f.cells(outRow, "D").Value = t.cells(r, "L").Value
            f.cells(outRow, "E").Value = t.cells(r, "M").Value
            f.cells(outRow, "F").Value = t.cells(r, "J").Value
            f.cells(outRow, "G").Value = statusVal
            f.cells(outRow, "H").Value = t.cells(r, "C").Value
            f.cells(outRow, "I").Value = t.cells(r, "S").Value
            f.cells(outRow, "J").Value = t.cells(r, "U").Value
            f.cells(outRow, "K").Value = t.cells(r, "V").Value
            f.cells(outRow, "L").Value = t.cells(r, "P").Value
            matchCount = matchCount + 1
        End If
ContinueLoop:
    Next r

    If matchCount = 0 Then MsgBox "No tenants matched your search.", vbInformation
End Sub

Public Sub RefreshGrid()
    ' Only restore formulas if the grid is currently formula-driven (i.e. not mid-search)
    RestoreGridFormulas
End Sub


Public Sub ResetForm()

    Dim f As Worksheet
    Set f = ThisWorkbook.Worksheets("TENANT FORM")

    f.Range("C5").ClearContents
    f.Range("C7").ClearContents
    f.Range("C9").ClearContents
    f.Range("C11").ClearContents
    f.Range("C13").ClearContents
    f.Range("C15").ClearContents
    f.Range("C17").ClearContents
    f.Range("C19").ClearContents

    f.Range("F5").ClearContents
    f.Range("F7").ClearContents
    f.Range("F9").ClearContents
    f.Range("F11").ClearContents
    f.Range("F13").ClearContents
    f.Range("F15").ClearContents
    f.Range("F17").ClearContents
    f.Range("F19").ClearContents

    f.Range("I5").ClearContents
    f.Range("I7").ClearContents
    f.Range("I9").ClearContents

    f.Range("H12").Value = "Tenant ID photo / lease scan"

End Sub
