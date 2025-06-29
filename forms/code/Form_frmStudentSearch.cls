VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStudentSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'=== Initialize form ===
Private Sub Form_Load()
    Me.cboSearchField = ""
    Me.txtSearchText = ""
End Sub

'=== Populate the list with top 100 records ===
Private Sub SetupListBox()
    With Me.lstStudents
        .RowSourceType = "Table/Query"
        .ColumnCount = 7
        .BoundColumn = 1
        .ColumnHeads = False
        .ColumnWidths = "1cm;3cm;3cm;2.5cm;2.5cm;2.5cm;3cm"
        .RowSource = _
            "SELECT TOP 100 p.PersonID, p.FirstName, p.LastName, " & _
            "Format(p.BirthDate,'mm/dd/yyyy') AS DOB, " & _
            "Format(p.BirthTime,'hh:nn:ss') AS TOB, " & _
            "Nz(p.Initials,'') AS Initials, " & _
            "l.City & IIf(Len(l.[StateProvince])>0,', ' & l.[StateProvince],'') & ', ' & l.Country AS BirthPlace " & _
            "FROM tblPeople AS p " & _
            "LEFT JOIN tblLocations AS l ON p.LocationID = l.LocationID " & _
            "ORDER BY p.LastName, p.FirstName;"
    End With
End Sub

'=== Search button ===
Private Sub btnSearch_Click()
    On Error GoTo ErrHandler
    
    Dim strSQL   As String
    Dim strWhere As String
    Dim db       As DAO.Database
    Dim rs       As DAO.Recordset
    Dim txt      As String
    
    txt = Nz(Me.txtSearchText, "")
    If Len(Trim(txt)) = 0 Then
        SetupListBox
        Exit Sub
    End If
    
    ' Base SELECT
    strSQL = _
      "SELECT p.PersonID, p.FirstName, p.LastName, " & _
      "Format(p.BirthDate,'mm/dd/yyyy') AS DOB, " & _
      "Format(p.BirthTime,'hh:nn:ss') AS TOB, " & _
      "Nz(p.Initials,'') AS Initials, " & _
      "l.City & IIf(Len(l.[StateProvince])>0,', ' & l.[StateProvince],'') & ', ' & l.Country AS BirthPlace " & _
      "FROM tblPeople AS p " & _
      "LEFT JOIN tblLocations AS l ON p.LocationID = l.LocationID "
    
    ' Build WHERE clause
    Select Case Me.cboSearchField
        Case "ID", "PersonID"
            ' exact match on the numeric key
            If IsNumeric(txt) Then
                strWhere = "WHERE p.PersonID = " & CLng(txt) & " "
            Else
                MsgBox "Please enter a numeric ID.", vbExclamation
                Exit Sub
            End If
            
        Case "LastName"
            strWhere = "WHERE p.LastName LIKE '*" & Replace(txt, "'", "''") & "*' "
            
        Case "FirstName"
            strWhere = "WHERE p.FirstName LIKE '*" & Replace(txt, "'", "''") & "*' "
            
        Case "FullName"
            strWhere = "WHERE (p.FirstName & ' ' & p.LastName) LIKE '*" & Replace(txt, "'", "''") & "*' "
            
        Case "BirthDate"
            If IsDate(txt) Then
                strWhere = "WHERE p.BirthDate = #" & Format(CDate(txt), "mm\/dd\/yyyy") & "# "
            Else
                strWhere = "WHERE Format(p.BirthDate,'mm/dd/yyyy') LIKE '*" & Replace(txt, "'", "''") & "*' "
            End If
            
        Case "Location"
            strWhere = _
              "WHERE l.City LIKE '*" & Replace(txt, "'", "''") & "*' OR " & _
                    "l.[StateProvince] LIKE '*" & Replace(txt, "'", "''") & "*' OR " & _
                    "l.Country LIKE '*" & Replace(txt, "'", "''") & "*' "
                    
        Case "Country"
            strWhere = "WHERE l.Country LIKE '*" & Replace(txt, "'", "''") & "*' "
            
        Case Else
            strWhere = ""
    End Select
    
    ' Finalize and run
    strSQL = strSQL & strWhere & "ORDER BY p.LastName, p.FirstName;"
    Me.lstStudents.RowSource = strSQL
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rs.recordCount = 0 Then
        MsgBox "No students found matching your search criteria.", vbInformation, "No Results"
    End If

CleanExit:
    On Error Resume Next
    rs.Close: Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error during search: " & Err.Number & " - " & Err.Description, vbExclamation, "Search Error"
    Resume CleanExit
End Sub


'=== Clear button ===
Private Sub btnClear_Click()
    Me.txtSearchText = Null
    Me.cboSearchField = "LastName"
    SetupListBox
End Sub

'=== Add New Student ===
Private Sub btnAddNew_Click()
    DoCmd.OpenForm "frmStudentNew", DataMode:=acFormAdd
End Sub

'=== Edit Selected Student ===
Private Sub btnEdit_Click()
    Dim pid As Long
    pid = Nz(Me.lstStudents.value, 0)
    If pid = 0 Then
        MsgBox "Please select a student to edit.", vbInformation, "No Selection"
        Exit Sub
    End If
    DoCmd.OpenForm _
      "frmStudentEdit", _
      DataMode:=acFormEdit, _
      openArgs:=CStr(pid)
End Sub

'=== View Natal Chart ===
Private Sub btnViewChart_Click()
    Dim pid As Long
    pid = Nz(Me.lstStudents.value, 0)
    If pid = 0 Then
        MsgBox "Please select a student to view chart.", vbInformation, "No Selection"
        Exit Sub
    End If
    If Not ChartGenerated(pid) Then
        MsgBox "Natal chart not yet generated. Please generate it first.", vbInformation, "No Chart"
        Exit Sub
    End If
    DoCmd.OpenForm "frmChartViewer", WhereCondition:="PersonID=" & pid
End Sub

'=== Helper to check chart status ===
Private Function ChartGenerated(PersonID As Long) As Boolean
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT NatalChartGenerated FROM tblPeople WHERE PersonID=" & PersonID, _
        dbOpenSnapshot)
    ChartGenerated = (Not rs.EOF And rs!NatalChartGenerated)
    rs.Close: Set rs = Nothing
End Function

'=== Close Form ===
Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub Form_Open(Cancel As Integer)
    SetupListBox
End Sub

'=== Double-click list opens record ===
Private Sub lstStudents_DblClick(Cancel As Integer)
'    If Nz(Me.lstStudents.Value, 0) <> 0 Then
'        DoCmd.OpenForm "frmStudentEdit", WhereCondition:="PersonID=" & Me.lstStudents.Value
'    End If

    btnEdit_Click
End Sub

