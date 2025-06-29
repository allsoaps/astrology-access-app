VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEventSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
#Const DEBUG_MODE = False

'=== Initialize form ===
Private Sub Form_Load()
    Me.cboSearchField = ""
    Me.txtSearchText = ""
End Sub

'=== Populate the list with top 100 records ===
Private Sub SetupListBox()
    With Me.lstEvents
        .RowSourceType = "Table/Query"
        .ColumnCount = 7
        .BoundColumn = 1
        .ColumnHeads = False
        .ColumnWidths = ".70cm;6.5cm;2.25cm;2.35cm;1.95cm;1.75cm;1.75cm"
        .RowSource = _
            "Select e.EventID AS 'EventID', e.EventName AS 'Event Name', et.Event_Desc AS 'Event Type', e.EventChartGenerated AS 'Chart Ready?', e.EventDate AS 'Event Date', e.TotalScore AS 'Total Score', e.OverUnder AS 'Over/Under' " & _
            "FROM tblEvents as e LEFT JOIN lu_EventType as et on e.EventType = et.ID;"
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
        "Select e.EventID AS [Event ID], e.EventName AS [Event Name], et.Event_Desc AS [EventType], " & _
        "e.EventChartGenerated AS [Chart Ready?], e.EventDate AS [EventDate] " & _
        "FROM tblEvents as e LEFT JOIN lu_EventType as et on e.EventType = et.ID "
    
    ' Build WHERE clause
    Select Case Me.cboSearchField
        Case "ID", "EventID"
            ' exact match on the numeric key
            If IsNumeric(txt) Then
                strWhere = "WHERE e.EventID = " & CLng(txt) & " "
            Else
                MsgBox "Please enter a numeric EventID.", vbExclamation
                Exit Sub
            End If
            
        Case "EventName"
            strWhere = "WHERE e.EventName LIKE '*" & Replace(txt, "'", "''") & "*' "
            
        Case "EventType"
            strWhere = "WHERE et.Event_Desc LIKE '*" & Replace(txt, "'", "''") & "*' "
            
        Case "EventDate"
            If IsDate(txt) Then
                strWhere = "WHERE e.EventDate = #" & Format(CDate(txt), "mm\/dd\/yyyy") & "# "
            Else
                strWhere = "WHERE Format(e.EventDate,'mm/dd/yyyy') LIKE '*" & Replace(txt, "'", "''") & "*' "
            End If
            
        Case Else
            strWhere = ""
    End Select
    
    ' Finalize and run
    strSQL = strSQL & strWhere & "ORDER BY e.EventID;"
    Me.lstEvents.RowSource = strSQL
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rs.recordCount = 0 Then
        MsgBox "No events found matching your search criteria.", vbInformation, "No Results"
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
    Me.cboSearchField = "EventType"
    SetupListBox
End Sub

'=== Add New Event ===
Private Sub btnAddNew_Click()
    DoCmd.OpenForm "frmEventNew", DataMode:=acFormAdd
End Sub

Private Sub btnEdit_Click()
    Dim pid As Long
    pid = Nz(Me.lstEvents.value, 0)
    If pid = 0 Then
        MsgBox "Please select an event to edit.", vbInformation, "No Selection"
        Exit Sub
    End If
    DoCmd.OpenForm _
      "frmEventEdit", _
      DataMode:=acFormEdit, _
      openArgs:=CStr(pid)
End Sub

'=== View Natal Chart ===
Private Sub btnViewChart_Click()
'    Dim pid As Long
'    pid = Nz(Me.lstEvents.Value, 0)
'    If pid = 0 Then
'        MsgBox "Please select an event to view chart.", vbInformation, "No Selection"
'        Exit Sub
'    End If
'    If Not ChartGenerated(pid) Then
'        MsgBox "Natal chart not yet generated. Please generate it first.", vbInformation, "No Chart"
'        Exit Sub
'    End If
'    DoCmd.OpenForm "frmChartViewer", WhereCondition:="PersonID=" & pid

'Private Sub cmdViewAspectGrid_Click()
    Dim chartID As Long
    
    ' Get current chart ID (adjust this based on your form)
'    chartID = Me.chartID
chartID = 11
    
    If chartID > 0 Then
        ShowChartAspectGrid chartID
    Else
        MsgBox "Please select a chart first.", vbInformation
    End If
'End Sub

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

'=== Double-click list opens record ===
Private Sub lstEvents_DblClick(Cancel As Integer)
    btnEdit_Click
End Sub

Private Sub Form_Open(Cancel As Integer)
    SetupListBox
End Sub

' Add this code to frmEventSearch after the existing button click events

Private Sub btnViewers_Click()
    On Error GoTo ErrorHandler
    
    ' Get the selected EventID from the list box
    Dim selectedEventID As Long
    selectedEventID = Nz(Me.lstEvents.value, 0)
    
    If selectedEventID = 0 Then
        MsgBox "Please select an event from the list to assign viewers.", vbInformation, "No Selection"
        Exit Sub
    End If
    
    ' Get the Event Name for the selected event
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim eventName As String
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT EventName FROM tblEvents WHERE EventID = " & selectedEventID, dbOpenSnapshot)
    
    If Not rs.EOF Then
        eventName = Nz(rs!eventName, "")
    Else
        eventName = ""
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    ' Pass the EventID and EventName to the viewer assignment form via OpenArgs
    ' The format is "EventID|EventName" (same format used in other forms)
    Dim openArgs As String
    openArgs = selectedEventID & "|" & eventName
    
    ' Open the viewer assignment form (frmViewerAssignList) in Dialog mode
    ' This form allows users to assign/unassign viewers and shows current assignments
    DoCmd.OpenForm "frmViewerAssignList", acViewNormal, , , , acDialog, openArgs
    
    ' Optionally refresh the list after viewer assignment changes
    ' (in case any event data was updated)
    Me.lstEvents.Requery
    
    Exit Sub

ErrorHandler:
    ' Clean up resources
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    
    MsgBox "An unexpected error occurred while trying to open the viewer assignment form." & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Error"
End Sub

