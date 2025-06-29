VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssignedViewerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Form: frmAssignedViewerReport
' Purpose: Shows a read-only list of viewers assigned to an event and their session count.
' This version uses the most reliable VBA-centric method to build the data.

Private m_EventID As Long
Private m_EventName As String
Private Const TEMP_TABLE_NAME As String = "tmpViewerReport"

'================'
' FORM EVENTS    '
'================'

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Set default title
    Me.lblTitle.Caption = "Assigned Viewers"
    
    ' Parse OpenArgs to get EventID and EventName
    If Not IsNull(Me.openArgs) And Len(Me.openArgs) > 0 Then
        Dim args() As String
        args = Split(Me.openArgs, "|")
        
        If UBound(args) >= 0 Then m_EventID = CLng(args(0))
        If UBound(args) >= 1 Then
            m_EventName = args(1)
            If Len(m_EventName) > 0 Then Me.lblTitle.Caption = "Viewers for: " & m_EventName
        End If
        
        ' Load the data now that we have the EventID
        If m_EventID > 0 Then
            BuildAndShowViewerData
        End If
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error loading assigned viewers report: " & Err.Description, vbCritical, "Form Load Error"
End Sub

Private Sub Form_Close()
    ' Ensure the temporary table is deleted when the form closes
    CleanupTempTable
End Sub

'================'
' BUTTON EVENTS  '
'================'

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.name
End Sub

'================'
' HELPER SUBS    '
'================'

Private Sub BuildAndShowViewerData()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsTemp As DAO.Recordset
    Dim sql As String
    
    ' Use a dictionary to stage the data in memory
    Dim sessionCounts As Object 'Scripting.Dictionary
    Set sessionCounts = CreateObject("Scripting.Dictionary")
    
    Dim recordsAdded As Long
    
    Set db = CurrentDb()
    recordsAdded = 0
    
    ' --- STAGE 1: Get ALL session counts for this event in ONE pass ---
    ' This is efficient and avoids running DCount in a loop.
    sql = "SELECT PersonID, Count(SessionID) AS NumSessions " & _
          "FROM tblSessions " & _
          "WHERE EventID = " & m_EventID & " " & _
          "GROUP BY PersonID"
          
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Load the session counts into our fast in-memory dictionary
    If Not rs.EOF Then
        Do While Not rs.EOF
            sessionCounts.Add key:=rs!PersonID, Item:=rs!NumSessions
            rs.MoveNext
        Loop
    End If
    rs.Close

    ' --- STAGE 2: Get the list of ASSIGNED viewers ---
    ' This query is now very simple and reliable.
    sql = "SELECT p.PersonID, p.FirstName, p.LastName " & _
          "FROM tblPeople AS p INNER JOIN tblAssignments AS a ON p.PersonID = a.PersonID " & _
          "WHERE a.EventID = " & m_EventID & " " & _
          "ORDER BY p.LastName, p.FirstName;"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' Check if there are any assigned viewers
    If rs.EOF Then
        ' No viewers found, prompt to assign them
        HandleNoViewers
        GoTo CleanExit
    End If

    ' --- STAGE 3: Build the temporary table ---
    CleanupTempTable
    db.Execute "CREATE TABLE " & TEMP_TABLE_NAME & " " & _
               "(FirstName TEXT(50), LastName TEXT(50), SessionCount LONG);"
    
    Set rsTemp = db.OpenRecordset(TEMP_TABLE_NAME, dbOpenDynaset)
    
    ' Loop through the assigned viewers, look up their session count from the dictionary,
    ' and write the final record to our temp table.
    Do While Not rs.EOF
        Dim currentCount As Long
        
        ' Check if the person exists in our session counts dictionary
        If sessionCounts.Exists(rs!PersonID) Then
            currentCount = sessionCounts(rs!PersonID)
        Else
            ' If not, their count is 0
            currentCount = 0
        End If
        
        rsTemp.AddNew
            rsTemp!FirstName = rs!FirstName
            rsTemp!LastName = rs!LastName
            rsTemp!sessionCount = currentCount
        rsTemp.Update
        
        recordsAdded = recordsAdded + 1
        rs.MoveNext
    Loop
    rsTemp.Close

    ' --- STAGE 4: Bind the form to the completed temp table ---
    Me.RecordSource = "SELECT * FROM " & TEMP_TABLE_NAME & " ORDER BY LastName, FirstName;"

CleanExit:
    On Error Resume Next
    rs.Close
    rsTemp.Close
    Set rs = Nothing
    Set rsTemp = Nothing
    Set db = Nothing
    Set sessionCounts = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error building viewer data: " & Err.Description, vbCritical, "Error"
    Resume CleanExit
End Sub

Private Sub HandleNoViewers()
    ' This sub contains the logic for when no viewers are assigned
    Me.Visible = False
    
    Dim response As VbMsgBoxResult
    response = MsgBox("No viewers have been assigned to this event." & vbCrLf & vbCrLf & _
                      "Would you like to assign viewers now?", _
                      vbQuestion + vbYesNo, "Assign Viewers?")
                      
    If response = vbYes Then
        Dim openArgs As String
        openArgs = m_EventID & "|" & m_EventName
        DoCmd.OpenForm "frmViewerAssignList", acViewNormal, , , , acDialog, openArgs
    End If
    
    ' Close this form
    DoCmd.Close acForm, Me.name
End Sub

Private Sub CleanupTempTable()
    ' Deletes the temporary table. Use error resume in case it doesn't exist.
    On Error Resume Next
    CurrentDb.Execute "DROP TABLE " & TEMP_TABLE_NAME & ";"
    On Error GoTo 0
End Sub


