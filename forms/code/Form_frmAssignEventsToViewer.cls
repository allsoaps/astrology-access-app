VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAssignEventsToViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Form-level variables
Private m_PersonID As Long
Private Const TEMP_TABLE_NAME As String = "tmpEventAssignments"

'================'
'  FORM EVENTS   '
'================'

Private Sub Form_Load()
    On Error GoTo Err_Load
    
    ' Get the PersonID passed from the calling form
    If Not IsNull(Me.openArgs) Then
        m_PersonID = CLng(Me.openArgs)
    Else
        MsgBox "No viewer specified. Cannot assign events.", vbCritical, "Error"
        DoCmd.Close acForm, Me.name
        Exit Sub
    End If
    
    ' Set up the temporary data source for the subform
    SetupTempDataSource
    
    Exit Sub
    
Err_Load:
    MsgBox "Error loading event assignment form: " & Err.Description, vbCritical
    CleanupTempTable
End Sub

Private Sub Form_Close()
    ' Always delete the temporary table when the form closes
    CleanupTempTable
End Sub

'================'
' BUTTON EVENTS  '
'================'

Private Sub btnSaveAssignments_Click()
    On Error GoTo Err_Save
    
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' Force the subform to save any pending checkbox changes to the temp table
    If Me.subformEvents.Form.Dirty Then
        Me.subformEvents.Form.Dirty = False
    End If
    
    ' First, delete all existing assignments for this person
    db.Execute "DELETE FROM tblAssignments WHERE PersonID = " & m_PersonID, dbFailOnError
    
    ' Now, insert new assignments based on the checked items in the temp table
    db.Execute "INSERT INTO tblAssignments ( PersonID, EventID, DateCreated, DateUpdated ) " & _
               "SELECT " & m_PersonID & ", EventID, Now(), Now() " & _
               "FROM " & TEMP_TABLE_NAME & " " & _
               "WHERE IsAssigned = True;", dbFailOnError
               
    MsgBox "Event assignments have been updated.", vbInformation, "Save Successful"
    DoCmd.Close acForm, Me.name
    
    Set db = Nothing
    Exit Sub
    
Err_Save:
    MsgBox "An error occurred saving assignments: " & Err.Description, vbExclamation
    Set db = Nothing
End Sub

Private Sub btnCancel_Click()
    ' No need to ask to save, as all changes are in a temp table
    DoCmd.Close acForm, Me.name
End Sub

'================'
' HELPER SUBS    '
'================'

Private Sub SetupTempDataSource()
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' 1. Ensure old temp table is gone
    CleanupTempTable
    
    ' 2. Create the temp table structure
    db.Execute "CREATE TABLE " & TEMP_TABLE_NAME & " (EventID LONG PRIMARY KEY, IsAssigned YESNO);"
    
    ' 3. Populate it with all events, initially unchecked
    db.Execute "INSERT INTO " & TEMP_TABLE_NAME & " ( EventID, IsAssigned ) " & _
               "SELECT EventID, False FROM tblEvents;"
               
    ' 4. Update the temp table, checking the boxes for events this person is already assigned to
    db.Execute "UPDATE " & TEMP_TABLE_NAME & " INNER JOIN tblAssignments " & _
               "ON " & TEMP_TABLE_NAME & ".EventID = tblAssignments.EventID " & _
               "SET " & TEMP_TABLE_NAME & ".IsAssigned = True " & _
               "WHERE tblAssignments.PersonID = " & m_PersonID & ";"
               
    ' 5. Create the final SQL to bind the subform
    Dim sql As String
    sql = "SELECT t.IsAssigned, e.EventName, e.EventDate " & _
          "FROM tblEvents AS e INNER JOIN " & TEMP_TABLE_NAME & " AS t ON e.EventID = t.EventID " & _
          "ORDER BY e.EventDate DESC;"
          
    ' 6. Set the subform's RecordSource
    Me.subformEvents.Form.RecordSource = sql
    
    Set db = Nothing
End Sub

Private Sub CleanupTempTable()
    On Error Resume Next
    CurrentDb.Execute "DROP TABLE " & TEMP_TABLE_NAME & ";"
    On Error GoTo 0
End Sub


