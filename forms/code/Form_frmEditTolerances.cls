VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEditTolerances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Name for our temporary, local table
Private Const TEMP_TABLE_NAME As String = "tmpTolerances"

'================'
' FORM EVENTS    '
'================'

Private Sub Form_Load()
    On Error GoTo Err_Load
    
    ' This subroutine will create a temporary copy of the data
    ' and bind the form to it.
    SetupTempDataSource
    
    ' --- NEW CODE: PROGRAMMATICALLY RESIZE THE FORM ---
    ' This is the most reliable way to eliminate extra whitespace.
    
    Dim headerHeight As Long
    Dim detailHeight As Long
    Dim footerHeight As Long
    Dim recordCount As Long
    Dim totalHeight As Long
    Dim windowBorder As Long
    
    ' Get the height of each section in Twips (1440 twips = 1 inch)
    headerHeight = Me.Section(acHeader).Height
    detailHeight = Me.Section(acDetail).Height
    footerHeight = Me.Section(acFooter).Height
    
    ' Get a reliable count of records
    recordCount = Me.RecordsetClone.recordCount
    
    ' Add a little extra for the window's title bar and borders
    windowBorder = 350 ' Adjust as needed
    
    ' Calculate the total height required
    totalHeight = headerHeight + footerHeight + (detailHeight * recordCount) + windowBorder + 600
    
    ' Use MoveSize to set the form's height precisely
    DoCmd.MoveSize , , , totalHeight
    ' --- END NEW CODE ---
    
    Exit Sub
    
Err_Load:
    MsgBox "An error occurred loading the form: " & Err.Description, vbCritical, "Load Error"
    ' Clean up and close if loading fails
    CleanupTempTable
    DoCmd.Close acForm, Me.name
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' This event fires right before Access saves a changed record TO THE TEMP TABLE.
    ' We still want to update the DateChanged field in our temp table.
    Me.DateChanged = Now()
End Sub

Private Sub Form_Close()
    ' No matter how the form is closed, always delete the temporary table.
    CleanupTempTable
End Sub


'================'
' BUTTON EVENTS  '
'================'

Private Sub btnSave_Click()
    On Error GoTo Err_Save
    
    ' 1. Make sure any final change to the current record is saved to the temp table.
    If Me.Dirty Then
        Me.Dirty = False
    End If
    
    ' 2. Now, run an update query to copy the changes from the temp table
    ' back to the real tblTolerances.
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' This query joins the real table to the temp table on their ID
    ' and updates the two fields we care about.
    Dim sql As String
    sql = "UPDATE tblTolerances INNER JOIN " & TEMP_TABLE_NAME & " " & _
          "ON tblTolerances.ID = " & TEMP_TABLE_NAME & ".ID " & _
          "SET tblTolerances.ToleranceScore = " & TEMP_TABLE_NAME & ".ToleranceScore, " & _
          "    tblTolerances.DateChanged = " & TEMP_TABLE_NAME & ".DateChanged;"
          
    db.Execute sql, dbFailOnError
    
    Set db = Nothing
    
    MsgBox "Tolerance settings have been saved.", vbInformation, "Save Successful"
    DoCmd.Close acForm, Me.name
    
    Exit Sub
    
Err_Save:
    MsgBox "An error occurred while trying to save: " & Err.Description, vbExclamation, "Save Error"
End Sub

Private Sub btnCancel_Click()
    ' If the user has made any changes, the form will be "dirty".
    If Me.Dirty Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You have unsaved changes. Do you want to save them before closing?", _
                          vbQuestion + vbYesNoCancel, "Save Changes?")
                          
        Select Case response
            Case vbYes
                ' User wants to save. Call the save button's logic.
                btnSave_Click
                
            Case vbNo
                ' User wants to discard changes. Simply close the form.
                DoCmd.Close acForm, Me.name
                
            Case vbCancel
                ' User canceled the close operation. Do nothing and stay on the form.
                Exit Sub
        End Select
    Else
        ' If the form is not dirty, just close.
        DoCmd.Close acForm, Me.name
    End If
End Sub


'================'
' HELPER SUBS    '
'================'

Private Sub SetupTempDataSource()
    Dim db As DAO.Database
    Set db = CurrentDb()
    
    ' 1. Make sure a temp table from a previous crash doesn't exist.
    CleanupTempTable
    
    ' 2. Create a new temporary table by copying the structure AND data
    ' from the real tblTolerances.
    db.Execute "SELECT * INTO " & TEMP_TABLE_NAME & " FROM tblTolerances;"
    
    ' 3. Set the form's RecordSource to the new temporary table.
    Me.RecordSource = TEMP_TABLE_NAME
End Sub

Private Sub CleanupTempTable()
    ' Deletes the temporary table.
    On Error Resume Next
    CurrentDb.Execute "DROP TABLE " & TEMP_TABLE_NAME & ";"
    On Error GoTo 0
End Sub


