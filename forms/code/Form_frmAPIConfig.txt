VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAPIConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Solar Flare Form Module — Fixed Form_Load binding and LoadAPISettings procedure
Option Compare Database
Option Explicit

' Runs when the form loads — make sure On Load is set to [Event Procedure]
Private Sub Form_Load()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb()

    ' LocationIQ API key
    Set rs = db.OpenRecordset("SELECT SettingValue FROM tblSettings WHERE SettingName = 'LocationIQ_API_Key'")
    If Not rs.EOF Then Me.txtLocationIQKey = rs!SettingValue
    rs.Close

    ' Moon Phase API URL
    Set rs = db.OpenRecordset("SELECT SettingValue FROM tblSettings WHERE SettingName = 'Moon_Phase_API_URL'")
    If Not rs.EOF Then Me.txtMoonPhaseURL = rs!SettingValue
    rs.Close

    Set rs = Nothing
    Set db = Nothing
End Sub

' Save button handler — unchanged
Private Sub btnSaveAPIConfig_Click()
    If SaveAPISettings() Then
        MsgBox "API settings saved successfully.", vbInformation, "Success"
        DoCmd.Close acForm, Me.name
    End If
End Sub

' Persists settings back to tblSettings — unchanged
Private Function SaveAPISettings() As Boolean
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM tblSettings WHERE SettingName = 'LocationIQ_API_Key'")
    If rs.EOF Then
        rs.AddNew
        rs!settingName = "LocationIQ_API_Key"
        rs!SettingCategory = "API"
        rs!SettingDescription = "API Key for LocationIQ geocoding service"
        rs!DataType = "Text"
    Else
        rs.Edit
    End If
    rs!SettingValue = Me.txtLocationIQKey
    rs.Update: rs.Close

    Set rs = db.OpenRecordset("SELECT * FROM tblSettings WHERE SettingName = 'Moon_Phase_API_URL'")
    If rs.EOF Then
        rs.AddNew
        rs!settingName = "Moon_Phase_API_URL"
        rs!SettingCategory = "API"
        rs!SettingDescription = "URL for Moon Phase API service"
        rs!DataType = "Text"
    Else
        rs.Edit
    End If
    rs!SettingValue = Me.txtMoonPhaseURL
    rs.Update: rs.Close

    SaveAPISettings = True
    Exit Function

ErrorHandler:
    MsgBox "Error saving settings: " & Err.Description, vbExclamation, "Error"
    SaveAPISettings = False
End Function

' Cancel button — unchanged
Private Sub btnCancel_Click()
    If Me.Dirty Then
        Dim response As Integer
        response = MsgBox("Do you want to save your changes before closing?", vbQuestion + vbYesNoCancel, "Save Changes?")
        Select Case response
            Case vbYes: btnSaveAPIConfig_Click: If Me.Dirty Then Exit Sub
            Case vbCancel: Exit Sub
        End Select
    End If
    DoCmd.Close acForm, Me.name
End Sub

' --- Instructions:
' 1. In the form's Property Sheet ? Event ? On Load, choose [Event Procedure].
' 2. Remove any stray text from the On Load property.
' 3. Compile the VBA project (Debug ? Compile).
' 4. Test the form; the settings fields should populate on load.
