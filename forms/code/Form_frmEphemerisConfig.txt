VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEphemerisConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnBrowseEphe_Click()
    ' Open folder browser dialog for Ephemeris Path
    Dim folderPath As String
    folderPath = BrowseForFolder("Select Ephemeris Files Folder")
    
    If folderPath <> "" Then
        Me.EphePath = folderPath
    End If
End Sub

Private Sub btnBrowseDLL_Click()
    ' Open folder browser dialog for DLL Path
    Dim folderPath As String
    folderPath = BrowseForFolder("Select DLL Folder")
    
    If folderPath <> "" Then
        Me.DLLPath = folderPath
    End If

End Sub

Private Sub btnConfigCheck_Click()
    ' Check db entries resolve for Swiss Ephemeris
  InitSwissEph

  If Not InitSwissEph() Then
     MsgBox "Initialization Failed", vbInformation, "Config Check Failed"
  Else
     MsgBox "Initializtion Successful", vbInformation, "Success"
  End If
End Sub

Private Sub btnSaveSwiss_Click()
    ' Save any unsaved changes
    If Me.Dirty Then
        Me.Dirty = False
    End If
    
    ' Save coordinate system setting
    SaveCoordinateSystemSetting
    
    ' Perform validation if needed
    If Not DirExists(Me.EphePath) Then
        MsgBox "Warning: Ephemeris folder path doesn't exist or is not accessible.", vbExclamation, "Path Validation"
    End If
    
    If Not DirExists(Me.DLLPath) Then
        MsgBox "Warning: DLL folder path doesn't exist or is not accessible.", vbExclamation, "Path Validation"
    End If
    
    ' Show confirmation message
    MsgBox "Swiss Ephemeris settings saved successfully.", vbInformation, "Save Complete"
    
    ' Close the form and return to previous form
    DoCmd.Close acForm, Me.name
End Sub

' Helper function to check if directory exists
Private Function DirExists(ByVal path As String) As Boolean
    On Error Resume Next
    DirExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

' Helper function to check if file exists
Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(path) <> "")
    On Error GoTo 0
End Function

Private Sub btnCancel_Click()
    ' Cancel any changes
    If Me.Dirty Then
        If MsgBox("Discard changes?", vbQuestion + vbYesNo, "Cancel") = vbYes Then
            Me.Undo
        Else
            Exit Sub
        End If
    End If
    
    ' Close the form
    DoCmd.Close acForm, Me.name
End Sub

' Helper function to browse for folder
Private Function BrowseForFolder(Optional prompt As String = "Select Folder") As String
    Dim shell As Object
    Dim folder As Object
    
    On Error Resume Next
    
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, prompt, 0)
    
    If Not folder Is Nothing Then
        BrowseForFolder = folder.Self.path
    Else
        BrowseForFolder = ""
    End If
    
    Set folder = Nothing
    Set shell = Nothing
End Function

' Helper function to browse for file
Private Function BrowseForFile(Optional fileFilter As String = "All Files (*.*)|*.*") As String
    Dim fd As Object
    
    ' 3 is the value of msoFileDialogFilePicker
    Set fd = Application.FileDialog(3)
    
    With fd
        .filters.Clear
        .Title = "Select File"
        
        ' Parse filter string (similar to Common Dialog)
        Dim filters() As String
        filters = Split(fileFilter, "|")
        
        Dim i As Integer
        For i = 0 To UBound(filters) Step 2
            If i + 1 <= UBound(filters) Then
                .filters.Add filters(i), filters(i + 1)
            End If
        Next i
        
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseForFile = .SelectedItems(1)
        Else
            BrowseForFile = ""
        End If
    End With
    
    Set fd = Nothing
End Function

Private Sub Form_Current()
    ' Load coordinate system setting when form opens or moves to a record
    If Not Me.NewRecord Then
        LoadCoordinateSystemSetting
    End If
End Sub

Private Sub LoadCoordinateSystemSetting()
    On Error GoTo ErrorHandler

    Dim CoordinateSystem As String
    CoordinateSystem = modSwissItems.GetCoordinateSystemSetting()

    ' Set the option group value based on database value
    ' Assuming: 1 = Geocentric, 2 = Heliocentric
    If UCase(CoordinateSystem) = "HELIOCENTRIC" Then
        Me.grpCoordinateSystem.value = 2  ' Replace with your actual Option Group name
    Else
        Me.grpCoordinateSystem.value = 1  ' Geocentric (default)
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Error in LoadCoordinateSystemSetting: " & Err.Description
    ' Default to Geocentric
    Me.grpCoordinateSystem.value = 1
End Sub

' Updated function to get coordinate system from Option Group
Public Function GetCoordinateSystem() As String
    ' Check the Option Group value
    If Me.grpCoordinateSystem.value = 2 Then
        GetCoordinateSystem = "Heliocentric"
    Else
        GetCoordinateSystem = "Geocentric"
    End If
End Function

' You can also add an event for when the user changes the selection
Private Sub grpCoordinateSystem_AfterUpdate()
    ' Optional: Provide immediate feedback when user changes selection
    ' Debug.Print "Coordinate system changed to: " & GetCoordinateSystem()
End Sub

' Form load - set up defaults if needed
Private Sub Form_Load()
    ' If no records exist, create default
    If Me.NewRecord Then
        ' Set default paths based on application path
        Dim appPath As String
        appPath = CurrentProject.path
        
        Me.EphePath = appPath & "\ephe"
        Me.DLLPath = appPath & "\swedll64.dll"
        Me.DefaultHouseSystem = 1  ' Placidus
        Me.DefaultOrbs = True
        Me.IncludeAsteroids = False
        Me.UseTopocentric = False
        Me.UseTrue = True
        Me.optGeocentric.value = True
    End If
End Sub

' Set the coordinate system
Public Sub SetCoordinateSystem(systemType As String)
    If UCase(systemType) = "HELIOCENTRIC" Then
        Me.optHeliocentric.value = True
        Me.optGeocentric.value = False
    Else
        Me.optGeocentric.value = True
        Me.optHeliocentric.value = False
    End If
End Sub

' Save setting to database (add to your existing save routine)
Private Sub SaveCoordinateSystemSetting()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tblSwissEphSettings", dbOpenDynaset)
    
    If rs.recordCount > 0 Then
        rs.MoveFirst
        rs.Edit
    Else
        rs.AddNew
    End If
    
    rs!CoordinateSystem = GetCoordinateSystem()
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub


' Helper function to get coordinate system setting from database
Private Function GetCoordinateSystemSetting() As String
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim result As String
    
    Set db = CurrentDb()
    
    ' Try to get the setting from the database
    sql = "SELECT CoordinateSystem FROM tblSwissEphSettings"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        result = Nz(rs!CoordinateSystem, "Geocentric")
    Else
        ' No record found, return default
        result = "Geocentric"
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    GetCoordinateSystemSetting = result
    Exit Function
    
ErrorHandler:
    ' If there's any error accessing the database, default to Geocentric
    GetCoordinateSystemSetting = "Geocentric"
    
    ' Clean up - Remove the rs.state check since DAO doesn't have State property
    If Not rs Is Nothing Then
        On Error Resume Next  ' Ignore errors when closing
        rs.Close
        On Error GoTo 0      ' Resume normal error handling
        Set rs = Nothing
    End If
    Set db = Nothing
    
#If DEBUG_MODE Then
    Debug.Print "Error getting coordinate system setting, defaulting to Geocentric: " & Err.Description
#End If
End Function
