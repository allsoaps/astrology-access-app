VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmStudentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Form-level variables to track changes and control saving
Private m_OriginalValues As New Collection
Private m_AllowSave As Boolean
Private m_FormLoading As Boolean
Private m_PersonID As Long  ' Track the current PersonID


' === Form Events ===

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Only allow updates if explicitly triggered by Save button
    If Not m_AllowSave Then
        Cancel = True
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, response As Integer)
    ' Handle common form errors gracefully
    Select Case DataErr
        Case 2169 ' You can't save this record at this time
            response = acDataErrContinue
            MsgBox "Unable to save record. Please check your data and try again.", vbExclamation, "Save Error"
        Case Else
            response = acDataErrDisplay
    End Select
End Sub

' === Button Events ===

Private Sub btnViewSessions_Click()
    On Error GoTo ErrorHandler
    
    ' Check for a valid PersonID. This form is for editing, so it must exist.
    If IsNull(Me.txtPersonID.value) Or Me.txtPersonID.value = 0 Then
        MsgBox "Cannot view sessions for an unsaved student. Please save this record first.", vbExclamation, "Save Required"
        Exit Sub
    End If
    
    ' Check for unsaved changes before proceeding
    If HasChanges() Then
        MsgBox "You have unsaved changes. Please save the record before viewing sessions.", vbInformation, "Save Changes"
        Exit Sub
    End If
    
    ' Pass the context to the Session Manager form.
    ' Format: "PersonID=123"
    Dim openArgs As String
    openArgs = "PersonID=" & Me.txtPersonID.value
    
    ' Open the Session Manager form
    DoCmd.OpenForm "frmSessionManager", acNormal, , , , , openArgs
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred opening the Session Manager: " & Err.Description, vbCritical, "Error"
End Sub

Private Sub btnSave_Click()
    DoCmd.SetWarnings False
    
    On Error GoTo SaveError
    
    ' Validate required fields
    If Not ValidateForm() Then
        Exit Sub
    End If
    
    ' Check if any changes were made
    If Not HasChanges() Then
        MsgBox "No changes detected. Record was not updated.", vbInformation, "No Changes"
        Exit Sub
    End If
    
    ' Save the data to the database
    SaveStudentData
    
    ' Show success message
    MsgBox "Student information has been saved successfully.", vbInformation, "Save Successful"
    
    ' Refresh the original values after successful save
    StoreOriginalValues
    
    ' Update chart controls in case anything chart-related changed
    UpdateChartControls
    
    DoCmd.SetWarnings True
    Exit Sub

SaveError:
    DoCmd.SetWarnings True
    MsgBox "Error saving student information: " & Err.Description, vbCritical, "Save Error"
End Sub

Private Sub SaveStudentData()
    ' Save data from unbound controls to database
    On Error GoTo SaveDataError
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim LocationID As Long
    
    Set db = CurrentDb()
    
    ' First, handle the location
    LocationID = GetOrCreateLocationID()
    
    ' Update the person record
    sql = "SELECT * FROM tblPeople WHERE PersonID = " & Me.txtPersonID.value
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    
    If Not rs.EOF Then
        rs.Edit
        rs!FirstName = ProperCase(Me.txtFirstName)
        rs!LastName = ProperCase(Me.txtLastName)
        rs!BirthDate = Me.dtBirthDate.value
        rs!BirthTime = Me.dtBirthTime.value
        rs!LocationID = LocationID
        rs!Initials = Mid(ProperCase(Me.txtFirstName), 1, 1) & Mid(ProperCase(Me.txtLastName), 1, 1)
        rs!DateUpdated = Now()
        rs.Update
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

SaveDataError:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Private Function GetOrCreateLocationID() As Long
    ' Get existing location ID or create new location record
    On Error GoTo LocationIDError
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim newLocationID As Long
    
    Set db = CurrentDb()
    
    ' Check if location already exists
    sql = "SELECT LocationID FROM tblLocations WHERE " & _
          "City = '" & Replace(Me.txtCity.value, "'", "''") & "' AND " & _
          "Country = '" & Replace(Me.cboCountry.value, "'", "''") & "'"
    
    If Not IsNull(Me.txtStateProv.value) And Len(Trim(Me.txtStateProv.value)) > 0 Then
        sql = sql & " AND [StateProvince] = '" & Replace(Me.txtStateProv.value, "'", "''") & "'"
    Else
        sql = sql & " AND ([StateProvince] IS NULL OR [StateProvince] = '')"
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' --- PROACTIVE FIX for Quirk #9 ---
    ' Force Access to populate RecordCount before checking it
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    ' --- END FIX ---
    
    If rs.recordCount > 0 Then
        ' Location exists, return the ID
        GetOrCreateLocationID = rs!LocationID
    Else
        ' Create new location using dynaset recordset
        rs.Close
        Set rs = db.OpenRecordset("tblLocations", dbOpenDynaset)
        rs.AddNew
        rs!City = ProperCase(Me.txtCity.value)
        rs![StateProvince] = UCase(Me.txtStateProv.value)
        rs!Country = Me.cboCountry.value
        rs!latitude = Me.txtLatitude.value
        rs!longitude = Me.txtLongitude.value
        rs!DateCreated = Now()
        rs!DateUpdated = Now()
        rs.Update
        
        ' IMPORTANT: Move to the record we just created to get the ID
        rs.Bookmark = rs.LastModified
        newLocationID = rs!LocationID
        GetOrCreateLocationID = newLocationID
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
LocationIDError:
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
        On Error GoTo 0
    End If
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, "Error managing location: " & Err.Description
End Function

Private Sub btnCancel_Click()
    On Error GoTo CancelError
    
    ' Check if there are unsaved changes
    If HasChanges() Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You have unsaved changes. Are you sure you want to cancel?" & vbCrLf & _
                          "All changes will be lost.", vbQuestion + vbYesNo, "Confirm Cancel")
        
        If response = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Revert any unsaved changes
    If Me.Dirty Then
        Me.Undo
    End If
    
    ' Show confirmation message
    MsgBox "Changes have been cancelled. No updates were saved to the database.", _
           vbInformation, "Changes Cancelled"
    
    ' Close the form
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

CancelError:
    MsgBox "Error cancelling changes: " & Err.Description, vbCritical, "Cancel Error"
End Sub

Private Sub btnClose_Click()
    On Error GoTo CloseError
    
    ' Check if there are unsaved changes
    If HasChanges() And Me.Dirty Then
        Dim response As VbMsgBoxResult
        response = MsgBox("You have unsaved changes. Do you want to save them before closing?", _
                          vbQuestion + vbYesNoCancel, "Unsaved Changes")
        
        Select Case response
            Case vbYes
                ' Save changes
                Call btnSave_Click
                If HasChanges() Then  ' If save failed, don't close
                    Exit Sub
                End If
            Case vbNo
                ' Don't save, just close
                If Me.Dirty Then
                    Me.Undo
                End If
                MsgBox "Changes were not saved.", vbInformation, "Not Saved"
            Case vbCancel
                ' Don't close
                Exit Sub
        End Select
    End If
    
    ' Close the form
    DoCmd.Close acForm, Me.name, acSaveNo
    Exit Sub

CloseError:
    MsgBox "Error closing form: " & Err.Description, vbCritical, "Close Error"
End Sub

' === Utility Functions ===

Private Sub StoreOriginalValues()
    ' Store original values for comparison
    Set m_OriginalValues = New Collection
    
    On Error Resume Next
    
    ' Store the current values of all bound controls
    m_OriginalValues.Add Me.txtFirstName.value, "FirstName"
    m_OriginalValues.Add Me.txtLastName.value, "LastName"
    m_OriginalValues.Add Me.dtBirthDate.value, "BirthDate"
    m_OriginalValues.Add Me.dtBirthTime.value, "BirthTime"
    m_OriginalValues.Add Me.txtCity.value, "City"
    m_OriginalValues.Add Me.txtStateProv.value, "StateProvince"
    m_OriginalValues.Add Me.cboCountry.value, "Country"
    m_OriginalValues.Add Me.txtLatitude.value, "Latitude"
    m_OriginalValues.Add Me.txtLongitude.value, "Longitude"
    
    On Error GoTo 0
End Sub

Private Function HasChanges() As Boolean
    ' Compare current values with original values
    On Error Resume Next
    
    HasChanges = False
    
    ' Check each field for changes
    If Nz(Me.txtFirstName.value, "") <> Nz(m_OriginalValues("FirstName"), "") Then HasChanges = True
    If Nz(Me.txtLastName.value, "") <> Nz(m_OriginalValues("LastName"), "") Then HasChanges = True
    If Nz(Me.dtBirthDate.value, #1/1/1900#) <> Nz(m_OriginalValues("BirthDate"), #1/1/1900#) Then HasChanges = True
    If Nz(Me.dtBirthTime.value, #12:00:00 AM#) <> Nz(m_OriginalValues("BirthTime"), #12:00:00 AM#) Then HasChanges = True
    If Nz(Me.txtCity.value, "") <> Nz(m_OriginalValues("City"), "") Then HasChanges = True
    If Nz(Me.txtStateProv.value, "") <> Nz(m_OriginalValues("StateProvince"), "") Then HasChanges = True
    If Nz(Me.cboCountry.value, "") <> Nz(m_OriginalValues("Country"), "") Then HasChanges = True
    If Nz(Me.txtLatitude.value, 0) <> Nz(m_OriginalValues("Latitude"), 0) Then HasChanges = True
    If Nz(Me.txtLongitude.value, 0) <> Nz(m_OriginalValues("Longitude"), 0) Then HasChanges = True
    
    On Error GoTo 0
End Function

Private Function ValidateForm() As Boolean
    ' Validate required fields
    ValidateForm = True
    
    ' Check required fields
    If Len(Trim(Nz(Me.txtFirstName.value, ""))) = 0 Then
        MsgBox "First Name is required.", vbExclamation, "Validation Error"
        Me.txtFirstName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Len(Trim(Nz(Me.txtLastName.value, ""))) = 0 Then
        MsgBox "Last Name is required.", vbExclamation, "Validation Error"
        Me.txtLastName.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If IsNull(Me.dtBirthDate.value) Then
        MsgBox "Birth Date is required.", vbExclamation, "Validation Error"
        Me.dtBirthDate.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Validate birth date is not in the future
    If Me.dtBirthDate.value > Date Then
        MsgBox "Birth Date cannot be in the future.", vbExclamation, "Validation Error"
        Me.dtBirthDate.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Validate location fields
    If Len(Trim(Nz(Me.txtCity.value, ""))) = 0 Then
        MsgBox "City is required.", vbExclamation, "Validation Error"
        Me.txtCity.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    If Len(Trim(Nz(Me.cboCountry.value, ""))) = 0 Then
        MsgBox "Country is required.", vbExclamation, "Validation Error"
        Me.cboCountry.SetFocus
        ValidateForm = False
        Exit Function
    End If
    
    ' Validate coordinates if present
    If Not IsNull(Me.txtLatitude.value) Then
        If Me.txtLatitude.value < -90 Or Me.txtLatitude.value > 90 Then
            MsgBox "Latitude must be between -90 and 90 degrees.", vbExclamation, "Validation Error"
            Me.txtLatitude.SetFocus
            ValidateForm = False
            Exit Function
        End If
    End If
    
    If Not IsNull(Me.txtLongitude.value) Then
        If Me.txtLongitude.value < -180 Or Me.txtLongitude.value > 180 Then
            MsgBox "Longitude must be between -180 and 180 degrees.", vbExclamation, "Validation Error"
            Me.txtLongitude.SetFocus
            ValidateForm = False
            Exit Function
        End If
    End If
End Function

Private Sub UpdateDateUpdated()
    ' Update the DateUpdated field if it exists
    On Error Resume Next
    
    Dim db As DAO.Database
    Dim sql As String
    
    Set db = CurrentDb()
    
    ' Update the DateUpdated field in tblPeople
    sql = "UPDATE tblPeople SET DateUpdated = Now() WHERE PersonID = " & Me.txtPersonID.value
    db.Execute sql
    
    Set db = Nothing
    On Error GoTo 0
End Sub

' === Control Events for Change Tracking ===

Private Sub FirstName_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub LastName_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub BirthDate_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub BirthTime_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub City_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub StateProvince_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub Country_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub Latitude_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub Longitude_AfterUpdate()
    MarkFormDirty
End Sub

Private Sub MarkFormDirty()
    ' Ensure the form knows it has been modified
    If Not m_FormLoading Then
        Me.Dirty = True
    End If
End Sub

' === Location and Natal Chart Integration ===

Private Sub UpdateLocationInfo()
    ' This function can be called after coordinates are retrieved
    ' to update the location in tblLocations and link to the person
    
    On Error GoTo LocationError
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim LocationID As Long
    Dim sql As String
    
    Set db = CurrentDb()
    
    ' Check if location already exists
    sql = "SELECT LocationID FROM tblLocations WHERE " & _
          "City = '" & Replace(Me.txtCity.value, "'", "''") & "' AND " & _
          "Country = '" & Replace(Me.cboCountry.value, "'", "''") & "'"
    
    If Not IsNull(Me.txtStateProv.value) And Len(Trim(Me.txtStateProv.value)) > 0 Then
        sql = sql & " AND [StateProvince] = '" & Replace(Me.txtStateProv.value, "'", "''") & "'"
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    ' --- PROACTIVE FIX for Quirk #9 ---
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    ' --- END FIX ---
    
    If rs.recordCount > 0 Then
        ' Location exists, get the ID
        LocationID = rs!LocationID
    Else
        ' Create new location
        rs.Close
        Set rs = db.OpenRecordset("tblLocations", dbOpenDynaset)
        rs.AddNew
        rs!City = ProperCase(Me.txtCity.value)
        rs![StateProvince] = UCase(Me.txtStateProv.value)
        rs!Country = Me.cboCountry.value
        rs!latitude = Me.txtLatitude.value
        rs!longitude = Me.txtLongitude.value
        rs!DateCreated = Now()
        rs.Update
        ' Get the new ID after updating
        rs.Bookmark = rs.LastModified
        LocationID = rs!LocationID
    End If
    
    ' Update the person's location
    rs.Close
    sql = "UPDATE tblPeople SET LocationID = " & LocationID & " WHERE PersonID = " & Me.txtPersonID.value
    db.Execute sql
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

LocationError:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    MsgBox "Error updating location information: " & Err.Description, vbCritical, "Location Error"
End Sub


Private Sub btnGetCoordinates_Click()
    Dim strCity As String
    Dim strState As String
    Dim strCountry As String
    Dim strAddress As String
    Dim strResult As String
    Dim dblLat As Double
    Dim dblLng As Double

   
    ' Get values from form controls
    strCity = ProperCase(Nz(Me.txtCity, ""))
    strState = UCase(Nz(Me.txtStateProv, ""))
    strCountry = Nz(Me.cboCountry, "")
    
    ' Basic validation - ensure required fields are filled
    If Trim(strCity) = "" Or Trim(strCountry) = "" Then
        MsgBox "City and Country are required fields to get coordinates.", vbExclamation, "Missing Information"
        Exit Sub
    End If
    
    ' Special case: USA requires state
    If UCase(Trim(strCountry)) = "USA" And Trim(strState) = "" Then
        MsgBox "State is required for locations in the USA.", vbExclamation, "Missing Information"
        Exit Sub
    End If
    
    ' Format the address for the API
    If Trim(strState) = "" Then
        strAddress = Trim(strCity) & ", " & Trim(strCountry)
    Else
        strAddress = Trim(strCity) & ", " & Trim(strState) & ", " & Trim(strCountry)
    End If
    
    ' Show hourglass cursor
    DoCmd.Hourglass True
    
    ' Call the API function to get coordinates
    strResult = GetLatLong_LocationIQ(strAddress)
    
    ' Process the result
    If Left(strResult, 5) = "Lat: " Then
        ' Parse the latitude and longitude from the result
        dblLat = val(Mid(strResult, 6, InStr(strResult, ", Lng:") - 6))
        dblLng = val(Mid(strResult, InStr(strResult, ", Lng:") + 7))
        
        ' Update the form fields
        Me.txtLatitude = dblLat
        Me.txtLongitude = dblLng
        
        ' Mark form as having changes
        MarkFormDirty
        
        MsgBox "Coordinates retrieved successfully!", vbInformation, "Success"
    Else
        ' Show error message
        MsgBox "Error retrieving coordinates: " & strResult, vbExclamation, "API Error"
    End If
    
    ' Turn off hourglass cursor
    DoCmd.Hourglass False
End Sub


' UPDATE YOUR EXISTING Form_Load TO ADD THESE LINES
Private Sub Form_Load()
    ' Set up the Country combo box
    Me.cboCountry.RowSource = "SELECT Country_Name FROM lu_Country ORDER BY Country_Name"
    Me.cboCountry.ColumnCount = 1
    Me.cboCountry.BoundColumn = 1
    Me.cboCountry.ColumnWidths = "2.5in"
    Me.cboCountry.LimitToList = True
    
    m_FormLoading = True
    m_AllowSave = False
    m_PersonID = 0  ' ADD THIS LINE
    
    ' Load data if PersonID was passed via OpenArgs
    If Not IsNull(Me.openArgs) And Len(Me.openArgs) > 0 Then
        LoadStudentData CLng(Me.openArgs)
    End If
    
    ' Store original values for comparison
    StoreOriginalValues
    
    ' Initialize chart controls
    UpdateChartControls
    
    m_FormLoading = False
End Sub


Private Sub LoadStudentData(PersonID As Long)
    ' Load student data from database into unbound controls
    On Error GoTo LoadError
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb()
    
    '  Store the PersonID for later use
    m_PersonID = PersonID
    
    ' Query to get person and location data
    sql = "SELECT p.*, l.City, l.[StateProvince] AS StateProvince, l.Country, " & _
          "l.Latitude, l.Longitude " & _
          "FROM tblPeople p LEFT JOIN tblLocations l ON p.LocationID = l.LocationID " & _
          "WHERE p.PersonID = " & PersonID
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ' Populate the form controls
        Me.txtPersonID.value = rs!PersonID
        Me.txtFirstName.value = Nz(rs!FirstName, "")
        Me.txtLastName.value = Nz(rs!LastName, "")
        Me.dtBirthDate.value = rs!BirthDate
        Me.dtBirthTime.value = rs!BirthTime
        Me.txtCity.value = Nz(rs!City, "")
        Me.txtStateProv.value = Nz(rs!StateProvince, "")
        Me.cboCountry.value = Nz(rs!Country, "")
        Me.txtLatitude.value = Nz(rs!latitude, 0)
        Me.txtLongitude.value = Nz(rs!longitude, 0)
        
        ' Update natal chart checkbox if it exists
        On Error Resume Next
        Me.chkNatalChartGenerated.value = Nz(rs!NatalChartGenerated, False)
        On Error GoTo LoadError
    Else
        MsgBox "Student record not found.", vbExclamation, "Record Not Found"
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

LoadError:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    MsgBox "Error loading student data: " & Err.Description, vbCritical, "Load Error"
End Sub

' CHART MANAGEMENT
Private Sub UpdateChartControls()
    Dim chartExists As Boolean
    
    ' Check if we have a valid PersonID and if chart exists
    If m_PersonID > 0 Then
        chartExists = GetChartStatus(m_PersonID)
    Else
        chartExists = False
    End If
    
    ' Update checkbox state
    Me.chkNatalChartGenerated.value = chartExists
    
    ' Update button states
    Me.btnGenerateNatalChart.Enabled = (m_PersonID > 0)
    Me.btnViewNatalChart.Enabled = chartExists And (m_PersonID > 0)
    
    ' Update button caption based on chart existence
    If chartExists Then
        Me.btnGenerateNatalChart.Caption = "Regenerate Natal Chart"
    Else
        Me.btnGenerateNatalChart.Caption = "Generate Natal Chart"
    End If
End Sub

Private Function GetChartStatus(PersonID As Long) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb()
    sql = "SELECT NatalChartGenerated FROM tblPeople WHERE PersonID = " & PersonID
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        GetChartStatus = Nz(rs!NatalChartGenerated, False)
    Else
        GetChartStatus = False
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetChartStatus = False
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
End Function

Private Sub btnGenerateNatalChart_Click()
    On Error GoTo ErrorHandler
    
    ' Validate that we have a PersonID
    If m_PersonID = 0 Then
        MsgBox "Cannot generate natal chart for an unsaved student.", vbExclamation, "Save Required"
        Exit Sub
    End If
    
    ' First ensure all data is saved
    If HasChanges() Then
        MsgBox "Please save any changes before generating the natal chart.", vbInformation, "Save Required"
        Exit Sub
    End If
    
    ' Validate required data for chart generation
    If IsNull(Me.dtBirthDate.value) Or IsNull(Me.dtBirthTime.value) Then
        MsgBox "Birth Date and Birth Time are required for natal chart generation.", vbExclamation, "Missing Data"
        Exit Sub
    End If
    
    If IsNull(Me.txtLatitude.value) Or IsNull(Me.txtLongitude.value) Then
        MsgBox "Location coordinates are required for natal chart generation." & vbCrLf & _
               "Please use the 'Get Coordinates' button first.", vbExclamation, "Missing Coordinates"
        Exit Sub
    End If
    
    ' Show confirmation for regeneration
    Dim chartExists As Boolean
    chartExists = GetChartStatus(m_PersonID)
    
    If chartExists Then
        If MsgBox("A natal chart already exists for this student. Do you want to regenerate it?", _
                  vbQuestion + vbYesNo, "Regenerate Chart") = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Change cursor to hourglass
    DoCmd.Hourglass True
    
    ' Call the WORKING function from modSimpleChart
    Dim success As Boolean
    success = modSimpleChart.GenerateNatalChart(m_PersonID)
    
    ' Turn off hourglass
    DoCmd.Hourglass False
    
    If success Then
        ' Update the controls to reflect the new chart status
        UpdateChartControls
        MsgBox "Natal chart generated successfully!", vbInformation, "Chart Generated"
    Else
        MsgBox "Failed to generate natal chart. Please check the error messages.", vbExclamation, "Generation Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    DoCmd.Hourglass False
    MsgBox "Error generating natal chart: " & Err.Description, vbExclamation, "Error"
End Sub

'Private Sub btnViewNatalChart_Click()
'    On Error GoTo ErrorHandler
'
'    ' Validate that we have a PersonID and chart exists
'    If m_PersonID = 0 Then
'        MsgBox "No student record found.", vbExclamation, "No Record"
'        Exit Sub
'    End If
'
'    ' Double-check that chart exists
'    If Not GetChartStatus(m_PersonID) Then
'        MsgBox "No natal chart found for this student. Please generate a chart first.", vbExclamation, "No Chart"
'        Exit Sub
'    End If
'
'    ' Placeholder for chart viewing - will be implemented later
'    MsgBox "Chart viewing functionality will be implemented when chart display is ready." & vbCrLf & _
'           "PersonID: " & m_PersonID, vbInformation, "Chart View"
'
'    ' TODO: Implement chart viewing when ShowChartAspectGrid function is created
'    ' Example: modCharts.ShowChartAspectGrid m_PersonID, "Natal"
'
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "Error opening natal chart view: " & Err.Description, vbExclamation, "Error"
'End Sub

Private Sub btnViewNatalChart_Click()
    On Error GoTo ErrorHandler
    
    ' Validate that we have a PersonID and chart exists
    If m_PersonID = 0 Then
        MsgBox "No student record found.", vbExclamation, "No Record"
        Exit Sub
    End If
    
    ' Double-check that chart exists
    If Not GetChartStatus(m_PersonID) Then
        MsgBox "No natal chart found for this student. Please generate a chart first.", vbExclamation, "No Chart"
        Exit Sub
    End If
    
    ' Get the ChartID for this person's natal chart
    Dim chartID As Long
    chartID = GetNatalChartID(m_PersonID)
    
    If chartID = 0 Then
        MsgBox "Unable to find natal chart in database. Please regenerate the chart.", vbExclamation, "Chart Not Found"
        Exit Sub
    End If
    
    ' Open the aspect grid with the natal chart
    DoCmd.OpenForm "frmAspectGrid", acNormal, , , , , CStr(chartID)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error opening natal chart view: " & Err.Description, vbExclamation, "Error"
End Sub


Private Function GetNatalChartID(PersonID As Long) As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb()
    sql = "SELECT ChartID FROM tblCharts WHERE PersonID = " & PersonID & " AND ChartType = 'Natal'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        GetNatalChartID = rs!chartID
    Else
        GetNatalChartID = 0
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    GetNatalChartID = 0
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
End Function
