VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEventNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
#Const DEBUG_MODE = False

' Swiss Ephemeris declarations
' Constants for planets and calculation flags
Private Const SE_SUN As Long = 0
Private Const SE_MOON As Long = 1
Private Const SE_MERCURY As Long = 2
Private Const SE_VENUS As Long = 3
Private Const SE_MARS As Long = 4
Private Const SE_JUPITER As Long = 5
Private Const SE_SATURN As Long = 6
Private Const SE_URANUS As Long = 7
Private Const SE_NEPTUNE As Long = 8
Private Const SE_PLUTO As Long = 9

' Calculation flags
Private Const SEFLG_SPEED As Long = 256 ' Return speed values
Private Const SEFLG_SWIEPH As Long = 2 ' Use Swiss Ephemeris

' Swiss Ephemeris DLL declarations
#If VBA7 Then
    #If Win64 Then
        ' 64-bit declarations
        Private Declare PtrSafe Function swe_calc_ut Lib "swedll64.dll" ( _
            ByVal tjd_ut As Double, _
            ByVal ipl As Long, _
            ByVal iflag As Long, _
            ByRef xx As Double, _
            ByRef serr As String) As Long
            
        Private Declare PtrSafe Function swe_set_ephe_path Lib "swedll64.dll" ( _
            ByVal path As String) As Long
            
        Private Declare PtrSafe Function swe_close Lib "swedll64.dll" () As Long
        
        Private Declare PtrSafe Function swe_houses Lib "swedll64.dll" ( _
            ByVal tjd_ut As Double, _
            ByVal geolat As Double, _
            ByVal geolon As Double, _
            ByVal hsys As Long, _
            ByRef cusps As Double, _
            ByRef ascmc As Double) As Long
    #Else
        ' 32-bit declarations with PtrSafe (VBA7 but not 64-bit)
        Private Declare PtrSafe Function swe_calc_ut Lib "swedll32.dll" ( _
            ByVal tjd_ut As Double, _
            ByVal ipl As Long, _
            ByVal iflag As Long, _
            ByRef xx As Double, _
            ByRef serr As String) As Long
            
        Private Declare PtrSafe Function swe_set_ephe_path Lib "swedll32.dll" ( _
            ByVal path As String) As Long
            
        Private Declare PtrSafe Function swe_close Lib "swedll32.dll" () As Long
        
        Private Declare PtrSafe Function swe_houses Lib "swedll32.dll" ( _
            ByVal tjd_ut As Double, _
            ByVal geolat As Double, _
            ByVal geolon As Double, _
            ByVal hsys As Long, _
            ByRef cusps As Double, _
            ByRef ascmc As Double) As Long
    #End If
#Else
    ' 32-bit declarations (pre-VBA7)
    'Private Declare Function swe_calc_ut Lib "swedll32.dll" ( _
    '   ByVal tjd_ut As Double, _
    '   ByVal ipl As Long, _
    '   ByVal iflag As Long, _
    '   ByRef xx As Double, _
    '   ByRef serr As String) As Long
        
   'Private Declare Function swe_set_ephe_path Lib "swedll32.dll" ( _
   '    ByVal path As String) As Long
        
   'Private Declare Function swe_close Lib "swedll32.dll" () As Long
    
   'Private Declare Function swe_houses Lib "swedll32.dll" ( _
   '    ByVal tjd_ut As Double, _
   '    ByVal geolat As Double, _
   '    ByVal geolon As Double, _
   '    ByVal hsys As Long, _
   '    ByRef cusps As Double, _
   '    ByRef ascmc As Double) As Long
#End If
    
    ' Form variables
    Dim mlocationID As Long
    Dim mLocationExists As Boolean
    Dim mEventID As Long       ' Add this to store the EventID
    Dim mRecordSaved As Boolean ' Add this to track if the record has been saved

Private Function CurrentUser() As String
    ' Get current user name
    CurrentUser = Environ$("USERNAME")
End Function

Private Sub btnViewers_Click()
    On Error GoTo ErrorHandler
    
    ' Make sure the event is saved first
    If mEventID = 0 Then
        MsgBox "Please save the event first before assigning viewers.", vbInformation, "Save Required"
        Exit Sub
    End If
    
    ' Open the viewer assignment form, passing EventID and EventName in OpenArgs
    ' Format: "EventID|EventName"
    Dim openArgs As String
    openArgs = CStr(mEventID) & "|" & Nz(Me.txtEventName.value, "")
    
    DoCmd.OpenForm "frmViewerAssignList", acNormal, , , , acDialog, openArgs
    
    Exit Sub

ErrorHandler:
    MsgBox "Error opening viewer assignment form: " & Err.Description, vbCritical, "Error"
End Sub

' Helper function to check if a form is loaded
Private Function IsLoaded(formName As String) As Boolean
    On Error Resume Next
    IsLoaded = (Forms(formName).name = formName)
    On Error GoTo 0
End Function

Private Sub Form_Load()
    ' Set up the Country combo box
    Me.cboCountry.RowSource = "SELECT Country_Name FROM lu_Country ORDER BY Country_Name"
    Me.cboCountry.ColumnCount = 1
    Me.cboCountry.BoundColumn = 1
    Me.cboCountry.ColumnWidths = "2.5in"
    Me.cboCountry.LimitToList = True  ' This prevents users from entering values not in the list
    
    ' Set up the Event Type combo box
    Me.cboEventType.RowSource = "SELECT * FROM [lu_EventType] ORDER BY [Event_Desc];"
    Me.cboEventType.ColumnCount = 4
    Me.cboEventType.BoundColumn = 1
    Me.cboEventType.ColumnWidths = "0in;0in;0in;2.5in"
    Me.cboEventType.LimitToList = True  ' This prevents users from entering values not in the list

    ' Initialize form
    Me.dtEventDate = Date
    mLocationExists = False
    mEventID = 0           ' Initialize EventID
    mRecordSaved = False   ' Initialize record saved flag
    
    ' Initially hide team fields until we know if this is a sports event
    Me.cboHomeTeam.Visible = False
    Me.cboVisitingTeam.Visible = False
    Me.lblHomeTeam.Visible = False
    Me.lblVisitingTeam.Visible = False
    
    ' Set focus to first field
    Me.txtEventName.SetFocus
End Sub


Private Sub btnGetCoordinates_Click()
    Dim strCity As String
    Dim strStateProv As String
    Dim strCountry As String
    Dim strAddress As String
    Dim strResult As String
    Dim dblLat As Double
    Dim dblLng As Double
    
    ' Get values from form controls
    strCity = ProperCase(Nz(Me.txtCity, ""))
    strStateProv = UCase(Nz(Me.txtStateProv, ""))
    strCountry = Nz(Me.cboCountry, "")
    
    ' Basic validation - ensure required fields are filled
    If Trim(strCity) = "" Or Trim(strCountry) = "" Then
        MsgBox "City and Country are required fields to get coordinates.", vbExclamation, "Missing Information"
        If Trim(strCity) = "" Then
            Me.txtCity.SetFocus
        Else
            Me.cboCountry.SetFocus
        End If
        Exit Sub
    End If
    
    ' Special case: USA requires state
    If UCase(Trim(strCountry)) = "USA" And Trim(strStateProv) = "" Then
        MsgBox "State is required for locations in the USA.", vbExclamation, "Missing Information"
        Me.txtStateProv.SetFocus
        Exit Sub
    End If
    
    ' Check if location already exists in database
    mlocationID = modUtilities.FindLocationID(strCity, strStateProv, strCountry)
  
    If mlocationID > 0 Then
        ' Location exists, get coordinates from database
        mLocationExists = True
        FillLocationCoordinates mlocationID
        MsgBox "Location found in database. Coordinates loaded.", vbInformation, "Success"
    Else
        ' Location doesn't exist, get coordinates from API
        ' Format the address for the API
        If Trim(strStateProv) = "" Then
            strAddress = Trim(strCity) & ", " & Trim(strCountry)
        Else
            strAddress = Trim(strCity) & ", " & Trim(strStateProv) & ", " & Trim(strCountry)
        End If
        
        ' Show hourglass cursor
        DoCmd.Hourglass True
        
        ' Call the API function to get coordinates - with single address parameter
        strResult = GetLatLong_LocationIQ(strAddress)
        
        ' Process the result
        If Left(strResult, 5) = "Lat: " Then
            ' Parse the latitude and longitude from the result
            dblLat = val(Mid(strResult, 6, InStr(strResult, ", Lng:") - 6))
            dblLng = val(Mid(strResult, InStr(strResult, ", Lng:") + 7))
            
            ' Display coordinates on form
            Me.txtLatitude.value = dblLat
            Me.txtLongitude.value = dblLng
            
            ' Add location to database and get the new LocationID
            mlocationID = modUtilities.AddNewLocation(strCity, strStateProv, strCountry, dblLat, dblLng)
            mLocationExists = True
            
            MsgBox "Coordinates retrieved successfully!", vbInformation, "Success"
        Else
            ' Show error message
            MsgBox "Error retrieving coordinates: " & strResult, vbExclamation, "API Error"
        End If
        
        ' Turn off hourglass cursor
        DoCmd.Hourglass False
    End If
End Sub

'Private Sub btnGenerateEventChart_Click()
'    ' Generate astrological chart for the event
'    Dim eventDate As Date
'    Dim eventTime As Date
'    Dim lat As Double
'    Dim lng As Double
'    Dim Success As Boolean
'
'    ' Check if we have an EventID stored in the module-level variable
'    If mEventID = 0 Then
'        MsgBox "Please save the event record first before generating a chart.", vbExclamation, "Save Required"
'        Exit Sub
'    End If
'
'#If DEBUG_MODE Then
' Debug.Print "mEventID: " & mEventID
'#End If
'
'    ' Validate required fields
'    If Not ValidateRequiredFields() Then
'        Exit Sub
'    End If
'
'    ' Get values from form
'    eventDate = Me.dtEventDate.value
'    eventTime = Me.dtEventTime.value
'    lat = val(Me.txtLatitude.value)
'    lng = val(Me.txtLongitude.value)
'
'    ' Call function to generate chart with the valid EventID
'    Success = GenerateEventChart(mEventID, eventDate, eventTime, lat, lng, mlocationID)
'
'    If Success Then
'        Me.chkEventChartGenerated = True
'        MsgBox "Event chart generated successfully!", vbInformation, "Success"
'    Else
'        MsgBox "Error generating event chart. Please try again.", vbExclamation, "Error"
'    End If
'End Sub

Private Sub btnClose_Click()
    ' Check if form has unsaved changes
    If Me.Dirty Then
        Dim response As Integer
        response = MsgBox("Do you want to save your changes before closing?", vbQuestion + vbYesNoCancel, "Save Changes?")
        
        If response = vbYes Then
            ' Call the save button's click event
            btnSave_Click
            
            ' If still dirty, there was an error saving - don't close
            If Me.Dirty Then
                Exit Sub
            End If
        ElseIf response = vbCancel Then
            ' User canceled - don't close
            Exit Sub
        End If
    End If
    
    ' Close the form
    DoCmd.Close acForm, Me.name
End Sub

Private Sub btnSave_Click()
    ' Save event to database without closing the form
    On Error GoTo ErrorHandler
    
    ' Validate required fields
    If Not ValidateRequiredFields() Then
        Exit Sub
    End If
    
    ' Ensure we have a location
    If mlocationID = 0 Then
        MsgBox "Please get coordinates for the event location first.", vbExclamation, "Missing Location"
        Exit Sub
    End If
    
    ' Get values directly from form controls
    Dim strEventName As String
    Dim strEventType As String
    Dim dtEventDate As Date
    Dim strEventTime As String
    Dim blnChartGenerated As Boolean
    Dim lngEventID As Long
    Dim numTotalScore As Double
    Dim numOverUnder As Double
    
    ' Get team values - these may be null if not applicable
    Dim varHomeTeam As Variant
    Dim varVisitingTeam As Variant
    
    ' Get values with error handling
    On Error Resume Next
    strEventName = Nz(Me.txtEventName.value, "")
    strEventType = Nz(Me.cboEventType.value, "")
    If IsDate(Me.dtEventDate.value) Then
        dtEventDate = Me.dtEventDate.value
    Else
        dtEventDate = Date
    End If
    strEventTime = Nz(Me.dtEventTime.value, "")
    blnChartGenerated = Nz(Me.chkEventChartGenerated.value, False)
    
    ' Get numeric values if controls exist, otherwise default to 0
    numTotalScore = 0
    numOverUnder = 0
    If Not IsNull(Me.Controls("numTotalScore")) Then
        numTotalScore = Nz(Me.numTotalScore.value, 0)
    End If
    If Not IsNull(Me.Controls("numOverUnder")) Then
        numOverUnder = Nz(Me.numOverUnder.value, 0)
    End If
    
    ' Get team values if controls are visible
    If Me.cboHomeTeam.Visible Then
        varHomeTeam = Me.cboHomeTeam.value
    Else
        varHomeTeam = Null
    End If
    
    If Me.cboVisitingTeam.Visible Then
        varVisitingTeam = Me.cboVisitingTeam.value
    Else
        varVisitingTeam = Null
    End If
    
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim response As Integer
    
    ' Get workspace and database
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb()
    
    ' Check if we already have an EventID stored in module-level variable
    If mEventID > 0 Then
        ' We've already saved this event - perform an update
        strSQL = "UPDATE tblEvents SET " & _
                 "EventName = '" & Replace(strEventName, "'", "''") & "', " & _
                 "EventType = '" & Replace(strEventType, "'", "''") & "', " & _
                 "EventDate = #" & Format(dtEventDate, "mm/dd/yyyy") & "#, " & _
                 "EventTime = #" & Format(strEventTime, "hh:nn:ss") & "#, " & _
                 "LocationID = " & mlocationID & ", " & _
                 "TotalScore = " & numTotalScore & ", " & _
                 "OverUnder = " & numOverUnder & ", " & _
                 "EventChartGenerated = " & IIf(blnChartGenerated, "True", "False") & ", " & _
                 "Home = " & IIf(IsNull(varHomeTeam), "NULL", varHomeTeam) & ", " & _
                 "Visitor = " & IIf(IsNull(varVisitingTeam), "NULL", varVisitingTeam) & ", " & _
                 "DateUpdated = #" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "# " & _
                 "WHERE EventID = " & mEventID
        
        db.Execute strSQL, dbFailOnError
        MsgBox "Event has been updated.", vbInformation, "Update Successful"
        Exit Sub
    End If
    
    ' Start transaction for new record
    ws.BeginTrans
    
    ' Check for duplicate events
    strSQL = "SELECT EventID FROM tblEvents WHERE " & _
             "EventType = '" & Replace(strEventType, "'", "''") & "' AND " & _
             "EventDate = #" & Format(dtEventDate, "mm/dd/yyyy") & "# AND " & _
             "EventTime = #" & Format(strEventTime, "hh:nn:ss") & "# AND " & _
             "LocationID = " & mlocationID
             
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If Not rs.EOF Then
        ' Found a duplicate event
        lngEventID = rs!EventID
        rs.Close
        
        ' Rollback transaction
        ws.Rollback
        
        response = MsgBox("An event with the same type, date, time, and location already exists. " & _
                        "Would you like to update the existing event?", _
                        vbQuestion + vbYesNo, "Duplicate Event")
        
        If response = vbYes Then
            ' Update existing event (no transaction needed for single statement)
            strSQL = "UPDATE tblEvents SET " & _
                     "EventName = '" & Replace(strEventName, "'", "''") & "', " & _
                     "EventChartGenerated = " & IIf(blnChartGenerated, "True", "False") & ", " & _
                     "Home = " & IIf(IsNull(varHomeTeam), "NULL", varHomeTeam) & ", " & _
                     "Visitor = " & IIf(IsNull(varVisitingTeam), "NULL", varVisitingTeam) & ", " & _
                     "TotalScore = " & numTotalScore & ", " & _
                     "OverUnder = " & numOverUnder & ", " & _
                     "DateUpdated = #" & Format(Now, "mm/dd/yyyy hh:nn:ss") & "# " & _
                     "WHERE EventID = " & lngEventID
                     
            db.Execute strSQL, dbFailOnError
            
            ' Store the EventID in module-level variable
            mEventID = lngEventID
            
            MsgBox "Existing event has been updated.", vbInformation, "Update Successful"
        Else
            ' User chose not to update - exit
            MsgBox "No changes were made.", vbInformation, "Save Cancelled"
        End If
        
        Set rs = Nothing
        Set db = Nothing
        Set ws = Nothing
        Exit Sub
    End If
    
    ' Close recordset if it's open
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        On Error GoTo ErrorHandler
    End If
    
    ' No duplicate found, add new record
    ' Fixed SQL statement - removed extra comma and added proper field handling
    strSQL = "INSERT INTO tblEvents (EventName, EventType, EventDate, EventTime, " & _
             "LocationID, TotalScore, OverUnder, EventChartGenerated, Home, Visitor, DateCreated, " & _
             "DateUpdated, CreatedBy) " & _
             "VALUES ('" & Replace(strEventName, "'", "''") & "', '" & _
             Replace(strEventType, "'", "''") & "', #" & _
             Format(dtEventDate, "mm/dd/yyyy") & "#, #" & _
             Format(strEventTime, "hh:nn:ss") & "#, " & _
             mlocationID & ", " & _
             numTotalScore & ", " & _
             numOverUnder & ", " & _
             IIf(blnChartGenerated, "True", "False") & ", " & _
             IIf(IsNull(varHomeTeam), "NULL", varHomeTeam) & ", " & _
             IIf(IsNull(varVisitingTeam), "NULL", varVisitingTeam) & ", #" & _
             Format(Now, "mm/dd/yyyy hh:nn:ss") & "#, #" & _
             Format(Now, "mm/dd/yyyy hh:nn:ss") & "#, '" & _
             CurrentUser() & "')"
    
    Debug.Print "SQL: " & strSQL
    
    ' Execute the SQL statement
    db.Execute strSQL, dbFailOnError
    
    ' Get the new EventID
    Set rs = db.OpenRecordset("SELECT MAX(EventID) AS MaxID FROM tblEvents WHERE CreatedBy = '" & CurrentUser() & "'", dbOpenSnapshot)
    If Not rs.EOF And Not IsNull(rs!MaxID) Then
        lngEventID = rs!MaxID
    End If
    rs.Close
    
    ' Commit the transaction
    ws.CommitTrans
    
    ' Store the EventID in a module-level variable
    mEventID = lngEventID
    
#If DEBUG_MODE Then
    Debug.Print "New EventID saved: " & mEventID
#End If
    
    ' Show success message with EventID
    MsgBox "Event saved successfully! You can now generate a chart if needed." & vbCrLf & _
           "EventID: " & mEventID, vbInformation, "Success"
    
    Set rs = Nothing
    Set db = Nothing
    Set ws = Nothing
    Exit Sub
    
ErrorHandler:
    ' Rollback transaction on error
    If Not ws Is Nothing Then ws.Rollback
    MsgBox "Error saving event: " & Err.Number & " - " & Err.Description, vbCritical, "Error"
#If DEBUG_MODE Then
    Debug.Print "Error in btnSave_Click: " & Err.Number & " - " & Err.Description
#End If
    
    ' Clean up
    If Not rs Is Nothing Then
        On Error Resume Next
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
    Set ws = Nothing
End Sub

Private Sub btnCancel_Click()
    ' Cancel and close the form
    If Me.Dirty Then
        If MsgBox("Do you want to discard your changes?", vbQuestion + vbYesNo, "Confirm Cancel") = vbYes Then
            DoCmd.Close acForm, Me.name
        End If
    Else
        DoCmd.Close acForm, Me.name
    End If
End Sub

Private Function ValidateRequiredFields() As Boolean
    ' Validate that all required fields have values
    
    ' Event Name - Optional prior to 5/31/2025
    If (IsNull(Me.txtEventName) Or Me.txtEventName = "") And _
       (Not IsNull(Me.dtEventDate) And Me.dtEventDate > DateSerial(2025, 5, 31)) Then
        MsgBox "Please enter an Event Name.", vbExclamation, "Required Field"
        Me.txtEventName.SetFocus
        ValidateRequiredFields = False
        Exit Function
    End If
    
    ' Event Type
    If IsNull(Me.cboEventType) Then
        MsgBox "Please select an Event Type.", vbExclamation, "Required Field"
        Me.cboEventType.SetFocus
        ValidateRequiredFields = False
        Exit Function
    End If
    
    ' Event Date
    If IsNull(Me.dtEventDate) Then
        MsgBox "Please enter an Event Date.", vbExclamation, "Required Field"
        Me.dtEventDate.SetFocus
        ValidateRequiredFields = False
        Exit Function
    End If
    
    ' Event Time
    If IsNull(Me.dtEventTime) Then
        MsgBox "Please enter an Event Time.", vbExclamation, "Required Field"
        Me.dtEventTime.SetFocus
        ValidateRequiredFields = False
        Exit Function
    End If
    
    ' Location
    If IsNull(Me.txtCity) Or Me.txtCity = "" Then
        MsgBox "Please enter a City for the event location.", vbExclamation, "Required Field"
        Me.txtCity.SetFocus
        ValidateRequiredFields = False
        Exit Function
    End If
    
    'sherrie: add logic to check StateProv IF cboCountry = 'USA' 5/17
    
    If IsNull(Me.cboCountry) Then
        MsgBox "Please select a Country for the event location.", vbExclamation, "Required Field"
        Me.cboCountry.SetFocus
        ValidateRequiredFields = False
        Exit Function
    End If
    
    ValidateRequiredFields = True
End Function

Private Sub FillLocationCoordinates(LocationID As Long)
    ' Fill latitude and longitude fields from database location
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb()
    sql = "SELECT Latitude, Longitude FROM tblLocations WHERE LocationID=" & LocationID
    
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.EOF Then
        Me.txtLatitude.value = rs!latitude
        Me.txtLongitude.value = rs!longitude
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

Private Function GenerateEventChart(EventID As Long, eventDate As Date, eventTime As Date, lat As Double, lng As Double, LocationID As Long) As Boolean
    ' Calculate planetary positions and save chart data
    On Error GoTo ErrorHandler
    
    ' Setup variables
    Dim db As DAO.Database
    Dim rsChart As DAO.Recordset
    Dim rsPositions As DAO.Recordset
    Dim eventDateTime As Date
    Dim julianDay As Double
    Dim chartID As Long
    Dim i As Integer
    Dim BodyID As Long
    Dim result As Long
    Dim xx(6) As Double
    Dim serr As String
    Dim iflag As Long
    
    ' Debug
#If DEBUG_MODE Then
    Debug.Print "Starting chart generation for EventID: " & EventID
#End If
    
    ' Combine date and time
    eventDateTime = CDate(Format(eventDate, "yyyy-mm-dd") & " " & Format(eventTime, "hh:nn:ss"))
#If DEBUG_MODE Then
    Debug.Print "Event date/time: " & eventDateTime
#End If
    
    ' Convert to Julian Day
    julianDay = modUtilities.DateToJulianDay(eventDateTime)
    Debug.Print "Julian Day: " & julianDay
    
    ' Get database reference
    Set db = CurrentDb()
    
    ' First check if we can talk to the Swiss Ephemeris DLL
    If Not InitSwissEph() Then
        MsgBox "Failed to initialize Swiss Ephemeris. Please check DLL installation.", vbCritical, "Error"
        GenerateEventChart = False
        Exit Function
    End If
    
    ' Create or retrieve chart record
    Set rsChart = db.OpenRecordset("SELECT ChartID FROM tblCharts WHERE EventID = " & EventID, dbOpenSnapshot)
    
    If rsChart.EOF Then
#If DEBUG_MODE Then
        Debug.Print "Creating new chart record"
#End If
        rsChart.Close
        Set rsChart = db.OpenRecordset("tblCharts", dbOpenDynaset)
        rsChart.AddNew
        rsChart!PersonID = Null  ' Not a person chart
        rsChart!EventID = EventID
        rsChart!chartType = "Event"
        rsChart!chartDate = eventDate
        rsChart!chartTime = eventTime
        rsChart!LocationID = LocationID
        rsChart!DateCreated = Now()
        rsChart.Update
        
        ' Get the new ChartID
        rsChart.Bookmark = rsChart.LastModified
        chartID = rsChart!chartID
    Else
        chartID = rsChart!chartID
        rsChart.Close
        
        ' Delete any existing positions
        db.Execute "DELETE FROM tblChartPositions WHERE ChartID = " & chartID
    End If
#If DEBUG_MODE Then
    Debug.Print "Using ChartID: " & chartID
#End If
    
    ' Set calculation flags
    iflag = SEFLG_SWIEPH Or SEFLG_SPEED  ' Use Swiss Ephemeris & include speed calculations
    
    ' Open recordset for chart positions
    Set rsPositions = db.OpenRecordset("tblChartPositions", dbOpenDynaset)
    
    ' Calculate positions for each celestial body
    For i = 0 To 10  ' Sun through Pluto, plus North Node
        ' Set bodyID based on Swiss Ephemeris ID
        BodyID = i + 1  ' Adjust to match tblCelestialBodies IDs (Sun=1, Moon=2, etc.)
        
        ' Initialize error string
        serr = vbNullString
        
        ' Calculate position
#If DEBUG_MODE Then
        Debug.Print "Calculating position for body " & i
#End If
        result = swe_calc_ut(julianDay, i, iflag, xx(0), serr)
        
        If result >= 0 Then
#If DEBUG_MODE Then
            Debug.Print "Body " & i & " longitude: " & xx(0)
#End If
            
            ' Save position to database
            rsPositions.AddNew
            rsPositions!chartID = chartID
            rsPositions!BodyID = BodyID
            rsPositions!longitude = xx(0)
            rsPositions!latitude = xx(1)
            rsPositions!distance = xx(2)
            rsPositions!LongitudeSpeed = xx(3)
            
            ' Set LatitudeSpeed if the field exists
            On Error Resume Next
            rsPositions!LatitudeSpeed = xx(4)
            On Error GoTo ErrorHandler
            
            ' Determine sign (1-12 for Aries through Pisces)
            rsPositions!SignID = (Int(xx(0) / 30) Mod 12) + 1
            
            ' Determine if retrograde (negative longitude speed)
            rsPositions!IsRetrograde = (xx(3) < 0)
            
            rsPositions.Update
        Else
#If DEBUG_MODE Then
            Debug.Print "Failed to calculate position for body " & i & ": " & Trim$(serr)
#End If
        End If
    Next i
    
    ' Calculate house cusps and angles (Ascendant, Midheaven, etc.)
    Dim cusps(13) As Double ' Array for house cusps
    Dim ascmc(10) As Double ' Array for special points (Asc, MC, etc.)
    Dim hsys As String
    
    hsys = "P" ' Placidus houses
    
    ' Initialize error string
    serr = vbNullString
    
    ' Calculate houses
#If DEBUG_MODE Then
    Debug.Print "Calculating houses"
#End If
    result = swe_houses(julianDay, lat, lng, Asc(hsys), cusps(0), ascmc(0))
    
    If result >= 0 Then
        ' Add records for angles (Ascendant, MC, etc.)
        ' Ascendant (ASC)
        rsPositions.AddNew
        rsPositions!chartID = chartID
        rsPositions!BodyID = 15 ' ID for Ascendant in tblCelestialBodies
        rsPositions!longitude = ascmc(0)
        rsPositions!SignID = (Int(ascmc(0) / 30) Mod 12) + 1
        rsPositions.Update
        
        ' Midheaven (MC)
        rsPositions.AddNew
        rsPositions!chartID = chartID
        rsPositions!BodyID = 16 ' ID for Midheaven in tblCelestialBodies
        rsPositions!longitude = ascmc(1)
        rsPositions!SignID = (Int(ascmc(1) / 30) Mod 12) + 1
        rsPositions.Update
        
        ' Descendant (DSC) - opposite of ASC
        rsPositions.AddNew
        rsPositions!chartID = chartID
        rsPositions!BodyID = 17 ' ID for Descendant in tblCelestialBodies
        rsPositions!longitude = (ascmc(0) + 180) Mod 360
        rsPositions!SignID = (Int((ascmc(0) + 180) / 30) Mod 12) + 1
        rsPositions.Update
        
        ' Imum Coeli (IC) - opposite of MC
        rsPositions.AddNew
        rsPositions!chartID = chartID
        rsPositions!BodyID = 18 ' ID for Imum Coeli in tblCelestialBodies
        rsPositions!longitude = (ascmc(1) + 180) Mod 360
        rsPositions!SignID = (Int((ascmc(1) + 180) / 30) Mod 12) + 1
        rsPositions.Update
        
        ' Add records for house cusps if needed
        ' Uncomment this section if you want to store house cusps
        ' For i = 1 To 12
        '     rsPositions.AddNew
        '     rsPositions!ChartID = chartID
        '     rsPositions!bodyID = 18 + i ' IDs for house cusps (assuming they exist in tblCelestialBodies)
        '     rsPositions!Longitude = cusps(i)
        '     rsPositions!SignID = (Int(cusps(i) / 30) Mod 12) + 1
        '     rsPositions.Update
        ' Next i
    Else
#If DEBUG_MODE Then
        Debug.Print "Failed to calculate houses"
#End If
    End If
    
    ' Update the EventChartGenerated flag in the Events table
    db.Execute "UPDATE tblEvents SET EventChartGenerated = True WHERE EventID = " & EventID
    
    ' Clean up
    rsPositions.Close
    Set rsPositions = Nothing
    Set db = Nothing
    
    GenerateEventChart = True
#If DEBUG_MODE Then
    Debug.Print "Chart generation completed successfully"
#End If
    
    Exit Function
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error in GenerateEventChart: " & Err.Number & " - " & Err.Description
#End If
    MsgBox "Error generating chart: " & Err.Description, vbExclamation, "Error"
    GenerateEventChart = False
    
    ' Clean up
    If Not rsPositions Is Nothing Then
        On Error Resume Next
        rsPositions.Close
        Set rsPositions = Nothing
    End If
    If Not rsChart Is Nothing Then
        On Error Resume Next
        rsChart.Close
        Set rsChart = Nothing
    End If
    Set db = Nothing
End Function


Private Sub cboEventType_AfterUpdate()
    ' This event handler runs after the user selects a value in the Event Type dropdown
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsEventType As DAO.Recordset
    Dim strEventType As String
    Dim strEventCode As String
    Dim strSQL As String
    
    ' Get the selected event type ID
    Dim lngEventTypeID As Long
    lngEventTypeID = Me.cboEventType.value
    
    ' Clear both team combo boxes first
    Me.cboHomeTeam.RowSource = ""
    Me.cboVisitingTeam.RowSource = ""
    
    ' If no selection, exit
    If IsNull(lngEventTypeID) Then
        Me.cboHomeTeam.Visible = False
        Me.cboVisitingTeam.Visible = False
        Me.lblHomeTeam.Visible = False
        Me.lblVisitingTeam.Visible = False
        Exit Sub
    End If
    
    Set db = CurrentDb()
    
    ' Get the Event_Type and Event_Code for this selection
    strSQL = "SELECT Event_Type, Event_Code FROM lu_EventType WHERE ID = " & lngEventTypeID
    Set rsEventType = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If Not rsEventType.EOF Then
        strEventType = Nz(rsEventType!Event_Type, "")
        strEventCode = Nz(rsEventType!Event_Code, "")
        
        ' If this is a Sports event, populate team combos
        If strEventType = "Sports" Then
            ' Make team dropdowns visible
            Me.cboHomeTeam.Visible = True
            Me.cboVisitingTeam.Visible = True
            Me.lblHomeTeam.Visible = True
            Me.lblVisitingTeam.Visible = True
            
            ' Set up row source for team combo boxes
            strSQL = "SELECT ID, Team FROM lu_Team WHERE League = '" & Replace(strEventCode, "'", "''") & "' ORDER BY Team"
            
            ' Set the row source for both combo boxes
            Me.cboHomeTeam.RowSource = strSQL
            Me.cboVisitingTeam.RowSource = strSQL
            
            ' Set up column properties
            Me.cboHomeTeam.ColumnCount = 2
            Me.cboHomeTeam.BoundColumn = 1
            Me.cboHomeTeam.ColumnWidths = "0in;2in"
            
            Me.cboVisitingTeam.ColumnCount = 2
            Me.cboVisitingTeam.BoundColumn = 1
            Me.cboVisitingTeam.ColumnWidths = "0in;2in"
        Else
            ' Not a sports event, hide team fields
            Me.cboHomeTeam.Visible = False
            Me.cboVisitingTeam.Visible = False
            Me.lblHomeTeam.Visible = False
            Me.lblVisitingTeam.Visible = False
        End If
    End If
    
    rsEventType.Close
    Set rsEventType = Nothing
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error in cboEventType_AfterUpdate: " & Err.Number & " - " & Err.Description
#End If
End Sub

Private Sub btnGenerateChart_Click()
    modChartCalculation.GenerateEventChartForEvent (mEventID)
End Sub
