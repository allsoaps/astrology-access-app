Attribute VB_Name = "modTransitChart"
' Module: modTransitChart
' Generate Transit Chart comparing natal positions to current sky positions
' Uses existing Swiss Ephemeris functions from modSwissItems and modSimpleChart

Public Function GenerateTransitChart(PersonID As Long, SessionID As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsSession As DAO.Recordset
    Dim rsLocation As DAO.Recordset
    Dim rsNatal As DAO.Recordset
    Dim rsTransitAspects As DAO.Recordset
    
    Dim sessionDate As Date
    Dim sessionTime As Date
    Dim sessionDateTime As Date
    Dim latitude As Double
    Dim longitude As Double
    Dim julianDay As Double
    
    Set db = CurrentDb()
    
    Debug.Print "=== GENERATING TRANSIT CHART ==="
    Debug.Print "PersonID: " & PersonID & ", SessionID: " & SessionID
    
    ' 1. Get Session Details
    Set rsSession = db.OpenRecordset("SELECT * FROM tblSessions WHERE SessionID = " & SessionID)
    If rsSession.EOF Then
        Debug.Print "ERROR: SessionID " & SessionID & " not found"
        GenerateTransitChart = False
        Exit Function
    End If
    
    sessionDate = rsSession!sessionDate
    sessionTime = rsSession!SessionStartTime
    sessionDateTime = sessionDate + timeValue(sessionTime)
    
    Debug.Print "Session DateTime: " & sessionDateTime
    
    ' 2. Get Location Details
    Set rsLocation = db.OpenRecordset("SELECT * FROM tblLocations WHERE LocationID = " & rsSession!LocationID)
    If rsLocation.EOF Then
        Debug.Print "ERROR: LocationID " & rsSession!LocationID & " not found"
        GenerateTransitChart = False
        Exit Function
    End If
    
    latitude = rsLocation!latitude
    longitude = rsLocation!longitude
    
    Debug.Print "Location: " & rsLocation!City & " (" & latitude & ", " & longitude & ")"
    
    ' 3. Get Natal Chart Positions
    Set rsNatal = db.OpenRecordset( _
        "SELECT cp.*, cb.BodyName FROM (tblChartPositions cp " & _
        "INNER JOIN tblCharts c ON cp.ChartID = c.ChartID) " & _
        "INNER JOIN tblCelestialBodies cb ON cp.BodyID = cb.BodyID " & _
        "WHERE c.PersonID = " & PersonID & " AND c.ChartType = 'Natal' " & _
        "ORDER BY cp.BodyID")
    
    If rsNatal.EOF Then
        Debug.Print "ERROR: No natal chart found for PersonID " & PersonID
        GenerateTransitChart = False
        Exit Function
    End If
    
    Debug.Print "Loading natal positions..."
    
    ' Debug: Check what's in the natal recordset
    Debug.Print "Natal recordset field count: " & rsNatal.Fields.Count
    If Not rsNatal.EOF Then
        rsNatal.MoveFirst
        Debug.Print "First natal record - BodyID: " & rsNatal!BodyID & ", Longitude: " & rsNatal!longitude
        Debug.Print "Available fields: "
        Dim fld As DAO.Field
        For Each fld In rsNatal.Fields
            Debug.Print "  " & fld.name
        Next fld
    End If
    
    ' 4. Calculate Julian Day for transit time
    julianDay = modUtilities.DateToJulianDay(sessionDateTime)
    Debug.Print "Transit Julian Day: " & julianDay
    
    ' Initialize Swiss Ephemeris
    If Not modSwissItems.InitSwissEph() Then
        Debug.Print "Error: Failed to initialize Swiss Ephemeris"
        GenerateTransitChart = False
        Exit Function
    End If
    
    ' 5. Clear any existing transit aspects for this combination
    db.Execute "DELETE FROM tblTransitAspects WHERE PersonID = " & PersonID & " AND SessionID = " & SessionID
    
    ' 6. Calculate Transit-to-Natal Aspects
    Debug.Print "Calculating transit aspects..."
    
    Set rsTransitAspects = db.OpenRecordset("tblTransitAspects", dbOpenDynaset)
    
    Dim aspectCount As Integer: aspectCount = 0
    
    ' Get active planets for transit calculation
    Dim db2 As DAO.Database
    Dim rsTransitPlanets As DAO.Recordset
    Set db2 = CurrentDb()
    Set rsTransitPlanets = db2.OpenRecordset( _
        "SELECT BodyID, BodyName, SwissEphID FROM tblCelestialBodies " & _
        "WHERE IsActive = True AND BodyType IN ('Planet', 'Asteroid', 'Node') " & _
        "ORDER BY DisplayOrder")
    
    ' Loop through each transit planet
    Do While Not rsTransitPlanets.EOF
        Debug.Print "Calculating transit " & rsTransitPlanets!BodyName & "..."
        
        ' Calculate current position using existing function
        Dim transitData As Variant
        transitData = modSwissItems.GetCompletePlanetData(julianDay, rsTransitPlanets!swissEphID)
        
        If IsArray(transitData) And transitData(0) <> -999 Then
            Dim transitLongitude As Double
            transitLongitude = CDbl(transitData(0))
            
            Debug.Print "Transit " & rsTransitPlanets!BodyName & ": " & Format(transitLongitude, "0.00") & "°"
            
            ' Compare this transit planet to all natal planets
            ' Create fresh natal recordset for each transit planet to avoid cursor issues
            Dim rsNatalCompare As DAO.Recordset
            Set rsNatalCompare = db.OpenRecordset( _
                "SELECT cp.*, cb.BodyName FROM (tblChartPositions cp " & _
                "INNER JOIN tblCharts c ON cp.ChartID = c.ChartID) " & _
                "INNER JOIN tblCelestialBodies cb ON cp.BodyID = cb.BodyID " & _
                "WHERE c.PersonID = " & PersonID & " AND c.ChartType = 'Natal' " & _
                "ORDER BY cp.BodyID")
            
            Do While Not rsNatalCompare.EOF
                
                ' Calculate aspect between transit planet and natal planet
                Dim aspectResult As Variant
                aspectResult = CalculateTransitAspect( _
                    rsTransitPlanets!BodyID, transitLongitude, _
                    rsNatalCompare!BodyID, rsNatalCompare!longitude)
                
                ' If valid aspect found, save it
                If IsArray(aspectResult) And CBool(aspectResult(0)) = True Then
                    rsTransitAspects.AddNew
                    rsTransitAspects!PersonID = PersonID
                    rsTransitAspects!SessionID = SessionID
                    rsTransitAspects!transitBodyID = rsTransitPlanets!BodyID
                    rsTransitAspects!natalBodyID = rsNatalCompare!BodyID
                    rsTransitAspects!aspectID = CInt(aspectResult(1))  ' AspectID
                    rsTransitAspects!ExactDegree = CDbl(aspectResult(2))  ' ExactDegree
                    rsTransitAspects!OrbitApplying = CBool(aspectResult(3))  ' IsApplying
                    rsTransitAspects!transitLongitude = transitLongitude
                    rsTransitAspects!natalLongitude = rsNatalCompare!longitude
                    rsTransitAspects!Notes = "Transit " & rsTransitPlanets!BodyName & " to Natal " & rsNatalCompare!BodyName
                    rsTransitAspects.Update
                    
                    aspectCount = aspectCount + 1
                    
                    Debug.Print "ASPECT: Transit " & rsTransitPlanets!BodyName & _
                        " " & GetAspectName(CInt(aspectResult(1))) & _
                        " Natal " & rsNatalCompare!BodyName & _
                        " (Orb: " & Format(aspectResult(2), "0.00") & "°, " & _
                        IIf(CBool(aspectResult(3)), "Applying", "Separating") & ")"
                End If
                
                rsNatalCompare.MoveNext
            Loop
            
            rsNatalCompare.Close
            Set rsNatalCompare = Nothing
        Else
            Debug.Print "Failed to calculate transit " & rsTransitPlanets!BodyName
        End If
        
        rsTransitPlanets.MoveNext
    Loop
    
    Debug.Print "=== TRANSIT CHART COMPLETE ==="
    Debug.Print "Total aspects found: " & aspectCount
    
    ' Cleanup
    rsSession.Close
    rsLocation.Close
    rsNatal.Close
    rsTransitAspects.Close
    rsTransitPlanets.Close
    Set db = Nothing
    Set db2 = Nothing
    
    GenerateTransitChart = True
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in GenerateTransitChart: " & Err.Number & " - " & Err.Description
    GenerateTransitChart = False
    
    ' Cleanup on error
    On Error Resume Next
    If Not rsSession Is Nothing Then rsSession.Close
    If Not rsLocation Is Nothing Then rsLocation.Close
    If Not rsNatal Is Nothing Then rsNatal.Close
    If Not rsTransitAspects Is Nothing Then rsTransitAspects.Close
    If Not rsTransitPlanets Is Nothing Then rsTransitPlanets.Close
    Set db = Nothing
    Set db2 = Nothing
End Function

' Function to calculate aspect between transit planet and natal planet
' Returns array: [IsValidAspect, AspectID, ExactDegree, IsApplying]
Private Function CalculateTransitAspect(transitBodyID As Integer, transitLongitude As Double, _
                                      natalBodyID As Integer, natalLongitude As Double) As Variant
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rsAspects As DAO.Recordset
    
    Set db = CurrentDb()
    Set rsAspects = db.OpenRecordset("SELECT * FROM tblAspects ORDER BY AspectID")
    
    ' Calculate angular difference
    Dim angularDiff As Double
    angularDiff = Abs(transitLongitude - natalLongitude)
    If angularDiff > 180 Then angularDiff = 360 - angularDiff
    
    ' Check each aspect type for a match
    Do While Not rsAspects.EOF
        Dim targetAngle As Double: targetAngle = rsAspects!angle
        Dim orb As Double
        
        ' Use existing orb calculation logic from modSimpleChart
        orb = GetTransitOrb(rsAspects!OrbitMajor, rsAspects!OrbitMinor, transitBodyID, natalBodyID)
        
        ' Check if within orb
        Dim aspectDiff As Double
        If targetAngle = 0 Then
            ' Conjunction - check both 0° and 360°
            aspectDiff = IIf(Abs(angularDiff - 0) < Abs(angularDiff - 360), Abs(angularDiff - 0), Abs(angularDiff - 360))
        Else
            aspectDiff = Abs(angularDiff - targetAngle)
        End If
        
        If aspectDiff <= orb Then
            ' Found valid aspect - return success array
            CalculateTransitAspect = Array(True, CInt(rsAspects!aspectID), CDbl(aspectDiff), True)
            rsAspects.Close
            Set db = Nothing
            Exit Function
        End If
        
        rsAspects.MoveNext
    Loop
    
    ' No aspect found
    CalculateTransitAspect = Array(False, CInt(0), CDbl(0), False)
    
    rsAspects.Close
    Set db = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in CalculateTransitAspect: " & Err.Number & " - " & Err.Description
    CalculateTransitAspect = Array(False, CInt(0), CDbl(0), False)
End Function

' Helper function for transit orbs (simplified version of existing logic)
Private Function GetTransitOrb(majorOrb As Double, minorOrb As Double, _
                              transitBodyID As Integer, natalBodyID As Integer) As Double
    ' Simplified orb logic - use major orb for important planets
    Dim isMajorTransit As Boolean: isMajorTransit = (transitBodyID <= 10) ' Sun through Pluto
    Dim isMajorNatal As Boolean: isMajorNatal = (natalBodyID <= 10)
    
    If isMajorTransit And isMajorNatal Then
        GetTransitOrb = majorOrb
    Else
        GetTransitOrb = minorOrb
    End If
End Function

' Helper to get aspect name
Private Function GetAspectName(aspectID As Integer) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT AspectName FROM tblAspects WHERE AspectID = " & aspectID)
    
    If Not rs.EOF Then
        GetAspectName = rs!AspectName
    Else
        GetAspectName = "Unknown"
    End If
    
    rs.Close
    Set db = Nothing
End Function

' Test function to run the transit calculation
Public Sub TestTransitChart()
    Debug.Print "=== TESTING TRANSIT CHART ==="
    
    Dim success As Boolean
    success = GenerateTransitChart(15, 19)
    
    If success Then
        Debug.Print "Transit chart generated successfully!"
        
        ' Show results
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        
        Set db = CurrentDb()
        Set rs = db.OpenRecordset( _
            "SELECT ta.*, cb1.BodyName AS TransitPlanet, cb2.BodyName AS NatalPlanet, asp.AspectName " & _
            "FROM ((tblTransitAspects ta " & _
            "INNER JOIN tblCelestialBodies cb1 ON ta.TransitBodyID = cb1.BodyID) " & _
            "INNER JOIN tblCelestialBodies cb2 ON ta.NatalBodyID = cb2.BodyID) " & _
            "INNER JOIN tblAspects asp ON ta.AspectID = asp.AspectID " & _
            "WHERE ta.PersonID = 15 AND ta.SessionID = 19 " & _
            "ORDER BY ta.ExactDegree")
        
        Debug.Print "TRANSIT ASPECTS FOUND:"
        Do While Not rs.EOF
            Debug.Print "Transit " & rs!TransitPlanet & " " & rs!AspectName & _
                       " Natal " & rs!NatalPlanet & " (Orb: " & Format(rs!ExactDegree, "0.00") & "°)"
            rs.MoveNext
        Loop
        
        rs.Close
        Set db = Nothing
    Else
        Debug.Print "Transit chart generation failed!"
    End If
End Sub

