Attribute VB_Name = "modSwissItems"
Option Compare Database
Option Explicit

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

' Constants for Swiss Ephemeris flags
Private Const SEFLG_GEOCENTRIC As Long = 0      ' Default
Private Const SEFLG_HELIOCENTRIC As Long = 8    ' Heliocentric flag

' Module-level variables for DLL paths
Private m_DLLPath As String
Private m_EphePath As String
Private m_IsInitialized As Boolean


' Swiss Ephemeris DLL declarations
#If VBA7 Then
    #If Win64 Then
        ' 64-bit declarations
        Private Declare PtrSafe Function swe_calc_ut Lib "swedll64.dll" ( _
            ByVal tjd_ut As Double, _
            ByVal ipl As Long, _
            ByVal iflag As Long, _
            ByRef xx As Double, _
            ByVal serr As String) As Long
            
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
            ByVal serr As String) As Long
            
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
'    Private Declare Function swe_calc_ut Lib "swedll32.dll" ( _
'        ByVal tjd_ut As Double, _
'        ByVal ipl As Long, _
'        ByVal iflag As Long, _
'        ByRef xx As Double, _
'        ByVal serr As String) As Long
'
'    Private Declare Function swe_set_ephe_path Lib "swedll32.dll" ( _
'        ByVal path As String) As Long
'
'    Private Declare Function swe_close Lib "swedll32.dll" () As Long
'
'    Private Declare Function swe_houses Lib "swedll32.dll" ( _
'        ByVal tjd_ut As Double, _
'        ByVal geolat As Double, _
'        ByVal geolon As Double, _
'        ByVal hsys As Long, _
'        ByRef cusps As Double, _
'        ByRef ascmc As Double) As Long
#End If
' Add these constants to the top of your modSwissItems module
' (after the existing planet constants)

' Additional planet/body constants for chart calculation
Public Const SE_CHIRON As Long = 15
Public Const SE_CERES As Long = 17
Public Const SE_MEAN_NODE As Long = 11
Public Const SE_TRUE_NODE As Long = 11   ' Same as mean node for basic usage

' House system constants
Public Const SE_PLACIDUS As Long = 80    ' ASCII value of 'P'
Public Const SE_KOCH As Long = 75        ' ASCII value of 'K'
Public Const SE_EQUAL As Long = 69       ' ASCII value of 'E'
Public Const SE_WHOLE_SIGN As Long = 87  ' ASCII value of 'W'

' Make the existing private constants public so chart module can use them
Public Const SEFLG_SPEED_PUBLIC As Long = 256
Public Const SEFLG_SWIEPH_PUBLIC As Long = 2
Public Const SEFLG_HELIOCENTRIC_PUBLIC As Long = 8

' Also add this public function to get the Swiss Ephemeris constants
Public Function GetSwissEphConstant(planetName As String) As Long
    Select Case UCase(planetName)
        Case "SUN": GetSwissEphConstant = SE_SUN
        Case "MOON": GetSwissEphConstant = SE_MOON
        Case "MERCURY": GetSwissEphConstant = SE_MERCURY
        Case "VENUS": GetSwissEphConstant = SE_VENUS
        Case "MARS": GetSwissEphConstant = SE_MARS
        Case "JUPITER": GetSwissEphConstant = SE_JUPITER
        Case "SATURN": GetSwissEphConstant = SE_SATURN
        Case "URANUS": GetSwissEphConstant = SE_URANUS
        Case "NEPTUNE": GetSwissEphConstant = SE_NEPTUNE
        Case "PLUTO": GetSwissEphConstant = SE_PLUTO
        Case "CHIRON": GetSwissEphConstant = SE_CHIRON
        Case "CERES": GetSwissEphConstant = SE_CERES
        Case "MEAN_NODE", "NORTH_NODE": GetSwissEphConstant = SE_MEAN_NODE
        Case Else: GetSwissEphConstant = -1 ' Invalid
    End Select
End Function

Public Function InitSwissEph() As Boolean
    On Error GoTo ErrorHandler
    
    ' Reset initialization flag
    m_IsInitialized = False
    
    ' Get DLL and ephemeris paths from database settings
    If Not LoadPathsFromDatabase() Then
        MsgBox "Could not load Swiss Ephemeris paths from database. Please check tblSwissEphSettings.", vbExclamation
        Exit Function
    End If
    
    ' Verify DLL file exists
    If Not VerifyDLLExists() Then
        Exit Function
    End If
    
    ' Verify ephemeris path exists
    If Not VerifyEphePathExists() Then
        Exit Function
    End If
    
    ' Test if the DLL is accessible
    Dim xx(6) As Double
    Dim serr As String
    'Dim serr As String * 255
    Dim result As Long
    
    ' Initialize error string properly
    'serr = vbNullString
    'serr = " "
    serr = String(255, vbNullChar)
    
    ' Set the ephemeris path
    Call swe_set_ephe_path(m_EphePath)
    
    ' Try to get the Sun's position for current date
    Dim julDay As Double
    julDay = modUtilities.DateToJulianDay(Now)
    
    ' Use simple flags (no speed)
    'result = swe_calc_ut(julDay, SE_SUN, 0, xx(0), serr)
    'result = SafeCalcUT(julDay, SE_SUN, 0, xx(0), serr)
    
    Dim sunPosition As Double
    
    If SafeCalcUT(julDay, SE_SUN, 0, sunPosition) Then
        m_IsInitialized = True
        InitSwissEph = True
    Else
        Debug.Print "Swiss Ephemeris initialization failed"
        MsgBox "Swiss Ephemeris calculation failed", vbExclamation, "Error"
        InitSwissEph = False
    End If
    Exit Function
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error in InitSwissEph: " & Err.Number & " - " & Err.Description
#End If
    
    ' Provide more specific error messages
    Select Case Err.Number
        Case 48, 53 ' File not found errors
            MsgBox "Swiss Ephemeris DLL not found. Please ensure the DLL files are in the correct location:" & vbCrLf & _
                   "Expected: " & GetExpectedDLLPath() & vbCrLf & vbCrLf & _
                   "Current setting: " & m_DLLPath, vbExclamation, "DLL Not Found"
        Case Else
            MsgBox "Failed to initialize Swiss Ephemeris: " & Err.Description, vbExclamation, "Error"
    End Select
    
    InitSwissEph = False
End Function

Private Function LoadPathsFromDatabase() As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb()
    
    ' Get paths from settings table
    sql = "SELECT EphePath, DLLPath FROM tblSwissEphSettings WHERE SettingID = 1"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        m_EphePath = Nz(rs!EphePath, "")
        m_DLLPath = Nz(rs!DLLPath, "")

        ' If no paths are set, use defaults relative to the database location
        If m_EphePath = "" Then
            m_EphePath = GetDefaultEphePath()
        End If
        If m_DLLPath = "" Then
            m_DLLPath = GetDefaultDLLPath()
        End If
        
        LoadPathsFromDatabase = True
    Else
        ' No settings record found, create default one
        CreateDefaultSettings
        LoadPathsFromDatabase = LoadPathsFromDatabase() ' Recursive call after creating defaults
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Exit Function
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error loading paths from database: " & Err.Description
#End If
    ' Use default paths as fallback
    m_EphePath = GetDefaultEphePath()
    m_DLLPath = GetDefaultDLLPath()
    LoadPathsFromDatabase = True
    
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
End Function

Private Function GetDefaultEphePath() As String
    ' Default to Resources\ephe subdirectory relative to database
    Dim dbPath As String
    dbPath = CurrentProject.path
    GetDefaultEphePath = dbPath & "\Resources\ephe"
End Function

Private Function GetDefaultDLLPath() As String
    ' Default to Resources subdirectory relative to database
    Dim dbPath As String
    dbPath = CurrentProject.path
    GetDefaultDLLPath = dbPath & "\Resources\sweph\bin"
End Function

Private Function GetExpectedDLLPath() As String
    #If VBA7 Then
        #If Win64 Then
            GetExpectedDLLPath = m_DLLPath & "\swedll64.dll"
        #Else
            GetExpectedDLLPath = m_DLLPath & "\swedll32.dll"
        #End If
    #Else
        GetExpectedDLLPath = m_DLLPath & "\swedll32.dll"
    #End If

End Function

Private Function VerifyDLLExists() As Boolean
    Dim dllFullPath As String
    dllFullPath = GetExpectedDLLPath()
    
    If Dir(dllFullPath) = "" Then
        MsgBox "Swiss Ephemeris DLL not found at: " & vbCrLf & dllFullPath & vbCrLf & vbCrLf & _
               "Please ensure the DLL files are copied to the Resources folder.", vbExclamation, "DLL Not Found"
        VerifyDLLExists = False
    Else
        VerifyDLLExists = True
    End If
End Function

Private Function VerifyEphePathExists() As Boolean
    If Dir(m_EphePath, vbDirectory) = "" Then
        MsgBox "Ephemeris data path not found: " & vbCrLf & m_EphePath & vbCrLf & vbCrLf & _
               "Please ensure the ephemeris data files are in the Resources\ephe folder.", vbExclamation, "Ephemeris Path Not Found"
        VerifyEphePathExists = False
    Else
        VerifyEphePathExists = True
    End If
End Function

Private Sub CreateDefaultSettings()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim sql As String
    
    Set db = CurrentDb()
    
    sql = "INSERT INTO tblSwissEphSettings (SettingID, EphePath, DLLPath, DefaultHouseSystem, DefaultOrbs, IncludeAsteroids, UseTopocentric, UseTrue) " & _
          "VALUES (1, '" & GetDefaultEphePath() & "', '" & GetDefaultDLLPath() & "', 1, True, False, False, True)"
    
    db.Execute sql
    
    Set db = Nothing
#If DEBUG_MODE Then
    Debug.Print "Created default Swiss Ephemeris settings"
#End If
    
    Exit Sub
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error creating default settings: " & Err.Description
#End If
    Set db = Nothing
End Sub

Public Function GetPlanetPosition(planetID As Long, dateTime As Date) As Double
    On Error GoTo ErrorHandler
    
    ' Ensure Swiss Ephemeris is initialized
    If Not m_IsInitialized Then
        If Not InitSwissEph() Then
            GetPlanetPosition = -999 ' Error indicator
            Exit Function
        End If
    End If
    
    Dim xx(6) As Double ' Array to receive the coordinates
    Dim serr As String
    Dim julianDay As Double
    Dim result As Long
    Dim flags As Long
    Dim CoordinateSystem As String
    
    ' Initialize error string
    'serr = vbNullString
    serr = String(255, vbNullChar)

    
    ' Get coordinate system setting from database
    CoordinateSystem = GetCoordinateSystemSetting()
    
    ' Set base flags
    flags = SEFLG_SPEED Or SEFLG_SWIEPH
    
    ' Add heliocentric flag if needed
    If CoordinateSystem = "Heliocentric" Then
        flags = flags Or SEFLG_HELIOCENTRIC
    End If
    ' Note: Geocentric is the default (no additional flag needed)
    
    ' Convert date to Julian day number
    julianDay = modUtilities.DateToJulianDay(dateTime)
    
    ' Calculate planet position with appropriate flags
    result = swe_calc_ut(julianDay, planetID, flags, xx(0), serr)
    
    If result < 0 Then
#If DEBUG_MODE Then
        Debug.Print "Error calculating planet position (" & CoordinateSystem & "): " & Trim$(serr)
#End If
        GetPlanetPosition = -999 ' Error indicator
    Else
        ' Return longitude in degrees (xx(0) contains longitude)
        GetPlanetPosition = xx(0)
#If DEBUG_MODE Then
        Debug.Print "Planet " & planetID & " position (" & CoordinateSystem & "): " & xx(0) & " degrees"
#End If
    End If
    
    Exit Function
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error in GetPlanetPosition: " & Err.Description
#End If
    GetPlanetPosition = -999
End Function



Public Function GetHouseCusps(julianDay As Double, latitude As Double, longitude As Double) As Variant
    On Error GoTo ErrorHandler
    
    ' Ensure Swiss Ephemeris is initialized
    If Not m_IsInitialized Then
        If Not InitSwissEph() Then
            GetHouseCusps = Array() ' Return empty array on error
            Exit Function
        End If
    End If
    
    Dim cusps(13) As Double ' Array for house cusps (1-12, plus extras)
    Dim ascmc(10) As Double ' Array for special points (Asc, MC, etc.)
    Dim result As Long
    Dim i As Integer
    Dim resultArray(12) As Double ' Return array for cusps 1-12
    
    ' Calculate houses using Placidus system (P)
    result = swe_houses(julianDay, latitude, longitude, Asc("P"), cusps(0), ascmc(0))
    
    If result >= 0 Then
        ' Copy cusps 1-12 to result array
        For i = 1 To 12
            resultArray(i - 1) = cusps(i)
        Next i
        GetHouseCusps = resultArray
    Else
#If DEBUG_MODE Then
        Debug.Print "Error calculating house cusps"
#End If
        GetHouseCusps = Array() ' Return empty array on error
    End If
    
    Exit Function
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error in GetHouseCusps: " & Err.Description
#End If
    GetHouseCusps = Array()
End Function

Public Function GetAscendantMidheaven(julianDay As Double, latitude As Double, longitude As Double) As Variant
    On Error GoTo ErrorHandler
    
    ' Ensure Swiss Ephemeris is initialized
    If Not m_IsInitialized Then
        If Not InitSwissEph() Then
            GetAscendantMidheaven = Array(-999, -999) ' Return error indicators
            Exit Function
        End If
    End If
    
    Dim cusps(13) As Double ' Array for house cusps
    Dim ascmc(10) As Double ' Array for special points
    Dim result As Long
    
    ' Calculate houses to get Ascendant and Midheaven
    result = swe_houses(julianDay, latitude, longitude, Asc("P"), cusps(0), ascmc(0))
    
    If result >= 0 Then
        ' ascmc(0) = Ascendant, ascmc(1) = Midheaven
        GetAscendantMidheaven = Array(ascmc(0), ascmc(1))
    Else
#If DEBUG_MODE Then
        Debug.Print "Error calculating Ascendant/Midheaven"
#End If
        GetAscendantMidheaven = Array(-999, -999)
    End If
    
    Exit Function
    
ErrorHandler:
#If DEBUG_MODE Then
    Debug.Print "Error in GetAscendantMidheaven: " & Err.Description
#End If
    GetAscendantMidheaven = Array(-999, -999)
End Function

' Cleanup function - call this when closing the application
Public Sub CloseSwissEph()
    On Error Resume Next
    Call swe_close
    m_IsInitialized = False
#If DEBUG_MODE Then
    Debug.Print "Swiss Ephemeris closed"
#End If
End Sub

' Property to check if Swiss Ephemeris is initialized
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

' Properties to get current paths
Public Property Get EphePath() As String
    EphePath = m_EphePath
End Property

Public Property Get DLLPath() As String
    DLLPath = m_DLLPath
End Property

' === Safe function to set the ephemeris path ===
Public Function SafeSetEphePath(EphePath As String, Optional showMsg As Boolean = True) As Boolean
    On Error GoTo Fail

    If Dir(EphePath, vbDirectory) = "" Then
        If showMsg Then MsgBox "Invalid Ephemeris Path: " & EphePath, vbCritical
        SafeSetEphePath = False
        Exit Function
    End If

    swe_set_ephe_path EphePath
    SafeSetEphePath = True
    Exit Function

Fail:
    If showMsg Then MsgBox "Error calling swe_set_ephe_path: " & Err.Description, vbCritical
    SafeSetEphePath = False
End Function

' === Safe wrapper for swe_calc_ut ===
Public Function SafeCalcUT(jd As Double, planet As Long, flags As Long, ByRef lonOut As Double, Optional showMsg As Boolean = True) As Boolean
    On Error GoTo Fail

    Dim xx(5) As Double
    Dim serr As String
    Dim result As Long

    ' Initialize the error string buffer - THIS WAS MISSING
    serr = String(255, vbNullChar)

    result = swe_calc_ut(jd, planet, flags, xx(0), serr)

    If result < 0 Then
        If showMsg Then MsgBox "Swiss Ephemeris Error: " & Trim$(serr), vbExclamation
        SafeCalcUT = False
    Else
        lonOut = xx(0)
        SafeCalcUT = True
    End If
    Exit Function

Fail:
    If showMsg Then MsgBox "Unexpected error in swe_calc_ut: " & Err.Description, vbCritical
    SafeCalcUT = False
End Function

' Helper function to get coordinate system setting from database
Public Function GetCoordinateSystemSetting() As String
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

'#######
' Add this function to your modSwissItems module
' Returns complete planetary data: longitude, latitude, distance, speeds
Public Function GetCompletePlanetData(julianDay As Double, planetID As Long) As Variant
    On Error GoTo ErrorHandler
    
    ' Ensure Swiss Ephemeris is initialized
    If Not m_IsInitialized Then
        If Not InitSwissEph() Then
            GetCompletePlanetData = Array(-999, -999, -999, -999, -999, -999) ' Error indicators
            Exit Function
        End If
    End If
    
    Dim xx(5) As Double ' Array to receive all coordinates
    Dim serr As String
    Dim result As Long
    Dim flags As Long
    
    ' Initialize error string buffer
    serr = String(255, vbNullChar)
    
    ' Set flags to get coordinates with speed (for retrograde detection)
    flags = SEFLG_SWIEPH Or SEFLG_SPEED
    
    ' Add heliocentric flag if needed based on settings
    If GetCoordinateSystemSetting() = "Heliocentric" Then
        flags = flags Or SEFLG_HELIOCENTRIC
    End If
    
    ' Calculate planet position with all data
    result = swe_calc_ut(julianDay, planetID, flags, xx(0), serr)
    
    If result < 0 Then
        Debug.Print "Error calculating complete planet data: " & Trim$(serr)
        GetCompletePlanetData = Array(-999, -999, -999, -999, -999, -999) ' Error indicators
    Else
        ' Return all 6 values: longitude, latitude, distance, lon_speed, lat_speed, dist_speed
        GetCompletePlanetData = Array(xx(0), xx(1), xx(2), xx(3), xx(4), xx(5))
        
    On Error Resume Next
    
    Debug.Print "Planet " & planetID & " calculated successfully. Lon=" & Format(xx(0), "0.00") & "°"

    On Error GoTo ErrorHandler
    
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in GetCompletePlanetData: " & Err.Description
    GetCompletePlanetData = Array(-999, -999, -999, -999, -999, -999)
End Function
