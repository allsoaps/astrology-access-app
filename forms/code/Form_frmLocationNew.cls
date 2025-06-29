VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocationNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Load()
    ' Set up the Country combo box
    Me.Country.RowSource = "SELECT Country_Name FROM lu_Country ORDER BY Country_Name"
    Me.Country.ColumnCount = 1
    Me.Country.BoundColumn = 1
    Me.Country.ColumnWidths = "2.5in"
    Me.Country.LimitToList = True  ' This prevents users from entering values not in the list
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
    strCity = Nz(Me.City, "")
    strState = Nz(Me.StateProv, "")
    strCountry = Nz(Me.Country, "")
    
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
        
        MsgBox "Coordinates retrieved successfully!", vbInformation, "Success"
    Else
        ' Show error message
        MsgBox "Error retrieving coordinates: " & strResult, vbExclamation, "API Error"
    End If
    
    ' Turn off hourglass cursor
    DoCmd.Hourglass False
End Sub

Private Sub btnSave_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim strCity As String
    Dim strState As String
    Dim strCountry As String
    Dim dblLatitude As Double
    Dim dblLongitude As Double
    Dim locationExists As Boolean
    
    ' Get values from form controls
    strCity = ProperCase(Nz(Me.City, ""))
    strState = UCase(Nz(Me.StateProv, ""))
    strCountry = Nz(Me.Country, "")
    dblLatitude = Nz(Me.txtLatitude, 0)
    dblLongitude = Nz(Me.txtLongitude, 0)
    
    ' Basic validation - ensure required fields are filled
    If Trim(strCity) = "" Or Trim(strCountry) = "" Then
        MsgBox "City and Country are required fields.", vbExclamation, "Missing Information"
        Exit Sub
    End If
    
    ' Special case: USA requires state
    If UCase(Trim(strCountry)) = "USA" And Trim(strState) = "" Then
        MsgBox "State is required for locations in the USA.", vbExclamation, "Missing Information"
        'Exit Sub
    ElseIf UCase(Trim(strCountry)) <> "USA" And Trim(strState) <> "" Then
        MsgBox "State is BLANK for locations outside the USA.", vbExclamation, "Remove Information"
        Exit Sub
    End If
    
    ' Check if location already exists
    locationExists = (DCount("*", "tblLocations", _
                      "City=""" & Replace(strCity, """", """""") & """ AND " & _
                      "Country=""" & Replace(strCountry, """", """""") & """" & _
                      IIf(Trim(strState) <> "", " AND [StateProv]=""" & Replace(strState, """", """""") & """", "")) > 0)
    
    If locationExists Then
        ' Location already exists
        MsgBox "This location (" & strCity & ", " & IIf(strState <> "", strState & ", ", "") & _
               strCountry & ") already exists in the database.", vbInformation, "Duplicate Location"
    Else
        ' Add new location
        Set db = CurrentDb()
        Set rs = db.OpenRecordset("tblLocations", dbOpenDynaset)
        rs.AddNew
        rs!City = strCity
        rs![StateProvince] = strState
        rs!Country = strCountry
        rs!latitude = dblLatitude
        rs!longitude = dblLongitude
        rs!DateCreated = Now()
        rs!DateUpdated = Now()
        rs.Update
        rs.Close
        
        MsgBox "Location added successfully.", vbInformation, "Success"
        DoCmd.Close acForm, Me.name
    End If
    
    Set rs = Nothing
    Set db = Nothing
End Sub
