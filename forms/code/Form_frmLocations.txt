VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnUpdCoord_Click()
    RunUpdateMissingCoordinates
End Sub


' Function to update missing latitude and longitude in tblLocations
' Uses the GetLatLong_LocationIQ function from modAPIFunctions
Public Function UpdateMissingCoordinates() As String
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strResult As String
    Dim latLngStr As String
    Dim strCity As String
    Dim strState As String
    Dim strCountry As String
    Dim strAddress As String
    Dim lat As Double
    Dim lng As Double
    Dim locationsUpdated As Integer
    Dim locationsWithErrors As Integer
    Dim resultMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Initialize counters
    locationsUpdated = 0
    locationsWithErrors = 0
    resultMsg = ""
    
    ' Get reference to current database
    Set db = CurrentDb()
    
    ' Query for all locations with missing latitude or longitude
    strSQL = "SELECT * FROM tblLocations WHERE Latitude = 0 OR Longitude = 0"
    Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    ' Check if any records were found
    If rst.EOF Then
        UpdateMissingCoordinates = "No locations with missing coordinates found."
        rst.Close
        Set rst = Nothing
        Set db = Nothing
        Exit Function
    End If
    
    ' Loop through each record and update the coordinates
    Do While Not rst.EOF
    
    strCity = rst!City
    strState = Nz(rst![StateProvince], "")
    strCountry = rst!Country
    
    ' Format the address for the API
    If Trim(strState) = "" Then
        strAddress = Trim(strCity) & ", " & Trim(strCountry)
    Else
        strAddress = Trim(strCity) & ", " & Trim(strState) & ", " & Trim(strCountry)
    End If
    
    ' Show hourglass cursor
    'DoCmd.Hourglass True
    
    ' Call the API function to get coordinates
    latLngStr = GetLatLong_LocationIQ(strAddress)
    
        ' Check if we got a valid result
        If Left(latLngStr, 5) = "Lat: " Then
            ' Parse the latitude and longitude from the result
            lat = val(Mid(latLngStr, 6, InStr(latLngStr, ", Lng:") - 6))
            lng = val(Mid(latLngStr, InStr(latLngStr, ", Lng:") + 7))
            
            ' Update the record if valid coordinates were returned
            If lat <> 0 And lng <> 0 Then
                rst.Edit
                rst!latitude = lat
                rst!longitude = lng
                rst!DateUpdated = Now()
                rst.Update
                
                resultMsg = resultMsg & "Updated: " & rst!City & ", " & _
                             Nz(rst!StateProvince, "") & ", " & rst!Country & _
                             " - Lat: " & lat & ", Lng: " & lng & vbCrLf
                             

                             
                locationsUpdated = locationsUpdated + 1
            Else
                resultMsg = resultMsg & "Error: Could not get valid coordinates for " & _
                             rst!City & ", " & Nz(rst!state, "") & ", " & rst!Country & vbCrLf
                locationsWithErrors = locationsWithErrors + 1
            End If
        Else
            ' Log the error
            resultMsg = resultMsg & "Error: " & latLngStr & " for " & _
                         rst!City & ", " & Nz(rst![state], "") & ", " & rst!Country & vbCrLf
            locationsWithErrors = locationsWithErrors + 1
        End If
        
        ' Add delay to avoid hitting API rate limits
    DoEvents
        Sleep 1000  ' Wait 1 second between API calls
        
        rst.MoveNext
    Loop
    
    ' Close recordset
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    
    ' Return summary of results
    UpdateMissingCoordinates = "Updated " & locationsUpdated & " locations, " & _
                              locationsWithErrors & " errors." & vbCrLf & vbCrLf & resultMsg
    
    Exit Function
    
ErrorHandler:
    UpdateMissingCoordinates = "Error: " & Err.Number & " - " & Err.Description
    
    ' Clean up
    If Not rst Is Nothing Then
        On Error Resume Next  ' In case rst is already closed
        rst.Close
        On Error GoTo ErrorHandler  ' Restore error handling
        Set rst = Nothing
    End If
    Set db = Nothing
End Function

' Wrapper function that can be called from a button or other UI element
Public Sub RunUpdateMissingCoordinates()
    Dim result As String
    
    ' Display hourglass cursor
    DoCmd.Hourglass True
    
    ' Call the main function
    result = UpdateMissingCoordinates()
    
    ' Turn off hourglass
    DoCmd.Hourglass False
    
    ' Display results
    MsgBox result, vbInformation, "Update Coordinates Results"
End Sub



