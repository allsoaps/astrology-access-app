Attribute VB_Name = "modAPIFunctions"
Option Compare Database
Option Explicit

Function GetLatLong_LocationIQ(ByVal address As String) As String

    Dim http As Object
    Dim json As Object
    Dim url As String
    Dim apiKey As String
    Dim result As String
    
    ' Retrieve the API key from the settings table
    apiKey = GetLocationIQAPIKey()
    url = "https://us1.locationiq.com/v1/search.php?key=" & apiKey & "&q=" & URLEncode(address) & "&format=json"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status = 200 Then
        Set json = JsonConverter.ParseJson(http.responseText)
        If json.Count > 0 Then
            result = "Lat: " & json(1)("lat") & ", Lng: " & json(1)("lon")
        Else
            result = "No result"
        End If
    Else
        result = "Error: " & http.Status
    End If

    GetLatLong_LocationIQ = result
End Function


' Retrieve the LocationIQ API key from tblSettings
Function GetLocationIQAPIKey() As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim apiKey As String

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT SettingValue FROM tblSettings WHERE SettingName = 'LocationIQ_API_Key'")

    If Not rs.EOF Then
        apiKey = rs!SettingValue
    Else
        apiKey = ""
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    GetLocationIQAPIKey = apiKey
End Function
' Helper function to encode spaces and symbols in URL
Function URLEncode(ByVal Text As String) As String
    Dim i As Long
    Dim CharCode As Integer
    Dim Char As String
    Dim OutStr As String

    For i = 1 To Len(Text)
        Char = Mid$(Text, i, 1)
        CharCode = Asc(Char)
        Select Case CharCode
            Case 48 To 57, 65 To 90, 97 To 122
                OutStr = OutStr & Char
            Case Else
                OutStr = OutStr & "%" & Hex(CharCode)
        End Select
    Next i
    URLEncode = OutStr
End Function

' Module: modMoonPhase
' Requires: JsonConverter.bas from https://github.com/VBA-tools/VBA-JSON

Public Function GetMoonPhase(lat As Double, lon As Double, moonDate As String) As String
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim json As Object
    
    On Error GoTo ErrorHandler
    
    url = "https://api.open-meteo.com/v1/astronomy?latitude=" & lat & _
          "&longitude=" & lon & "&date=" & moonDate

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status = 200 Then
        response = http.responseText
        Set json = JsonConverter.ParseJson(response)
        GetMoonPhase = json("moon_phase")
    Else
        GetMoonPhase = "Error: HTTP " & http.Status
    End If
    
    Exit Function

ErrorHandler:
    GetMoonPhase = "Error: " & Err.Description
End Function




