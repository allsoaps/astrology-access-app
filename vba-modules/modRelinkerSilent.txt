Attribute VB_Name = "modRelinkerSilent"
' Module: modRelinkerSilent
Option Compare Database
Option Explicit

Public Function SilentRelinkToBackend() As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim backendPath As String

'UNCOMMENT before DELIVERY

    'backendPath = CurrentProject.path & "\Resources\SolarFlare_be.accdb"
    If Dir(backendPath) = "" Then Exit Function

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 And Left(tdf.name, 4) <> "MSys" Then
            On Error Resume Next
            tdf.Connect = ";DATABASE=" & backendPath
            tdf.RefreshLink
            Err.Clear
            On Error GoTo 0
        End If
    Next tdf

    Set tdf = Nothing
    Set db = Nothing

    SilentRelinkToBackend = True
End Function

