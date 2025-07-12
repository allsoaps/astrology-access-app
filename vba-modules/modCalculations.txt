Attribute VB_Name = "modCalculations"
Option Compare Database
Option Explicit

Private Function GetNextImpressionNumber(SessionID As Long) As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Set db = CurrentDb()
    strSQL = "SELECT MAX(ImpressionNumber) AS MaxImpressionNum FROM tblImpressions " & _
             "WHERE SessionID = " & SessionID
    
    Set rs = db.OpenRecordset(strSQL)
    
    If rs.EOF Or IsNull(rs!MaxImpressionNum) Then
        GetNextImpressionNumber = 1
    Else
        GetNextImpressionNumber = rs!MaxImpressionNum + 1
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Private Function GetNextSessionNumber(studentID As Long, EventID As Long) As Integer
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Set db = CurrentDb()
    strSQL = "SELECT MAX(SessionNumber) AS MaxSessionNum FROM tblSessions " & _
             "WHERE StudentID = " & studentID & " AND EventID = " & EventID
    
    Set rs = db.OpenRecordset(strSQL)
    
    If rs.EOF Or IsNull(rs!MaxSessionNum) Then
        GetNextSessionNumber = 1
    Else
        GetNextSessionNumber = rs!MaxSessionNum + 1
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function



