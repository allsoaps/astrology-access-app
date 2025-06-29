VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAspectGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
' START OF MODULE: frmAspectGrid
'===========================================================

' =================================================================
' frmAspectGrid - Chart Display Form
' =================================================================
' This form displays planetary positions and aspects for a given chart
' Called from Session Manager with ChartID parameter

' Form-level variables
Private m_ChartID As Long
Private m_ChartType As String
Private m_ChartTitle As String

' =================================================================
' FORM EVENTS
' =================================================================

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== frmAspectGrid Form_Load ==="
    Debug.Print "OpenArgs value: " & Nz(Me.openArgs, "NULL")
    Debug.Print "OpenArgs is null: " & IsNull(Me.openArgs)
    
    ' Get ChartID from OpenArgs parameter
    If Not IsNull(Me.openArgs) And Me.openArgs <> "" Then
        If IsNumeric(Me.openArgs) Then
            m_ChartID = CLng(Me.openArgs)
            Debug.Print "frmAspectGrid loaded with ChartID: " & m_ChartID
            
            ' Load and display the chart data
            LoadChartData
        Else
            MsgBox "Invalid chart ID received: " & Me.openArgs, vbCritical, "Chart Display Error"
            DoCmd.Close acForm, Me.name
        End If
    Else
        MsgBox "No chart specified to display. OpenArgs is null or empty.", vbExclamation, "Chart Display Error"
        DoCmd.Close acForm, Me.name
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading chart display: " & Err.Description, vbCritical, "Form Load Error"
    Debug.Print "ERROR in Form_Load: " & Err.Number & " - " & Err.Description
    DoCmd.Close acForm, Me.name
End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'
'    ' Auto-resize panels when form is resized
'    'If Me.WindowWidth > 6000 Then ' Minimum width check
'        ' Left panel: 45% of form width
'        'Me.lstPlanetPositions.Width = Me.WindowWidth * 0.45
'
'        '' Right panel: 50% of form width, positioned after left panel
'        'Me.lstAspects.Left = Me.lstPlanetPositions.Left + Me.lstPlanetPositions.Width + 200
'        'Me.lstAspects.Width = Me.WindowWidth * 0.5
'    'End If
'End Sub

' =================================================================
' CHART DATA LOADING
' =================================================================

Private Sub LoadChartData()
    On Error GoTo ErrorHandler
    
    Debug.Print "Loading chart data for ChartID: " & m_ChartID
    
    ' Load chart metadata
    LoadChartMetadata
    
    ' Format grid cells (do this once, early)
    FormatGridHeaders
    FormatGridCells
    
    ' ADD THIS LINE:
    PopulateGridHeaders
    
    ' Load planetary positions
    'LoadPlanetaryPositions
    
    ' Load aspects
    'LoadAspects
    
    ' Populate aspect grid
    PopulateAspectGrid
    
    ' Update form title
    Me.Caption = "Chart Display - " & m_ChartTitle
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error loading chart data: " & Err.Description, vbCritical, "Data Load Error"
    Debug.Print "ERROR in LoadChartData: " & Err.Description
End Sub

Private Sub LoadChartMetadata()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    Set db = CurrentDb()
    
    ' Get chart information
    sql = "SELECT c.*, p.FirstName, p.LastName, e.EventName " & _
          "FROM (tblCharts c " & _
          "LEFT JOIN tblPeople p ON c.PersonID = p.PersonID) " & _
          "LEFT JOIN tblEvents e ON c.EventID = e.EventID " & _
          "WHERE c.ChartID = " & m_ChartID
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        m_ChartType = Nz(rs!chartType, "Unknown")
        
        ' Build chart title based on chart type
        Select Case m_ChartType
            Case "Natal"
                m_ChartTitle = "Natal Chart - " & Nz(rs!FirstName, "") & " " & Nz(rs!LastName, "")
            Case "Event"
                m_ChartTitle = "Event Chart - " & Nz(rs!eventName, "Unknown Event")
            Case "Session"
                m_ChartTitle = "Session Chart - " & Format(rs!chartDate, "mm/dd/yyyy") & " " & Format(rs!chartTime, "h:nn AM/PM")
            Case Else
                m_ChartTitle = m_ChartType & " Chart - " & Format(rs!chartDate, "mm/dd/yyyy")
        End Select
        
        ' Update labels on form
        Me.lblChartTitle.Caption = m_ChartTitle
        Me.lblChartDate.Caption = "Date: " & Format(rs!chartDate, "dddd, mmmm d, yyyy") & " at " & Format(rs!chartTime, "h:nn AM/PM")
 '       Me.lblLocation.Caption =
    Else
        m_ChartTitle = "Chart Not Found"
        Me.lblChartTitle.Caption = "Chart ID " & m_ChartID & " not found"
    End If
    
    rs.Close
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in LoadChartMetadata: " & Err.Description
    m_ChartTitle = "Error Loading Chart"
End Sub

'Private Sub LoadPlanetaryPositions()
'    On Error GoTo ErrorHandler
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim sql As String
'    Dim positionText As String
'
'    Set db = CurrentDb()
'
'    ' Clear existing items
'    'Me.lstPlanetPositions.RowSource = ""
'
'    ' Get planetary positions with names and signs
'    sql = "SELECT cp.*, cb.BodyName, cb.Symbol, zs.SignName " & _
'          "FROM ((tblChartPositions cp " & _
'          "INNER JOIN tblCelestialBodies cb ON cp.BodyID = cb.BodyID) " & _
'          "LEFT JOIN tblZodiacSigns zs ON cp.SignID = zs.SignID) " & _
'          "WHERE cp.ChartID = " & m_ChartID & " " & _
'          "ORDER BY cb.DisplayOrder"
'
'    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
'
'    ' Build value list for listbox
'    Dim valueList As String
'    valueList = ""
'
'    Do While Not rs.EOF
'        ' Format: "Sun ? 197°13' Libra (R)"
'        Dim degrees As Integer, minutes As Integer
'        degrees = Int(rs!longitude)
'        minutes = Round((rs!longitude - degrees) * 60)
'
'        positionText = rs!BodyName & " " & _
'                      Nz(rs!Symbol, "") & " " & _
'                      degrees & "°" & Format(minutes, "00") & "' " & _
'                      Nz(rs!SignName, "") & _
'                      IIf(rs!IsRetrograde, " (R)", "")
'
'        If valueList <> "" Then valueList = valueList & ";"
'        valueList = valueList & positionText
'
'        rs.MoveNext
'    Loop
'
'    ' Set the listbox row source
'    'Me.lstPlanetPositions.RowSource = valueList
'
'    rs.Close
'    Set db = Nothing
'
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error in LoadPlanetaryPositions: " & Err.Description
'End Sub

'Private Sub LoadAspects()
'    On Error GoTo ErrorHandler
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim sql As String
'    Dim aspectText As String
'
'    Set db = CurrentDb()
'
'    ' Clear existing items
'    'Me.lstAspects.RowSource = ""
'
'    ' Get aspects with planet names and aspect names
'    sql = "SELECT ca.*, cb1.BodyName AS Body1Name, cb1.Symbol AS Body1Symbol, " & _
'          "cb2.BodyName AS Body2Name, cb2.Symbol AS Body2Symbol, " & _
'          "asp.AspectName, asp.Symbol AS AspectSymbol " & _
'          "FROM (((tblChartAspects ca " & _
'          "INNER JOIN tblCelestialBodies cb1 ON ca.Body1ID = cb1.BodyID) " & _
'          "INNER JOIN tblCelestialBodies cb2 ON ca.Body2ID = cb2.BodyID) " & _
'          "INNER JOIN tblAspects asp ON ca.AspectID = asp.AspectID) " & _
'          "WHERE ca.ChartID = " & m_ChartID & " " & _
'          "ORDER BY ca.ExactDegree"
'
'    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
'
'    ' Build value list for listbox
'    Dim valueList As String
'    valueList = ""
'
'    Do While Not rs.EOF
'        ' Format: "Sun ? Jupiter 0°27' (Applying)"
'        Dim degrees As Integer, minutes As Integer
'        degrees = Int(rs!ExactDegree)
'        minutes = Round((rs!ExactDegree - degrees) * 60)
'
'        aspectText = rs!Body1Name & " " & _
'                    Nz(rs!aspectSymbol, rs!AspectName) & " " & _
'                    rs!Body2Name & " " & _
'                    degrees & "°" & Format(minutes, "00") & "' " & _
'                    IIf(rs!OrbitApplying, "(Applying)", "(Separating)")
'
'        If valueList <> "" Then valueList = valueList & ";"
'        valueList = valueList & aspectText
'
'        rs.MoveNext
'    Loop
'
'    ' Set the listbox row source
'    'Me.lstAspects.RowSource = valueList
'
'    rs.Close
'    Set db = Nothing
'
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error in LoadAspects: " & Err.Description
'    'Me.lstAspects.RowSource = "Error loading aspects"
'End Sub

' =================================================================
' BUTTON EVENTS
' =================================================================

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub btnRefresh_Click()
    LoadChartData
End Sub

Private Sub btnPrint_Click()
    On Error GoTo ErrorHandler
    
    ' Print the form
    DoCmd.PrintOut acPrintAll
    Exit Sub
    
ErrorHandler:
    MsgBox "Error printing chart: " & Err.Description, vbCritical, "Print Error"
End Sub

'===========================================================
' END OF MODULE: frmAspectGrid
'===========================================================

' ADD THIS NEW METHOD:
Private Sub PopulateGridHeaders()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Integer
    
    Set db = CurrentDb()
    
    ' Get celestial bodies in display order
    sql = "SELECT BodyID, BodyName, Symbol, DisplayOrder " & _
          "FROM tblCelestialBodies " & _
          "WHERE IsActive = True " & _
          "ORDER BY DisplayOrder"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    i = 1
    Do While Not rs.EOF And i <= 18
        ' Populate horizontal header
        Me.Controls("header_" & i).Caption = Nz(rs!Symbol, Left(rs!BodyName, 2))
        
        ' Populate vertical header
        Me.Controls("sidebar_" & i).Caption = Nz(rs!Symbol, Left(rs!BodyName, 2))
        
        Debug.Print "Header " & i & ": " & rs!BodyName & " (" & Nz(rs!Symbol, "No Symbol") & ")"
        
        rs.MoveNext
        i = i + 1
    Loop
    
    ' Clear any remaining headers if we have fewer than 18 bodies
    Do While i <= 18
        Me.Controls("header_" & i).Caption = ""
        Me.Controls("sidebar_" & i).Caption = ""
        i = i + 1
    Loop
    
    rs.Close
    Set db = Nothing
    
    Debug.Print "Grid headers populated successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PopulateGridHeaders: " & Err.Description
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Set db = Nothing
End Sub

'Private Sub PopulateAspectGrid()
'    On Error GoTo ErrorHandler
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim rsBodyMap As DAO.Recordset
'    Dim sql As String
'    Dim i As Integer, j As Integer
'
'    ' First, clear all grid cells
'    For i = 1 To 18
'        For j = 1 To 18
'            Me.Controls("cell_" & i & "_" & j).value = ""
'        Next j
'    Next i
'
'    Set db = CurrentDb()
'
'    ' Create a mapping from BodyID to DisplayOrder position
'    Dim bodyPositions As New Collection
'    sql = "SELECT BodyID, DisplayOrder FROM tblCelestialBodies WHERE IsActive = True ORDER BY DisplayOrder"
'    Set rsBodyMap = db.OpenRecordset(sql, dbOpenSnapshot)
'
'    Dim pos As Integer
'    pos = 1
'    Do While Not rsBodyMap.EOF And pos <= 18
'        bodyPositions.Add pos, CStr(rsBodyMap!BodyID)
'        Debug.Print "Body mapping: BodyID " & rsBodyMap!BodyID & " -> Position " & pos
'        rsBodyMap.MoveNext
'        pos = pos + 1
'    Loop
'    rsBodyMap.Close
'
'    ' Get aspects for this chart
'    sql = "SELECT ca.*, asp.Symbol AS AspectSymbol " & _
'          "FROM tblChartAspects ca " & _
'          "INNER JOIN tblAspects asp ON ca.AspectID = asp.AspectID " & _
'          "WHERE ca.ChartID = " & m_ChartID
'
'    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
'
'    Do While Not rs.EOF
'        Dim body1Pos As Integer, body2Pos As Integer
'        Dim aspectText As String
'        Dim degrees As Integer
'
'        ' Get positions for both bodies
'        On Error Resume Next
'        body1Pos = bodyPositions(CStr(rs!body1ID))
'        body2Pos = bodyPositions(CStr(rs!body2ID))
'        On Error GoTo ErrorHandler
'
'        If body1Pos > 0 And body2Pos > 0 And body1Pos <= 18 And body2Pos <= 18 Then
'            ' Format the aspect text: Symbol + "  " + A/S + Degrees
'            degrees = Round(rs!ExactDegree)
'            aspectText = Nz(rs!aspectSymbol, "?") & "  " & _
'            IIf(rs!OrbitApplying, "A", "S") & Format(degrees, "00")
'
'            ' Populate both cells (aspect works both ways)
'            Me.Controls("cell_" & body1Pos & "_" & body2Pos).value = aspectText
'            Me.Controls("cell_" & body2Pos & "_" & body1Pos).value = aspectText
'
'            Debug.Print "Aspect: Body " & rs!body1ID & " (pos " & body1Pos & ") " & _
'                       rs!aspectSymbol & " Body " & rs!body2ID & " (pos " & body2Pos & ") = " & aspectText
'        End If
'
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set db = Nothing
'
'    Debug.Print "Aspect grid populated successfully"
'
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error in PopulateAspectGrid: " & Err.Description
'    If Not rs Is Nothing Then
'        rs.Close
'        Set rs = Nothing
'    End If
'    If Not rsBodyMap Is Nothing Then
'        rsBodyMap.Close
'        Set rsBodyMap = Nothing
'    End If
'    Set db = Nothing
'End Sub

Private Sub PopulateAspectGrid()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsBodyMap As DAO.Recordset
    Dim sql As String
    Dim i As Integer, j As Integer
    
    ' First, clear all grid cells
    For i = 1 To 18
        For j = 1 To 18
            Me.Controls("cell_" & i & "_" & j).value = ""
            ' Reset to default white background
            Me.Controls("cell_" & i & "_" & j).backColor = RGB(255, 255, 255)
        Next j
    Next i
    
    Set db = CurrentDb()
    
    ' Create a mapping from BodyID to DisplayOrder position
    Dim bodyPositions As New Collection
    sql = "SELECT BodyID, DisplayOrder FROM tblCelestialBodies WHERE IsActive = True ORDER BY DisplayOrder"
    Set rsBodyMap = db.OpenRecordset(sql, dbOpenSnapshot)
    
    Dim pos As Integer
    pos = 1
    Do While Not rsBodyMap.EOF And pos <= 18
        bodyPositions.Add pos, CStr(rsBodyMap!BodyID)
        Debug.Print "Body mapping: BodyID " & rsBodyMap!BodyID & " -> Position " & pos
        rsBodyMap.MoveNext
        pos = pos + 1
    Loop
    rsBodyMap.Close
    
    ' Get aspects for this chart with aspect type info
    sql = "SELECT ca.*, asp.Symbol AS AspectSymbol, asp.AspectType " & _
          "FROM tblChartAspects ca " & _
          "INNER JOIN tblAspects asp ON ca.AspectID = asp.AspectID " & _
          "WHERE ca.ChartID = " & m_ChartID
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    Do While Not rs.EOF
        Dim body1Pos As Integer, body2Pos As Integer
        Dim aspectText As String
        Dim degrees As Integer
        Dim cellColor As Long
        
        ' Get positions for both bodies
        On Error Resume Next
        body1Pos = bodyPositions(CStr(rs!body1ID))
        body2Pos = bodyPositions(CStr(rs!body2ID))
        On Error GoTo ErrorHandler
        
        If body1Pos > 0 And body2Pos > 0 And body1Pos <= 18 And body2Pos <= 18 Then
            
            degrees = Round(rs!ExactDegree)
            
             ' Split exact degree into degrees and minutes
            Dim wholeDegrees As Integer, minutes As Integer
            wholeDegrees = Int(rs!ExactDegree)
            minutes = Round((rs!ExactDegree - wholeDegrees) * 60)
            
            ' Format: Symbol + "  " + Degrees + A/S + Minutes
            aspectText = Nz(rs!aspectSymbol, "?") & " " & _
                        wholeDegrees & _
                        IIf(rs!OrbitApplying, "A", "S") & _
                        Format(minutes, "00")
            
            ' Set color based on aspect type
            Select Case UCase(Nz(rs!aspectType, ""))
                Case "MAJOR"
                    cellColor = RGB(255, 240, 240)  ' Light red for major aspects
                Case "MINOR"
                    cellColor = RGB(240, 255, 240)  ' Light green for minor aspects
                Case Else
                    cellColor = RGB(240, 240, 255)  ' Light blue for other aspects
            End Select
            
            ' Populate both cells (aspect works both ways)
            Me.Controls("cell_" & body1Pos & "_" & body2Pos).value = aspectText
            Me.Controls("cell_" & body1Pos & "_" & body2Pos).backColor = cellColor
            
            Me.Controls("cell_" & body2Pos & "_" & body1Pos).value = aspectText
            Me.Controls("cell_" & body2Pos & "_" & body1Pos).backColor = cellColor
            
            Debug.Print "Aspect: Body " & rs!body1ID & " (pos " & body1Pos & ") " & _
                       rs!aspectSymbol & " Body " & rs!body2ID & " (pos " & body2Pos & ") = " & aspectText & _
                       " (" & Nz(rs!aspectType, "Unknown") & ")"
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set db = Nothing
    
    Debug.Print "Aspect grid populated successfully with color coding"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in PopulateAspectGrid: " & Err.Description
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not rsBodyMap Is Nothing Then
        rsBodyMap.Close
        Set rsBodyMap = Nothing
    End If
    Set db = Nothing
End Sub

'Private Sub FormatGridCells()
'    On Error GoTo ErrorHandler
'
'    Dim i As Integer, j As Integer
'    Dim cellControl As Control
'
'    ' Format all grid cells
'    For i = 1 To 18
'        For j = 1 To 18
'            Set cellControl = Me.Controls("cell_" & i & "_" & j)
'
'            ' Set font properties
'            cellControl.FontName = "Arial"  ' Sans-serif font
'            cellControl.FontSize = 8
'            cellControl.TextAlign = 2       ' Center alignment (0=General, 1=Left, 2=Center, 3=Right)
'            cellControl.FontWeight = 400    ' Normal weight
'
'            ' Set background and border
'            cellControl.BackStyle = 1       ' Normal (opaque)
'            cellControl.backColor = RGB(255, 255, 255)  ' White background
'            cellControl.BorderStyle = 1     ' Solid border
'            cellControl.BorderWidth = 1     ' Thin border
'            cellControl.BorderColor = RGB(192, 192, 192)  ' Light gray border
'
'        Next j
'    Next i
'
'    ' Format headers too
'    For i = 1 To 18
'        ' Horizontal headers
'        Set cellControl = Me.Controls("header_" & i)
'        cellControl.FontName = "Arial"
'        cellControl.FontSize = 10
'        cellControl.FontWeight = 700    ' Bold
'        cellControl.TextAlign = 2       ' Center
'        cellControl.backColor = RGB(240, 240, 240)  ' Light gray background
'
'        ' Vertical headers
'        Set cellControl = Me.Controls("sidebar_" & i)
'        cellControl.FontName = "Arial"
'        cellControl.FontSize = 10
'        cellControl.FontWeight = 700    ' Bold
'        cellControl.TextAlign = 2       ' Center
'        cellControl.backColor = RGB(240, 240, 240)  ' Light gray background
'    Next i
'
'    Debug.Print "Grid cells formatted successfully"
'
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error in FormatGridCells: " & Err.Description
'End Sub

Private Sub FormatGridCells()
    On Error GoTo ErrorHandler
    
    Dim i As Integer, j As Integer
    Dim cellControl As Control
    
    ' Format all grid cells (these are TextBox controls)
    For i = 1 To 18
        For j = 1 To 18
            Set cellControl = Me.Controls("cell_" & i & "_" & j)
            
            With cellControl
                ' Set font properties
                .FontName = "Arial"      ' Sans-serif font
                .FontSize = 8
                .TextAlign = 2           ' Center alignment
                .FontWeight = 400        ' Normal weight
                
                ' Set background and border
                .BackStyle = 1           ' Normal (opaque)
                .backColor = RGB(255, 255, 255)  ' White background
                .BorderStyle = 1         ' Solid border
                .BorderWidth = 1         ' Thin border
                .BorderColor = RGB(192, 192, 192)  ' Light gray border
                .ForeColor = RGB(0, 0, 0)          ' Black text
                
                ' Make read-only for display
                .Locked = True
                .TabStop = False
            End With
        Next j
    Next i
    
    Debug.Print "Grid cells formatted successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in FormatGridCells: " & Err.Description & " at cell " & i & "_" & j
End Sub

Private Sub FormatGridHeaders()
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim headerControl As Control
    
    ' Format headers (these are Label controls, not TextBox controls)
    For i = 1 To 18
        ' Horizontal headers
        
        Set headerControl = Me.Controls("header_" & i)
        With headerControl
            .FontName = "Calibri"
            .FontSize = 14
            .FontWeight = 550        ' Bold
            .TextAlign = 2           ' Center alignment
            .BackStyle = 1           ' Normal (opaque)
            .backColor = RGB(240, 240, 240)  ' Light gray background
            .ForeColor = RGB(0, 0, 0)        ' BLACK TEXT - explicitly set
            .BorderStyle = 1         ' Solid border
            .BorderWidth = 1         ' Thin border
            .BorderColor = RGB(128, 128, 128)  ' Gray border
        End With
        
        ' Vertical headers
        Set headerControl = Me.Controls("sidebar_" & i)
        With headerControl
            .FontName = "Calibri"
            .FontSize = 14
            .FontWeight = 550        ' Bold
            .TextAlign = 2           ' Center alignment
            .BackStyle = 1           ' Normal (opaque)
            .backColor = RGB(240, 240, 240)  ' Light gray background
            .ForeColor = RGB(0, 0, 0)        ' Black text
            .BorderStyle = 1         ' Solid border
            .BorderWidth = 1         ' Thin border
            .BorderColor = RGB(128, 128, 128)  ' Gray border
        End With
    Next i
    
    Debug.Print "Grid headers formatted successfully"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in FormatGridHeaders: " & Err.Description & " at header " & i
End Sub

Private Sub btnDiagnostic_Click()
    Dim i As Integer
    Dim headerControl As Control
    
    ' Check the formatting of each header
    For i = 1 To 6  ' Just check first 6 headers
        Set headerControl = Me.Controls("header_" & i)
        Debug.Print "Header " & i & ": Caption=" & headerControl.Caption & _
                   ", ForeColor=" & headerControl.ForeColor & _
                   ", FontWeight=" & headerControl.FontWeight & _
                   ", FontName=" & headerControl.FontName
    Next i
End Sub
