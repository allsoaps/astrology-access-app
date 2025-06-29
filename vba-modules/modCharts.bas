Attribute VB_Name = "modCharts"
Option Compare Database
Option Explicit

Public Sub BuildAspectGridForm()
    Const strFormName As String = "frmAspectGrid"
    Dim frm As Access.Form
    
    ' 1) Open the pre-made form in Design view
    DoCmd.OpenForm strFormName, acDesign
    Set frm = Forms(strFormName)
    

    ' 2) Build the grid & labels on this open form
    BuildAspectGrid frm
    BuildPlanetaryLabels frm

    ' 3) Add the Close button if it doesn't already exist
    If Not ControlExists(strFormName, "cmdClose") Then
        Dim btn As Control
        Set btn = CreateControl( _
            formName:=strFormName, _
            ControlType:=acCommandButton, _
            Section:=acDetail, _
            Left:=450, _
            Top:=400)
        With btn
            .name = "cmdClose"
            .Caption = "Close"
            .Width = 1000
            .Height = 300
            .OnClick = "=CloseAspectGrid()"
        End With
    End If

    ' 4) Save & re-open the form for use
    DoCmd.Save acForm, strFormName
    DoCmd.Close acForm, strFormName, acSaveYes
    DoCmd.OpenForm strFormName, acNormal
End Sub

'================================================================
' Helper: Return True if frmName.Controls contains a controlName
'================================================================
Private Function ControlExists(frmName As String, ctrlName As String) As Boolean
    On Error Resume Next
    ControlExists = Not Forms(frmName)(ctrlName) Is Nothing
End Function

' Build the 18×18 grid of disabled textboxes on the open form
Private Sub BuildAspectGrid(frm As Access.Form)
    Dim row As Integer, col As Integer
    Dim txtBox As Control

    ' new, readable cell dimensions:
    Const boxW As Integer = 640    ' ~0.44" wide
    Const boxH As Integer = 640    ' ~0.44" tall

    ' 1" margin = 1440 twips; start your grid after that
    Const left0 As Integer = 1440
    Const top0  As Integer = 1440

    For row = 1 To 18
        For col = 1 To 18
            Set txtBox = CreateControl( _
                formName:=frm.name, _
                ControlType:=acTextBox, _
                Section:=acDetail, _
                Left:=left0 + (col - 1) * boxW, _
                Top:=top0 + (row - 1) * boxH)

            With txtBox
                .name = "cell_" & row & "_" & col
                .Width = boxW
                .Height = boxH
                .TextAlign = 2       ' center
                .Enabled = False     ' read-only
                If row = col Then
                    .backColor = RGB(220, 220, 220)
                End If
            End With
        Next col
    Next row
End Sub

' Build the header row and sidebar labels from tblCelestialBodies
Private Sub BuildPlanetaryLabels(frm As Access.Form)
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim i As Integer, lbl As Control

    ' same cell dims & margins as above
    Const boxW As Integer = 640
    Const boxH As Integer = 640
    Const left0 As Integer = 1440
    Const top0  As Integer = 1440

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
      "SELECT Symbol FROM tblCelestialBodies ORDER BY DisplayOrder")

    ' header row (above the grid)
    i = 1
    Do While Not rs.EOF And i <= 18
        Set lbl = CreateControl( _
            formName:=frm.name, _
            ControlType:=acLabel, _
            Section:=acDetail, _
            Left:=left0 + (i - 1) * boxW, _
            Top:=top0 - boxH)             ' one cell-height above
        With lbl
            .name = "header_" & i
            .Caption = rs!Symbol
            .Width = boxW
            .Height = boxH
            .TextAlign = 2
            .FontSize = 12
        End With
        i = i + 1: rs.MoveNext
    Loop

    ' sidebar (one cell-width to the left)
    rs.MoveFirst: i = 1
    Do While Not rs.EOF And i <= 18
        Set lbl = CreateControl( _
            formName:=frm.name, _
            ControlType:=acLabel, _
            Section:=acDetail, _
            Left:=left0 - boxW, _
            Top:=top0 + (i - 1) * boxH)
        With lbl
            .name = "sidebar_" & i
            .Caption = rs!Symbol
            .Width = boxW
            .Height = boxH
            .TextAlign = 2
            .FontSize = 12
        End With
        i = i + 1: rs.MoveNext
    Loop

    rs.Close
End Sub

'===========================
' Draw the 18×18 textbox grid
'===========================
Public Sub CreateAspectGrid()
    Dim frm     As Form
    Dim txtBox  As Control
    Dim row     As Integer, col As Integer
    Dim leftPos As Integer, topPos As Integer
    Dim boxW    As Integer, boxH    As Integer

    DoCmd.OpenForm "frmAspectGrid", acDesign
    Set frm = Forms("frmAspectGrid")

    boxW = 40: boxH = 25
    leftPos = 50: topPos = 50

    For row = 1 To 18
        For col = 1 To 18
            Set txtBox = CreateControl( _
                formName:=frm.name, _
                ControlType:=acTextBox, _
                Section:=acDetail, _
                Left:=leftPos + (col - 1) * boxW, _
                Top:=topPos + (row - 1) * boxH)

            With txtBox
                .name = "cell_" & row & "_" & col
                .Width = boxW
                .Height = boxH
                .TextAlign = 2        ' center
                .Enabled = False      ' read-only
                If row = col Then .backColor = RGB(220, 220, 220)
            End With
        Next
    Next

    DoCmd.Save acForm, frm.name
    DoCmd.Close acForm, frm.name
End Sub

'==============================
' Add header & sidebar labels
'==============================
Public Sub AddPlanetaryLabels()
    Dim frm       As Form
    Dim lbl       As Control
    Dim db        As DAO.Database
    Dim rs        As DAO.Recordset
    Dim i         As Integer
    Dim leftPos   As Integer, topPos   As Integer
    Dim boxW      As Integer, boxH     As Integer

    DoCmd.OpenForm "frmAspectGrid", acDesign
    Set frm = Forms("frmAspectGrid")

    boxW = 40: boxH = 25
    leftPos = 50: topPos = 25

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT BodyID, Symbol FROM tblCelestialBodies ORDER BY DisplayOrder")

    ' Top row
    i = 1
    Do While Not rs.EOF And i <= 18
        Set lbl = CreateControl( _
            formName:=frm.name, _
            ControlType:=acLabel, _
            Section:=acDetail, _
            Left:=leftPos + (i - 1) * boxW, _
            Top:=topPos - 25)
        With lbl
            .name = "header_" & i
            .Caption = rs!Symbol
            .Width = boxW
            .Height = 20
            .TextAlign = 2
            .FontSize = 12
        End With
        i = i + 1: rs.MoveNext
    Loop

    ' Left column
    rs.MoveFirst: i = 1
    Do While Not rs.EOF And i <= 18
        Set lbl = CreateControl( _
            formName:=frm.name, _
            ControlType:=acLabel, _
            Section:=acDetail, _
            Left:=leftPos - 25, _
            Top:=topPos + (i - 1) * boxH)
        With lbl
            .name = "sidebar_" & i
            .Caption = rs!Symbol
            .Width = 20
            .Height = boxH
            .TextAlign = 2
            .FontSize = 12
        End With
        i = i + 1: rs.MoveNext
    Loop

    rs.Close: Set rs = Nothing: Set db = Nothing

    DoCmd.Save acForm, frm.name
    DoCmd.Close acForm, frm.name
End Sub



'=======================================
' Populate & display for a given ChartID
'=======================================
Public Sub ShowChartAspectGrid(chartID As Long)
    PopulateAspectGrid chartID
End Sub

Public Sub PopulateAspectGrid(chartID As Long)
    Dim db    As DAO.Database
    Dim rs    As DAO.Recordset
    Dim rso   As DAO.Recordset
    Dim frm   As Form
    Dim dict  As New Scripting.Dictionary
    Dim row   As Integer, col As Integer
    Dim b1    As Integer, b2    As Integer
    Dim sym   As String, deg    As Double
    Dim info  As String

    DoCmd.OpenForm "frmAspectGrid", acNormal
    Set frm = Forms("frmAspectGrid")

    ClearAspectGrid

    Set db = CurrentDb

    ' Build BodyID?position dict
    Set rso = db.OpenRecordset("SELECT BodyID, DisplayOrder FROM tblCelestialBodies ORDER BY DisplayOrder")
    
    row = 1
    Do While Not rso.EOF And row <= 18
        Dim keyID As Long
        keyID = rso!BodyID
        
        If Not dict.Exists(keyID) Then
            dict.Add keyID, row
        End If
        
        row = row + 1
        rso.MoveNext
    Loop
    
    rso.Close

    ' Pull aspects
    Dim sql As String
    sql = "SELECT ca.*, cb1.BodyID AS Body1, cb2.BodyID AS Body2, " & _
          "a.Symbol, ca.ExactDegree " & _
          "FROM ((tblChartAspects AS ca " & _
          "INNER JOIN tblCelestialBodies AS cb1 ON ca.Body1ID=cb1.BodyID) " & _
          "INNER JOIN tblCelestialBodies AS cb2 ON ca.Body2ID=cb2.BodyID) " & _
          "INNER JOIN tblAspects         AS a   ON ca.AspectID = a.AspectID " & _
          "WHERE ca.ChartID=" & chartID & ";"
    
    
    Set rs = db.OpenRecordset(sql)
    Do While Not rs.EOF
        If dict.Exists(rs!Body1) And dict.Exists(rs!Body2) Then
            b1 = dict(rs!Body1)
            b2 = dict(rs!Body2)
            sym = rs!Symbol
            deg = rs!ExactDegree
            info = sym & " " & Format(deg, "0°")
            frm("cell_" & b1 & "_" & b2) = info
            If b1 <> b2 Then frm("cell_" & b2 & "_" & b1) = info
        End If
        rs.MoveNext
    Loop
       
    rs.Close

    ' Update caption
    Set rs = db.OpenRecordset("SELECT ChartType, ChartDate FROM tblCharts WHERE ChartID=" & chartID)
    If Not rs.EOF Then
        frm.Caption = "Aspect Grid: " & rs!chartType & " - " & Format(rs!chartDate, "mm/dd/yyyy")
    End If
    rs.Close

    Set db = Nothing

    ' Color-code
    ApplyAspectFormatting
End Sub

'=====================================
' Color-code cells by aspect type
'=====================================
Public Sub ApplyAspectFormatting()
    Dim frm   As Form
    Dim db    As DAO.Database
    Dim rs    As DAO.Recordset
    Dim dict  As New Scripting.Dictionary
    Dim row   As Integer, col As Integer
    Dim val   As String, sym As String
    Dim ctl   As Control

' using Segoe UI Symbol character set in table data

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT SymbolCode, AspectType FROM tblAspects")
    Do While Not rs.EOF
    'Debug.Print ChrW(rs!SymbolCode)
        'dict.Add rs!Symbol, rs!AspectType
        sym = ChrW(rs!SymbolCode)
        dict.Add sym, rs!aspectType
        rs.MoveNext
    Loop
    rs.Close

    Set frm = Forms("frmAspectGrid")
    For row = 1 To 18
        For col = 1 To 18
            Set ctl = frm("cell_" & row & "_" & col)
            val = Nz(ctl.value, "")
            If val <> "" Then
                sym = Left(val, InStr(val & " ", " ") - 1)
                If dict.Exists(sym) Then
                    Select Case dict(sym)
                        Case "Major"
                            Select Case sym
                                Case 9740: ctl.ForeColor = RGB(0, 128, 0)
                                Case 9741: ctl.ForeColor = RGB(192, 0, 0)
                                Case 9633: ctl.ForeColor = RGB(0, 128, 128)
                                Case 9651: ctl.ForeColor = RGB(0, 0, 192)
                                Case 9913: ctl.ForeColor = RGB(128, 0, 128)
                                Case Else: ctl.ForeColor = RGB(0, 0, 0)
                            End Select
                        Case "Minor"
                            ctl.ForeColor = RGB(128, 128, 128)
                        Case Else
                            ctl.ForeColor = RGB(0, 0, 0)
                    End Select
                    ctl.FontBold = True
                End If
            End If
        Next
    Next

    Set db = Nothing
End Sub

'===========================
' Close button callback
'===========================
Public Function CloseAspectGrid() As Long
    DoCmd.Close acForm, "frmAspectGrid"
    CloseAspectGrid = 0
End Function

Public Sub ClearAspectGrid()
    Dim r As Integer, c As Integer
    Dim frm As Form
    Set frm = Forms("frmAspectGrid")

    For r = 1 To 18
      For c = 1 To 18
        frm("cell_" & r & "_" & c).value = ""
      Next c
    Next r
End Sub



