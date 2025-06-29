VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAspectOrbConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnSave_Click()
    Dim majorOrb As Double, minorOrb As Double
    
    majorOrb = val(Me.txtMajorOrb)
    minorOrb = val(Me.txtMinorOrb)
    
    If majorOrb > 0 And minorOrb > 0 Then
        modFormIntegration.ApplyAspectOrbConfig majorOrb, minorOrb
        MsgBox "Orb configuration saved!", vbInformation
        DoCmd.Close acForm, Me.name
    Else
End If
End Sub

