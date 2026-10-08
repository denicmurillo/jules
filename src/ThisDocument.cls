VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    On Error GoTo ErrorHandler
    If ContentControl.Type = wdContentControlCheckBox Then
        If Left(ContentControl.Tag, 4) = "CHK_" Then
            modChecklist.ProcesarCambioCasilla ContentControl
        End If
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Error en evento ContentControlOnExit: " & Err.Description, vbCritical, "Error"
End Sub
