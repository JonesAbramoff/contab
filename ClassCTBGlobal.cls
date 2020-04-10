VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassCTBGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjCTB() As ClassContab

Dim objCTB As ClassContab
Dim lErro As Long

On Error GoTo Erro_gobjCTB

    If CtbGlob_objCTB Is Nothing Then
        
        Set objCTB = New ClassContab
        Set CtbGlob_objCTB = objCTB
        
    End If
    
    Set gobjCTB = CtbGlob_objCTB
    
    Exit Property
    
Erro_gobjCTB:

    Set gobjCTB = Nothing
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149251)

    End Select

    Exit Property

End Property

