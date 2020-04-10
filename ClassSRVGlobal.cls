VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSRVGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjSRV() As ClassSRV

Dim objSRV As ClassSRV, lErro As Long

On Error GoTo Erro_gobjSRV

    If SRVGlob_objSRV Is Nothing Then
        
        Set objSRV = New ClassSRV
        lErro = objSRV.Inicializa
        If lErro <> SUCESSO Then gError 183408
        Set SRVGlob_objSRV = objSRV
        
    End If
    
    Set gobjSRV = SRVGlob_objSRV
    
    Exit Property
    
Erro_gobjSRV:

    Set gobjSRV = Nothing
    
    Select Case gErr

        Case 183408
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183409)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()

    SRVGlob_Refs = SRVGlob_Refs + 1
    
End Sub


Private Sub Class_Terminate()

    SRVGlob_Refs = SRVGlob_Refs - 1
    If SRVGlob_Refs = 0 Then Set SRVGlob_objSRV = Nothing

End Sub

Public Sub Resetar()
    
    Set SRVGlob_objSRV = Nothing

End Sub



