VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFeGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjNFeInfo() As ClassNFeInfo
Dim objNFeInfo As ClassNFeInfo, lErro As Long

On Error GoTo Erro_gobjNFe

    If NFeGlob_objNFe Is Nothing Then
    
        Set objNFeInfo = New ClassNFeInfo
        lErro = objNFeInfo.Inicializa
        If lErro <> SUCESSO Then Error 27728
        Set NFeGlob_objNFe = objNFeInfo
        
    End If
    
    Set gobjNFeInfo = NFeGlob_objNFe
    
    Exit Property
    
Erro_gobjNFe:

    Set gobjNFeInfo = Nothing
    
    Select Case Err

        Case 27728
        
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149744)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    NFeGlob_Refs = NFeGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    NFeGlob_Refs = NFeGlob_Refs - 1
    If NFeGlob_Refs = 0 Then Set NFeGlob_objNFe = Nothing
End Sub

Public Sub Resetar()
    
    Set NFeGlob_objNFe = Nothing

End Sub


