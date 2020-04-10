VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSATGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjSATInfo() As ClassSATInfo
Dim objSATInfo As ClassSATInfo, lErro As Long

On Error GoTo Erro_gobjSAT

    If SatGlob_objSAT Is Nothing Then
    
        Set objSATInfo = New ClassSATInfo
        lErro = objSATInfo.Inicializa
        If lErro <> SUCESSO Then Error 27728
        Set SatGlob_objSAT = objSATInfo
        
    End If
    
    Set gobjSATInfo = SatGlob_objSAT
    
    Exit Property
    
Erro_gobjSAT:

    Set gobjSATInfo = Nothing
    
    Select Case Err

        Case 27728
        
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149744)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    SatGlob_Refs = SatGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    SatGlob_Refs = SatGlob_Refs - 1
    If SatGlob_Refs = 0 Then Set SatGlob_objSAT = Nothing
End Sub

Public Sub Resetar()
    
    Set SatGlob_objSAT = Nothing

End Sub

