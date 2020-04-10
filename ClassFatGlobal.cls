VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFatGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjFAT() As ClassFAT
Dim objFAT As ClassFAT, lErro As Long
On Error GoTo Erro_gobjFAT

    If FatGlob_objFAT Is Nothing Then
    
        Set objFAT = New ClassFAT
        lErro = objFAT.Inicializa
        If lErro <> SUCESSO Then Error 27728
        Set FatGlob_objFAT = objFAT
        
    End If
    
    Set gobjFAT = FatGlob_objFAT
    
    Exit Property
    
Erro_gobjFAT:

    Set gobjFAT = Nothing
    
    Select Case Err

        Case 27728
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149744)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    FatGlob_Refs = FatGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    FatGlob_Refs = FatGlob_Refs - 1
    If FatGlob_Refs = 0 Then Set FatGlob_objFAT = Nothing
End Sub

Public Sub Resetar()
    
    Set FatGlob_objFAT = Nothing

End Sub
