VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMATGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjMAT() As ClassMAT

Dim objMAT As ClassMAT, lErro As Long

On Error GoTo Erro_gobjMAT

    If MatGlob_objMAT Is Nothing Then
        
        Set objMAT = New ClassMAT
        
        lErro = objMAT.Inicializa
        If lErro <> SUCESSO Then Error 59266
        
        Set MatGlob_objMAT = objMAT
        
    End If
    
    Set gobjMAT = MatGlob_objMAT
    
    Exit Property
    
Erro_gobjMAT:

    Set gobjMAT = Nothing
    
    Select Case Err

        Case 59266
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151675)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    MatGlob_Refs = MatGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    MatGlob_Refs = MatGlob_Refs - 1
    If MatGlob_Refs = 0 Then Set MatGlob_objMAT = Nothing
End Sub

Public Sub Resetar()
    
    Set MatGlob_objMAT = Nothing

End Sub


