VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCPGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjCP() As ClassCP

Dim objCP As ClassCP, lErro As Long

On Error GoTo Erro_gobjCP

    If CpGlob_objCP Is Nothing Then
        
        Set objCP = New ClassCP
        lErro = objCP.Inicializa
        If lErro <> SUCESSO Then Error 33671
        Set CpGlob_objCP = objCP
        
    End If
    
    Set gobjCP = CpGlob_objCP
    
    Exit Property
    
Erro_gobjCP:

    Set gobjCP = Nothing
    
    Select Case Err

        Case 33671
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 147886)

    End Select

    Exit Property

End Property


Private Sub Class_Initialize()
    CpGlob_Refs = CpGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    CpGlob_Refs = CpGlob_Refs - 1
    If CpGlob_Refs = 0 Then Set CpGlob_objCP = Nothing
End Sub

Public Sub Resetar()
    
    Set CpGlob_objCP = Nothing

End Sub


