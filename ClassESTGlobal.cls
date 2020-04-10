VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassESTGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjEST() As ClassEST

Dim objEST As ClassEST, lErro As Long

On Error GoTo Erro_gobjEST

    If EstGlob_objEST Is Nothing Then
        
        Set objEST = New ClassEST
        lErro = objEST.Inicializa
        If lErro <> SUCESSO Then Error 33650
        Set EstGlob_objEST = objEST
        
    End If
    
    Set gobjEST = EstGlob_objEST
    
    Exit Property
    
Erro_gobjEST:

    Set gobjEST = Nothing
    
    Select Case Err

        Case 33650
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149481)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    EstGlob_Refs = EstGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    EstGlob_Refs = EstGlob_Refs - 1
    If EstGlob_Refs = 0 Then Set EstGlob_objEST = Nothing
End Sub

Public Sub Resetar()
    
    Set EstGlob_objEST = Nothing

End Sub

