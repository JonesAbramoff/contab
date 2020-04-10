VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTESGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjTES() As ClassTES

Dim objTES As ClassTES, lErro As Long

On Error GoTo Erro_gobjTES

    If TesGlob_objTES Is Nothing Then
        
        Set objTES = New ClassTES
        lErro = objTES.Inicializa
        If lErro <> SUCESSO Then Error 33708
        Set TesGlob_objTES = objTES
        
    End If
    
    Set gobjTES = TesGlob_objTES
    
    Exit Property
    
Erro_gobjTES:

    Set gobjTES = Nothing
    
    Select Case Err

        Case 33708
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153608)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    TesGlob_Refs = TesGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    TesGlob_Refs = TesGlob_Refs - 1
    If TesGlob_Refs = 0 Then Set TesGlob_objTES = Nothing
End Sub

Public Sub Resetar()
    
    Set TesGlob_objTES = Nothing

End Sub


