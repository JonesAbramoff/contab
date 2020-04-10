VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjCR() As ClassCR

Dim objCR As ClassCR, lErro As Long

On Error GoTo Erro_gobjCR

    If CrGlob_objCR Is Nothing Then
        
        Set objCR = New ClassCR
        lErro = objCR.Inicializa
        If lErro <> SUCESSO Then Error 33688
        Set CrGlob_objCR = objCR
        
    End If
    
    Set gobjCR = CrGlob_objCR
    
    Exit Property
    
Erro_gobjCR:

    Set gobjCR = Nothing
    
    Select Case Err

        Case 33688
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149221)

    End Select

    Exit Property

End Property




Private Sub Class_Initialize()

    CrGlob_Refs = CrGlob_Refs + 1

End Sub

Private Sub Class_Terminate()

    CrGlob_Refs = CrGlob_Refs - 1
    If CrGlob_Refs = 0 Then Set CrGlob_objCR = Nothing

End Sub

Public Sub Resetar()
    
    Set CrGlob_objCR = Nothing

End Sub


