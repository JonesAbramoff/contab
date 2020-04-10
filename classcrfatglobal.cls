VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get gobjCRFAT() As ClassCRFAT

Dim objCRFAT As ClassCRFAT, lErro As Long

On Error GoTo Erro_gobjCRFAT

    If CrFATGlob_objCRFAT Is Nothing Then
        
        Set objCRFAT = New ClassCRFAT
        lErro = objCRFAT.Inicializa
        If lErro <> SUCESSO Then Error 56654
        Set CrFATGlob_objCRFAT = objCRFAT
        
    End If
    
    Set gobjCRFAT = CrFATGlob_objCRFAT
    
    Exit Property
    
Erro_gobjCRFAT:

    Set gobjCRFAT = Nothing
    
    Select Case Err

        Case 56654
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148830)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()

    CrFATGlob_Refs = CrFATGlob_Refs + 1
    
End Sub


Private Sub Class_Terminate()

    CrFATGlob_Refs = CrFATGlob_Refs - 1
    If CrFATGlob_Refs = 0 Then Set CrFATGlob_objCRFAT = Nothing

End Sub

Public Sub Resetar()
    
    Set CrFATGlob_objCRFAT = Nothing

End Sub


