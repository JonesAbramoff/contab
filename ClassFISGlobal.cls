VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get gobjFIS() As ClassFIS

Dim objFis As ClassFIS, lErro As Long

On Error GoTo Erro_gobjFIS

    If FisGlob_objFis Is Nothing Then
        Set objFis = New ClassFIS
        lErro = objFis.Inicializa
        If lErro <> SUCESSO Then gError 80152
        Set FisGlob_objFis = objFis
    End If
    
    Set gobjFIS = FisGlob_objFis
    
    Exit Property
    
Erro_gobjFIS:

    Set gobjFIS = Nothing
    
    Select Case gErr

        Case 80152
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 150293)

    End Select

    Exit Property

End Property

Private Sub Class_Initialize()
    FisGlob_Refs = FisGlob_Refs + 1
End Sub

Private Sub Class_Terminate()
    FisGlob_Refs = FisGlob_Refs - 1
    If FisGlob_Refs = 0 Then Set FisGlob_objFis = Nothing
End Sub



