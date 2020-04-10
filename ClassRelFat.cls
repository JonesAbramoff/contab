VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DblParaTime(sHora As String, ByVal dValor As Double) As Long

Dim lErro As Long

On Error GoTo Erro_DblParaTime

    If dValor > 0 Then
        sHora = Left(CStr(CDate(dValor)), 5)
    Else
        sHora = "00:00"
    End If
    
    DblParaTime = SUCESSO
    
    Exit Function
    
Erro_DblParaTime:

    DblParaTime = gErr
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153084)
            
        End Select
        
    Exit Function

End Function
