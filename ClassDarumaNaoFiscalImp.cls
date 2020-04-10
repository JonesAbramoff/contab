VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDarumaNaoFiscalImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DARUMA_NAO_FISCAL_MAX_TEXTO = 2000

Private sBuffer As String
Public sPulaLinha As String

Public sDirTemp As String

Public Function ImprimeNormal(ByVal sTexto As String) As Integer

Dim iResult As Integer

    sTexto = sTexto & sPulaLinha
    
    iResult = 1
    
    If Len(sBuffer & sTexto) > DARUMA_NAO_FISCAL_MAX_TEXTO Then
        iResult = DescarregaBuffer
    End If
    sBuffer = sBuffer & sTexto
    
    ImprimeNormal = iResult
    
End Function

Public Function TerminaImpressao() As Integer

Dim iResult As Integer

    iResult = DescarregaBuffer
    
    TerminaImpressao = iResult

End Function

Private Function DescarregaBuffer() As Integer

Dim iResult As Integer

    iResult = 1
    If Len(sBuffer) <> 0 Then
    
        iResult = iImprimirTexto_DUAL_DarumaFramework(sBuffer, 0)
        sBuffer = ""
    
    End If
    
    DescarregaBuffer = iResult
    
End Function

Private Sub Class_Initialize()
    sPulaLinha = "<l>"
End Sub
