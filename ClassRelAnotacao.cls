VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelAnotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sUltID As String
Private sUltTitulo As String
Private iUltTipoDocOrigem As Integer

Private sUltTexto As String

Function Anotacao_ObterTrecho(sRetorno As String, ByVal iTipoDocOrigem As Integer, ByVal sID As String, ByVal sTitulo As String, ByVal iPosInicial As Integer, ByVal iTam As Integer) As Long
'ex de tipodocorigem: ANOTACAO_ORIGEM_NFISCAL

Dim lErro As Long, sTexto As String, objAnotacao As New ClassAnotacoes

On Error GoTo Erro_Anotacao_ObterTrecho

    If sUltID <> sID Or sUltTitulo <> sTitulo Or iUltTipoDocOrigem <> iTipoDocOrigem Then
    
        With objAnotacao
        
            .iTipoDocOrigem = iTipoDocOrigem
            .sID = sID
            .sTitulo = sTitulo
            
        End With
    
        lErro = CF("Anotacao_Le_ComTextoCompleto", objAnotacao)
        If lErro <> SUCESSO And lErro <> 119875 Then gError 130408
        
        sUltID = sID
        sUltTitulo = sTitulo
        iUltTipoDocOrigem = iTipoDocOrigem
        
        If lErro <> SUCESSO Then
            sUltTexto = ""
        Else
            sUltTexto = objAnotacao.sTextoCompleto
        End If
        
    End If
    
    sTexto = sUltTexto
        
    If iPosInicial < 1 Then iPosInicial = 1
    If iTam > 254 Then iTam = 255
    
    sRetorno = Mid(sTexto, iPosInicial, iTam)
    
    Anotacao_ObterTrecho = SUCESSO
     
    Exit Function
    
Erro_Anotacao_ObterTrecho:

    Anotacao_ObterTrecho = gErr
     
    Select Case gErr
          
        Case 130408
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152784)
     
    End Select
     
    Exit Function

End Function

