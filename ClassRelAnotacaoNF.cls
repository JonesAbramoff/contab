VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelAnotacaoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lUltNumIntDoc As Long
Private sUltTitulo As String

Private sUltTexto As String

Function NF_ObterAnotacao(sRetorno As String, ByVal lNumIntDoc As Long, ByVal sTitulo As String, ByVal iPosInicial As Integer, ByVal iTam As Integer) As Long

Dim lErro As Long, sTexto As String, objAnotacao As New ClassAnotacoes

On Error GoTo Erro_NF_ObterAnotacao

    If lNumIntDoc <> lUltNumIntDoc Or sUltTitulo <> sTitulo Then
    
        With objAnotacao
        
            .iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
            .sID = CStr(lNumIntDoc)
            .sTitulo = sTitulo
            
        End With
    
        lErro = CF("Anotacao_Le_ComTextoCompleto", objAnotacao)
        If lErro <> SUCESSO And lErro <> 119875 Then gError 130408
        
        lUltNumIntDoc = lNumIntDoc
        sUltTitulo = sTitulo
        
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
    
    NF_ObterAnotacao = SUCESSO
     
    Exit Function
    
Erro_NF_ObterAnotacao:

    NF_ObterAnotacao = gErr
     
    Select Case gErr
          
        Case 130408
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152784)
     
    End Select
     
    Exit Function

End Function
