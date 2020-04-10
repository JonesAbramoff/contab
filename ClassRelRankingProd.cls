VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelRankingProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private dValorFreteAnt As Double
Private lNumIntRelAnt As Long

Private dValorFreteCF As Double
Private dtDataCF As Date
Private iFilialCF As Integer
Private iCaixaCF As Integer

Public Function RelProdRanking_Obter_ValorFrete(dValorFrete As Double, ByVal iFilialEmpresa As Integer, ByVal dtDataDe As Date, ByVal dtDataAte As Date, ByVal lNumIntRel As Long) As Long
'Para obter o valor frete de uma filialempresa em um periodo

Dim lErro As Long

On Error GoTo Erro_RelProdRanking_Obter_ValorFrete

    If lNumIntRelAnt <> lNumIntRel Then

        lErro = CF("RelProdRanking_Obter_ValorFrete", iFilialEmpresa, dtDataDe, dtDataAte, lNumIntRel, dValorFrete)
        If lErro <> SUCESSO Then gError SUCESSO
        
        lNumIntRelAnt = lNumIntRel
        dValorFreteAnt = dValorFrete
    
    Else
    
        dValorFrete = dValorFreteAnt
    
    End If
    
    RelProdRanking_Obter_ValorFrete = SUCESSO

    Exit Function

Erro_RelProdRanking_Obter_ValorFrete:

    RelProdRanking_Obter_ValorFrete = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152929)

    End Select

    Exit Function

End Function

Public Function CF_Obter_ValorFrete(dValorFrete As Double, ByVal iFilialEmpresa As Integer, ByVal iCaixa As Integer, ByVal dtData As Date) As Long
'Para obter o valor frete de uma filialempresa em um periodo

Dim lErro As Long

On Error GoTo Erro_CF_Obter_ValorFrete

    If dtDataCF <> dtData Or iFilialCF <> iFilialEmpresa Or iCaixa <> iCaixaCF Then

        lErro = CF("CupomFiscal_Obter_ValorFrete", iFilialEmpresa, iCaixa, dtData, dValorFrete)
        If lErro <> SUCESSO Then gError SUCESSO
    
        dtDataCF = dtData
        iFilialCF = iFilialEmpresa
        iCaixaCF = iCaixa
        dValorFreteCF = dValorFrete
    
    Else
    
        dValorFrete = dValorFreteCF
    
    End If
    
    CF_Obter_ValorFrete = SUCESSO

    Exit Function

Erro_CF_Obter_ValorFrete:

    CF_Obter_ValorFrete = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 152929)

    End Select

    Exit Function

End Function
