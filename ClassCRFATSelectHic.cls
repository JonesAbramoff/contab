VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFATSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Function NFiscal_Le_InfoUsu(ByVal objNFiscal As ClassNFiscal) As Long

Dim lErro As Long, lComando As Long, dValorFrete1 As Double
Dim lSubConta As Long, lHistorico As Long 'Inserido por Wagner

On Error GoTo Erro_NFiscal_Le_InfoUsu

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124192
    
    '####################################
    'Alterado por Wagner
    lErro = Comando_Executar(lComando, "SELECT ValorFrete1, Historico, SubConta FROM NFiscal WHERE NumIntDoc = ? AND Status <> ? AND FilialEmpresa = ?", dValorFrete1, lHistorico, lSubConta, objNFiscal.lNumIntDoc, STATUS_EXCLUIDO, giFilialEmpresa)
    If lErro <> AD_SQL_SUCESSO Then gError 124193
    '####################################
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124194
    If lErro <> AD_SQL_SUCESSO Then gError 124195

    Set objNFiscal.objInfoUsu = New ClassNFiscalHic
        
    objNFiscal.objInfoUsu.dValorFrete1 = dValorFrete1
    
    '####################################
    'Inserido por Wagner
    objNFiscal.objInfoUsu.lHistorico = lHistorico
    objNFiscal.objInfoUsu.lSubConta = lSubConta
    '####################################
    
    Call Comando_Fechar(lComando)
    
    NFiscal_Le_InfoUsu = SUCESSO
     
    Exit Function
    
Erro_NFiscal_Le_InfoUsu:

    NFiscal_Le_InfoUsu = gErr
     
    Select Case gErr
          
        Case 124192
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 124193, 124194, 124195
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_NFISCAL_HIC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 179174)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function TituloPagar_Gera(ByVal objNFiscal As ClassNFiscal, ByVal sTipoDocCPR As String, lNumIntDocCPR As Long) As Long
'registra nota fiscal fatura a pagar a partir dos dados em objNFiscal vinculada a nota fiscal emitida pelos modulos de faturamento/estoque
'retorna em lNumIntDocCPR o numintdoc do registro em TitulosPag

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New ColParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim objTiTPagHic As New ClassNfPagTitPagHic ' Inserido por Wagner
Dim objNFiscalHic As ClassNFiscalHic 'Inserido por Wagner

On Error GoTo Erro_TituloPagar_Gera

    objTituloPagar.dOutrasDespesas = objNFiscal.dValorOutrasDespesas
    objTituloPagar.dtDataEmissao = objNFiscal.dtDataEmissao
    objTituloPagar.dValorFrete = objNFiscal.dValorFrete
    objTituloPagar.dValorProdutos = objNFiscal.dValorProdutos
    objTituloPagar.dValorSeguro = objNFiscal.dValorSeguro
    objTituloPagar.dValorTotal = objNFiscal.dValorTotal
    objTituloPagar.iFilial = objNFiscal.iFilialForn
    objTituloPagar.iFilialEmpresa = objNFiscal.iFilialEmpresa
    objTituloPagar.iNumParcelas = objNFiscal.ColParcelaPagar.Count
    objTituloPagar.iStatus = STATUS_LANCADO
    objTituloPagar.lFornecedor = objNFiscal.lFornecedor
    objTituloPagar.lNumTitulo = objNFiscal.lNumNotaFiscal
    objTituloPagar.sSiglaDocumento = sTipoDocCPR
    objTituloPagar.dtDataRegistro = gdtDataHoje
    objTituloPagar.iCondicaoPagto = objNFiscal.iCondicaoPagto
    
    objTituloPagar.dValorICMS = objNFiscal.objTributacaoNF.dICMSValor
    objTituloPagar.dValorICMSSubst = objNFiscal.objTributacaoNF.dICMSSubstValor
    objTituloPagar.dValorIPI = objNFiscal.objTributacaoNF.dIPIValor
    objTituloPagar.dValorIRRF = objNFiscal.objTributacaoNF.dIRRFValor
    objTituloPagar.iCreditoICMS = IIf(objNFiscal.objTributacaoNF.dICMSCredito > 0.01, 1, 0)
    objTituloPagar.iCreditoIPI = IIf(objNFiscal.objTributacaoNF.dIPICredito > 0.01, 1, 0)
    objTituloPagar.dValorINSS = objNFiscal.objTributacaoNF.dValorINSS
    objTituloPagar.iINSSRetido = objNFiscal.objTributacaoNF.iINSSRetido
    objTituloPagar.dPISRetido = objNFiscal.objTributacaoNF.dPISRetido
    objTituloPagar.dISSRetido = objNFiscal.objTributacaoNF.dISSRetido
    objTituloPagar.dCOFINSRetido = objNFiscal.objTributacaoNF.dCOFINSRetido
    objTituloPagar.dCSLLRetido = objNFiscal.objTributacaoNF.dCSLLRetido
    
    objTituloPagar.dTaxaMoeda = objNFiscal.dTaxaMoeda
    
    objTituloPagar.lNumIntDoc = objNFiscal.lNumIntDocCPR
    
    '########################################
    'Inserido por Wagner
    If Not (objNFiscal.objInfoUsu Is Nothing) Then
    
        Set objNFiscalHic = objNFiscal.objInfoUsu
    
        objTiTPagHic.lHistorico = objNFiscalHic.lHistorico
        objTiTPagHic.lSubConta = objNFiscalHic.lSubConta
        
        Set objTituloPagar.objInfoUsu = objTiTPagHic
    End If
    '########################################
    
    For Each objParcelaPagar In objNFiscal.ColParcelaPagar
        With objParcelaPagar
            colParcelas.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras
        End With
    Next
    
    lErro = CF("NFFatPag_Grava_BD", objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then Error 46197
    
    lNumIntDocCPR = objTituloPagar.lNumIntDoc
    
    TituloPagar_Gera = SUCESSO
    
    Exit Function
    
Erro_TituloPagar_Gera:

    TituloPagar_Gera = Err
    
    Select Case Err
    
        Case 46197
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 179175)
            
    End Select
    
    Exit Function

End Function
