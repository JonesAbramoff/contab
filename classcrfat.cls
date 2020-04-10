VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCRFAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvardPercJurosDiario As Double
Private mvardPercMulta As Double
Private mvariDescontoCodigo1 As Integer
Private mvariDescontoCodigo2 As Integer
Private mvariDescontoCodigo3 As Integer
Private mvardDescontoPerc1 As Double
Private mvardDescontoPerc2 As Double
Private mvardDescontoPerc3 As Double
Private mvariDescontoDias1 As Integer
Private mvariDescontoDias2 As Integer
Private mvariDescontoDias3 As Integer
Private mvarcolTiposDocInfo As New ColTipoDocInfo
Private mvarcolTiposDesconto As New AdmColCodigoNome
Private mvariISSNaBase As Integer
Private mvariIPIAbateDesconto As Integer
Private mvariCondPagtoSemCodigo As Integer
Private mvariDiasBloqueioAtraso As Integer
Private mvariUsaComissoesRegras As Integer
Private mvariUsaNFe As Integer
Private mvariGravaNFe As Integer
Private mvariCreditoVerificaLimite As Integer
Private mvarlFornEmp As Long
Private mvarlCliEmp As Long
Private mvarlFornecedorReceita As Long
Private mvarobjFilialEmpresa As AdmFiliais
Private mvariPermiteMultaSemAtraso As Integer
Private mvariValidaBordPagto As Integer
Private mvardTaxaFinanceira As Double
Private mvariBloquearPVPrecoBaixo As Integer
Private mvariBloquearPVMargemBaixa As Integer
Private mvariBloquearPVPrecoDefasado As Integer
Private mvarsDirBoletoGer As String
Private mvariUsaNFSE As Integer
Private mvariGravaNFSE As Integer
Private mvariProdRepetidoNF As Integer
Private mvariProdRepetidoPV As Integer
Private mvariProdRepetidoOV As Integer
Private mvariOVGravacaoVersao As Integer
Private mvariComisRetiraCreditosImp As Integer
Private mvariSeparaItensGradePrecoDif As Integer
Private mvariNaoTestarNFe As Integer
Private mvariVersaoNFe As Integer
Private mvariICMSSTComIEIsento As Integer
Private mvariNFImportacaoTribFlag01 As Integer
Private mvariNFImportacaoTribFlag02 As Integer
Private mvariNFImportacaoTribFlag03 As Integer
Private mvariNFImportacaoTribFlag04 As Integer
Private mvariNFImportacaoTribFlag05 As Integer
Private mvariNFImportacaoTribFlag06 As Integer
Private mvariNFImportacaoTribFlag07 As Integer
Private mvariNFImportacaoTribFlag08 As Integer
Private mvariNFImportacaoTribFlag09 As Integer
Private mvariNFImportacaoTribFlag10 As Integer
Private mvariNFImportacaoTribFlag11 As Integer
Private mvariNFImportacaoTribFlag12 As Integer
Private mvariNFImportacaoTribFlag13 As Integer
Private mvariNFImportacaoTribFlag14 As Integer
Private mvariNFImportacaoTribFlag15 As Integer
Private mvariNFTestaSomatorios As Integer
Private mvariCliAtrasoNaoGravaPV As Integer
Private mvariNFImportXml As Integer '0: nao pega nada, 1: pega tudo que puder
Private mvarsNFImportXmlProd As String 'contem a ordem dos criterios a serem usados para identificar o produto no corporator pelo que vem no xml
Private mvariTribIPINaoIncide As Integer '0: nao fazer nada, 1:colocar como 53 nao trib, 2: colocar como 99 outras
Private mvariTribICMSPFForaEst As Integer '0: nao fazer nada, 1:considerar lei 12.221/06 para vendas fora do estado para nao contribuinte
Private mvariTribCalcSTCobrAnt As Integer '0: nao calcula ST cobrada anteriormente, 1: calcula
Private mvarsExtensaoGerRelExp As String
Private mvariTitPagVerRepetSemSigla As Integer
Private mvardtResolucao13Senado As Date
Private mvardtEC872015 As Date
Private mvariEC872015Teste As Integer '0: nao é teste, 1:é versao de teste, ex. demo102
Private mvariEC872015Simples As Integer '0: nao calula, 1:calcula normalmente

Private mvariTotTribTipo As Integer
Private mvariTotTribCNPJ As Integer '0:com cnpj nao destaca, 1: destaca
Private mvariICMSProt21 As Integer '0: nao considera o protocolo 21, 1: usa a aliquota interestadual e coloca a diferenca p/a aliquota do destino como ST
Private mvariOtimizaTrib As Integer
Private mvariCriarPagSuspenso As Integer
Private mvariInfoAdicCclObrigatorio As Integer
Private mvariInfoAdicNatObrigatorio As Integer

Private mvariValidaDescMaxTabPreco As Integer
Private mvariUsaBloqAcessoPorTelaControle As Integer

Private mvariCalculaCFOPItem As Integer
Private mvariAcaoTesteCFOPTipoNF As Integer
Private mvarsNFeSistemaContingencia As String

Private mvariNaoTestarNFePesoVolume As Integer

Private mvariPrecoUnitarioFormPreco As Integer '1: calcula preco unitario usando planilha de formação de preços, 0:normal buscando em tabela de precos

Public iPisCofinsSimples99 As Integer '0: colocar 49, 1: colocar 99 como cst para pis e cofins qdo for venda no Simples
Public iVendaFuturaTributacao As Integer '0:tratamento padrao, 1:mesmo que seja 90 ou 99 é para zerar icms, pis e cofins em venda para entrega futura
Public iTrataSTEntEmCST00 As Integer
Public iCNABCEFNomeIniComE As Integer
Public dtIPIIniDescExcBase As Date
Public iICMSIncluiIPIBaseConsFinal As Integer
Public iCalcIPIDevolvido As Integer
Public iPISSimplesCalc As Integer

Public Property Get iPrecoUnitarioFormPreco() As Integer
    iPrecoUnitarioFormPreco = mvariPrecoUnitarioFormPreco
End Property

Public Property Let iPrecoUnitarioFormPreco(ByVal vData As Integer)
    mvariPrecoUnitarioFormPreco = vData
End Property

Public Property Get sNFeSistemaContingencia() As String
    sNFeSistemaContingencia = mvarsNFeSistemaContingencia
End Property

Public Property Let sNFeSistemaContingencia(ByVal vData As String)
    mvarsNFeSistemaContingencia = vData
End Property

Public Property Get iAcaoTesteCFOPTipoNF() As Integer
    iAcaoTesteCFOPTipoNF = mvariAcaoTesteCFOPTipoNF
End Property

Public Property Let iAcaoTesteCFOPTipoNF(ByVal vData As Integer)
    mvariAcaoTesteCFOPTipoNF = vData
End Property

Public Property Get iCalculaCFOPItem() As Integer
    iCalculaCFOPItem = mvariCalculaCFOPItem
End Property

Public Property Let iCalculaCFOPItem(ByVal vData As Integer)
    mvariCalculaCFOPItem = vData
End Property

Public Property Get iUsaBloqAcessoPorTelaControle() As Integer
    iUsaBloqAcessoPorTelaControle = mvariUsaBloqAcessoPorTelaControle
End Property

Public Property Let iUsaBloqAcessoPorTelaControle(ByVal vData As Integer)
    mvariUsaBloqAcessoPorTelaControle = vData
End Property

Public Property Get iValidaDescMaxTabPreco() As Integer
    iValidaDescMaxTabPreco = mvariValidaDescMaxTabPreco
End Property

Public Property Let iValidaDescMaxTabPreco(ByVal vData As Integer)
    mvariValidaDescMaxTabPreco = vData
End Property

Public Property Get iInfoAdicNatObrigatorio() As Integer
    iInfoAdicNatObrigatorio = mvariInfoAdicNatObrigatorio
End Property

Public Property Let iInfoAdicNatObrigatorio(ByVal vData As Integer)
    mvariInfoAdicNatObrigatorio = vData
End Property

Public Property Get iInfoAdicCclObrigatorio() As Integer
    iInfoAdicCclObrigatorio = mvariInfoAdicCclObrigatorio
End Property

Public Property Let iInfoAdicCclObrigatorio(ByVal vData As Integer)
    mvariInfoAdicCclObrigatorio = vData
End Property

Public Property Get iCriarPagSuspenso() As Integer
    iCriarPagSuspenso = mvariCriarPagSuspenso
End Property

Public Property Let iCriarPagSuspenso(ByVal vData As Integer)
    mvariCriarPagSuspenso = vData
End Property

Public Property Get iOtimizaTrib() As Integer
    iOtimizaTrib = mvariOtimizaTrib
End Property

Public Property Let iOtimizaTrib(ByVal vData As Integer)
    mvariOtimizaTrib = vData
End Property

Public Property Get iICMSProt21() As Integer
    iICMSProt21 = mvariICMSProt21
End Property

Public Property Let iICMSProt21(ByVal vData As Integer)
    mvariICMSProt21 = vData
End Property

Public Property Get iEC872015Teste() As Integer
    iEC872015Teste = mvariEC872015Teste
End Property

Public Property Let iEC872015Teste(ByVal vData As Integer)
    mvariEC872015Teste = vData
End Property

Public Property Get iEC872015Simples() As Integer
    iEC872015Simples = mvariEC872015Simples
End Property

Public Property Let iEC872015Simples(ByVal vData As Integer)
    mvariEC872015Simples = vData
End Property

Public Property Let dtResolucao13Senado(ByVal vData As Date)
    mvardtResolucao13Senado = vData
End Property

Public Property Get dtResolucao13Senado() As Date
    dtResolucao13Senado = mvardtResolucao13Senado
End Property

Public Property Let dtEC872015(ByVal vData As Date)
    mvardtEC872015 = vData
End Property

Public Property Get dtEC872015() As Date
    dtEC872015 = mvardtEC872015
End Property

Public Property Get iTitPagVerRepetSemSigla() As Integer
    iTitPagVerRepetSemSigla = mvariTitPagVerRepetSemSigla
End Property

Public Property Let iTitPagVerRepetSemSigla(ByVal vData As Integer)
    mvariTitPagVerRepetSemSigla = vData
End Property

Public Property Get sExtensaoGerRelExp() As String
    sExtensaoGerRelExp = mvarsExtensaoGerRelExp
End Property

Public Property Let sExtensaoGerRelExp(ByVal vData As String)
    mvarsExtensaoGerRelExp = vData
    gsExtensaoGerRelExp = mvarsExtensaoGerRelExp
End Property

Public Property Get iTribCalcSTCobrAnt() As Integer
    iTribCalcSTCobrAnt = mvariTribCalcSTCobrAnt
End Property

Public Property Let iTribCalcSTCobrAnt(ByVal vData As Integer)
    mvariTribCalcSTCobrAnt = vData
End Property

Public Property Get iTribIPINaoIncide() As Integer
    iTribIPINaoIncide = mvariTribIPINaoIncide
End Property

Public Property Let iTribIPINaoIncide(ByVal vData As Integer)
    mvariTribIPINaoIncide = vData
End Property

Public Property Get iTribICMSPFForaEst() As Integer
    iTribICMSPFForaEst = mvariTribICMSPFForaEst
End Property

Public Property Let iTribICMSPFForaEst(ByVal vData As Integer)
    mvariTribICMSPFForaEst = vData
End Property

Public Property Get sNFImportXmlProd() As String
   sNFImportXmlProd = mvarsNFImportXmlProd
End Property

Public Property Let sNFImportXmlProd(ByVal vData As String)
    mvarsNFImportXmlProd = vData
End Property

Public Property Get iNFImportXml() As Integer
    iNFImportXml = mvariNFImportXml
End Property

Public Property Let iNFImportXml(ByVal vData As Integer)
    mvariNFImportXml = vData
End Property

Public Property Get iNFTestaSomatorios() As Integer
    iNFTestaSomatorios = mvariNFTestaSomatorios
End Property

Public Property Let iNFTestaSomatorios(ByVal vData As Integer)
    mvariNFTestaSomatorios = vData
End Property

Public Property Get iNFImportacaoTribFlag10() As Integer
    iNFImportacaoTribFlag10 = mvariNFImportacaoTribFlag10
End Property

Public Property Let iNFImportacaoTribFlag10(ByVal vData As Integer)
    mvariNFImportacaoTribFlag10 = vData
End Property

Public Property Get iNFImportacaoTribFlag11() As Integer
    iNFImportacaoTribFlag11 = mvariNFImportacaoTribFlag11
End Property

Public Property Let iNFImportacaoTribFlag11(ByVal vData As Integer)
    mvariNFImportacaoTribFlag11 = vData
End Property

Public Property Get iNFImportacaoTribFlag15() As Integer
    iNFImportacaoTribFlag15 = mvariNFImportacaoTribFlag15
End Property

Public Property Let iNFImportacaoTribFlag15(ByVal vData As Integer)
    mvariNFImportacaoTribFlag15 = vData
End Property

Public Property Get iNFImportacaoTribFlag14() As Integer
    iNFImportacaoTribFlag14 = mvariNFImportacaoTribFlag14
End Property

Public Property Let iNFImportacaoTribFlag14(ByVal vData As Integer)
    mvariNFImportacaoTribFlag14 = vData
End Property

Public Property Get iNFImportacaoTribFlag13() As Integer
    iNFImportacaoTribFlag13 = mvariNFImportacaoTribFlag13
End Property

Public Property Let iNFImportacaoTribFlag13(ByVal vData As Integer)
    mvariNFImportacaoTribFlag13 = vData
End Property

Public Property Get iNFImportacaoTribFlag12() As Integer
    iNFImportacaoTribFlag12 = mvariNFImportacaoTribFlag12
End Property

Public Property Let iNFImportacaoTribFlag12(ByVal vData As Integer)
    mvariNFImportacaoTribFlag12 = vData
End Property

Public Property Get iNFImportacaoTribFlag09() As Integer
    iNFImportacaoTribFlag09 = mvariNFImportacaoTribFlag09
End Property

Public Property Let iNFImportacaoTribFlag09(ByVal vData As Integer)
    mvariNFImportacaoTribFlag09 = vData
End Property

Public Property Get iNFImportacaoTribFlag08() As Integer
    iNFImportacaoTribFlag08 = mvariNFImportacaoTribFlag08
End Property

Public Property Let iNFImportacaoTribFlag08(ByVal vData As Integer)
    mvariNFImportacaoTribFlag08 = vData
End Property

Public Property Get iNFImportacaoTribFlag07() As Integer
    iNFImportacaoTribFlag07 = mvariNFImportacaoTribFlag07
End Property

Public Property Let iNFImportacaoTribFlag07(ByVal vData As Integer)
    mvariNFImportacaoTribFlag07 = vData
End Property

Public Property Get iNFImportacaoTribFlag06() As Integer
    iNFImportacaoTribFlag06 = mvariNFImportacaoTribFlag06
End Property

Public Property Let iNFImportacaoTribFlag06(ByVal vData As Integer)
    mvariNFImportacaoTribFlag06 = vData
End Property

Public Property Get iNFImportacaoTribFlag05() As Integer
    iNFImportacaoTribFlag05 = mvariNFImportacaoTribFlag05
End Property

Public Property Let iNFImportacaoTribFlag05(ByVal vData As Integer)
    mvariNFImportacaoTribFlag05 = vData
End Property

Public Property Get iNFImportacaoTribFlag04() As Integer
    iNFImportacaoTribFlag04 = mvariNFImportacaoTribFlag04
End Property

Public Property Let iNFImportacaoTribFlag04(ByVal vData As Integer)
    mvariNFImportacaoTribFlag04 = vData
End Property

Public Property Get iNFImportacaoTribFlag03() As Integer
    iNFImportacaoTribFlag03 = mvariNFImportacaoTribFlag03
End Property

Public Property Let iNFImportacaoTribFlag03(ByVal vData As Integer)
    mvariNFImportacaoTribFlag03 = vData
End Property

Public Property Get iNFImportacaoTribFlag02() As Integer
    iNFImportacaoTribFlag02 = mvariNFImportacaoTribFlag02
End Property

Public Property Let iNFImportacaoTribFlag02(ByVal vData As Integer)
    mvariNFImportacaoTribFlag02 = vData
End Property

Public Property Get iNFImportacaoTribFlag01() As Integer
    iNFImportacaoTribFlag01 = mvariNFImportacaoTribFlag01
End Property

Public Property Let iNFImportacaoTribFlag01(ByVal vData As Integer)
    mvariNFImportacaoTribFlag01 = vData
End Property

Public Property Get iICMSSTComIEIsento() As Integer
    iICMSSTComIEIsento = mvariICMSSTComIEIsento
End Property

Public Property Let iICMSSTComIEIsento(ByVal vData As Integer)
    mvariICMSSTComIEIsento = vData
End Property

Public Property Get iVersaoNFE() As Integer
    iVersaoNFE = mvariVersaoNFe
End Property

Public Property Let iVersaoNFE(ByVal vData As Integer)
    mvariVersaoNFe = vData
End Property

Public Property Get iSeparaItensGradePrecoDif() As Integer
    iSeparaItensGradePrecoDif = mvariSeparaItensGradePrecoDif
End Property

Public Property Let iSeparaItensGradePrecoDif(ByVal vData As Integer)
    mvariSeparaItensGradePrecoDif = vData
End Property

Public Property Get iComisRetiraCreditosImp() As Integer
    iComisRetiraCreditosImp = mvariComisRetiraCreditosImp
End Property

Public Property Let iComisRetiraCreditosImp(ByVal vData As Integer)
    mvariComisRetiraCreditosImp = vData
End Property

Public Property Get iOVGravacaoVersao() As Integer
    iOVGravacaoVersao = mvariOVGravacaoVersao
End Property

Public Property Let iOVGravacaoVersao(ByVal vData As Integer)
    mvariOVGravacaoVersao = vData
End Property

Public Property Get sDirBoletoGer() As String
   sDirBoletoGer = mvarsDirBoletoGer
End Property

Public Property Let sDirBoletoGer(ByVal vData As String)
    mvarsDirBoletoGer = vData
End Property

Public Property Let dTaxaFinanceira(ByVal vData As Double)
    mvardTaxaFinanceira = vData
End Property

Public Property Get dTaxaFinanceira() As Double
    dTaxaFinanceira = mvardTaxaFinanceira
End Property

Public Property Get iBloquearPVPrecoBaixo() As Integer
    iBloquearPVPrecoBaixo = mvariBloquearPVPrecoBaixo
End Property

Public Property Let iBloquearPVPrecoBaixo(ByVal vData As Integer)
    mvariBloquearPVPrecoBaixo = vData
End Property

Public Property Get iBloquearPVPrecoDefasado() As Integer
    iBloquearPVPrecoDefasado = mvariBloquearPVPrecoDefasado
End Property

Public Property Let iBloquearPVPrecoDefasado(ByVal vData As Integer)
    mvariBloquearPVPrecoDefasado = vData
End Property
Public Property Get iValidaBordPagto() As Integer
    iValidaBordPagto = mvariValidaBordPagto
End Property

Public Property Let iValidaBordPagto(ByVal vData As Integer)
    mvariValidaBordPagto = vData
End Property

Public Property Get iPermiteMultaSemAtraso() As Integer
    iPermiteMultaSemAtraso = mvariPermiteMultaSemAtraso
End Property

Public Property Let iPermiteMultaSemAtraso(ByVal vData As Integer)
    mvariPermiteMultaSemAtraso = vData
End Property

Public Property Get iCreditoVerificaLimite() As Integer
    iCreditoVerificaLimite = mvariCreditoVerificaLimite
End Property

Public Property Let iCreditoVerificaLimite(ByVal vData As Integer)
    mvariCreditoVerificaLimite = vData
End Property

Public Property Get iProdRepetidoNF() As Integer
    iProdRepetidoNF = mvariProdRepetidoNF
End Property

Public Property Let iProdRepetidoNF(ByVal vData As Integer)
    mvariProdRepetidoNF = vData
End Property

Public Property Get iProdRepetidoPV() As Integer
    iProdRepetidoPV = mvariProdRepetidoPV
End Property

Public Property Let iProdRepetidoPV(ByVal vData As Integer)
    mvariProdRepetidoPV = vData
End Property

Public Property Get iProdRepetidoOV() As Integer
    iProdRepetidoOV = mvariProdRepetidoOV
End Property

Public Property Let iProdRepetidoOV(ByVal vData As Integer)
    mvariProdRepetidoOV = vData
End Property

Public Property Let lFornecedorReceita(ByVal vData As Long)
    mvarlFornecedorReceita = vData
End Property

Public Property Get lFornecedorReceita() As Long
    lFornecedorReceita = mvarlFornecedorReceita
End Property

Public Property Let lFornEmp(ByVal vData As Long)
    mvarlFornEmp = vData
End Property

Public Property Get lFornEmp() As Long
    lFornEmp = mvarlFornEmp
End Property

Public Property Let lCliEmp(ByVal vData As Long)
    mvarlCliEmp = vData
End Property

Public Property Get lCliEmp() As Long
    lCliEmp = mvarlCliEmp
End Property

Public Property Get iISSNaBase() As Integer
    iISSNaBase = mvariISSNaBase
End Property

Public Property Let iISSNaBase(ByVal vData As Integer)
    mvariISSNaBase = vData
End Property

Public Property Get iIPIAbateDesconto() As Integer
    iIPIAbateDesconto = mvariIPIAbateDesconto
End Property

Public Property Let iIPIAbateDesconto(ByVal vData As Integer)
    mvariIPIAbateDesconto = vData
End Property

Public Property Get iCondPagtoSemCodigo() As Integer
    iCondPagtoSemCodigo = mvariCondPagtoSemCodigo
End Property

Public Property Let iCondPagtoSemCodigo(ByVal vData As Integer)
    mvariCondPagtoSemCodigo = vData
End Property

Public Property Get objFilialEmpresa() As AdmFiliais

    Set objFilialEmpresa = mvarobjFilialEmpresa
    
End Property

Public Property Get colTiposDesconto() As AdmColCodigoNome
    Set colTiposDesconto = mvarcolTiposDesconto
End Property

Public Property Get colTiposDocInfo() As ColTipoDocInfo
    Set colTiposDocInfo = mvarcolTiposDocInfo
End Property

Public Property Get iCliAtrasoNaoGravaPV() As Integer
    iCliAtrasoNaoGravaPV = mvariCliAtrasoNaoGravaPV
End Property

Public Property Let iCliAtrasoNaoGravaPV(ByVal vData As Integer)
    mvariCliAtrasoNaoGravaPV = vData
End Property

Public Function Inicializa() As Long

Dim lErro As Long
Dim colConfigCRFAT As New ColCRFATConfig

On Error GoTo Erro_Inicializa

    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO1, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_CODIGO1
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO2, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_CODIGO2
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO3, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_CODIGO3
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS1, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_DIAS1
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS2, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_DIAS2
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS3, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_DIAS3
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL1, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_PERCENTUAL1
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL2, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_PERCENTUAL2
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL3, EMPRESA_TODA, "", 0, "", CRFATCFG_DESCONTO_PADRAO_PERCENTUAL3
    colConfigCRFAT.Add CRFATCFG_JUROS_PADRAO, EMPRESA_TODA, "", 0, "", CRFATCFG_JUROS_PADRAO
    colConfigCRFAT.Add CRFATCFG_MULTA_PADRAO, EMPRESA_TODA, "", 0, "", CRFATCFG_MULTA_PADRAO
    colConfigCRFAT.Add CRFATCFG_ISS_NA_BASE, EMPRESA_TODA, "", 0, "", CRFATCFG_ISS_NA_BASE
    colConfigCRFAT.Add CRFATCFG_IPI_ABATE_DESCONTO, EMPRESA_TODA, "", 0, "", CRFATCFG_IPI_ABATE_DESCONTO
    colConfigCRFAT.Add CRFATCFG_DIAS_BLOQUEIO_ATRASO, EMPRESA_TODA, "", 0, "", CRFATCFG_DIAS_BLOQUEIO_ATRASO
    colConfigCRFAT.Add CRFATCFG_USA_COMISSOES_REGRAS, EMPRESA_TODA, "", 0, "", CRFATCFG_USA_COMISSOES_REGRAS 'Por Leo em 15/03/02
    colConfigCRFAT.Add CRFATCFG_CREDITO_VERIFICA_LIMITE, EMPRESA_TODA, "", 0, "", CRFATCFG_CREDITO_VERIFICA_LIMITE 'tulio130103
    colConfigCRFAT.Add CRFATCFG_COD_FORNECEDOR_EMP, EMPRESA_TODA, "", 0, "", CRFATCFG_COD_FORNECEDOR_EMP
    colConfigCRFAT.Add CRFATCFG_COD_CLIENTE_EMP, EMPRESA_TODA, "", 0, "", CRFATCFG_COD_CLIENTE_EMP
    colConfigCRFAT.Add CRFATCFG_COD_FORNECEDOR_RECEITA, EMPRESA_TODA, "", 0, "", CRFATCFG_COD_FORNECEDOR_RECEITA 'correcao por tulio 22/07/02
    colConfigCRFAT.Add CRFATCFG_PERMITE_MULTA_SEM_ATRASO, EMPRESA_TODA, "", 0, "", CRFATCFG_PERMITE_MULTA_SEM_ATRASO
    colConfigCRFAT.Add CRFATCFG_CONDPAGTO_SEM_CODIGO, EMPRESA_TODA, "", 0, "", CRFATCFG_CONDPAGTO_SEM_CODIGO
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_USA_NFE, giFilialEmpresa, "", 0, "", CRFATCFG_USA_NFE
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_GRAVA_NFE, giFilialEmpresa, "", 0, "", CRFATCFG_GRAVA_NFE
    colConfigCRFAT.Add CRFATCFG_VALIDA_BORDERO_PAGTO, EMPRESA_TODA, "", 0, "", CRFATCFG_VALIDA_BORDERO_PAGTO
    colConfigCRFAT.Add CRFATCFG_TAXAFINANCEIRA_PADRAO, EMPRESA_TODA, "", 0, "", CRFATCFG_TAXAFINANCEIRA_PADRAO
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_PRECO_BAIXO, EMPRESA_TODA, "", 0, "", CRFATCFG_BLOQUEARPV_PRECO_BAIXO
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_MARGEM_BAIXA, EMPRESA_TODA, "", 0, "", CRFATCFG_BLOQUEARPV_MARGEM_BAIXA
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_PRECO_DEFASADO, EMPRESA_TODA, "", 0, "", CRFATCFG_BLOQUEARPV_PRECO_DEFASADO
    colConfigCRFAT.Add CRFATCFG_DIRETORIO_BOLETO_GERACAO, EMPRESA_TODA, "", 0, "", CRFATCFG_DIRETORIO_BOLETO_GERACAO
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_USA_NFSE, giFilialEmpresa, "", 0, "", CRFATCFG_USA_NFSE
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_GRAVA_NFSE, giFilialEmpresa, "", 0, "", CRFATCFG_GRAVA_NFSE
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_NF, EMPRESA_TODA, "", 0, "", CRFATCFG_PROD_REPETIDO_NF
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_PV, EMPRESA_TODA, "", 0, "", CRFATCFG_PROD_REPETIDO_PV
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_OV, EMPRESA_TODA, "", 0, "", CRFATCFG_PROD_REPETIDO_OV
    colConfigCRFAT.Add CRFATCFG_OV_GRAVACAO_VERSAO, EMPRESA_TODA, "", 0, "", CRFATCFG_OV_GRAVACAO_VERSAO
    colConfigCRFAT.Add CRFATCFG_COMISSAO_RETIRA_CREDITO_IMPOSTOS, EMPRESA_TODA, "", 0, "", CRFATCFG_COMISSAO_RETIRA_CREDITO_IMPOSTOS
    colConfigCRFAT.Add CRFATCFG_SEPARA_ITENSGRADE_PRECO_DIF, EMPRESA_TODA, "", 0, "", CRFATCFG_SEPARA_ITENSGRADE_PRECO_DIF
    colConfigCRFAT.Add CRFATCFG_NAO_TESTAR_NFE, EMPRESA_TODA, "", 0, "", CRFATCFG_NAO_TESTAR_NFE
    colConfigCRFAT.Add CRFATCFG_NAO_TESTAR_NFE_PESO_VOLUME, EMPRESA_TODA, "", 0, "", CRFATCFG_NAO_TESTAR_NFE_PESO_VOLUME
    colConfigCRFAT.Add CRFATCFG_VERSAO_LAYOUT_NFE, EMPRESA_TODA, "", 0, "", CRFATCFG_VERSAO_LAYOUT_NFE
    colConfigCRFAT.Add CRFATCFG_ICMSST_COM_IE_ISENTO, EMPRESA_TODA, "", 0, "", CRFATCFG_ICMSST_COM_IE_ISENTO
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG01, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG01
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG02, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG02
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG03, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG03
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG04, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG04
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG05, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG05
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG06, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG06
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG07, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG07
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG08, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG08
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG09, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG09
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG10, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG10
    
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG11, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG11
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG12, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG12
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG13, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG13
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG14, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG14
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_TRB_FLAG15, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_TRB_FLAG15
    
    colConfigCRFAT.Add CRFATCFG_NF_TESTA_SOMATORIOS, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_TESTA_SOMATORIOS
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_XML, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_XML
    colConfigCRFAT.Add CRFATCFG_NF_IMPORTACAO_XML_PROD, EMPRESA_TODA, "", 0, "", CRFATCFG_NF_IMPORTACAO_XML_PROD
    colConfigCRFAT.Add CRFATCFG_TRIB_IPI_NAO_INCIDE, EMPRESA_TODA, "", 0, "", CRFATCFG_TRIB_IPI_NAO_INCIDE
    colConfigCRFAT.Add CRFATCFG_TRIB_ICMS_PFFORAEST, EMPRESA_TODA, "", 0, "", CRFATCFG_TRIB_ICMS_PFFORAEST
    colConfigCRFAT.Add CRFATCFG_TRIB_CALCSTCOBRANT, EMPRESA_TODA, "", 0, "", CRFATCFG_TRIB_CALCSTCOBRANT
    
    colConfigCRFAT.Add CRFATCFG_CLI_ATRASO_DESTACAR, EMPRESA_TODA, "", 0, "", CRFATCFG_CLI_ATRASO_DESTACAR
    colConfigCRFAT.Add CRFATCFG_CLI_ATRASO_NAO_GRAVAR_PV, EMPRESA_TODA, "", 0, "", CRFATCFG_CLI_ATRASO_NAO_GRAVAR_PV
    
    colConfigCRFAT.Add CRFATCFG_EXTENSAO_GER_REL_EXPORTA, EMPRESA_TODA, "", 0, "", CRFATCFG_EXTENSAO_GER_REL_EXPORTA
    colConfigCRFAT.Add CRFATCFG_TITPAG_VERIFICA_REPETICAO_SEM_SIGLA, EMPRESA_TODA, "", 0, "", CRFATCFG_TITPAG_VERIFICA_REPETICAO_SEM_SIGLA
    
    colConfigCRFAT.Add CRFATCFG_DATA_RESOLUCAO_13_SENADO, EMPRESA_TODA, "", 0, "", CRFATCFG_DATA_RESOLUCAO_13_SENADO
    colConfigCRFAT.Add CRFATCFG_DATA_EC_87_2015, EMPRESA_TODA, "", 0, "", CRFATCFG_DATA_EC_87_2015
    
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_TOT_TRIB_TIPO, giFilialEmpresa, "", 0, "", CRFATCFG_TOT_TRIB_TIPO
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_TOT_TRIB_CNPJ, giFilialEmpresa, "", 0, "", CRFATCFG_TOT_TRIB_CNPJ
    
    colConfigCRFAT.Add CRFATCFG_ICMS_PROTOCOLO_21, EMPRESA_TODA, "", 0, "", CRFATCFG_ICMS_PROTOCOLO_21
    colConfigCRFAT.Add CRFATCFG_EC_87_2015_TESTE, EMPRESA_TODA, "", 0, "", CRFATCFG_EC_87_2015_TESTE
    colConfigCRFAT.Add CRFATCFG_EC_87_2015_SIMPLES, EMPRESA_TODA, "", 0, "", CRFATCFG_EC_87_2015_SIMPLES
    
    colConfigCRFAT.Add CRFATCFG_OTIMIZA_TRIBUTACAO, EMPRESA_TODA, "", 0, "", CRFATCFG_OTIMIZA_TRIBUTACAO
    
    colConfigCRFAT.Add CRFATCFG_CRIAR_PAG_SUSPENSO, EMPRESA_TODA, "", 0, "", CRFATCFG_CRIAR_PAG_SUSPENSO
    
    colConfigCRFAT.Add CRFATCFG_INFOADIC_NATUREZA_OBRIGATORIA, EMPRESA_TODA, "", 0, "", CRFATCFG_INFOADIC_NATUREZA_OBRIGATORIA
    colConfigCRFAT.Add CRFATCFG_INFOADIC_CCL_OBRIGATORIA, EMPRESA_TODA, "", 0, "", CRFATCFG_INFOADIC_CCL_OBRIGATORIA
    colConfigCRFAT.Add CRFATCFG_VALIDA_DESC_MAX_TAB_PRECO, EMPRESA_TODA, "", 0, "", CRFATCFG_VALIDA_DESC_MAX_TAB_PRECO
    
    colConfigCRFAT.Add CRFATCFG_USA_BLOQ_ACESSO_POR_TELA_CONTROLE, EMPRESA_TODA, "", 0, "", CRFATCFG_USA_BLOQ_ACESSO_POR_TELA_CONTROLE
    
    colConfigCRFAT.Add CRFATCFG_CALCULA_CFOP_ITEM, EMPRESA_TODA, "", 0, "", CRFATCFG_CALCULA_CFOP_ITEM
    
    colConfigCRFAT.Add CRFATCFG_ACAO_TESTE_CFOP_TIPONF, EMPRESA_TODA, "", 0, "", CRFATCFG_ACAO_TESTE_CFOP_TIPONF
    
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_NFE_SISTEMA_CONTINGENCIA, giFilialEmpresa, "", 0, "", CRFATCFG_NFE_SISTEMA_CONTINGENCIA
    
    colConfigCRFAT.Add CRFATCFG_PRECO_UNITARIO_FORM_PRECO, EMPRESA_TODA, "", 0, "", CRFATCFG_PRECO_UNITARIO_FORM_PRECO
    
    colConfigCRFAT.Add CRFATCFG_PISCOFINS_SIMPLES_99, EMPRESA_TODA, "", 0, "", CRFATCFG_PISCOFINS_SIMPLES_99
    colConfigCRFAT.Add CRFATCFG_VENDA_FUTURA_TRIBUTACAO, EMPRESA_TODA, "", 0, "", CRFATCFG_VENDA_FUTURA_TRIBUTACAO
    
    colConfigCRFAT.Add CRFATCFG_TRIB_TRATA_ST_ENTRADA_EM_CST00, EMPRESA_TODA, "", 0, "", CRFATCFG_TRIB_TRATA_ST_ENTRADA_EM_CST00
    
    colConfigCRFAT.Add CRFATCFG_CNAB_CEF_NOME_INI_COM_E, EMPRESA_TODA, "", 0, "", CRFATCFG_CNAB_CEF_NOME_INI_COM_E
    
    colConfigCRFAT.Add CRFATCFG_IPI_INICIO_EXCLUSAO_DESC_DA_BASE, EMPRESA_TODA, "", 0, "", CRFATCFG_IPI_INICIO_EXCLUSAO_DESC_DA_BASE
    colConfigCRFAT.Add CRFATCFG_ICMS_INCLUI_IPI_BASE_CONSUMIDOR_FINAL, EMPRESA_TODA, "", 0, "", CRFATCFG_ICMS_INCLUI_IPI_BASE_CONSUMIDOR_FINAL
    
    colConfigCRFAT.Add CRFATCFG_IPI_CALCULO_VLR_DEVOLVIDO, EMPRESA_TODA, "", 0, "", CRFATCFG_IPI_CALCULO_VLR_DEVOLVIDO
    colConfigCRFAT.Add CRFATCFG_PIS_TIPO_CALCULO_NO_SIMPLES, EMPRESA_TODA, "", 0, "", CRFATCFG_PIS_TIPO_CALCULO_NO_SIMPLES
    
    'Lê as configurações em CRFATConfig
    lErro = CF("CRFATConfig_Le_Configs", colConfigCRFAT)
    If lErro <> SUCESSO Then Error 56553
    
    iPISSimplesCalc = CInt(colConfigCRFAT.Item(CRFATCFG_PIS_TIPO_CALCULO_NO_SIMPLES).sConteudo)
    iCalcIPIDevolvido = CInt(colConfigCRFAT.Item(CRFATCFG_IPI_CALCULO_VLR_DEVOLVIDO).sConteudo)
    
    iICMSIncluiIPIBaseConsFinal = CInt(colConfigCRFAT.Item(CRFATCFG_ICMS_INCLUI_IPI_BASE_CONSUMIDOR_FINAL).sConteudo)
    dtIPIIniDescExcBase = StrParaDate(colConfigCRFAT.Item(CRFATCFG_IPI_INICIO_EXCLUSAO_DESC_DA_BASE).sConteudo)
    
    iCNABCEFNomeIniComE = CInt(colConfigCRFAT.Item(CRFATCFG_CNAB_CEF_NOME_INI_COM_E).sConteudo)
    
    iTrataSTEntEmCST00 = CInt(colConfigCRFAT.Item(CRFATCFG_TRIB_TRATA_ST_ENTRADA_EM_CST00).sConteudo)
    
    iVendaFuturaTributacao = CInt(colConfigCRFAT.Item(CRFATCFG_VENDA_FUTURA_TRIBUTACAO).sConteudo)
    iPisCofinsSimples99 = CInt(colConfigCRFAT.Item(CRFATCFG_PISCOFINS_SIMPLES_99).sConteudo)
    
    mvariPrecoUnitarioFormPreco = CInt(colConfigCRFAT.Item(CRFATCFG_PRECO_UNITARIO_FORM_PRECO).sConteudo)
    
    If giFilialEmpresa <> EMPRESA_TODA Then mvarsNFeSistemaContingencia = colConfigCRFAT.Item(CRFATCFG_NFE_SISTEMA_CONTINGENCIA).sConteudo
    
    mvariAcaoTesteCFOPTipoNF = CInt(colConfigCRFAT.Item(CRFATCFG_ACAO_TESTE_CFOP_TIPONF).sConteudo)
    
    mvariCalculaCFOPItem = CInt(colConfigCRFAT.Item(CRFATCFG_CALCULA_CFOP_ITEM).sConteudo)
    mvariUsaBloqAcessoPorTelaControle = CInt(colConfigCRFAT.Item(CRFATCFG_USA_BLOQ_ACESSO_POR_TELA_CONTROLE).sConteudo)
    
    mvariValidaDescMaxTabPreco = CInt(colConfigCRFAT.Item(CRFATCFG_VALIDA_DESC_MAX_TAB_PRECO).sConteudo)
    
    mvariInfoAdicCclObrigatorio = CInt(colConfigCRFAT.Item(CRFATCFG_INFOADIC_CCL_OBRIGATORIA).sConteudo)
    mvariInfoAdicNatObrigatorio = CInt(colConfigCRFAT.Item(CRFATCFG_INFOADIC_NATUREZA_OBRIGATORIA).sConteudo)
    
    mvariTitPagVerRepetSemSigla = CInt(colConfigCRFAT.Item(CRFATCFG_TITPAG_VERIFICA_REPETICAO_SEM_SIGLA).sConteudo)
    
    sExtensaoGerRelExp = colConfigCRFAT.Item(CRFATCFG_EXTENSAO_GER_REL_EXPORTA).sConteudo
    
    mvariPermiteMultaSemAtraso = CInt(colConfigCRFAT.Item(CRFATCFG_PERMITE_MULTA_SEM_ATRASO).sConteudo)
    mvariDiasBloqueioAtraso = CInt(colConfigCRFAT.Item(CRFATCFG_DIAS_BLOQUEIO_ATRASO).sConteudo) 'por leo em 26/02/02
    mvardPercJurosDiario = CDbl(colConfigCRFAT.Item(CRFATCFG_JUROS_PADRAO).sConteudo)
    mvardPercMulta = CDbl(colConfigCRFAT.Item(CRFATCFG_MULTA_PADRAO).sConteudo)
    mvariDescontoCodigo1 = CInt(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_CODIGO1).sConteudo)
    mvariDescontoCodigo2 = CInt(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_CODIGO2).sConteudo)
    mvariDescontoCodigo3 = CInt(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_CODIGO3).sConteudo)
    mvardDescontoPerc1 = CDbl(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_PERCENTUAL1).sConteudo)
    mvardDescontoPerc2 = CDbl(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_PERCENTUAL2).sConteudo)
    mvardDescontoPerc3 = CDbl(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_PERCENTUAL3).sConteudo)
    mvariDescontoDias1 = CInt(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_DIAS1).sConteudo)
    mvariDescontoDias2 = CInt(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_DIAS2).sConteudo)
    mvariDescontoDias3 = CInt(colConfigCRFAT.Item(CRFATCFG_DESCONTO_PADRAO_DIAS3).sConteudo)
    mvariISSNaBase = CInt(colConfigCRFAT.Item(CRFATCFG_ISS_NA_BASE).sConteudo)
    mvariIPIAbateDesconto = CInt(colConfigCRFAT.Item(CRFATCFG_IPI_ABATE_DESCONTO).sConteudo)
    mvariCondPagtoSemCodigo = CInt(colConfigCRFAT.Item(CRFATCFG_CONDPAGTO_SEM_CODIGO).sConteudo)
    mvariUsaComissoesRegras = CInt(colConfigCRFAT.Item(CRFATCFG_USA_COMISSOES_REGRAS).sConteudo) 'por leo em 15/03/02
    If giFilialEmpresa <> EMPRESA_TODA Then mvariUsaNFe = CInt(colConfigCRFAT.Item(CRFATCFG_USA_NFE).sConteudo)
    If giFilialEmpresa <> EMPRESA_TODA Then mvariGravaNFe = CInt(colConfigCRFAT.Item(CRFATCFG_GRAVA_NFE).sConteudo)
    mvariCreditoVerificaLimite = CInt(colConfigCRFAT.Item(CRFATCFG_CREDITO_VERIFICA_LIMITE).sConteudo) 'por tulio130103
    mvarlFornEmp = CLng(colConfigCRFAT.Item(CRFATCFG_COD_FORNECEDOR_EMP).sConteudo)
    mvarlCliEmp = CLng(colConfigCRFAT.Item(CRFATCFG_COD_CLIENTE_EMP).sConteudo)
    mvarlFornecedorReceita = CLng(colConfigCRFAT.Item(CRFATCFG_COD_FORNECEDOR_RECEITA).sConteudo)
    mvariValidaBordPagto = CInt(colConfigCRFAT.Item(CRFATCFG_VALIDA_BORDERO_PAGTO).sConteudo)
    mvariBloquearPVPrecoBaixo = CInt(colConfigCRFAT.Item(CRFATCFG_BLOQUEARPV_PRECO_BAIXO).sConteudo)
    mvariBloquearPVMargemBaixa = CInt(colConfigCRFAT.Item(CRFATCFG_BLOQUEARPV_MARGEM_BAIXA).sConteudo)
    mvariBloquearPVPrecoDefasado = CInt(colConfigCRFAT.Item(CRFATCFG_BLOQUEARPV_PRECO_DEFASADO).sConteudo)
    mvardTaxaFinanceira = CDbl(colConfigCRFAT.Item(CRFATCFG_TAXAFINANCEIRA_PADRAO).sConteudo)
    mvarsDirBoletoGer = colConfigCRFAT.Item(CRFATCFG_DIRETORIO_BOLETO_GERACAO).sConteudo
    If giFilialEmpresa <> EMPRESA_TODA Then mvariUsaNFSE = CInt(colConfigCRFAT.Item(CRFATCFG_USA_NFSE).sConteudo)
    If giFilialEmpresa <> EMPRESA_TODA Then mvariGravaNFSE = CInt(colConfigCRFAT.Item(CRFATCFG_GRAVA_NFSE).sConteudo)
    mvariProdRepetidoNF = CInt(colConfigCRFAT.Item(CRFATCFG_PROD_REPETIDO_NF).sConteudo)
    mvariProdRepetidoPV = CInt(colConfigCRFAT.Item(CRFATCFG_PROD_REPETIDO_PV).sConteudo)
    mvariProdRepetidoOV = CInt(colConfigCRFAT.Item(CRFATCFG_PROD_REPETIDO_OV).sConteudo)
    mvariOVGravacaoVersao = CInt(colConfigCRFAT.Item(CRFATCFG_OV_GRAVACAO_VERSAO).sConteudo)
    mvariComisRetiraCreditosImp = CInt(colConfigCRFAT.Item(CRFATCFG_COMISSAO_RETIRA_CREDITO_IMPOSTOS).sConteudo)
    mvariSeparaItensGradePrecoDif = CInt(colConfigCRFAT.Item(CRFATCFG_SEPARA_ITENSGRADE_PRECO_DIF).sConteudo)
    mvariNaoTestarNFe = CInt(colConfigCRFAT.Item(CRFATCFG_NAO_TESTAR_NFE).sConteudo)
    mvariNaoTestarNFePesoVolume = CInt(colConfigCRFAT.Item(CRFATCFG_NAO_TESTAR_NFE_PESO_VOLUME).sConteudo)
    mvariVersaoNFe = CInt(colConfigCRFAT.Item(CRFATCFG_VERSAO_LAYOUT_NFE).sConteudo)
    mvariICMSSTComIEIsento = CInt(colConfigCRFAT.Item(CRFATCFG_ICMSST_COM_IE_ISENTO).sConteudo)
    mvariNFImportacaoTribFlag01 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG01).sConteudo)
    mvariNFImportacaoTribFlag02 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG02).sConteudo)
    mvariNFImportacaoTribFlag03 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG03).sConteudo)
    mvariNFImportacaoTribFlag04 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG04).sConteudo)
    mvariNFImportacaoTribFlag05 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG05).sConteudo)
    mvariNFImportacaoTribFlag06 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG06).sConteudo)
    mvariNFImportacaoTribFlag07 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG07).sConteudo)
    mvariNFImportacaoTribFlag08 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG08).sConteudo)
    mvariNFImportacaoTribFlag09 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG09).sConteudo)
    mvariNFImportacaoTribFlag10 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG10).sConteudo)
    
    mvariNFImportacaoTribFlag11 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG11).sConteudo)
    mvariNFImportacaoTribFlag12 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG12).sConteudo)
    mvariNFImportacaoTribFlag13 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG13).sConteudo)
    mvariNFImportacaoTribFlag14 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG14).sConteudo)
    mvariNFImportacaoTribFlag15 = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_TRB_FLAG15).sConteudo)
    
    mvariNFTestaSomatorios = CInt(colConfigCRFAT.Item(CRFATCFG_NF_TESTA_SOMATORIOS).sConteudo)
    mvariNFImportXml = CInt(colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_XML).sConteudo)
    mvarsNFImportXmlProd = colConfigCRFAT.Item(CRFATCFG_NF_IMPORTACAO_XML_PROD).sConteudo
    mvariTribIPINaoIncide = CInt(colConfigCRFAT.Item(CRFATCFG_TRIB_IPI_NAO_INCIDE).sConteudo)
    mvariTribICMSPFForaEst = CInt(colConfigCRFAT.Item(CRFATCFG_TRIB_ICMS_PFFORAEST).sConteudo)
    mvariTribCalcSTCobrAnt = CInt(colConfigCRFAT.Item(CRFATCFG_TRIB_CALCSTCOBRANT).sConteudo)
    
    giCliAtrasoDestacar = CInt(colConfigCRFAT.Item(CRFATCFG_CLI_ATRASO_DESTACAR).sConteudo)
    mvariCliAtrasoNaoGravaPV = CInt(colConfigCRFAT.Item(CRFATCFG_CLI_ATRASO_NAO_GRAVAR_PV).sConteudo)
    
    mvardtResolucao13Senado = StrParaDate(colConfigCRFAT.Item(CRFATCFG_DATA_RESOLUCAO_13_SENADO).sConteudo)
    mvardtEC872015 = StrParaDate(colConfigCRFAT.Item(CRFATCFG_DATA_EC_87_2015).sConteudo)
    
    If giFilialEmpresa <> EMPRESA_TODA Then mvariTotTribTipo = CInt(colConfigCRFAT.Item(CRFATCFG_TOT_TRIB_TIPO).sConteudo)
    If giFilialEmpresa <> EMPRESA_TODA Then mvariTotTribCNPJ = CInt(colConfigCRFAT.Item(CRFATCFG_TOT_TRIB_CNPJ).sConteudo)
    
    mvariICMSProt21 = CInt(colConfigCRFAT.Item(CRFATCFG_ICMS_PROTOCOLO_21).sConteudo)
    mvariEC872015Teste = CInt(colConfigCRFAT.Item(CRFATCFG_EC_87_2015_TESTE).sConteudo)
    mvariEC872015Simples = CInt(colConfigCRFAT.Item(CRFATCFG_EC_87_2015_SIMPLES).sConteudo)
    mvariOtimizaTrib = CInt(colConfigCRFAT.Item(CRFATCFG_OTIMIZA_TRIBUTACAO).sConteudo)
    mvariCriarPagSuspenso = CInt(colConfigCRFAT.Item(CRFATCFG_CRIAR_PAG_SUSPENSO).sConteudo)
    
    If giFilialEmpresa <> EMPRESA_TODA Then
    
        Set mvarobjFilialEmpresa = New AdmFiliais
        
        'fazer a leitura
        mvarobjFilialEmpresa.iCodFilial = giFilialEmpresa
        lErro = CF("FilialEmpresa_Le", mvarobjFilialEmpresa)
        If lErro <> SUCESSO Then Error 59343
    
    End If

    'preenche colecao de tiposdocinfo
    lErro = CF("TiposDocInfo_Le", mvarcolTiposDocInfo)
    If lErro <> SUCESSO Then Error 56553
    
    'Lê o código e a descrição de todos os Tipos de Desconto
    lErro = CF("Cod_Nomes_Le", "TiposDeDesconto", "Codigo", "DescReduzida", STRING_TIPOSDEDESCONTO_DESCRICAO, mvarcolTiposDesconto)
    If lErro <> SUCESSO Then Error 26492

    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 26492
        Case 56553, 59343
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 148813)
        
    End Select

    Exit Function

End Function

Public Property Get iTotTribTipo() As Integer
    iTotTribTipo = mvariTotTribTipo
End Property

Public Property Let iTotTribTipo(ByVal vData As Integer)
    mvariTotTribTipo = vData
End Property

Public Property Get iTotTribCNPJ() As Integer
    iTotTribCNPJ = mvariTotTribCNPJ
End Property

Public Property Let iTotTribCNPJ(ByVal vData As Integer)
    mvariTotTribCNPJ = vData
End Property

Public Property Let dDescontoPerc1(ByVal vData As Double)
    mvardDescontoPerc1 = vData
End Property

Public Property Get dDescontoPerc1() As Double
    dDescontoPerc1 = mvardDescontoPerc1
End Property

Public Property Let dDescontoPerc2(ByVal vData As Double)
    mvardDescontoPerc2 = vData
End Property

Public Property Get dDescontoPerc2() As Double
    dDescontoPerc2 = mvardDescontoPerc2
End Property

Public Property Let dDescontoPerc3(ByVal vData As Double)
    mvardDescontoPerc3 = vData
End Property

Public Property Get dDescontoPerc3() As Double
    dDescontoPerc3 = mvardDescontoPerc3
End Property

Public Property Get iDescontoDias1() As Integer
    iDescontoDias1 = mvariDescontoDias1
End Property

Public Property Let iDescontoDias1(ByVal vData As Integer)
    mvariDescontoDias1 = vData
End Property

Public Property Get iDescontoDias2() As Integer
    iDescontoDias2 = mvariDescontoDias2
End Property

Public Property Let iDescontoDias2(ByVal vData As Integer)
    mvariDescontoDias2 = vData
End Property

Public Property Get iDescontoDias3() As Integer
    iDescontoDias3 = mvariDescontoDias3
End Property

Public Property Let iDescontoDias3(ByVal vData As Integer)
    mvariDescontoDias3 = vData
End Property

Public Property Get iDescontoCodigo1() As Integer
    iDescontoCodigo1 = mvariDescontoCodigo1
End Property

Public Property Let iDescontoCodigo1(ByVal vData As Integer)
    mvariDescontoCodigo1 = vData
End Property

Public Property Get iDescontoCodigo2() As Integer
    iDescontoCodigo2 = mvariDescontoCodigo2
End Property

Public Property Let iDescontoCodigo2(ByVal vData As Integer)
    mvariDescontoCodigo2 = vData
End Property

Public Property Get iDescontoCodigo3() As Integer
    iDescontoCodigo3 = mvariDescontoCodigo3
End Property

Public Property Let iDescontoCodigo3(ByVal vData As Integer)
    mvariDescontoCodigo3 = vData
End Property

Public Property Let dPercMulta(ByVal vData As Double)
    mvardPercMulta = vData
End Property

Public Property Get dPercMulta() As Double
    dPercMulta = mvardPercMulta
End Property

Public Property Let dPercJurosDiario(ByVal vData As Double)
    mvardPercJurosDiario = vData
End Property

Public Property Get dPercJurosDiario() As Double
    dPercJurosDiario = mvardPercJurosDiario
End Property

Function Gravar() As Long
'Grava no bd a configuração atual

Dim lErro As Long
Dim colConfigCRFAT As New ColCRFATConfig

On Error GoTo Erro_Gravar
    
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO1, EMPRESA_TODA, "", 0, CStr(mvariDescontoCodigo1)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO2, EMPRESA_TODA, "", 0, CStr(mvariDescontoCodigo2)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO3, EMPRESA_TODA, "", 0, CStr(mvariDescontoCodigo3)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS1, EMPRESA_TODA, "", 0, CStr(mvariDescontoDias1)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS2, EMPRESA_TODA, "", 0, CStr(mvariDescontoDias2)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS3, EMPRESA_TODA, "", 0, CStr(mvariDescontoDias3)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL1, EMPRESA_TODA, "", 0, CStr(mvardDescontoPerc1)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL2, EMPRESA_TODA, "", 0, CStr(mvardDescontoPerc2)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL3, EMPRESA_TODA, "", 0, CStr(mvardDescontoPerc3)
    colConfigCRFAT.Add CRFATCFG_JUROS_PADRAO, EMPRESA_TODA, "", 0, CStr(mvardPercJurosDiario)
    colConfigCRFAT.Add CRFATCFG_MULTA_PADRAO, EMPRESA_TODA, "", 0, CStr(mvardPercMulta)
    colConfigCRFAT.Add CRFATCFG_ISS_NA_BASE, EMPRESA_TODA, "", 0, CStr(mvariISSNaBase)
    colConfigCRFAT.Add CRFATCFG_IPI_ABATE_DESCONTO, EMPRESA_TODA, "", 0, CStr(mvariIPIAbateDesconto)
    colConfigCRFAT.Add CRFATCFG_CONDPAGTO_SEM_CODIGO, EMPRESA_TODA, "", 0, CStr(mvariCondPagtoSemCodigo)
    colConfigCRFAT.Add CRFATCFG_DIAS_BLOQUEIO_ATRASO, EMPRESA_TODA, "", 0, CStr(mvariDiasBloqueioAtraso) 'por leo em 26/02/02
    colConfigCRFAT.Add CRFATCFG_USA_COMISSOES_REGRAS, EMPRESA_TODA, "", 0, CStr(mvariUsaComissoesRegras)  'por leo em 15/13/02
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_USA_NFE, giFilialEmpresa, "", 0, CStr(mvariUsaNFe)
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_GRAVA_NFE, giFilialEmpresa, "", 0, CStr(mvariGravaNFe)
    colConfigCRFAT.Add CRFATCFG_CREDITO_VERIFICA_LIMITE, EMPRESA_TODA, "", 0, CStr(mvariCreditoVerificaLimite)  'tulio130102
    colConfigCRFAT.Add CRFATCFG_COD_FORNECEDOR_EMP, EMPRESA_TODA, "", 0, CStr(mvarlFornEmp)
    colConfigCRFAT.Add CRFATCFG_COD_CLIENTE_EMP, EMPRESA_TODA, "", 0, CStr(mvarlCliEmp)
    colConfigCRFAT.Add CRFATCFG_COD_FORNECEDOR_RECEITA, EMPRESA_TODA, "", 0, CStr(mvarlFornecedorReceita)
    colConfigCRFAT.Add CRFATCFG_TAXAFINANCEIRA_PADRAO, EMPRESA_TODA, "", 0, CStr(mvardTaxaFinanceira)
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_PRECO_BAIXO, EMPRESA_TODA, "", 0, CStr(mvariBloquearPVPrecoBaixo)
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_MARGEM_BAIXA, EMPRESA_TODA, "", 0, CStr(mvariBloquearPVMargemBaixa)
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_PRECO_DEFASADO, EMPRESA_TODA, "", 0, CStr(mvariBloquearPVPrecoDefasado)
    colConfigCRFAT.Add CRFATCFG_DIRETORIO_BOLETO_GERACAO, EMPRESA_TODA, "", 0, mvarsDirBoletoGer
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_USA_NFSE, giFilialEmpresa, "", 0, CStr(mvariUsaNFSE)
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_GRAVA_NFSE, giFilialEmpresa, "", 0, CStr(mvariGravaNFSE)
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_NF, EMPRESA_TODA, "", 0, CStr(mvariProdRepetidoNF)
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_PV, EMPRESA_TODA, "", 0, CStr(mvariProdRepetidoPV)
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_OV, EMPRESA_TODA, "", 0, CStr(mvariProdRepetidoOV)
    colConfigCRFAT.Add CRFATCFG_OV_GRAVACAO_VERSAO, EMPRESA_TODA, "", 0, CStr(mvariOVGravacaoVersao)
    colConfigCRFAT.Add CRFATCFG_VERSAO_LAYOUT_NFE, EMPRESA_TODA, "", 0, CStr(mvariVersaoNFe)
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_TOT_TRIB_TIPO, giFilialEmpresa, "", 0, CStr(mvariTotTribTipo)

    colConfigCRFAT.Add CRFATCFG_VALIDA_DESC_MAX_TAB_PRECO, EMPRESA_TODA, "", 0, CStr(mvariValidaDescMaxTabPreco)
    colConfigCRFAT.Add CRFATCFG_PRECO_UNITARIO_FORM_PRECO, EMPRESA_TODA, "", 0, CStr(mvariPrecoUnitarioFormPreco)

    colConfigCRFAT.Add CRFATCFG_PISCOFINS_SIMPLES_99, EMPRESA_TODA, "", 0, CStr(iPisCofinsSimples99)
    colConfigCRFAT.Add CRFATCFG_VENDA_FUTURA_TRIBUTACAO, EMPRESA_TODA, "", 0, CStr(iVendaFuturaTributacao)
    
    lErro = CF("CRFATConfig_Grava_Configs", colConfigCRFAT)
    If lErro <> SUCESSO Then Error 56554
    
    Gravar = SUCESSO
    
    Exit Function
    
Erro_Gravar:

    Gravar = Err
    
    Select Case Err
    
        Case 56554
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148814)
    
    End Select
        
    Exit Function
    
End Function

Function Gravar_Trans() As Long
'Grava no bd a configuração atual

Dim lErro As Long
Dim colConfigCRFAT As New ColCRFATConfig

On Error GoTo Erro_Gravar_Trans
    
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO1, EMPRESA_TODA, "", 0, CStr(mvariDescontoCodigo1)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO2, EMPRESA_TODA, "", 0, CStr(mvariDescontoCodigo2)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_CODIGO3, EMPRESA_TODA, "", 0, CStr(mvariDescontoCodigo3)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS1, EMPRESA_TODA, "", 0, CStr(mvariDescontoDias1)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS2, EMPRESA_TODA, "", 0, CStr(mvariDescontoDias2)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_DIAS3, EMPRESA_TODA, "", 0, CStr(mvariDescontoDias3)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL1, EMPRESA_TODA, "", 0, CStr(mvardDescontoPerc1)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL2, EMPRESA_TODA, "", 0, CStr(mvardDescontoPerc2)
    colConfigCRFAT.Add CRFATCFG_DESCONTO_PADRAO_PERCENTUAL3, EMPRESA_TODA, "", 0, CStr(mvardDescontoPerc3)
    colConfigCRFAT.Add CRFATCFG_JUROS_PADRAO, EMPRESA_TODA, "", 0, CStr(mvardPercJurosDiario)
    colConfigCRFAT.Add CRFATCFG_MULTA_PADRAO, EMPRESA_TODA, "", 0, CStr(mvardPercMulta)
    colConfigCRFAT.Add CRFATCFG_ISS_NA_BASE, EMPRESA_TODA, "", 0, CStr(mvariISSNaBase)
    colConfigCRFAT.Add CRFATCFG_IPI_ABATE_DESCONTO, EMPRESA_TODA, "", 0, CStr(mvariIPIAbateDesconto)
    colConfigCRFAT.Add CRFATCFG_CONDPAGTO_SEM_CODIGO, EMPRESA_TODA, "", 0, CStr(mvariCondPagtoSemCodigo)
    colConfigCRFAT.Add CRFATCFG_DIAS_BLOQUEIO_ATRASO, EMPRESA_TODA, "", 0, CStr(mvariDiasBloqueioAtraso) 'por leo em 25/02/02
    colConfigCRFAT.Add CRFATCFG_USA_COMISSOES_REGRAS, EMPRESA_TODA, "", 0, CStr(mvariUsaComissoesRegras) 'por leo em 15/03/02
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_USA_NFE, giFilialEmpresa, "", 0, CStr(mvariUsaNFe)
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_GRAVA_NFE, giFilialEmpresa, "", 0, CStr(mvariGravaNFe)
    colConfigCRFAT.Add CRFATCFG_CREDITO_VERIFICA_LIMITE, EMPRESA_TODA, "", 0, CStr(mvariCreditoVerificaLimite) 'tulio130103
    colConfigCRFAT.Add CRFATCFG_COD_FORNECEDOR_EMP, EMPRESA_TODA, "", 0, CStr(mvarlFornEmp)
    colConfigCRFAT.Add CRFATCFG_COD_CLIENTE_EMP, EMPRESA_TODA, "", 0, CStr(mvarlCliEmp)
    colConfigCRFAT.Add CRFATCFG_COD_FORNECEDOR_RECEITA, EMPRESA_TODA, "", 0, CStr(mvarlFornecedorReceita)
    colConfigCRFAT.Add CRFATCFG_TAXAFINANCEIRA_PADRAO, EMPRESA_TODA, "", 0, CStr(mvardTaxaFinanceira)
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_PRECO_BAIXO, EMPRESA_TODA, "", 0, CStr(mvariBloquearPVPrecoBaixo)
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_MARGEM_BAIXA, EMPRESA_TODA, "", 0, CStr(mvariBloquearPVMargemBaixa)
    colConfigCRFAT.Add CRFATCFG_BLOQUEARPV_PRECO_DEFASADO, EMPRESA_TODA, "", 0, CStr(mvariBloquearPVPrecoDefasado)
    colConfigCRFAT.Add CRFATCFG_DIRETORIO_BOLETO_GERACAO, EMPRESA_TODA, "", 0, mvarsDirBoletoGer
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_USA_NFSE, giFilialEmpresa, "", 0, CStr(mvariUsaNFSE)
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_GRAVA_NFSE, giFilialEmpresa, "", 0, CStr(mvariGravaNFSE)
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_NF, EMPRESA_TODA, "", 0, CStr(mvariProdRepetidoNF)
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_PV, EMPRESA_TODA, "", 0, CStr(mvariProdRepetidoPV)
    colConfigCRFAT.Add CRFATCFG_PROD_REPETIDO_OV, EMPRESA_TODA, "", 0, CStr(mvariProdRepetidoOV)
    colConfigCRFAT.Add CRFATCFG_OV_GRAVACAO_VERSAO, EMPRESA_TODA, "", 0, CStr(mvariOVGravacaoVersao)
    colConfigCRFAT.Add CRFATCFG_VERSAO_LAYOUT_NFE, EMPRESA_TODA, "", 0, CStr(mvariVersaoNFe)
    If giFilialEmpresa <> EMPRESA_TODA Then colConfigCRFAT.Add CRFATCFG_TOT_TRIB_TIPO, giFilialEmpresa, "", 0, CStr(mvariTotTribTipo)
    colConfigCRFAT.Add CRFATCFG_VALIDA_DESC_MAX_TAB_PRECO, EMPRESA_TODA, "", 0, CStr(mvariValidaDescMaxTabPreco)
    colConfigCRFAT.Add CRFATCFG_PRECO_UNITARIO_FORM_PRECO, EMPRESA_TODA, "", 0, CStr(mvariPrecoUnitarioFormPreco)
    
    colConfigCRFAT.Add CRFATCFG_PISCOFINS_SIMPLES_99, EMPRESA_TODA, "", 0, CStr(iPisCofinsSimples99)
    colConfigCRFAT.Add CRFATCFG_VENDA_FUTURA_TRIBUTACAO, EMPRESA_TODA, "", 0, CStr(iVendaFuturaTributacao)
    
    lErro = CF("CRFATConfig_Grava_Configs_Trans", colConfigCRFAT)
    If lErro <> SUCESSO Then Error 56555
    
    Gravar_Trans = SUCESSO
    
    Exit Function
    
Erro_Gravar_Trans:

    Gravar_Trans = Err
    
    Select Case Err
    
        Case 56555
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 148815)
    
    End Select
        
    Exit Function
    
End Function

Public Property Get iDiasBloqueioAtraso() As Integer 'por Leo em 26/02/02
    iDiasBloqueioAtraso = mvariDiasBloqueioAtraso
End Property

Public Property Let iDiasBloqueioAtraso(ByVal vData As Integer) 'por Leo em 26/02/02
    mvariDiasBloqueioAtraso = vData
End Property

Public Property Get iUsaComissoesRegras() As Integer 'por Leo em 15/03/02
    iUsaComissoesRegras = mvariUsaComissoesRegras
End Property

Public Property Let iUsaComissoesRegras(ByVal vData As Integer) 'por Leo em 13/03/02
    mvariUsaComissoesRegras = vData
End Property

Public Property Get iUsaNFe() As Integer
    iUsaNFe = mvariUsaNFe
End Property

Public Property Let iUsaNFe(ByVal vData As Integer)
    mvariUsaNFe = vData
End Property

Public Property Get iGravaNFe() As Integer
    iGravaNFe = mvariGravaNFe
End Property

Public Property Let iGravaNFe(ByVal vData As Integer)
    mvariGravaNFe = vData
End Property

Public Property Get iUsaNFSE() As Integer
    iUsaNFSE = mvariUsaNFSE
End Property

Public Property Let iUsaNFSE(ByVal vData As Integer)
    mvariUsaNFSE = vData
End Property

Public Property Get iGravaNFSE() As Integer
    iGravaNFSE = mvariGravaNFSE
End Property

Public Property Let iGravaNFSE(ByVal vData As Integer)
    mvariGravaNFSE = vData
End Property

Public Property Get iNaoTestarNFe() As Integer
    iNaoTestarNFe = mvariNaoTestarNFe
End Property

Public Property Let iNaoTestarNFe(ByVal vData As Integer)
    mvariNaoTestarNFe = vData
End Property

Public Property Get iNaoTestarNFePesoVolume() As Integer
    iNaoTestarNFePesoVolume = mvariNaoTestarNFePesoVolume
End Property

Public Property Let iNaoTestarNFePesoVolume(ByVal vData As Integer)
    mvariNaoTestarNFePesoVolume = vData
End Property

Public Property Get iBloquearPVMargemBaixa() As Integer
    iBloquearPVMargemBaixa = mvariBloquearPVMargemBaixa
End Property

Public Property Let iBloquearPVMargemBaixa(ByVal vData As Integer)
    mvariBloquearPVMargemBaixa = vData
End Property


