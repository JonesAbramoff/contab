VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolTiposMovEst As New ColTipoMovEst
Private mvariAceitaEstoqueNegativo As Integer
Private mvariAceitaKitProdGerencial As Integer
Private mvariOPDetalhamentoRCs As Integer
Private mvariEmpenhaEmbalagem As Integer
Private mvariTrataEstTercCliForn As Integer
Private mvariAceitaEANRepetido As Integer
Private mvariESTTestaEstoqueProduto As Integer
Private mvariESTTestaSaldos As Integer
Private mvariESTTestaEstoqueNegativo As Integer
Private mvariESTTestaIntAcao As Integer
Private mvarsESTTestaIntArqLog As String
Private mvariProdFatorAjuste As Integer
Private mvariMovEstVerificaBloqData As Integer
Private mvardtMovEstBloqDataLimite As Date
Private mvariOPTipoCalcNecessidade As Integer
Private mvariEstIniBloqAlterDataAnoPost As Integer
Private mvariEstIniBloqDataDifProd As Integer
Private mvariEstIniBloqDataDif As Integer
Private mvariEstIniBloqVlrZerado As Integer
Private mvariReqProdComPerda As Integer
Private mvariContabPelaDtEntXml As Integer

Public Property Let iContabPelaDtEntXml(ByVal vData As Integer)
    mvariContabPelaDtEntXml = vData
End Property

Public Property Get iContabPelaDtEntXml() As Integer
    iContabPelaDtEntXml = mvariContabPelaDtEntXml
End Property

Public Property Let iReqProdComPerda(ByVal vData As Integer)
    mvariReqProdComPerda = vData
End Property

Public Property Get iReqProdComPerda() As Integer
    iReqProdComPerda = mvariReqProdComPerda
End Property

Public Property Let iEstIniBloqVlrZerado(ByVal vData As Integer)
    mvariEstIniBloqVlrZerado = vData
End Property

Public Property Get iEstIniBloqVlrZerado() As Integer
    iEstIniBloqVlrZerado = mvariEstIniBloqVlrZerado
End Property

Public Property Let iEstIniBloqDataDif(ByVal vData As Integer)
    mvariEstIniBloqDataDif = vData
End Property

Public Property Get iEstIniBloqDataDif() As Integer
    iEstIniBloqDataDif = mvariEstIniBloqDataDif
End Property

Public Property Let iEstIniBloqDataDifProd(ByVal vData As Integer)
    mvariEstIniBloqDataDifProd = vData
End Property

Public Property Get iEstIniBloqDataDifProd() As Integer
    iEstIniBloqDataDifProd = mvariEstIniBloqDataDifProd
End Property

Public Property Let iEstIniBloqAlterDataAnoPost(ByVal vData As Integer)
    mvariEstIniBloqAlterDataAnoPost = vData
End Property

Public Property Get iEstIniBloqAlterDataAnoPost() As Integer
    iEstIniBloqAlterDataAnoPost = mvariEstIniBloqAlterDataAnoPost
End Property

Public Property Let iOPTipoCalcNecessidade(ByVal vData As Integer)
    mvariOPTipoCalcNecessidade = vData
End Property

Public Property Get iOPTipoCalcNecessidade() As Integer
    iOPTipoCalcNecessidade = mvariOPTipoCalcNecessidade
End Property

Public Property Let dtMovEstBloqDataLimite(ByVal vData As Date)
    mvardtMovEstBloqDataLimite = vData
End Property

Public Property Get dtMovEstBloqDataLimite() As Date
    dtMovEstBloqDataLimite = mvardtMovEstBloqDataLimite
End Property

Public Property Let iMovEstVerificaBloqData(ByVal vData As Integer)
    mvariMovEstVerificaBloqData = vData
End Property

Public Property Get iMovEstVerificaBloqData() As Integer
    iMovEstVerificaBloqData = mvariMovEstVerificaBloqData
End Property

Public Property Let iProdFatorAjuste(ByVal vData As Integer)
    mvariProdFatorAjuste = vData
End Property

Public Property Get iProdFatorAjuste() As Integer
    iProdFatorAjuste = mvariProdFatorAjuste
End Property

Public Property Let sESTTestaIntArqLog(ByVal vData As String)
    mvarsESTTestaIntArqLog = vData
End Property

Public Property Get sESTTestaIntArqLog() As String
    sESTTestaIntArqLog = mvarsESTTestaIntArqLog
End Property

Public Property Let iESTTestaIntAcao(ByVal vData As Integer)
    mvariESTTestaIntAcao = vData
End Property

Public Property Get iESTTestaIntAcao() As Integer
    iESTTestaIntAcao = mvariESTTestaIntAcao
End Property

Public Property Let iESTTestaEstoqueNegativo(ByVal vData As Integer)
    mvariESTTestaEstoqueNegativo = vData
End Property

Public Property Get iESTTestaEstoqueNegativo() As Integer
    iESTTestaEstoqueNegativo = mvariESTTestaEstoqueNegativo
End Property

Public Property Let iESTTestaSaldos(ByVal vData As Integer)
    mvariESTTestaSaldos = vData
End Property

Public Property Get iESTTestaSaldos() As Integer
    iESTTestaSaldos = mvariESTTestaSaldos
End Property

Public Property Let iESTTestaEstoqueProduto(ByVal vData As Integer)
    mvariESTTestaEstoqueProduto = vData
End Property

Public Property Get iESTTestaEstoqueProduto() As Integer
    iESTTestaEstoqueProduto = mvariESTTestaEstoqueProduto
End Property

Public Property Let iAceitaEANRepetido(ByVal vData As Integer)
    mvariAceitaEANRepetido = vData
End Property

Public Property Get iAceitaEANRepetido() As Integer
    iAceitaEANRepetido = mvariAceitaEANRepetido
End Property

Public Property Let iTrataEstTercCliForn(ByVal vData As Integer)
    mvariTrataEstTercCliForn = vData
End Property

Public Property Get iTrataEstTercCliForn() As Integer
    iTrataEstTercCliForn = mvariTrataEstTercCliForn
End Property

Public Property Let iEmpenhaEmbalagem(ByVal vData As Integer)
    mvariEmpenhaEmbalagem = vData
End Property

Public Property Get iEmpenhaEmbalagem() As Integer
    iEmpenhaEmbalagem = mvariEmpenhaEmbalagem
End Property

Public Property Let iOPDetalhamentoRCs(ByVal vData As Integer)
    mvariOPDetalhamentoRCs = vData
End Property

Public Property Get iOPDetalhamentoRCs() As Integer
    iOPDetalhamentoRCs = mvariOPDetalhamentoRCs
End Property

Public Property Get colTiposMovEst() As ColTipoMovEst
    Set colTiposMovEst = mvarcolTiposMovEst
End Property

Public Property Let iAceitaKitProdGerencial(ByVal vData As Integer)
    mvariAceitaKitProdGerencial = vData
End Property

Public Property Get iAceitaKitProdGerencial() As Integer
    
    iAceitaKitProdGerencial = mvariAceitaKitProdGerencial
    
End Property

Public Property Let iAceitaEstoqueNegativo(ByVal vData As Integer)
    mvariAceitaEstoqueNegativo = vData
End Property

Public Property Get iAceitaEstoqueNegativo() As Integer
    
    iAceitaEstoqueNegativo = mvariAceitaEstoqueNegativo
    
End Property

Public Function Inicializa() As Long

Dim lErro As Long, objMATConfig As New ClassMATConfig, colConfig As Object
Dim sConteudo As String

On Error GoTo Erro_Inicializa

    Set colConfig = CreateObject("GlobaisEST.ColESTConfig")
    
    colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, giFilialEmpresa, "", 0, "", ESTCFG_ACEITA_ESTOQUE_NEGATIVO
    colConfig.Add ESTCFG_ACEITA_KIT_PRODUTO_GERENCIAL, 0, "", 0, "", ESTCFG_ACEITA_KIT_PRODUTO_GERENCIAL
    colConfig.Add ESTCFG_OP_DETALHAMENTO_RCS, 0, "", 0, "", ESTCFG_OP_DETALHAMENTO_RCS
    colConfig.Add ESTCFG_OP_EMPENHA_EMBALAGEM, 0, "", 0, "", ESTCFG_OP_EMPENHA_EMBALAGEM
    colConfig.Add ESTCFG_TRATA_EST_TERC_CLI_FORN, EMPRESA_TODA, "", 0, "", ESTCFG_TRATA_EST_TERC_CLI_FORN
    colConfig.Add ESTCFG_ACEITA_EAN_REPETIDO, EMPRESA_TODA, "", 0, "", ESTCFG_ACEITA_EAN_REPETIDO
    colConfig.Add ESTCFG_TESTA_INT_ESTPRODUTO, EMPRESA_TODA, "", 0, "", ESTCFG_TESTA_INT_ESTPRODUTO
    colConfig.Add ESTCFG_TESTA_INT_ESTNEGATIVO, EMPRESA_TODA, "", 0, "", ESTCFG_TESTA_INT_ESTNEGATIVO
    colConfig.Add ESTCFG_TESTA_INT_SALDOS, EMPRESA_TODA, "", 0, "", ESTCFG_TESTA_INT_SALDOS
    colConfig.Add ESTCFG_TESTA_INT_ACAO, EMPRESA_TODA, "", 0, "", ESTCFG_TESTA_INT_ACAO
    colConfig.Add ESTCFG_TESTA_INT_DIR_ARQ_LOG, EMPRESA_TODA, "", 0, "", ESTCFG_TESTA_INT_DIR_ARQ_LOG
    colConfig.Add ESTCFG_PROD_FATOR_AJUSTE, EMPRESA_TODA, "", 0, "", ESTCFG_PROD_FATOR_AJUSTE
    colConfig.Add ESTCFG_MOVEST_VERIFICA_BLOQUEIO_DATA, EMPRESA_TODA, "", 0, "", ESTCFG_MOVEST_VERIFICA_BLOQUEIO_DATA
    colConfig.Add ESTCFG_MOVEST_BLOQUEIO_DATA_LIMITE, EMPRESA_TODA, "", 0, "", ESTCFG_MOVEST_BLOQUEIO_DATA_LIMITE
    colConfig.Add ESTCFG_OP_TIPO_CALC_NECESSIDADE, EMPRESA_TODA, "", 0, "", ESTCFG_OP_TIPO_CALC_NECESSIDADE
    
    colConfig.Add ESTCFG_BLOQUEIA_ALTER_DATA_ESTINI_ANOPOST, EMPRESA_TODA, "", 0, "", ESTCFG_BLOQUEIA_ALTER_DATA_ESTINI_ANOPOST
    colConfig.Add ESTCFG_BLOQUEIA_DATA_DIF_ESTINI_PROD, EMPRESA_TODA, "", 0, "", ESTCFG_BLOQUEIA_DATA_DIF_ESTINI_PROD
    colConfig.Add ESTCFG_BLOQUEIA_DATA_DIF_ESTINI, EMPRESA_TODA, "", 0, "", ESTCFG_BLOQUEIA_DATA_DIF_ESTINI
    colConfig.Add ESTCFG_BLOQUEIA_VALOR_ZERADO_ESTINI, EMPRESA_TODA, "", 0, "", ESTCFG_BLOQUEIA_VALOR_ZERADO_ESTINI
    colConfig.Add ESTCFG_REQPROD_CONSIDERA_PERCPERDA, EMPRESA_TODA, "", 0, "", ESTCFG_REQPROD_CONSIDERA_PERCPERDA
    colConfig.Add ESTCFG_CONTAB_PELA_DATA_ENTRADA_XML, EMPRESA_TODA, "", 0, "", ESTCFG_CONTAB_PELA_DATA_ENTRADA_XML

    'Lê as configurações em ESTConfig
    lErro = CF("ESTConfig_Le_Configs", colConfig)
    If lErro <> SUCESSO Then Error 33651
    
    mvariContabPelaDtEntXml = CInt(colConfig.Item(ESTCFG_CONTAB_PELA_DATA_ENTRADA_XML).sConteudo)
    
    mvariReqProdComPerda = CInt(colConfig.Item(ESTCFG_REQPROD_CONSIDERA_PERCPERDA).sConteudo)
    
    mvariEstIniBloqAlterDataAnoPost = StrParaInt(colConfig.Item(ESTCFG_BLOQUEIA_ALTER_DATA_ESTINI_ANOPOST).sConteudo)
    mvariEstIniBloqDataDifProd = StrParaInt(colConfig.Item(ESTCFG_BLOQUEIA_DATA_DIF_ESTINI_PROD).sConteudo)
    mvariEstIniBloqDataDif = StrParaInt(colConfig.Item(ESTCFG_BLOQUEIA_DATA_DIF_ESTINI).sConteudo)
    mvariEstIniBloqVlrZerado = StrParaInt(colConfig.Item(ESTCFG_BLOQUEIA_VALOR_ZERADO_ESTINI).sConteudo)
    
    
    mvariOPTipoCalcNecessidade = CInt(colConfig.Item(ESTCFG_OP_TIPO_CALC_NECESSIDADE).sConteudo)
    mvariAceitaEstoqueNegativo = CInt(colConfig.Item(ESTCFG_ACEITA_ESTOQUE_NEGATIVO).sConteudo)
    mvariOPDetalhamentoRCs = CInt(colConfig.Item(ESTCFG_OP_DETALHAMENTO_RCS).sConteudo)
    mvariAceitaKitProdGerencial = CInt(colConfig.Item(ESTCFG_ACEITA_KIT_PRODUTO_GERENCIAL).sConteudo)
    mvariEmpenhaEmbalagem = CInt(colConfig.Item(ESTCFG_OP_EMPENHA_EMBALAGEM).sConteudo)
    mvariTrataEstTercCliForn = CInt(colConfig.Item(ESTCFG_TRATA_EST_TERC_CLI_FORN).sConteudo)
    mvariAceitaEANRepetido = CInt(colConfig.Item(ESTCFG_ACEITA_EAN_REPETIDO).sConteudo)
    mvariESTTestaEstoqueProduto = CInt(colConfig.Item(ESTCFG_TESTA_INT_ESTPRODUTO).sConteudo)
    mvariESTTestaEstoqueNegativo = CInt(colConfig.Item(ESTCFG_TESTA_INT_ESTNEGATIVO).sConteudo)
    mvariESTTestaSaldos = CInt(colConfig.Item(ESTCFG_TESTA_INT_SALDOS).sConteudo)
    mvariESTTestaIntAcao = CInt(colConfig.Item(ESTCFG_TESTA_INT_ACAO).sConteudo)
    mvarsESTTestaIntArqLog = (colConfig.Item(ESTCFG_TESTA_INT_DIR_ARQ_LOG).sConteudo)
    mvariProdFatorAjuste = CInt(colConfig.Item(ESTCFG_PROD_FATOR_AJUSTE).sConteudo)
    mvariMovEstVerificaBloqData = CInt(colConfig.Item(ESTCFG_MOVEST_VERIFICA_BLOQUEIO_DATA).sConteudo)
    
    sConteudo = colConfig.Item(ESTCFG_MOVEST_BLOQUEIO_DATA_LIMITE).sConteudo
    If Len(Trim(sConteudo)) > 0 Then
        mvardtMovEstBloqDataLimite = StrParaDate(sConteudo)
    Else
        mvardtMovEstBloqDataLimite = DATA_NULA
    End If
    
    Call Init_Fest
        
    'preenche colecao de tiposmovest
    lErro = CF("TiposMovEstoque_Le", mvarcolTiposMovEst)
    If lErro <> SUCESSO Then Error 59267
    
    Inicializa = SUCESSO
     
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
     
    Select Case Err
          
        Case 59267
        
        Case 33651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTCONFIG", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 151557)
     
    End Select
     
    Exit Function

End Function
