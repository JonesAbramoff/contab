VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariAglutinaLancamPorDia As Integer
Private mvariGeraLoteAutomatico As Integer
Private mvariIncluiFreteNoCusto As Integer
Private mvariOPRepeticaoItem As Integer
Private mvariPrioridadeMaquina As Integer
Private mvariPrioridadeProduto As Integer
Private mvariPrioridadeProdutoAnc As Integer
Private mvariClasseUMTempo As Integer
Private mvariTemRepeticoesOper As Integer
Private mvardtDataInicioMRP As Date
Private mvariGeraReqCompraEmLote As Integer
Private mvariUsaPerdaInsumosKit As Integer
Private mvariFormaCalcPrecoEntCom As Integer
Private mvariInventarioCodBarrAuto As Integer
Private mvariValidaProdBase As Integer
Private mvariProdEntVerificaBxOP As Integer
Private mvariTrazPrecoUltCompr As Integer
Private mvariTrataEstTercCliForn As Integer
Private mvariRelPontoPedExibePPZerado As Integer
Private mvariExibeMOCursos As Integer
Private mvariSerieEletronicaPadrao As Integer
Private mvariProdExibeComissao As Integer
Private mvariTratamentoAberturaMesIni As Integer
Private mvariAltDataKitNaGravacao As Integer

Public Property Let iAltDataKitNaGravacao(ByVal vData As Integer)
    mvariAltDataKitNaGravacao = vData
End Property

Public Property Get iAltDataKitNaGravacao() As Integer
    iAltDataKitNaGravacao = mvariAltDataKitNaGravacao
End Property

Public Property Let iTratamentoAberturaMesIni(ByVal vData As Integer)
    mvariTratamentoAberturaMesIni = vData
End Property

Public Property Get iTratamentoAberturaMesIni() As Integer
    iTratamentoAberturaMesIni = mvariTratamentoAberturaMesIni
End Property

Public Property Let iProdExibeComissao(ByVal vData As Integer)
    mvariProdExibeComissao = vData
End Property

Public Property Get iProdExibeComissao() As Integer
    iProdExibeComissao = mvariProdExibeComissao
End Property

Public Property Let iSerieEletronicaPadrao(ByVal vData As Integer)
    mvariSerieEletronicaPadrao = vData
End Property

Public Property Get iSerieEletronicaPadrao() As Integer
    iSerieEletronicaPadrao = mvariSerieEletronicaPadrao
End Property

Public Property Let iExibeMOCursos(ByVal vData As Integer)
    mvariExibeMOCursos = vData
End Property

Public Property Get iExibeMOCursos() As Integer
    iExibeMOCursos = mvariExibeMOCursos
End Property

Public Property Let iRelPontoPedExibePPZerado(ByVal vData As Integer)
    mvariRelPontoPedExibePPZerado = vData
End Property

Public Property Get iRelPontoPedExibePPZerado() As Integer
    iRelPontoPedExibePPZerado = mvariRelPontoPedExibePPZerado
End Property

Public Property Let iTrataEstTercCliForn(ByVal vData As Integer)
    mvariTrataEstTercCliForn = vData
End Property

Public Property Get iTrataEstTercCliForn() As Integer
    iTrataEstTercCliForn = mvariTrataEstTercCliForn
End Property

Public Property Let iTrazPrecoUltCompr(ByVal vData As Integer)
    mvariTrazPrecoUltCompr = vData
End Property

Public Property Get iTrazPrecoUltCompr() As Integer
    iTrazPrecoUltCompr = mvariTrazPrecoUltCompr
End Property

Public Property Let iProdEntVerificaBxOP(ByVal vData As Integer)
    mvariProdEntVerificaBxOP = vData
End Property

Public Property Get iProdEntVerificaBxOP() As Integer
    iProdEntVerificaBxOP = mvariProdEntVerificaBxOP
End Property

Public Property Let iValidaProdBase(ByVal vData As Integer)
    mvariValidaProdBase = vData
End Property

Public Property Get iValidaProdBase() As Integer
    iValidaProdBase = mvariValidaProdBase
End Property

Public Property Let iInventarioCodBarrAuto(ByVal vData As Integer)
    mvariInventarioCodBarrAuto = vData
End Property

Public Property Get iInventarioCodBarrAuto() As Integer
    iInventarioCodBarrAuto = mvariInventarioCodBarrAuto
End Property

Public Property Let iFormaCalcPrecoEntCom(ByVal vData As Integer)
    mvariFormaCalcPrecoEntCom = vData
End Property

Public Property Get iFormaCalcPrecoEntCom() As Integer
    iFormaCalcPrecoEntCom = mvariFormaCalcPrecoEntCom
End Property

Public Property Let iUsaPerdaInsumosKit(ByVal vData As Integer)
    mvariUsaPerdaInsumosKit = vData
End Property

Public Property Get iUsaPerdaInsumosKit() As Integer
    iUsaPerdaInsumosKit = mvariUsaPerdaInsumosKit
End Property

Public Property Let iGeraReqCompraEmLote(ByVal vData As Integer)
    mvariGeraReqCompraEmLote = vData
End Property

Public Property Get iGeraReqCompraEmLote() As Integer
    iGeraReqCompraEmLote = mvariGeraReqCompraEmLote
End Property

Public Property Let iPrioridadeMaquina(ByVal vData As Integer)
    mvariPrioridadeMaquina = vData
End Property

Public Property Get iPrioridadeMaquina() As Integer
    iPrioridadeMaquina = mvariPrioridadeMaquina
End Property

Public Property Let iPrioridadeProduto(ByVal vData As Integer)
    mvariPrioridadeProduto = vData
End Property

Public Property Get iPrioridadeProduto() As Integer
    iPrioridadeProduto = mvariPrioridadeProduto
End Property

Public Property Let iPrioridadeProdutoAnc(ByVal vData As Integer)
    mvariPrioridadeProdutoAnc = vData
End Property

Public Property Get iPrioridadeProdutoAnc() As Integer
    iPrioridadeProdutoAnc = mvariPrioridadeProdutoAnc
End Property

Public Property Let iClasseUMTempo(ByVal vData As Integer)
    mvariClasseUMTempo = vData
End Property

Public Property Get iClasseUMTempo() As Integer
    iClasseUMTempo = mvariClasseUMTempo
End Property

Public Property Let iTemRepeticoesOper(ByVal vData As Integer)
    mvariTemRepeticoesOper = vData
End Property

Public Property Get iTemRepeticoesOper() As Integer
    iTemRepeticoesOper = mvariTemRepeticoesOper
End Property

Public Property Let dtDataInicioMRP(ByVal vData As Date)
    mvardtDataInicioMRP = vData
End Property

Public Property Get dtDataInicioMRP() As Date
    dtDataInicioMRP = mvardtDataInicioMRP
End Property
'---------------------------------------

Public Function Inicializa() As Long

Dim lErro As Long, colConfig As New ColESTConfig

On Error GoTo Erro_Inicializa

    colConfig.Add ESTCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, "", ESTCFG_AGLUTINA_LANCAM_POR_DIA
    colConfig.Add ESTCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, "", ESTCFG_GERA_LOTE_AUTOMATICO
    colConfig.Add ESTCFG_INCLUI_FRETE_CUSTO, EMPRESA_TODA, "", 0, "", ESTCFG_INCLUI_FRETE_CUSTO
    colConfig.Add ESTCFG_OP_VALIDA_REPETICAO_ITEM, EMPRESA_TODA, "", 0, "", ESTCFG_OP_VALIDA_REPETICAO_ITEM
    colConfig.Add ESTCFG_VALIDA_PRODUTO_BASE_CARGA, EMPRESA_TODA, "", 0, "", ESTCFG_VALIDA_PRODUTO_BASE_CARGA
    colConfig.Add ESTCFG_PRIORIDADE_MAQUINA, EMPRESA_TODA, "", 0, "", ESTCFG_PRIORIDADE_MAQUINA
    colConfig.Add ESTCFG_PRIORIDADE_PRODUTO, EMPRESA_TODA, "", 0, "", ESTCFG_PRIORIDADE_PRODUTO
    colConfig.Add ESTCFG_PRIORIDADE_PRODUTO_ANCESTRAL, EMPRESA_TODA, "", 0, "", ESTCFG_PRIORIDADE_PRODUTO_ANCESTRAL
    colConfig.Add ESTCFG_CLASSE_UM_TEMPO, EMPRESA_TODA, "", 0, "", ESTCFG_CLASSE_UM_TEMPO
    colConfig.Add ESTCFG_DATA_INICIO_MRP, EMPRESA_TODA, "", 0, "", ESTCFG_DATA_INICIO_MRP
    colConfig.Add ESTCFG_GERA_REQCOMPRA_EM_LOTE, EMPRESA_TODA, "", 0, "", ESTCFG_GERA_REQCOMPRA_EM_LOTE
    colConfig.Add ESTCFG_USA_PERDA_INSUMOS_KIT, EMPRESA_TODA, "", 0, "", ESTCFG_USA_PERDA_INSUMOS_KIT
    colConfig.Add ESTCFG_FORMA_CALC_PRECO_ENT_COM, EMPRESA_TODA, "", 0, "", ESTCFG_FORMA_CALC_PRECO_ENT_COM
    colConfig.Add ESTCFG_TEM_REPETICAO_OPERACAO, EMPRESA_TODA, "", 0, "", ESTCFG_TEM_REPETICAO_OPERACAO
    colConfig.Add ESTCFG_INVENTARIOCODBARRAUTO, EMPRESA_TODA, "", 0, "", ESTCFG_INVENTARIOCODBARRAUTO
    colConfig.Add ESTCFG_PROD_ENT_VERIFICA_BX_OP, EMPRESA_TODA, "", 0, "", ESTCFG_PROD_ENT_VERIFICA_BX_OP
    colConfig.Add ESTCFG_TRAZ_PRECO_ULT_COMPRA, EMPRESA_TODA, "", 0, "", ESTCFG_TRAZ_PRECO_ULT_COMPRA
    colConfig.Add ESTCFG_REL_PONTOPED_EXIBE_PPZERADO, EMPRESA_TODA, "", 0, "", ESTCFG_REL_PONTOPED_EXIBE_PPZERADO
    colConfig.Add ESTCFG_TIPOMAODEOBRA_EXIBE_ABA_CURSOS, EMPRESA_TODA, "", 0, "", ESTCFG_TIPOMAODEOBRA_EXIBE_ABA_CURSOS
    colConfig.Add ESTCFG_SERIE_ELETRONICA_PADRAO, EMPRESA_TODA, "", 0, "", ESTCFG_SERIE_ELETRONICA_PADRAO
    colConfig.Add ESTCFG_PRODUTO_EXIBE_COMISSAO, EMPRESA_TODA, "", 0, "", ESTCFG_PRODUTO_EXIBE_COMISSAO
    colConfig.Add ESTCFG_TRATAMENTO_ABERTURA_MES_INI, EMPRESA_TODA, "", 0, "", ESTCFG_TRATAMENTO_ABERTURA_MES_INI
    colConfig.Add ESTCFG_ALTERA_DATA_KIT_NA_GRAVACAO, EMPRESA_TODA, "", 0, "", ESTCFG_ALTERA_DATA_KIT_NA_GRAVACAO
    
    'Lê as configurações em ESTConfig
    lErro = CF("ESTConfig_Le_Configs", colConfig)
    If lErro <> SUCESSO Then Error 33651
    
    mvariAltDataKitNaGravacao = CInt(colConfig.Item(ESTCFG_ALTERA_DATA_KIT_NA_GRAVACAO).sConteudo)
    mvariTratamentoAberturaMesIni = CInt(colConfig.Item(ESTCFG_TRATAMENTO_ABERTURA_MES_INI).sConteudo)
    mvariProdExibeComissao = CInt(colConfig.Item(ESTCFG_PRODUTO_EXIBE_COMISSAO).sConteudo)
    mvariSerieEletronicaPadrao = CInt(colConfig.Item(ESTCFG_SERIE_ELETRONICA_PADRAO).sConteudo)
    mvariAglutinaLancamPorDia = CInt(colConfig.Item(ESTCFG_AGLUTINA_LANCAM_POR_DIA).sConteudo)
    mvariGeraLoteAutomatico = CInt(colConfig.Item(ESTCFG_GERA_LOTE_AUTOMATICO).sConteudo)
    mvariIncluiFreteNoCusto = CInt(colConfig.Item(ESTCFG_INCLUI_FRETE_CUSTO).sConteudo)
    mvariOPRepeticaoItem = CInt(colConfig.Item(ESTCFG_OP_VALIDA_REPETICAO_ITEM).sConteudo)
    mvariValidaProdBase = CInt(colConfig.Item(ESTCFG_VALIDA_PRODUTO_BASE_CARGA).sConteudo)
    mvariPrioridadeMaquina = CInt(colConfig.Item(ESTCFG_PRIORIDADE_MAQUINA).sConteudo)
    mvariPrioridadeProduto = CInt(colConfig.Item(ESTCFG_PRIORIDADE_PRODUTO).sConteudo)
    mvariPrioridadeProdutoAnc = CInt(colConfig.Item(ESTCFG_PRIORIDADE_PRODUTO_ANCESTRAL).sConteudo)
    mvariClasseUMTempo = CInt(colConfig.Item(ESTCFG_CLASSE_UM_TEMPO).sConteudo)
    mvardtDataInicioMRP = StrParaDate(colConfig.Item(ESTCFG_DATA_INICIO_MRP).sConteudo)
    mvariGeraReqCompraEmLote = CInt(colConfig.Item(ESTCFG_GERA_REQCOMPRA_EM_LOTE).sConteudo)
    mvariUsaPerdaInsumosKit = CInt(colConfig.Item(ESTCFG_USA_PERDA_INSUMOS_KIT).sConteudo)
    mvariFormaCalcPrecoEntCom = CInt(colConfig.Item(ESTCFG_FORMA_CALC_PRECO_ENT_COM).sConteudo)
    mvariTemRepeticoesOper = CInt(colConfig.Item(ESTCFG_TEM_REPETICAO_OPERACAO).sConteudo)
    mvariInventarioCodBarrAuto = CInt(colConfig.Item(ESTCFG_INVENTARIOCODBARRAUTO).sConteudo)
    mvariProdEntVerificaBxOP = CInt(colConfig.Item(ESTCFG_PROD_ENT_VERIFICA_BX_OP).sConteudo)
    mvariTrazPrecoUltCompr = CInt(colConfig.Item(ESTCFG_TRAZ_PRECO_ULT_COMPRA).sConteudo)
    mvariRelPontoPedExibePPZerado = CInt(colConfig.Item(ESTCFG_REL_PONTOPED_EXIBE_PPZERADO).sConteudo)
    mvariExibeMOCursos = CInt(colConfig.Item(ESTCFG_TIPOMAODEOBRA_EXIBE_ABA_CURSOS).sConteudo)
    
    Inicializa = SUCESSO
    
    Exit Function
    
Erro_Inicializa:

    Inicializa = Err
    
    Select Case Err
    
        Case 33651
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ESTCONFIG", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 149454)
        
    End Select

    Exit Function

End Function

Public Property Get iAglutinaLancamPorDia() As Integer
    iAglutinaLancamPorDia = mvariAglutinaLancamPorDia
End Property

Public Property Get iGeraLoteAutomatico() As Integer
    iGeraLoteAutomatico = mvariGeraLoteAutomatico
End Property

Public Property Let iAglutinaLancamPorDia(ByVal vData As Integer)
    mvariAglutinaLancamPorDia = vData
End Property

Public Property Let iGeraLoteAutomatico(ByVal vData As Integer)
    mvariGeraLoteAutomatico = vData
End Property

Public Property Let iIncluiFreteNoCusto(ByVal vData As Integer)
    mvariIncluiFreteNoCusto = vData
End Property

Public Property Get iIncluiFreteNoCusto() As Integer
    iIncluiFreteNoCusto = mvariIncluiFreteNoCusto
End Property

Public Property Let iOPRepeticaoItem(ByVal vData As Integer)
    mvariOPRepeticaoItem = vData
End Property

Public Property Get iOPRepeticaoItem() As Integer
    iOPRepeticaoItem = mvariOPRepeticaoItem
End Property

Function Gravar() As Long
'Grava no bd a configuração atual

Dim lErro As Long, colConfig As New ColESTConfig

On Error GoTo Erro_Gravar
    
    colConfig.Add ESTCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add ESTCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfig.Add ESTCFG_INCLUI_FRETE_CUSTO, EMPRESA_TODA, "", 0, CStr(mvariIncluiFreteNoCusto)
    colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, giFilialEmpresa, "", 0, CStr(gobjMAT.iAceitaEstoqueNegativo)
    colConfig.Add ESTCFG_OP_VALIDA_REPETICAO_ITEM, EMPRESA_TODA, "", 0, CStr(mvariOPRepeticaoItem)
    colConfig.Add ESTCFG_VALIDA_PRODUTO_BASE_CARGA, EMPRESA_TODA, "", 0, CStr(mvariValidaProdBase)
    colConfig.Add ESTCFG_PRIORIDADE_MAQUINA, EMPRESA_TODA, "", 0, CStr(mvariPrioridadeMaquina)
    colConfig.Add ESTCFG_PRIORIDADE_PRODUTO, EMPRESA_TODA, "", 0, CStr(mvariPrioridadeProduto)
    colConfig.Add ESTCFG_PRIORIDADE_PRODUTO_ANCESTRAL, EMPRESA_TODA, "", 0, CStr(mvariPrioridadeProdutoAnc)
    colConfig.Add ESTCFG_CLASSE_UM_TEMPO, EMPRESA_TODA, "", 0, CStr(mvariClasseUMTempo)
    colConfig.Add ESTCFG_DATA_INICIO_MRP, EMPRESA_TODA, "", 0, CStr(mvardtDataInicioMRP)
    colConfig.Add ESTCFG_GERA_REQCOMPRA_EM_LOTE, EMPRESA_TODA, "", 0, CStr(mvariGeraReqCompraEmLote)
    colConfig.Add ESTCFG_USA_PERDA_INSUMOS_KIT, EMPRESA_TODA, "", 0, CStr(mvariUsaPerdaInsumosKit)
    colConfig.Add ESTCFG_FORMA_CALC_PRECO_ENT_COM, EMPRESA_TODA, "", 0, CStr(mvariFormaCalcPrecoEntCom)
    colConfig.Add ESTCFG_TEM_REPETICAO_OPERACAO, EMPRESA_TODA, "", 0, CStr(mvariTemRepeticoesOper)
    colConfig.Add ESTCFG_INVENTARIOCODBARRAUTO, EMPRESA_TODA, "", 0, CStr(mvariInventarioCodBarrAuto)
    colConfig.Add ESTCFG_MOVEST_VERIFICA_BLOQUEIO_DATA, EMPRESA_TODA, "", 0, CStr(gobjMAT.iMovEstVerificaBloqData)
    colConfig.Add ESTCFG_MOVEST_BLOQUEIO_DATA_LIMITE, EMPRESA_TODA, "", 0, CStr(gobjMAT.dtMovEstBloqDataLimite)

    'Chama ESTConfig_Grava_Configs
    lErro = CF("ESTConfig_Grava_Configs", colConfig)
    If lErro <> SUCESSO Then Error 33652
    
    Gravar = SUCESSO
    
    Exit Function
    
Erro_Gravar:

    Gravar = Err
    
    Select Case Err
    
        Case 33652
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149455)
    
    End Select
        
    Exit Function
    
End Function

Function Gravar_Trans() As Long
'Grava no bd a configuração atual
'TEM QUE SER CHAMADA EM TRANSACAO

Dim lErro As Long, colConfig As New ColESTConfig

On Error GoTo Erro_Gravar_Trans
    
    colConfig.Add ESTCFG_AGLUTINA_LANCAM_POR_DIA, EMPRESA_TODA, "", 0, CStr(mvariAglutinaLancamPorDia)
    colConfig.Add ESTCFG_GERA_LOTE_AUTOMATICO, EMPRESA_TODA, "", 0, CStr(mvariGeraLoteAutomatico)
    colConfig.Add ESTCFG_INCLUI_FRETE_CUSTO, EMPRESA_TODA, "", 0, CStr(mvariIncluiFreteNoCusto)
    colConfig.Add ESTCFG_ACEITA_ESTOQUE_NEGATIVO, giFilialEmpresa, "", 0, CStr(gobjMAT.iAceitaEstoqueNegativo)
    colConfig.Add ESTCFG_OP_VALIDA_REPETICAO_ITEM, EMPRESA_TODA, "", 0, CStr(mvariOPRepeticaoItem)
    colConfig.Add ESTCFG_VALIDA_PRODUTO_BASE_CARGA, EMPRESA_TODA, "", 0, CStr(mvariValidaProdBase)
    colConfig.Add ESTCFG_PRIORIDADE_MAQUINA, EMPRESA_TODA, "", 0, CStr(mvariPrioridadeMaquina)
    colConfig.Add ESTCFG_PRIORIDADE_PRODUTO, EMPRESA_TODA, "", 0, CStr(mvariPrioridadeProduto)
    colConfig.Add ESTCFG_PRIORIDADE_PRODUTO_ANCESTRAL, EMPRESA_TODA, "", 0, CStr(mvariPrioridadeProdutoAnc)
    colConfig.Add ESTCFG_CLASSE_UM_TEMPO, EMPRESA_TODA, "", 0, CStr(mvariClasseUMTempo)
    colConfig.Add ESTCFG_DATA_INICIO_MRP, EMPRESA_TODA, "", 0, CStr(mvardtDataInicioMRP)
    colConfig.Add ESTCFG_GERA_REQCOMPRA_EM_LOTE, EMPRESA_TODA, "", 0, CStr(mvariGeraReqCompraEmLote)
    colConfig.Add ESTCFG_USA_PERDA_INSUMOS_KIT, EMPRESA_TODA, "", 0, CStr(mvariUsaPerdaInsumosKit)
    colConfig.Add ESTCFG_FORMA_CALC_PRECO_ENT_COM, EMPRESA_TODA, "", 0, CStr(mvariFormaCalcPrecoEntCom)
    colConfig.Add ESTCFG_TEM_REPETICAO_OPERACAO, EMPRESA_TODA, "", 0, CStr(mvariTemRepeticoesOper)
    colConfig.Add ESTCFG_INVENTARIOCODBARRAUTO, EMPRESA_TODA, "", 0, CStr(mvariInventarioCodBarrAuto)
    colConfig.Add ESTCFG_MOVEST_VERIFICA_BLOQUEIO_DATA, EMPRESA_TODA, "", 0, CStr(gobjMAT.iMovEstVerificaBloqData)
    colConfig.Add ESTCFG_MOVEST_BLOQUEIO_DATA_LIMITE, EMPRESA_TODA, "", 0, CStr(gobjMAT.dtMovEstBloqDataLimite)
    
    lErro = CF("ESTConfig_Grava_Configs_Trans", colConfig)
    If lErro <> SUCESSO Then Error 41817
    
    Gravar_Trans = SUCESSO
    
    Exit Function
    
Erro_Gravar_Trans:

    Gravar_Trans = Err
    
    Select Case Err
    
        Case 41817
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 149456)
    
    End Select
        
    Exit Function
    
End Function
