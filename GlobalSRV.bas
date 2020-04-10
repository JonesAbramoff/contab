Attribute VB_Name = "GlobalSRV"
Option Explicit

Public Const STRING_OS_CODIGO = 9



Public Const STATUSOS_ABERTA = 0
Public Const STATUSOS_BAIXADA = 1

'StatusItemOS
Public Const STATUSITEMOS_ABERTA = 1
Public Const STATUSITEMOS_BAIXADA = 2

Public Const STRING_ITEMOS_OBS = 250

Public SRVGlob_Refs As Integer
Public SRVGlob_objSRV As Object

Public Const STRING_MO_NOME = 50
Public Const STRING_MO_NOMERED = 20
Public Const STRING_MO_OBS = 250


Public Const SRVCFG_AGLUTINA_LANCAM_POR_DIA = "AGLUTINA_LANCAM_POR_DIA"
Public Const SRVCFG_GERA_LOTE_AUTOMATICO = "GERA_LOTE_AUTOMATICO"
Public Const SRVCFG_VALIDA_GARANTIA = "VALIDA_GARANTIA"
Public Const SRVCFG_VALIDA_MANUTENCAO = "VALIDA_MANUTENCAO"
Public Const SRVCFG_GARANTIA_AUTOMATICA_SOLICITACAO = "GARANTIA_AUTOMATICA_SOLICITACAO"
Public Const SRVCFG_CONTRATO_AUTOMATICO_SOLICITACAO = "CONTRATO_AUTOMATICO_SOLICITACAO"
Public Const SRVCFG_VERIFICA_LOTE = "VERIFICA_LOTE"

Public Const SRING_TIPOGARANTIA_DESCRICAO = 250

'Número máximo de Solicitacoes de Serviço em uma tela de Solicitação
Public Const NUM_MAXIMO_SOLICITACOES = 200
Public Const STRING_SOLICITACAO = 250

Public Const NUM_MAXIMO_GARANTIA_SERVICOS = 100
Public Const NUM_MAXIMO_NUM_SERIE = 100

Public Const NUM_MAXIMO_PRODUTOSRV = 100

Public Const NUM_MAXIMO_PRODSOLICSRV = 100

'para tabela SRVConfig
Public Const STRING_SRVCONFIG_CODIGO = 50
Public Const STRING_SRVCONFIG_DESCRICAO = 150
Public Const STRING_SRVCONFIG_CONTEUDO = 255

Public Const VALIDA_GARANTIA = 1
Public Const NAO_VALIDA_GARANTIA = 0

Public Const VALIDA_MANUTENCAO = 1
Public Const NAO_VALIDA_MANUTENCAO = 0

Public Const GARANTIA_AUTOMATICA_SOLICITACAO = 1
Public Const NAO_GARANTIA_AUTOMATICA_SOLICITACAO = 0

Public Const CONTRATO_AUTOMATICO_SOLICITACAO = 1
Public Const NAO_CONTRATO_AUTOMATICO_SOLICITACAO = 0

Public Const VERIFICA_LOTE = 1
Public Const NAO_VERIFICA_LOTE = 0

Public Const GARANTIA_TOTAL = 1
Public Const GARANTIA_ATIVA = 1

Type typeSolicSRV
    iFilialEmpresa As Integer
    lCodigo As Long
    dtData As Date
    dHora As Double
    lCliente As Long
    iFilial As Integer
    lNumIntDoc As Long
    iVendedor As Integer
    iAtendente As Integer
    lClienteBenef As Long
    iFilialClienteBenef As Integer
    iPrazo As Integer
    iPrazoTipo As Integer
    dtDataEntrega As Date
    sObs As String
    lTipo As Long
    lFase As Long
End Type

Type typeItensSolicSRV
    lNumIntDoc As Long
    lNumIntSolicSRV As Long
    sProduto As String
    sProdutoDesc As String
    dtDataVenda As Date
    sServico As String
    sServicoDesc As String
    sUM As String
    dQuantidade As Double
    sLote As String
    iFilialOP As Integer
    sSolicitacao As String
    lGarantia As Long
    sManutencao As String
    sContrato As String
    iStatusItem As Integer
    sReparo As String
    dtDataBaixa As Date
End Type

Type typeSRVConfig
    sCodigo As String
    iFilialEmpresa As Integer
    sDescricao As String
    iTipo As Integer
    sConteudo As String
End Type

Type typeGarantia
    lNumIntDoc As Long
    lCodigo As Long
    iFilialEmpresa As Integer
    sProduto As String
    dtDataVenda As Date
    dQuantidade As Double
    sLote As String
    iFilialOP As Integer
    sSerie As String
    lNumNotaFiscal As Long
    lTipoGarantia As Long
    iGarantiaTotal As Integer
    iGarantiaTotalPrazo As Integer
    lFornecedor As Long
    iFilialFornecedor As Integer
    lCliFabr As Long
    iFilialCliFabr As Integer
End Type

Type typeGarantiaNumSerie
    lNumIntGarantia As Long
    sNumSerie As String
End Type

Type typeGarantiaProduto
    lNumIntGarantia As Long
    sProduto As String
    iPrazo As Integer
End Type

Type typeGarantiaContratoSRV
    lNumIntItensOrcSRV As Long
    lNumIntItensContratoSRV As Long
    lNumIntGarantia As Long
    dQuantidade As Double
    lGarantiaCod As Long
    sContratoCod As String
End Type

Type typeTipoGarantia
    lNumIntDoc As Long
    lCodigo As Long
    sDescricao As String
    iPrazoPadrao As Integer
    iGarantiaTotal As Integer
    iGarantiaTotalPrazo As Integer
    colTipoGarantiaProduto As New Collection
    objTela As Object
End Type

Type typeTipoGarantiaProduto
    lNumIntTipoGarantia As Long
    sProduto As String
    iPrazo As Integer
End Type

Type typeProdSolicSRV
    lNumIntDoc As Long
    lNumIntItensOrcSRV As Long
    dQuantidade As Double
    sProduto As String
    sServicoOrcSRV As String
    sLote As String
    iFilialOP As Integer
    lGarantia As Long
    sContrato As String
End Type

Type typeItemOS
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    sCodigo As String
    iItem As Integer
    sServico As String
    sCcl As String
    sSiglaUM As String
    dQuantidade As Double
    dtDataInicio As Date
    dtDataFim As Date
    iPrioridade As Integer
    iStatusItem As Integer
    iClasseUM As Integer
    sVersao As String
    sDescricao As String
    sObservacao As String
End Type

Type typeOS
    iFilialEmpresa As Integer
    sCodigo As String
    dtDataEmissao As Date
    iNumItens As Integer
    iNumItensBaixados As Integer
    lCodigoNumerico As Long
    iStatus As Integer
    lCodSolSRV As Long
    sProduto As String
    sLote As String
    iFilialOP As Integer
    lCodPedSRV As Long
    iFilialPedSRV As Integer
    lTipo As Long
End Type

Type typeItemOSOperacoes
    lNumIntDoc As Long
    lNumIntDocItemOS As Long
    iSeq As Integer
    lNumIntDocCompet As Long
    lNumIntDocCT As Long
    sObservacao As String
    iSeqPai As Integer
    iSeqArvore As Integer
    iNivel As Integer
    iSeqRoteiro As Integer
    iSeqRoteiroPai As Integer
    iNivelRoteiro As Integer
    sServico As String
    sVersao As String
    iIgnoraTaxaProducao As Integer
    iConsideraCarga As Integer
    iOrigem As Integer
    lNumIntDocOPerOrigem As Long
    iNumMaxMaqPorOper As Integer
    iNumRepeticoes As Integer
End Type

Type typeItemOSOperacoesPecas
    lNumIntDocOper As Long
    sProduto As String
    dQuantidade As Double
    sUMProduto As String
    iComposicao As Integer
    dPercentualPerda As Double
    dCustoStandard As Double
    sVersaoKitComp As String
End Type

Type typeItemOSOperacoesMaquinas
    lNumIntDocOper As Long
    lNumIntDocMaq As Long
    dHoras As Double
    iQuantidade As Integer
End Type

Type typeMO
    iAtivo As Integer
    lCodigo As Long
    sNome As String
    sNomeReduzido As String
    sObservacao As String
    iTipo As Integer
End Type

Type typeItemOSOperacoesMO
    lNumIntDocOper As Long
    lCodigoMO As Long
    dHoras As Double
End Type


Type typeItensDeContratoSrv
    lNumIntDoc As Long
    lCodigo As Long
    iFilialEmpresa As Integer
    lNumIntItemContrato As Long
    dQuantidade As Double
    sLote As String
    iFilialOP As Integer
    iGarantiaTotal As Integer
    lTipoGarantia As Long
    dtDataContratoIni As Date
    dtDataContratoFim As Date
    sProduto As String
    sServico As String
    colNumSerie As New Collection
    colProduto As New Collection
    sCodigoContrato As String
End Type


Type typeItensDeContratoSrvNumSerie
    lNumIntItemContratoSrv As Long
    sNumSerie As String
End Type

Type typeItensDeContratoSrvProd
    lNumIntItemContratoSrv As Long
    sProduto As String
End Type

Type typeProdutoSRV
    lNumIntProdSolicSRV As Long
    sProduto As String
    dQuantidade As Double
    lGarantia As Long
    sContrato As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRoteiroSRV
    lNumIntDoc As Long
    sServico As String
    sVersao As String
    sDescricao As String
    dtDataCriacao As Date
    dtDataUltModificacao As Date
    dQuantidade As Double
    sUM As String
    sAutor As String
    iDuracao As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRoteiroSRVOper
    lNumIntDoc As Long
    lNumIntDocRotSRV As Long
    iSeq As Integer
    lNumIntDocCompet As Long
    lNumIntDocCT As Long
    sObservacao As String
    iSeqPai As Integer
    iSeqArvore As Integer
    iNivel As Integer
    iPosicaoArvore As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRoteiroSRVOperMP
    lNumIntDoc As Long
    lNumIntDocOper As Long
    iSeq As Integer
    sProduto As String
    dQuantidade As Double
    sUM As String
    sVersao As String
    iComposicao As Integer
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRoteiroSRVOperMO
    lNumIntDoc As Long
    lNumIntDocOper As Long
    iSeq As Integer
    iCodMO As Integer
    dHoras As Double
    iQtd As Integer
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRoteiroSRVOperMaq
    lNumIntDoc As Long
    lNumIntDocOper As Long
    iSeq As Integer
    iCodMaq As Integer
    iFilialEmpMaq As Integer
    dHoras As Double
    iQtd As Integer
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeOSApMaq
    lCodigoAp As Long
    iFilialEmpresa As Integer
    iSeq As Integer
    lNumIntDocMaq As Long
    dHorasGastas As Double
    iQuantidade As Integer
    sOS As String
    sProdutoOS As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeOSApMO
    lCodigoAp As Long
    iFilialEmpresa As Integer
    iSeq As Integer
    lCodigoMO As Long
    dHorasGastas As Double
    sOS As String
    sProdutoOS As String
End Type

Type typeOSAp
    iFilialEmpresa As Integer
    lCodigo As Long
    lCodigoMovEst As Long
    dtData As Date
    lCliente As Long
    dtHora As Date
    lNumIntDoc As Long
End Type
