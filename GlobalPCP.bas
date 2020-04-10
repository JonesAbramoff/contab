Attribute VB_Name = "GlobalPCP"
Option Explicit

Public Const TELAGRAFICO_FUNCAO_INICIALIZA_GRIDAUX = "Inicializa_GridAux"
Public Const TELAGRAFICO_FUNCAO_PREENCHE_GRIDAUX = "Preenche_GridAux"

'*** Constantes ***
Public Const ZOOM_100 = 1
Public Const ZOOM_50 = 2
Public Const ZOOM_25 = 3

Public Const NUM_MAX_TENTATIVAS_PROD_DATA_INI = 3
Public Const PRODUCAO_PV_DELAY = 1 'Em dias

Public Const ZOOM_100_PORCENT = 1
Public Const ZOOM_50_PORCENT = 2
Public Const ZOOM_25_PORCENT = 4

Public Const TELAGRAFICOIMPITENS_TIPO_TEXT = 1
Public Const TELAGRAFICOIMPITENS_TIPO_LINE = 2
Public Const TELAGRAFICOIMPITENS_TIPO_SETA = 3
Public Const TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_COLUNA = 4
Public Const TELAGRAFICOIMPITENS_TIPO_TEXT_FIXO_LINHA = 5

Public Const STRING_MAQUINA_DESCRICAO = 50
Public Const STRING_MAQUINA_NOMEREDUZIDO = 20
Public Const NUM_MAX_MAQUINAS = 9999

Public Const TELA_GRAFICO_ICONE_INICIO = 1
Public Const TELA_GRAFICO_ICONE_FIM = 2
Public Const TELA_GRAFICO_ICONE_INICIO_E_FIM = 3

'constantes para carregamento da Combo Recurso em Máquinas e Centros de Trabalho
Public Const ITEMCT_RECURSO_MAQUINA = 1
Public Const STRING_ITEMCT_RECURSO_MAQUINA = "Máquina"
Public Const ITEMCT_RECURSO_HABILIDADE = 2
Public Const STRING_ITEMCT_RECURSO_HABILIDADE = "Habilidade"
Public Const ITEMCT_RECURSO_PROCESSO = 3
Public Const STRING_ITEMCT_RECURSO_PROCESSO = "Processo"

'constantes para carregamento da Combo Tipo em Taxas de Produção
Public Const ITEM_TIPO_TAXAPRODUCAO_VARIAVEL = 1
Public Const STRING_ITEM_TIPO_TAXAPRODUCAO_VARIAVEL = "Variável"
Public Const ITEM_TIPO_TAXAPRODUCAO_FIXO = 2
Public Const STRING_ITEM_TIPO_TAXAPRODUCAO_FIXO = "Fixo"

'constantes para tratar Taxa de Consumo em Maquinas e Taxas de Produção
Public Const STRING_BARRA_SEPARADORA = "/"
Public Const TAXA_CONSUMO_TEMPO_PADRAO = "h"

Public Const TAXA_ATIVA = 1
Public Const TAXA_INATIVA = 0

'constantes para o Relatório da Árvore de Produtos
Public Const NUMERONIVEIS_RELATORIO_RETRATO = 8
Public Const CARACTER_DA_STRING_DE_IDENTACAO = "."
Public Const NUMERO_DE_CARACTERES_POR_IDENTACAO = 3

'constantes para Roteiros de Fabricação
Public Const STRING_AUTOR_ROT_FAB = 10
Public Const NIVEL_MAXIMO_OPERACOES = 20

'constantes para Competências
Public Const STRING_COMPETENCIA_NOMERED = 20
Public Const STRING_COMPETENCIA_DESCRICAO = 50
Public Const NUM_MAX_COMPETENCIAS = 99999

'constantes para Centro de Trabalho
Public Const STRING_CENTRODETRABALHO_NOMERED = 20
Public Const STRING_CENTRODETRABALHO_DESCRICAO = 50
Public Const NUM_MAX_CENTROSDETRABALHOS = 99999
Public Const DOMINGO = 1
Public Const SEGUNDA = 2
Public Const TERCA = 3
Public Const QUARTA = 4
Public Const QUINTA = 5
Public Const SEXTA = 6
Public Const SABADO = 7
Public Const HORAS_DO_DIA = 24

'usada na gravação do Centro de Trabalho
Public Const CTCOMPETENCIA_NAO_VALIDA_EXCLUSAO_TODAS = False

Public Const ORIGEM_KIT = 0
Public Const ORIGEM_ROTEIRO = 1

Public Const INSUMO_COMPRADO = "C"
Public Const INSUMO_PRODUZIDO = "P"

'constantes para Tipo de Mao-de-Obra
Public Const STRING_TIPO_MO_DESCRICAO = 50
Public Const STRING_TIPO_MO_OBSERVACAO = 50
Public Const NUM_MAX_TIPOMO = 9999

'constantes para Custeio de Roteiro de Fabricacao
Public Const STRING_CUSTEIO_NOMEREDUZIDO = 20
Public Const STRING_CUSTEIO_DESCRICAO = 50

'##################################################
'Inserido por Wagner
Public Const PO_STATUS_OK = 1
Public Const PO_STATUS_SOBRECARGA = 2
Public Const PO_STATUS_FALTAMATERIAL = 3
Public Const PO_STATUS_AMBOS = 4

Public Const PO_STATUS_NOME_OK = "OK"
Public Const PO_STATUS_NOME_SOBRECARGA = "SobreCarga no CT"
Public Const PO_STATUS_NOME_FALTAMATERIAL = "Falta Material"
Public Const PO_STATUS_NOME_AMBOS = "SobreCarga e Falta Material"

Public Const MRP_ACERTA_POR_DATA_INICIO = 1
Public Const MRP_ACERTA_POR_DATA_FIM = 2

Public Const STRING_PMP_VERSAO = 10

Public Const NUM_MAX_DIAS_PRODUCAO = 1000

Public Const SIMULACAO_ESTOQUE_TIPO_PREVVENDA = 1
Public Const SIMULACAO_ESTOQUE_TIPO_COMSUMO = 2
Public Const SIMULACAO_ESTOQUE_TIPO_PRODUCAO = 3
Public Const SIMULACAO_ESTOQUE_TIPO_PREVCOMPRA = 4
Public Const SIMULACAO_ESTOQUE_TIPO_SALDOATUAL = 5
Public Const SIMULACAO_ESTOQUE_TIPO_SALDODISPONIVEL = 6
Public Const SIMULACAO_ESTOQUE_TIPO_PRODUCAOINI = 7
Public Const SIMULACAO_ESTOQUE_TIPO_PREVCOMPRAINI = 8

Public Const SIMULACAO_ESTOQUE_REAL = 1
Public Const SIMULACAO_ESTOQUE_PROJETADO = 2
Public Const SIMULACAO_ESTOQUE_CONSOLIDADO = 3
Public Const SIMULACAO_ESTOQUE_OUTROS = 4

Public Const SIMULACAO_ESTOQUE_TIPO_ORDEM_ENTRADA = 1
Public Const SIMULACAO_ESTOQUE_TIPO_ORDEM_SAIDA = 2
Public Const SIMULACAO_ESTOQUE_TIPO_ORDEM_OUTRA = 3
'##################################################

'*** Types ***

Public Type typeMaquinas
    iCodigo As Integer
    iFilialEmpresa As Integer
    lNumIntDoc As Long
    sNomeReduzido As String
    sDescricao As String
    dTempoMovimentacao As Double
    dTempoPreparacao As Double
    dTempoDescarga As Double
    iRecurso As Integer
    dCustoHora As Double
    sProduto As String
    dPeso As Double
    dLargura As Double
    dComprimento As Double
    dEspessura As Double
End Type

Type typeCentrodeTrabalho
    lNumIntDoc As Long
    lCodigo As Long
    iFilialEmpresa As Integer
    sNomeReduzido As String
    sDescricao As String
    dCargaMin As Double
    dCargaMax As Double
    iTurnos As Integer
    dHorasTurno As Double
    iDiaisUteis(1 To 7) As Integer
End Type

Type typeCompetencias
    lNumIntDoc As Long
    lCodigo As Long
    sNomeReduzido As String
    sDescricao As String
    lNumIntDocCT As Long
    iPadrao As Integer
End Type

Type typeTaxaDeProducao
    lNumIntDoc As Long
    sProduto As String
    lNumIntDocMaq As Long
    lNumIntDocCompet As Long
    dLoteMax As Double
    dLoteMin As Double
    dLotePadrao As Double
    dTempoPreparacao As Double
    dTempoMovimentacao As Double
    dTempoDescarga As Double
    iTipo As Integer
    dQuantidade As Double
    sUMProduto As String
    dTempoOperacao As Double
    sUMTempo As String
    iAtivo As Integer
    dtData As Date
    dtDataDesativacao As Date
End Type

Type typeMaquinasInsumos
    lNumIntDocMaq As Long
    sProduto As String
    dQuantidade As Double
    sUMProduto As String
    sUMTempo As String
End Type

Type typeMaquinaOperadores
    lNumIntDoc As Long
    lNumIntDocMaq As Long
    iTipoMaoDeObra As Integer
    iQuantidade As Integer
    dPercentualUso As Double
End Type

Type typeRoteirosDeFabricacao
    lNumIntDoc As Long
    sProdutoRaiz As String
    sVersao As String
    sDescricao As String
    dtDataCriacao As Date
    dtDataUltModificacao As Date
    dQuantidade As Double
    sUM As String
    sAutor As String
    iComposicao As Integer
    dPercentualPerda As Double
    dCustoStandard As Double
    colOperacoes As Collection
    iNumMaxMaqPorOper As Integer
End Type

Type typeOperacoes
    lNumIntDoc As Long
    lNumIntDocRotFabr As Long
    iSeq As Integer
    lNumIntDocCompet As Long
    lNumIntDocCT As Long
    sObservacao As String
    iIgnoraTaxaProducao As Integer
    iSeqPai As Integer
    iSeqArvore As Integer
    iNivel As Integer
    iPosicaoArvore As Integer
    colOperacaoInsumos As Collection
    objOperacoesTempo As Object
    iNumMaxMaqPorOper As Integer
    iNumRepeticoes As Integer
End Type

Type typeOperacoesTempo
    lNumIntDocOperacao As Long
    dLoteMax As Double
    dLoteMin As Double
    dLotePadrao As Double
    dTempoPreparacao As Double
    dTempoMovimentacao As Double
    dTempoDescarga As Double
    dTempoOperacao As Double
    sUMTempo As String
    iTipo As Integer
    lNumIntDocMaq As Long
End Type

Type typeOperacaoInsumos
    lNumIntDocOper As Long
    sProduto As String
    dQuantidade As Double
    sUMProduto As String
    iComposicao As Integer
    dPercentualPerda As Double
    dCustoStandard As Double
    sVersaoKitComp As String
End Type

Type typeCTCompetencias
    lNumIntDocCT As Long
    lNumIntDocCompet As Long
End Type

Type typeCTMaquinas
    lNumIntDocMaq As Long
    lNumIntDocCT As Long
    iQuantidade As Integer
End Type

Type typeTiposDeMaodeObra
    iCodigo As Integer
    sDescricao As String
    sObservacao As String
    dCustoHora As Double
    sProduto As String
End Type

Type typeOrdemProducaoOperacoes
    lNumIntDoc As Long
    lNumIntDocItemOP As Long
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
    sProduto As String
    sVersao As String
    iIgnoraTaxaProducao As Integer
    iConsideraCarga As Integer
    iOrigem As Integer
    lNumIntDocOperOrigem As Long
    iNumMaxMaqPorOper As Integer
    iNumRepeticoes As Integer
End Type

Type typeOrdemProducaoInsumos
    lNumIntDocOper As Long
    sProduto As String
    dQuantidade As Double
    sUMProduto As String
    iComposicao As Integer
    dPercentualPerda As Double
    dCustoStandard As Double
    sVersaoKitComp As String
End Type

Type typePlanoOperacional
    lNumIntDoc As Long
    lNumIntDocPOPai As Long
    lNumIntDocPMP As Long
    lNumIntDocOper As Long
    iNivel As Integer
    iSeq As Integer
    sCodOPOrigem As String
    sProduto As String
    sVersao As String
    dQuantidade As Double
    sUM As String
    lNumIntDocCT As Long
    dtDataInicio As Date
    dtDataFim As Date
End Type

Type typePOMaquinas
    lNumIntDoc As Long
    lNumIntDocPO As Long
    lNumIntDocMaq As Long
    iQuantidade As Integer
    dtData As Date
    dHorasMaquina As Double
    lNumIntDocTxProd As Long
End Type

Type typePMPItens
    lNumIntDoc As Long
    lCodGeracao As Long
    sProduto As String
    sVersao As String
    dQuantidade As Double
    sUM As String
    dtDataNecessidade As Date
    sCodOPOrigem As String
    lCliente As Long
    iFilialCli As Integer
    iFilialEmpresa As Integer
    iPrioridade As Integer
End Type

Type typePMP
    lCodGeracao As Long
    dtDataGeracao As Date
    sVersao As String
End Type

'---------------------------------------

Type typeCTMaqProgDisp
    lNumIntDoc As Long
    lNumIntDocCT As Long
    lNumIntDocMaq As Long
    dtData As Date
    iQuantidade As Integer
    sObservacao As String
End Type

Type typeCTMaqProgTurno
    lNumIntDoc As Long
    lNumIntDocCT As Long
    lNumIntDocMaq As Long
    dtData As Date
    sObservacao As String
End Type

Type typeApontamentoProducao
    lNumIntDocPO As Long
    dtData As Date
    dPercConcluido As Double
    dQuantidade As Double
    iConcluido As Integer
    sObservacao As String
End Type

Type typeCusteioRoteiro
    lNumIntDoc As Long
    lCodigo As Long
    sNomeReduzido As String
    sDescricao As String
    sProduto As String
    sVersao As String
    sUMedida As String
    dQuantidade As Double
    dtDataCusteio As Date
    dtDataValidade As Date
    dCustoTotalInsumosKit As Double
    dCustoTotalInsumosMaq As Double
    dCustoTotalMaoDeObra As Double
    dPrecoTotalRoteiro As Double
    sObservacao As String
End Type

Type typeCusteioRotInsumosKit
    lNumIntDoc As Long
    lNumIntDocCusteioRot As Long
    iSeq As Integer
    sProduto As String
    sUMedida As String
    dQuantidade As Double
    dCustoUnitarioCalculado As Double
    dCustoUnitarioInformado As Double
    sObservacao As String
End Type

Type typeCusteioRotInsumosMaq
    lNumIntDoc As Long
    lNumIntDocCusteioRot As Long
    iSeq As Integer
    sProduto As String
    sUMedida As String
    dQuantidade As Double
    dCustoUnitarioCalculado As Double
    dCustoUnitarioInformado As Double
    sObservacao As String
End Type

Type typeCusteioRotMaoDeObra
    lNumIntDoc As Long
    lNumIntDocCusteioRot As Long
    iSeq As Integer
    iCodMO As Integer
    sUMedida As String
    dQuantidade As Double
    dCustoUnitarioCalculado As Double
    dCustoUnitarioInformado As Double
    sObservacao As String
End Type

Type typeCTOperadores
    iCodTipoMO As Integer
    lNumIntDocCT As Long
    iQuantidade As Integer
End Type

'###################################
'Inserido por Wagner 07/11/2005
'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCTMaquinasParadas
    lNumIntDoc As Long
    lCodigo As Long
    iFilialEmpresa As Integer
    dtData As Date
    lNumIntDocCT As Long
    lNumIntDocMaq As Long
    iTipo As Integer
    dHoras As Double
    iQtdMaquinas As Integer
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeItemOPOperacoesMaquinas
    lNumIntDoc As Long
    lNumIntDocOper As Long
    lNumIntDocMaq As Long
    dtData As Date
    dHoras As Double
    iQuantidade As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeItemOPOperacoesMO
    lNumIntDoc As Long
    lNumIntDocItemOPMaq As Long
    iTipoMO As Integer
    dHoras As Double
    iQuantidade As Integer
End Type
'###################################

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCertificados
    lCodigo As Long
    sDescricao As String
    sSigla As String
    lValidade As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCursos
    lNumIntDoc As Long
    lCodigo As Long
    iFilialEmpresa As Integer
    sDetalhamento As String
    sResponsavel As String
    dtDataInicio As Date
    dtDataConclusao As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCursoCertificados
    lNumIntDocCurso As Long
    lCodCertificado As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCursoMO
    lNumIntDocCurso As Long
    iCodMO As Integer
    sAvaliacao As String
    iAprovado As Integer
End Type


