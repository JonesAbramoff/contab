Attribute VB_Name = "GlobalEST"
Option Explicit

Public Const STRING_PRESTSERV_NOME = 50
Public Const STRING_PRESTSERV_NOMERED = 20

Public Const RECEBIMENTO_MATERIAL_FCOM As String = "RecebMaterialFCom"

'Constantes que identificam rotina batch
Public Const ROTINA_CUSTO_MEDIO_PRODUCAO_BATCH = 1
Public Const ROTINA_ATUALIZA_INVLOTE_BATCH = 2
Public Const ROTINA_REPROCESSAMENTO_MOVEST_BATCH = 3
Public Const ROTINA_REPROCESSAMENTO_TESTAINT_BATCH = 4

Public Const FORMA_CALC_PESO_ENT_COM_MEDIA = 0
Public Const FORMA_CALC_PESO_ENT_COM_ULT = 1

Public Const ROTINA_GERACONTRATOCOBRANCA_BATCH = 5

'Títulos da tela de acompanhamento da rotina batch
Public Const TITULO_TELABATCH_CUSTO_MEDIO_PRODUCAO = "Custo Médio de Produção"
Public Const TITULO_TELABATCH_ATUALIZA_INVLOTE = "Atualização de Inventário Lote"
Public Const TITULO_TELABATCH_REPROCESSAMENTO_MOVEST = "Reprocessamento de Movimentos de Estoque"
Public Const TITULO_TELABATCH_REPROCESSAMENTO_TESTAINT = "Teste de Integridade do Reprocessamento"
Public Const TITULO_TELABATCH_GERACONTRATOCOBRANCA = "Faturamento de Contratos"

'Tipos de Transferencia (Movimento de Estoque)
Public Const TRANSF_DISPONIVEL_STRING = "Disponível"
Public Const TRANSF_DEFEITUOSO_STRING = "Defeituoso"
Public Const TRANSF_INDISPONIVEL_STRING = "Não Disponível"
Public Const TRANSF_OUTRAS_TERC_STRING = "Outras de 3os"
Public Const TRANSF_BENEF_TERC_STRING = "Beneficiado de 3os."

'Para tabela ESTConfig
Public Const STRING_ESTCONFIG_CODIGO = 50
Public Const STRING_ESTCONFIG_DESCRICAO = 150
Public Const STRING_ESTCONFIG_CONTEUDO = 255

Public Const STRING_PRIORIDADE_MAQUINAS = "Maquinas"
Public Const STRING_PRIORIDADE_PRODUTOS = "Produtos"
Public Const STRING_PRIORIDADE_PRODUTOS_ANC = "Produtos_Anc"

Public Const STRING_GRIDPRIORIDADE_MAQUINAS = "Maquina"
Public Const STRING_GRIDPRIORIDADE_PRODUTOS = "Produto"
Public Const STRING_GRIDPRIORIDADE_PRODUTOS_ANC = "Produto Ancestral"

Public Const NUM_PRIORIDADES_SELECAO = 3
Public Const NAO_GERA_REQCOMPRA_EM_LOTE = 0
Public Const GERA_REQCOMPRA_EM_LOTE = 1

Public Const ESTCFG_PRIORIDADE_MAQUINA = "PRIORIDADE_MAQUINA"
Public Const ESTCFG_PRIORIDADE_PRODUTO = "PRIORIDADE_PRODUTO"
Public Const ESTCFG_PRIORIDADE_PRODUTO_ANCESTRAL = "PRIORIDADE_PRODUTO_ANCESTRAL"
Public Const ESTCFG_CLASSE_UM_TEMPO = "CLASSE_UM_TEMPO"
Public Const ESTCFG_DATA_INICIO_MRP = "DATA_INICIO_OPERACOES_MRP"
Public Const ESTCFG_TEM_REPETICAO_OPERACAO = "TEM_REPETICAO_OPERACAO"
Public Const ESTCFG_GERA_REQCOMPRA_EM_LOTE = "ESTCFG_GERA_REQCOMPRA_EM_LOTE"
Public Const ESTCFG_USA_PERDA_INSUMOS_KIT = "ESTCFG_USA_PERDA_INSUMOS_KIT"
Public Const ESTCFG_FORMA_CALC_PRECO_ENT_COM = "FORMA_CALC_PRECO_ENT_COM"

Public Const INVENTARIOCODBARRAAUTO_SIM = 1
Public Const INVENTARIOCODBARRAAUTO_NAO = 0

Public Const ESTCFG_INVENTARIOCODBARRAUTO = "ESTCFG_INVENTARIOCODBARRAUTO"
Public Const ESTCFG_PROD_ENT_VERIFICA_BX_OP = "PROD_ENT_VERIFICA_BX_OP"
Public Const ESTCFG_TRAZ_PRECO_ULT_COMPRA = "TRAZ_PRECO_ULT_COMPRA"
Public Const ESTCFG_REL_PONTOPED_EXIBE_PPZERADO = "REL_PONTOPED_EXIBE_PPZERADO"
Public Const ESTCFG_TIPOMAODEOBRA_EXIBE_ABA_CURSOS = "TIPOMAODEOBRA_EXIBE_ABA_CURSOS"
Public Const ESTCFG_SERIE_ELETRONICA_PADRAO = "SERIE_ELETRONICA_PADRAO"
Public Const ESTCFG_PRODUTO_EXIBE_COMISSAO = "PRODUTO_EXIBE_COMISSAO"
Public Const ESTCFG_TRATAMENTO_ABERTURA_MES_INI = "TRATAMENTO_ABERTURA_MES_INI"
Public Const ESTCFG_ALTERA_DATA_KIT_NA_GRAVACAO = "ALTERA_DATA_KIT_NA_GRAVACAO"

'usado na leirua de escaninho p/ 3ºs
Public Const ESCANINHO_TERCEIROS = 1

Type typeESTConfig
    sCodigo As String
    iFilialEmpresa As Integer
    sDescricao As String
    iTipo As Integer
    sConteudo As String
End Type

'constantes que delimitam o numero de caracteres dos codigos de barra
Public Const PRODUTO_CODBARRAS_MIN = 8
Public Const PRODUTO_CODBARRAS_MAX = 14

Type typePrestServ
    lCodigo As Long
    sNome As String
    sNomeReduzido As String
    lFornecedor As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeImportDIXml
    sNumeroDI As String
    dPesoBruto As Double
    dPesoLiquido As Double
    dtDataDesembaraco As Date
    dtDataRegistro As Date
    dtDataChegada As Date
    lOperacaoCod As Long
    sOperacaoDesc As String
    lPaisProcedenciaCod As Long
    sPaisProcedenciaNome As String
    lUrfEntradaCod As Long
    sUrfEntradaNome As String
    dtDataEmbarque As Date
    sEmbarqueLocal As String
    dFreteCollect As Double
    dFreteEmTerritorioNacional As Double
    lFreteMoeda As Long
    sFreteMoedaNome As String
    dFretePrepaid As Double
    dFreteTotalDolares As Double
    dFreteTotalMoeda As Double
    dFreteTotalReais As Double
    dDescargaTotalDolares As Double
    dDescargaTotalReais As Double
    dEmbarqueTotalDolares As Double
    dEmbarqueTotalReais As Double
    lModDespachoCod As Long
    sModDespachoNome As String
    lSeguroMoeda As Long
    sSeguroMoedaNome As String
    dSeguroTotalDolares As Double
    dSeguroTotalMoeda As Double
    dSeguroTotalReais As Double
    lTotalAdicoes As Long
    lUrfDespachoCod As Long
    sUrfDespachoNome As String
    lViaTransporteCod As Long
    sViaTransporteMultimodal As String
    sViaTransporteNome As String
    sNomeTransportador As String
    lPaisTransportadorCod As Long
    sPaisTransportadorNome As String
    lRecintoAduaneiroCod  As Long
    sRecintoAduaneiroNome As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeImportADIXml
    sNumeroDI As String
    lNumeroAdicao As Long
    lNumeroLI As Long
    dCofinsPercRed As Double
    dCofinsBC As Double
    lCofinsRT As Long
    sCofinsRTNome As String
    dCofinsAliq As Double
    dCofinsAliqQtde As Double
    dCofinsAliqVlr As Double
    dCofinsReduzida As Double
    dCofinsValorDevido As Double
    dCofinsValorRecolher As Double
    dPisPercRed As Double
    dPisBC As Double
    lPisRT As Long
    sPisRTNome As String
    dPisAliq As Double
    dPisAliqQtde As Double
    dPisAliqVlr As Double
    dPisReduzida As Double
    dPisValorDevido As Double
    dPisValorRecolher As Double
    lCondVendMoedaCod As Long
    sCondVendMoedaCodNome As String
    dCondVendValorMoeda As Double
    dCondVendValorReais As Double
    sAplicacao As String
    sNCM As String
    sNCMNome As String
    dPesoLiquido As Double
    dDCRReducao As Double
    sDCRID As String
    dDCRValorDevido As Double
    dDCRValorDolar As Double
    dDCRValorReal As Double
    dDCRValorRecolher As Double
    sNomeFornecedor As String
    lFreteMoedaCod As Long
    sFreteMoedaNome As String
    dFreteValorMoeda As Double
    dFreteValorReal As Double
    dIIBC As Double
    lIIRT As Long
    sIIRTNome As String
    dIIAliq As Double
    dIIPercReducao As Double
    dIIAliqReduzida As Double
    dIIValorCalc As Double
    dIIValorDevido As Double
    dIIValorRecolher As Double
    dIIValorReduzido As Double
    lIPIRT As Long
    sIPIRTNome As String
    dIPIAliq As Double
    dIPIAliqEsp As Double
    dIPIAliqEspVlr As Double
    dIPIAliqRed As Double
    dIPIValorDevido As Double
    dIPIValorRecolher As Double
    lSeguroMoeda As Long
    dSeguroValorMoeda As Double
    dSeguroValorReais As Double
    dFreteInternacionalValorReais As Double
    dSeguroInternacionalValorReais As Double
    dCondVendaValorTotal As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeImportMADIXml
    sNumeroDI As String
    lNumeroAdicao As Long
    lSeq As Long
    sDescricao As String
    sUM As String
    dQuantidade As Double
    dValorUnitario As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCEST
    sCodigo As String
    sNCM As String
    sDescricao As String
End Type
