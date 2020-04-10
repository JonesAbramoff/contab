Attribute VB_Name = "GlobalTRB"
Option Explicit

Public Const DATA_INICIO_SIMPLES_NACIONAL = #7/1/2007#
Public Const DATA_PIS_NOVO_CALC = #10/10/2013#
Public Const DATA_INICIO_LEI_12741 = #6/8/2013#

Public Const LEI_12741_TIPO_NAO_CALCULA = 0
Public Const LEI_12741_TIPO_AUTOMATICO = 1 'IBPT
Public Const LEI_12741_TIPO_SIMPLES = 2

'Limites de código de Natureza
Public Const NATUREZA_ENTRADA_COD_INICIAL = "1000"
Public Const NATUREZA_ENTRADA_COD_FINAL = "4999"
Public Const NATUREZA_SAIDA_COD_INICIAL = "5000"
Public Const NATUREZA_SAIDA_COD_FINAL = "9999"

Public Const NUMERO_PRIMEIRO_TIPOTRIB_USUARIO = 500

Public Const NAO_CONTRIBUINTE_ISS = 0
Public Const CONTRIBUINTE_ISS = 1

Public Const EXCECAO_ICMS_TIPO_APLICACAO_TODOS = 0
Public Const EXCECAO_ICMS_TIPO_APLICACAO_CONSUMIDOR_FINAL = 2
Public Const EXCECAO_ICMS_TIPO_APLICACAO_NAO_CONSUMIDOR_FINAL = 1

Public Const STRING_TIPO_ICMS_DESCRICAO = 50
Public Const STRING_TIPO_IPI_DESCRICAO = 50
Public Const STRING_TIPO_ISS_DESCRICAO = 50
Public Const STRING_TIPO_ISS_CST = 1
Public Const STRING_TIPO_PIS_COFINS_DESCRICAO = 150
Public Const STRING_TIPO_ICMS_CODIGO = 4
Public Const STRING_TIPO_IPI_CODIGO = 4
Public Const STRING_TIPO_PIS_COFINS_CODIGO = 4
Public Const STRING_IPICODENQ_DESC = 500

Public Const TIPOTRIB_SEMCREDDEB = 0
Public Const TIPOTRIB_CREDITA = 1
Public Const TIPOTRIB_DEBITA = 2

Public Const PRODUTO_ORIGEM_NACIONAL = 0
Public Const PRODUTO_ORIGEM_IMP_DIRETA = 1
Public Const PRODUTO_ORIGEM_EST_MERC_INTERNO = 2
Public Const PRODUTO_ORIGEM_NAC_COM_IMP = 3
Public Const PRODUTO_ORIGEM_NAC_PPB = 4
Public Const PRODUTO_ORIGEM_NAC_POUCO_IMP = 5
Public Const PRODUTO_ORIGEM_IMP_DIR_SEM_SIMILAR = 6
Public Const PRODUTO_ORIGEM_IMP_MERC_INT_SEM_SIMILAR = 7
Public Const PRODUTO_ORIGEM_NAC_COM_IMP2 = 8

Public Const PRODUTO_ORIGEM_NACIONAL_TEXTO = "Nacional, exceto as indicadas nos códigos 3, 4, 5 e 8"
Public Const PRODUTO_ORIGEM_IMP_DIRETA_TEXTO = "Estrangeira - Importação direta, exceto a indicada no código 6"
Public Const PRODUTO_ORIGEM_EST_MERC_INTERNO_TEXTO = "Estrangeira - Adquirida no mercado interno, exceto a indicada no código 7"
Public Const PRODUTO_ORIGEM_NAC_COM_IMP_TEXTO = "Nacional, mercadoria ou bem com Conteúdo de Importação superior a 40% e inferior ou igual a 70%"
Public Const PRODUTO_ORIGEM_NAC_PPB_TEXTO = "Nacional, cuja produção tenha sido feita em conformidade com os processos produtivos básicos de que tratam o Decreto-Lei nº 288/67, e as Leis nºs 8.248/91, 8.387/91, 10.176/01 e 11.484/07"
Public Const PRODUTO_ORIGEM_NAC_POUCO_IMP_TEXTO = "Nacional, mercadoria ou bem com Conteúdo de Importação inferior ou igual a 40%"
Public Const PRODUTO_ORIGEM_IMP_DIR_SEM_SIMILAR_TEXTO = "Estrangeira - Importação direta, sem similar nacional, constante em lista de Resolução CAMEX"
Public Const PRODUTO_ORIGEM_IMP_MERC_INT_SEM_SIMILAR_TEXTO = "Estrangeira - Adquirida no mercado interno, sem similar nacional, constante em lista de Resolução CAMEX"
Public Const PRODUTO_ORIGEM_NAC_COM_IMP2_TEXTO = "Nacional, mercadoria ou bem com Conteúdo de Importação superior a 70%"

Public Const ICMS_MODALIDADE_MARGEM = 0
Public Const ICMS_MODALIDADE_PAUTA = 1
Public Const ICMS_MODALIDADE_TABELA = 2
Public Const ICMS_MODALIDADE_VALOR = 3

Public Const ICMS_MODALIDADE_MARGEM_TEXTO = "Margem Valor Agregado (%)"
Public Const ICMS_MODALIDADE_PAUTA_TEXTO = "Pauta (Valor)"
Public Const ICMS_MODALIDADE_TABELA_TEXTO = "Preço Tabelado Máximo (Valor)"
Public Const ICMS_MODALIDADE_VALOR_TEXTO = "Valor da Operação"

Public Const ICMS_ST_MODALIDADE_TABELA = 0
Public Const ICMS_ST_MODALIDADE_LISTA_NEG = 1
Public Const ICMS_ST_MODALIDADE_LISTA_POS = 2
Public Const ICMS_ST_MODALIDADE_LISTA_NEU = 3
Public Const ICMS_ST_MODALIDADE_MARGEM = 4
Public Const ICMS_ST_MODALIDADE_PAUTA = 5

Public Const ICMS_ST_MODALIDADE_TABELA_TEXTO = "Preço tabelado ou máximo sugerido"
Public Const ICMS_ST_MODALIDADE_LISTA_NEG_TEXTO = "Lista Negativa (Valor)"
Public Const ICMS_ST_MODALIDADE_LISTA_POS_TEXTO = "Lista Positiva (Valor)"
Public Const ICMS_ST_MODALIDADE_LISTA_NEU_TEXTO = "Lista Neutra (Valor)"
Public Const ICMS_ST_MODALIDADE_MARGEM_TEXTO = "Margem Valor Agregado (%)"
Public Const ICMS_ST_MODALIDADE_PAUTA_TEXTO = "Pauta (Valor)"

Public Const TRIB_TIPO_CALCULO_VALOR = 0
Public Const TRIB_TIPO_CALCULO_PERCENTUAL = 1

Public Const TRIB_TIPO_CALCULO_PERMITE_NADA = 0
Public Const TRIB_TIPO_CALCULO_PERMITE_PERC = 1
Public Const TRIB_TIPO_CALCULO_PERMITE_VALOR = 2
Public Const TRIB_TIPO_CALCULO_PERMITE_AMBOS = 3
Public Const TRIB_TIPO_CALCULO_PERMITE_AMBOS_MANUAL = 4

Public Const TIPO_TRIB_TIPO_CALCULO_DESABILITADO = 0
Public Const TIPO_TRIB_TIPO_CALCULO_PERCENTUAL = 1
Public Const TIPO_TRIB_TIPO_CALCULO_VALOR = 2
Public Const TIPO_TRIB_TIPO_CALCULO_ESCOLHA = 3
Public Const TIPO_TRIB_TIPO_CALCULO_ESCOLHA_MANUAL = 4

Public Const TRIB_TIPO_CALCULO_VALOR_TEXTO = "Valor"
Public Const TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO = "Em Percentual"

Public Const TIPOTRIB_PRIORIDADE_CLIENTE_PRODUTO = 0
Public Const TIPOTRIB_PRIORIDADE_CLIENTE = 1
Public Const TIPOTRIB_PRIORIDADE_PRODUTO = 2

'##################################################
''Inserido por Wagner 29/09/05
Public Const TIPOTRIB_PRIORIDADE_FORNECEDOR_PRODUTO = 3
Public Const TIPOTRIB_PRIORIDADE_FORNECEDOR = 4

Public Const ICMSEXCECOES_TIPOCLIFORN_CLIENTE = 0
Public Const ICMSEXCECOES_TIPOCLIFORN_FORNECEDOR = 1
'##################################################

Public Const TIPOTRIB_PERMITE_ALIQUOTA = 1
Public Const TIPOTRIB_PERMITE_MARGLUCRO = 1
Public Const TIPOTRIB_PERMITE_REDUCAOBASE = 1

'tipos de contribuinte IPI
Public Const NAO_CONTRIBUINTE_IPI = 0
Public Const CONTRIBUINTE_IPI_NORMAL = 1
Public Const CONTRIBUINTE_IPI_50_PCT = 2

Public Const STRING_EXCECAO_TRIB_FUNDAMENTACAO = 150

Public Const PAIS_BRASIL = 1
Public Const PAIS_BRASIL_NOME As String = "Brasil"

Public Const TRIBUTO_NAO_INCIDE = 0
Public Const TRIBUTO_INCIDE = 1

Public Const ICMS_NAO_INSCRITA = 0
Public Const ICMS_INSCRITA = 1

Public Const FRETE_EMITENTE = 1
Public Const FRETE_DESTINATARIO = 2
Public Const FRETE_TERCEIROS = 3
Public Const FRETE_SEM = 4

'Public Const VAR_PREENCH_VAZIO = 0              'ainda nao foi preenchida
'Public Const VAR_PREENCH_AUTOMATICO = 1         'preenchto segundo calculo do sistema
'Public Const VAR_PREENCH_MANUAL = 2             'preenchto pelo usuario

'ver tabela TiposTribICMS
Public Const ICMS_TIPO_NAO_TRIBUTADO = 0
Public Const ICMS_TIPO_NORMAL = 1
Public Const ICMS_TIPO_ISENTO = 2
Public Const ICMS_TIPO_SUSPENSO = 3
Public Const ICMS_TIPO_RED_BASE_E_SUBST = 4
Public Const ICMS_TIPO_COM_DIFERIMENTO = 5
Public Const ICMS_TIPO_COM_SUBST_TRIB = 6
Public Const ICMS_TIPO_RED_BASE = 7
Public Const ICMS_TIPO_COBR_SUBST_ANT = 8
Public Const ICMS_TIPO_OUTRAS = 99

'ver tabela TiposTribIPI
Public Const IPI_TIPO_NAO_TRIBUTADO = 0
Public Const IPI_TIPO_NORMAL = 1
Public Const IPI_TIPO_50PCTO = 2
Public Const IPI_TIPO_ISENTO = 3
Public Const IPI_TIPO_ALIQUOTA_ZERO = 4
Public Const IPI_TIPO_SUSPENSO = 5
Public Const IPI_TIPO_OUTROS = 6

Public Const MOV_ORIG_INTERNA = 0
Public Const MOV_ORIG_INTERESTADUAL = 1
Public Const MOV_ORIG_INTERNACIONAL = 2

Public Const TIPO_TRIBUTACAO_DESCRICAO = 100

Public Const ITEM_TIPO_NORMAL = 0
Public Const ITEM_TIPO_FRETE = 1
Public Const ITEM_TIPO_SEGURO = 2
Public Const ITEM_TIPO_DESCONTO = 3
Public Const ITEM_TIPO_OUTRAS_DESP = 4

Public Const EXCECAO_PIS_COFINS_TIPO_AMBOS = 0
Public Const EXCECAO_PIS_COFINS_TIPO_PIS = 1
Public Const EXCECAO_PIS_COFINS_TIPO_COFINS = 2

Public Const STRING_EXCECAO_FUNDAMENTACAO = 150

Type typeTributacaoTipo
    iTipo As Integer
    sDescricao As String
    iEntrada As Integer
    iICMSIncide As Integer
    iICMSTipo As Integer
    iICMSBaseComIPI As Integer
    iICMSCredita As Integer
    iIPICredita As Integer
    iIPIIncide As Integer
    iIPITipo As Integer
    iIPIFrete As Integer
    iIPIDestaca As Integer
    iISSIncide As Integer
    iIRIncide As Integer
    dIRAliquota As Double
    iINSSIncide As Integer
    dINSSRetencaoMinima As Double
    dINSSAliquota As Double
    iPISCredita As Integer
    iPISRetencao As Integer
    iISSRetencao As Integer
    iCOFINSCredita As Integer
    iCOFINSRetencao As Integer
    iCSLLRetencao As Integer
    iISSTipo As Integer
    iPISTipo As Integer
    iCOFINSTipo As Integer
    iICMSSimplesTipo As Integer
    iRegimeTributario As Integer
    sNatBCCred As String
    iISSIndExigibilidade As Integer
    sIPICodEnq As String
End Type

Type typeICMSExcecao
    sEstadoDestino As String
    sEstadoOrigem As String
    sCategoriaProduto As String
    sCategoriaProdutoItem As String
    sCategoriaCliente As String
    sCategoriaClienteItem As String
    iTipo As Integer
    dPercRedBaseCalculo As Double
    dAliquota As Double
    dPercMargemLucro As Double
    sFundamentacao As String
    iPrioridade As Integer
    '################################
    'Inserido por Wagner
    sCategoriaFornecedor As String
    sCategoriaFornecedorItem As String
    iTipoCliForn As Integer
    dPercRedBaseCalculoSubst As Double
    '################################
    iUsaPauta As Integer
    dValorPauta As Double
    iTipoSimples As Integer
    iGrupoOrigemMercadoria As Integer
    dICMSPercFCP As Double
    iTipoAplicacao As Integer
    iICMSSTBaseDupla As Integer
    dtICMSSTBaseDuplaIni As Date
    scBenef As String
    iICMSMotivo As Integer
End Type

Type typeIPIExcecao
    sCategoriaCliente As String
    sCategoriaClienteItem As String
    sCategoriaProduto As String
    sCategoriaProdutoItem As String
    iTipo As Integer
    dPercRedBaseCalculo As Double
    dAliquota As Double
    dPercMargemLucro As Double
    sFundamentacao As String
    iPrioridade As Integer
    iTipoCalculo As Integer
    dAliquotaRS As Double
End Type

Type typeNatOpPadrao
    lCodigo As Long
    iTipoOperacao As Integer
    iTipoAtividadeEmp As Integer
    iTipoTribEmp As Integer
    sCFOPEmp As String
    iTipoAtividadeExt As Integer
    iTipoTribExt As Integer
    sCFOPExt As String
    iPadrao As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTipoTribPISCOFINS
    iTipo As Integer
    sDescricao As String
    iTipoCalculo As Integer
    iVersaoNFE As Integer
    iEntrada As Integer
    iSaida As Integer
End Type

Public Const PISCOFINSEXCECOES_TIPOCLIFORN_CLIENTE = 0
Public Const PISCOFINSEXCECOES_TIPOCLIFORN_FORNECEDOR = 1

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePISCOFINSExcecao
    sCategoriaCliente As String
    sCategoriaClienteItem As String
    sCategoriaProduto As String
    sCategoriaProdutoItem As String
    iTipo As Integer
    iTipoPIS As Integer
    iTipoCOFINS As Integer
    iPISTipoCalculo As Integer
    iCOFINSTipoCalculo As Integer
    dAliquotaPisRS As Double
    dAliquotaPisPerc As Double
    dAliquotaCofinsRS As Double
    dAliquotaCofinsPerc As Double
    sFundamentacao As String
    iPrioridade As Integer
    sCategoriaFornecedor As String
    sCategoriaFornecedorItem As String
    iTipoCliForn As Integer
    iTipoPISE As Integer
    iTipoCOFINSE As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeIPICodEnquadramento
    sCodigo As String
    sGrupoCST As String
    sDescCompleta As String
    iTipoIPI As Integer
End Type

