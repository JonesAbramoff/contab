Attribute VB_Name = "GlobalFAT"
Option Explicit

Public Const VERIFICA_CREDITO_CLIENTE = 0
Public Const NAO_VERIFICA_CREDITO_CLIENTE = 1

Public Const PRODUTO_CAT_COD_SERVICO_NFE = "CODSERVICONFE"

Public Const MAPAENTREGA_TIPODOC_NF = 0
Public Const MAPAENTREGA_TIPODOC_PV = 1

Public Const PRECO_GRUPO_TIPO_VALOR = 0
Public Const PRECO_GRUPO_TIPO_PERCENTUAL = 1

Public Const STRING_RPS_SERIE = 5
Public Const STRING_RPS_TIPO = 5
Public Const STRING_RPS_SITUACAO = 1
Public Const STRING_RPS_CGC = 14
Public Const STRING_RPS_INSCEST = 15
Public Const STRING_RPS_INSCMUNI = 15
Public Const STRING_RPS_RAZAOSOCIAL = 115
Public Const STRING_RPS_ENDERECO = 103
Public Const STRING_RPS_ENDTIPO = 3
Public Const STRING_RPS_ENDNUM = 10
Public Const STRING_RPS_ENDCOMP = 60
Public Const STRING_RPS_BAIRRO = 72
Public Const STRING_RPS_CIDADE = 50
Public Const STRING_RPS_UF = 2
Public Const STRING_RPS_CEP = 8
Public Const STRING_RPS_EMAIL = 80
Public Const STRING_RPS_DISCRIMINACAO = 1000

Public Const NFE_MAX_ITEM_DESCRICAO = 120

Public Const TIPO_REGISTRO_TRANSPORTADORA = 1
Public Const TIPO_REGISTRO_NATUREZAOP = 2
Public Const TIPO_REGISTRO_PRODUTO = 3
Public Const TIPO_REGISTRO_PRODUTOCATEGORIA = 4
Public Const TIPO_REGISTRO_CLIENTE = 5
Public Const TIPO_REGISTRO_FILIALCLIENTE = 6
Public Const TIPO_REGISTRO_FORNECEDOR = 7
Public Const TIPO_REGISTRO_FILIALFORNECEDOR = 8
Public Const TIPO_REGISTRO_NOTAFISCAL = 9
Public Const TIPO_REGISTRO_PARCELAPAG = 10
Public Const TIPO_REGISTRO_PARCELAREC = 11
Public Const TIPO_REGISTRO_LANCAMENTO = 12
Public Const TIPO_REGISTRO_LANPENDENTE = 13
Public Const TIPO_REGISTRO_ITEMNOTAFISCAL = 14
Public Const TIPO_REGISTRO_ITEMNFGRADE = 15
Public Const TIPO_REGISTRO_ALOCACAOITEMNF = 16
Public Const TIPO_REGISTRO_LOCALIZACAOITEMNFGRADE = 17

Public Const VENDEDOR_VINCULO_AUTONOMO = 1
Public Const VENDEDOR_VINCULO_EMPREGADO = 2
Public Const VENDEDOR_VINCULO_EMPRESA = 3

Public Const STRING_CONHECIMENTOFRETE_CALCULOATE = 20
Public Const STRING_CONHECIMENTOFRETE_COLETA = 50
Public Const STRING_CONHECIMENTOFRETE_ENTREGA = 50
Public Const STRING_CONHECIMENTOFRETE_DESTINATARIO = 50
Public Const STRING_CONHECIMENTOFRETE_REMETENTE = 50
Public Const STRING_CONHECIMENTOFRETE_MARCA = 20
Public Const STRING_CONHECIMENTOFRETE_NATUREZACARGA = 30
Public Const STRING_CONHECIMENTOFRETE_NOTAS = 100

Public Const STRING_ROTA_CODIGO = 20
Public Const STRING_ROTA_DESCRICAO = 250
Public Const STRING_ROTAPONTO_OBS = 250

Public Const RPS_CIDADE_SAO_PAULO = 1
Public Const RPS_CIDADE_RECIFE = 2
Public Const RPS_CIDADE_VOLTA_REDONDA = 3
Public Const RPS_CIDADE_BARUERI = 4
Public Const RPS_CIDADE_TAUBATE = 5

'Constantes para inicialização das strings de ComissoesRegras
Public Const STRING_COMISSOESREGRAS_REGRA = 255
Public Const STRING_COMISSOESREGRAS_VALORBASE = 255
Public Const STRING_COMISSOESREGRAS_PERCCOMISSAO = 255

'Constante para rastreamento
Public Const ESCANINHO_HABILITADO = 1
Public Const ESCANINHO_DESABILITADO = 0

'Constantes para inicialização das string de MnemonicoComissoes e MnemonicoComissoesAux
Public Const STRING_MNEMONICOCOMISSOES_MNEMONICO = 25
Public Const STRING_MNEMONICOCOMISSOES_DESCRICAO = 255
Public Const STRING_MNEMONICOCOMISSOES_NOMEGRID = 50

'Constantes
Public Const NFISCAL_PAGO = 1
Public Const NFISCAL_PAGO_PARCIAL = 2
Public Const NFISCAL_NAO_PAGO = 3

Public Const COD_TABELA_PRECO_LIGHT = 1
Public Const NUM_PROX_RELPRAZOPAGTO As String = "NUM_PROX_RELPRAZOPAGTO"

Public Const FATCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Public Const FATCONFIG_GERA_LOTE_AUTOMATICO = 1
Public Const FATCONFIG_INVENTARIOCODBARRAAUTO = 2

Public Const IMPRIME_NOTA_FISCAL = 1
Public Const NAO_IMPRIME_NOTA_FISCAL = 0

'N.Fiscal Devolução
Public Const FORN_INATIVO = 0
Public Const FORN_ATIVO = 1

'Alocação Produto
Public Const NENHUMA_SELECAO = -1
Public Const SELECAO_OK = 0
Public Const CANCELA_ACIMA_DA_RESERVADA = 1
Public Const NAO_RESERVAR_PRODUTO = 2

'indica que um crédito foi aprovado ou não
Public Const CREDITO_APROVADO = 1
Public Const CREDITO_RECUSADO = 0

'indica o numero de lComandos do array de NFiscalFatura
Public Const NUM_MAX_LCOMANDO_NFISCALFAT = NUM_MAX_LCOMANDO_MOVESTOQUE + 55

'indica o numero de lComandos do array de NFiscal
Public Const NUM_MAX_LCOMANDO_NFISCAL = NUM_MAX_LCOMANDO_MOVESTOQUE + 55

Public Const PEDIDOVENDA_FATURA_INTEGRAL = 1 'indica que só fatura o pedido de venda integralmente

'Utilizado em Geracao de Nota Fiscal.
Public Const PEDIDO_NAO_CRIA_NFISCAL = 1 'Indica que o pedido não irá gerar nota fiscal
Public Const PEDIDO_CRIA_NFISCAL = 0 'indica que o pedido irá gerar nota fiscal

Public Const MOTIVO_NAOGERADA_POR_BLOQUEIO = 1 'Indica que a nota não foi gerada porque o pedido possui um bloqueio que o impeça de ser faturado.
Public Const MOTIVO_NAOGERADA_POR_FALTAESTOQUE = 2 'Indica que a nota não foi gerada porque o pedido não possuia um estoque para atende-lo.
Public Const MOTIVO_NAOGERADA_OUTROS = 3 'Indica que a nota não foi gerada por outros motivos (Erro Interno).
Public Const MOTIVO_NAOGERADA_POR_BLOQUEIO_CREDITO = 4 'Indica que a nota não foi gerada porque o pedido possui um bloqueio de crédito que impede o seu faturamento.
Public Const MOTIVO_NAOGERADA_POR_FALTACREDITO = 5 'Indica que a nota não foi gerada por falta de liberação de crédito.
Public Const MOTIVO_NAOGERADA_BLOQUEIO_DIAS_ATRASO = 6 'Indica que a nota não foi gerada porque o pedido possui um bloqueio de dias de atraso que impede o seu faturamento.

Public Const MOTIVO_NAOGERADA_DESCRICAO_BLOQUEIO = "Não gerou nota fiscal por possuir bloqueio."
Public Const MOTIVO_NAOGERADA_DESCRICAO_FALTA_ESTOQUE = "Não gerou nota fiscal por falta de estoque."
Public Const MOTIVO_NAOGERADA_DESCRICAO_OUTROS = "Erro Interno."
Public Const MOTIVO_NAOGERADA_DESCRICAO_BLOQUEIO_CREDITO = "Não gerou nota fiscal por possuir bloqueio de crédito."
Public Const MOTIVO_NAOGERADA_DESCRICAO_FALTACREDITO = "Não gerou nota fiscal por falta de liberação de crédito."

'Constantes

'para tabela PrevVenda
Public Const STRING_PREVVENDA_CODIGO = 10

'para tabela FATConfig
Public Const STRING_FATCONFIG_CODIGO = 50
Public Const STRING_FATCONFIG_DESCRICAO = 150
Public Const STRING_FATCONFIG_CONTEUDO = 255


'
''chaves em FATConfig
'Public Const FATCFG_PEDIDO_RESERVA_AUTOMATICA = "PEDIDO_RESERVA_AUTOMATICA"
'Public Const FATCFG_NFISCAL_ALOCA_AUTOMATICA = "NFISCAL_ALOCA_AUTOMATICA"
'Public Const FATCFG_PEDIDO_VENDA_EDITA_COMISSAO = "PEDIDO_VENDA_EDITA_COMISSAO"
'Public Const FATCFG_NFISCAL_EDITA_COMISSAO = "NFISCAL_EDITA_COMISSAO"
'Public Const FATCFG_NIVEL_TABELAS_PRECOS = "NIVEL_TABELAS_PRECOS"
'Public Const FATCFG_TIPO_CUSTEIO_ESTOQUE = "TIPO_CUSTEIO_ESTOQUE"
'Public Const FATCFG_EXIGE_DATA_SAIDA_NF = "EXIGE_DATA_SAIDA_NF"
'Public Const FATCFG_COND_PAGTO_PADRAO = "COND_PAGTO_PADRAO"
'Public Const FATCFG_AGLUTINA_LANCAM_POR_DIA = "AGLUTINA_LANCAM_POR_DIA"
'Public Const FATCFG_GERA_LOTE_AUTOMATICO = "GERA_LOTE_AUTOMATICO"
'Public Const FATCFG_NUM_PROX_CANAL_VENDA = "NUM_PROX_CANAL_VENDA"
'Public Const FATCFG_NUM_PROX_PEDIDO_VENDA = "NUM_PROX_PEDIDO_VENDA"
'Public Const FATCFG_NUM_PROX_TIPO_DE_BLOQUEIO = "NUM_PROX_TIPO_DE_BLOQUEIO"
'
'Public Const CFGFAT_TABELAS_PRECOS_FILIAL = 0
'Public Const CFGFAT_TABELAS_PRECOS_EMPRESA = 1
'
'Public Const CFGFAT_NAO_EXIGE_DATA_SAIDA_NF = 0
'Public Const CFGFAT_EXIGE_DATA_SAIDA_NF = 1
'
'Public Const CFGFAT_CUSTO_EST_STANDARD = 0
'Public Const CFGFAT_CUSTO_EST_MEDIA_FIXA = 1
'Public Const CFGFAT_CUSTO_EST_MEDIA_MOVEL = 2

Public Const PEDVENDA_NAO_RESERVA_AUTOMATICA = 0
Public Const PEDVENDA_RESERVA_AUTOMATICA = 1

Public Const NFISCAL_NAO_ALOCA_AUTOMATICA = 0
Public Const NFISCAL_ALOCA_AUTOMATICA = 1

Public Const PEDVENDA_NAO_EDITA_COMISSAO = 0
Public Const PEDVENDA_EDITA_COMISSAO = 1

Public Const NFISCAL_NAO_EDITA_COMISSAO = 0
Public Const NFISCAL_EDITA_COMISSAO = 1

'Public Const NAO_AGLUTINA_LANCAM_POR_DIA = 0
'Public Const AGLUTINA_LANCAM_POR_DIA = 1

'Public Const NAO_GERA_LOTE_AUTOMATICO = 0
'Public Const GERA_LOTE_AUTOMATICO = 1

Public Const PEDVENDA_VINCULADO_NF = 1

'Public Const BLOQUEIO_ESTOQUE_PARCIAL = 1000
'Public Const BLOQUEIO_ESTOQUE_NAO_RESERVA = 1001

Public Const PEDVENDA_PODE_FATURAR_PARCIAL = 0
Public Const PEDVENDA_FATURAR_SOMENTE_INTEGRAL = 1

Public Const ITEM_PEDVENDA_ABERTO = 0
Public Const ITEM_PEDVENDA_ATENDIDO = 1

'Tipos de NFiscal (Entrada/Saida)
Public Const NF_SAIDA = 0
Public Const NF_ENTRADA = 1

'para tabela de canais de venda
Public Const STRING_CANAL_VENDA_NOME = 50
Public Const STRING_CANAL_VENDA_NOME_REDUZIDO = 20
'Public Const STRING_CANAL_VENDA_DESCRICAO = 150

'para tabela de tipos de bloqueio
Public Const STRING_TIPO_BLOQUEIO_NOME_REDUZIDO = 20
Public Const STRING_TIPO_BLOQUEIO_DESCRICAO = 100

'Para Grid de Ítens de Pedido de Venda, NFiscal
Public Const NUM_MAXIMO_ITENS = 700

Public Const NUM_MAX_BLOQUEIOS_LIBERACAO = 1000

'Tipos de Formularios de N. Fiscal
Public Const TIPO_FORMULARIO_NFISCAL = 1
Public Const TIPO_FORMULARIO_NFISCAL_FATURA = 2
Public Const TIPO_FORMULARIO_NFISCAL_SERVICO = 3
Public Const TIPO_FORMULARIO_NFISCAL_FATURA_SERVICO = 4
Public Const TIPO_FORMULARIO_NFISCAL_FRETE = 5
Public Const TIPO_FORMULARIO_NFISCAL_FATURA_FRETE = 6

'Incluido por Jorge Specian
'-----------------------------------------
'para tela de Projeto
Public Const NUM_MAX_PROJETOS = 99999

Public Const STRING_PROJETO_NOMERED = 20
Public Const STRING_PROJETO_DESCRICAO = 50
Public Const STRING_PROJETO_RESPONSAVEL = 20
Public Const STRING_PROJETO_OBSERVACAO = 255

Public Const NUMMAX_DECIMAIS_VARIACAO = 2

'conteudo da Combo Destino
Public Const ITEMDEST_ORCAMENTO_DE_VENDA = 1
Public Const STRING_ITEMDEST_ORCAMENTO_DE_VENDA = "Orçamento de Venda"
Public Const ITEMDEST_PEDIDO_DE_VENDA = 2
Public Const STRING_ITEMDEST_PEDIDO_DE_VENDA = "Pedido de  Venda"
Public Const ITEMDEST_ORDEM_DE_PRODUCAO = 3
Public Const STRING_ITEMDEST_ORDEM_DE_PRODUCAO = "Ordem de Produção"
Public Const ITEMDEST_NFISCAL_SIMPLES = 4
Public Const STRING_ITEMDEST_NFISCAL_SIMPLES = "NFiscal Simples"
Public Const ITEMDEST_NFISCAL_FATURA = 5
Public Const STRING_ITEMDEST_NFISCAL_FATURA = "NFiscal Fatura"
Public Const ITEMDEST_ORDEM_DE_SERVICO = 6
Public Const STRING_ITEMDEST_ORDEM_DE_SERVICO = "Ordem de Serviço"

'constantes para formação do Status
Public Const STRING_STATUS_EXPORTADO = "EXPORTADO"
Public Const STRING_STATUS_NAO_EXPORTADO = "NÃO EXPORTADO"
'-----------------------------------------

Public Const NUM_MAX_ITENS_KITVENDA = 100

Public Const MAPBLOQGEN_TIPOTELA_PEDIDOSRV = 1
Public Const MAPBLOQGEN_TIPOTELA_ORCAMENTOSRV = 2

Public Const STRING_MAPBLOQGEN_NOMETELAEDITADOCBLOQ = 50
Public Const STRING_MAPBLOQGEN_NOMECLASSEDOCBLOQ = 50
Public Const STRING_MAPBLOQGEN_CLASSENOMECAMPOCHAVE = 50
Public Const STRING_MAPBLOQGEN_NOMEBROWSECHAVE = 50
Public Const STRING_MAPBLOQGEN_NOMETABELABLOQUEIOS = 50
Public Const STRING_MAPBLOQGEN_NOMETELATESTAPERMISSAO = 50
Public Const STRING_MAPBLOQGEN_NOMEFUNCLIBERACUST = 50
Public Const STRING_MAPBLOQGEN_PROJETOCLASSEDOCBLOQ = 50
Public Const STRING_MAPBLOQGEN_NOMEVIEWLEBLOQUEIOS = 50
Public Const STRING_MAPBLOQGEN_TABELANOMECAMPOCHAVE = 50
Public Const STRING_MAPBLOQGEN_NOMEFUNCLEDOC = 50
Public Const STRING_MAPBLOQGEN_NOMECOLECAOBLOQDOC = 50
Public Const STRING_MAPBLOQGEN_CLASSEDOCNOMEQTD = 50
Public Const STRING_MAPBLOQGEN_CLASSEDOCNOMEQTDRESERVADA = 50
Public Const STRING_MAPBLOQGEN_CLASSEDOCNOMEUM = 50
Public Const STRING_MAPBLOQGEN_CLASSEDOCNOMECOLITEM = 50
Public Const STRING_MAPBLOQGEN_CLASSEDOCNOMEPRODUTO = 50
Public Const STRING_MAPBLOQGEN_CLASSEDOCNOMEVALOR = 50
Public Const STRING_MAPBLOQGEN_DESCRICAO = 50

Public Const STRING_TIPOBLOQGEN_NOMEREDUZIDO = 20
Public Const STRING_TIPOBLOQGEN_DESCRICAO = 100
Public Const STRING_TIPOBLOQGEN_NOMEFUNCTRATATIPO = 50
Public Const STRING_TIPOBLOQGEN_NOMEFUNCGERATIPO = 50
Public Const STRING_TIPOBLOQGEN_NOMEFUNCTRATAGRAVARESERVA = 50

Public Const GERACAONFE_NAO_ENVIADOS = 1
Public Const GERACAONFE_NAO_ACEITOS = 2
Public Const GERACAONFE_AMBOS = 3

Public Const STRING_NFE_CSTAT = 3
Public Const STRING_NFE_NREC = 15


'******* TYPES...

Type typeIPIExcecao
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
End Type

Type typeTipoPedido
    sSigla As String
    sDescricao As String
    iVinculadoNF As Integer
End Type

Type typeFATConfig
    sCodigo As String
    iFilialEmpresa As Integer
    sDescricao As String
    iTipo As Integer
    sConteudo As String
End Type

'Criado por Leo em 06/09/02
Type typePrevVendaMensal
    iFilialEmpresa As Integer
    sCodigo As String
    iAno As Integer
    iCodRegiao As Integer
    lCliente As Long
    iFilial As Integer
    sProduto As String
    dQuantidade1 As Double
    dvalor1 As Double
    dtDataAtualizacao1 As Date
    dQuantidade2 As Double
    dvalor2 As Double
    dtDataAtualizacao2 As Date
    dQuantidade3 As Double
    dvalor3 As Double
    dtDataAtualizacao3 As Date
    dQuantidade4 As Double
    dValor4 As Double
    dtDataAtualizacao4 As Date
    dQuantidade5 As Double
    dValor5 As Double
    dtDataAtualizacao5 As Date
    dQuantidade6 As Double
    dValor6 As Double
    dtDataAtualizacao6 As Date
    dQuantidade7 As Double
    dValor7 As Double
    dtDataAtualizacao7 As Date
    dQuantidade8 As Double
    dValor8 As Double
    dtDataAtualizacao8 As Date
    dQuantidade9 As Double
    dValor9 As Double
    dtDataAtualizacao9 As Date
    dQuantidade10 As Double
    dValor10 As Double
    dtDataAtualizacao10 As Date
    dQuantidade11 As Double
    dValor11 As Double
    dtDataAtualizacao11 As Date
    dQuantidade12 As Double
    dValor12 As Double
    dtDataAtualizacao12 As Date
End Type

Type typePrevVenda
    iFilialEmpresa As Integer
    sCodigo As String
    dtDataPrevisao As Date
    dtDataInicio As Date
    dtDataFim As Date
    iCodRegiao As Integer
    sProduto As String
    iAlmoxarifado As Integer
    dQuantidade As Double
    dValor As Double
End Type


Type typeLiberacaoCredito
    sCodUsuario As String
    dLimiteOperacao As Double
    dLimiteMensal As Double
End Type

Type typeValorLiberadoCredito
    sCodUsuario As String
    iAno As Integer
    adValorLiberado(1 To 12) As Double
End Type

'Type usado para manipulação dos registros de MnemonicoComissoes
Type typeMnemonicoComissoes
    lNumIntDoc As Long
    sMnemonico As String
    sDescricao As String
    sProjetoBrowser As String
    sClasseBrowser As String
    sNomeBrowser As String
    sPropertyBrowser As String
    sGrid As String
    iTipo As Integer
    iNumParam As Integer
    iParam1 As Integer
    iParam2 As Integer
    iParam3 As Integer
    lNumIntDocOrigem As Long
End Type

Type typeCustoDirFabr
    iFilialEmpresa As Integer
    iAno As Integer
    sCodigoPrevVenda As String
    dtData As Date
    dCustoTotal As Double
    dQuantFator1 As Double
    dCustoFator1 As Double
    dQuantFator2 As Double
    dCustoFator2 As Double
    dQuantFator3 As Double
    dCustoFator3 As Double
    dQuantFator4 As Double
    dCustoFator4 As Double
    dQuantFator5 As Double
    dCustoFator5 As Double
    dQuantFator6 As Double
    dCustoFator6 As Double
    iMesIni As Integer
    iMesFim As Integer
End Type

Type typeCustoDirFabrProd
    iFilialEmpresa As Integer
    iAno As Integer
    sProduto As String
    dtData As Date
    dQuantPrevista As Double
    dQuantFator1 As Double
    dQuantFator2 As Double
    dQuantFator3 As Double
    dQuantFator4 As Double
    dQuantFator5 As Double
    dQuantFator6 As Double
End Type

Type typePrevVendaMensal2
    iFilialEmpresa As Integer
    sCodigo As String
    iAno As Integer
    iCodRegiao As Integer
    lCliente As Long
    iFilial As Integer
    sProduto As String
    adQuantidade(1 To 12) As Double
    adValor(1 To 12) As Double
    adtDataAtualizacao(1 To 12) As Date
End Type

Type typeCustoEmbMP
    iFilialEmpresa As Integer
    sProduto As String
    dtDataAtualizacao As Date
    dCusto As Double
    dAliquotaICMS As Double
    dFretePorKg As Double
    iCondicaoPagto As Integer
    iCondicaoPagtoInf As Integer
    iAliquotaICMSInf As Integer
    iFretePorKGInf As Integer
End Type

Public Const CUSTOFIXOPROD_AUTOMATICO = 1
Public Const CUSTOFIXOPROD_MANUAL = 0
Public Const CUSTOFIXOPROD_AUTOMATICO_MANUAL = 2

Type typeCustoFixoProd
    iFilialEmpresa As Integer
    dtDataReferencia As Date
    sProduto As String
    dCusto As Double
    dCustoCalculado As Double
    iAutomatico As Integer
End Type

'TipoFreteFP
Public Const STRING_TIPO_FRETE_NOME_REDUZIDO = 25
Public Const STRING_TIPO_FRETE_DESCRICAO = 50

Type typeTipoFreteFP
    iFilialEmpresa As Integer
    sDescricao As String
    sNomeReduzido As String
    iCodigo As Integer
    dPreco As Double
    dtDataAtualizacao As Date
End Type

Public Const NUM_MAX_PRODUTOS_DVVCLIENTE = 200

Type typeDVVCliente
    iFilialEmpresa As Integer
    lCodCliente As Long
    iCodFilial As Integer
    iTipoFrete As Integer
End Type

Type typeDVVClienteProd
    iFilialEmpresa As Integer
    lCodCliente As Long
    iCodFilial As Integer
    sProduto As String
    dPercDVV As Double
    iPaletizacao As Integer
End Type

Public Const FORMACAO_PRECO_ROTINA_CUSTOSDIRETOS = 1
Public Const FORMACAO_PRECO_ROTINA_CUSTOFIXO = 2
Public Const FORMACAO_PRECO_ROTINA_CALCPRECO = 3
Public Const FORMACAO_PRECO_ANALISE_MARGCONTR = 4
Public Const FORMACAO_PRECO_ANALISE_MARGCONTR_REL = 5
Public Const FORMACAO_PRECO_REL_COMP_CONSUMO = 6


Public Const FORMACAO_PRECO_QTDECALCPRECO = 1000

Public Const CUSTO_DIRETO_PRODUCAO_LOCAL = 1
Public Const CUSTO_DIRETO_COMPRA_LOCAL = 2
Public Const CUSTO_DIRETO_PRODUCAO_TRANSFERIDA = 3
Public Const CUSTO_DIRETO_COMPRA_TRANSFERIDA = 4

Public Const STRING_VEICULO_DESCRICAO = 100
Public Const STRING_VEICULO_PLACA = 20
Public Const STRING_VEICULO_PLACA_UF = 2

Public Const STRING_MAPA_RESPONSAVEL = 100

'****** CAMPOSGENERICOS / CAMPOSGENERICOSVALORES **********
Type typeCamposGenericos
    lCodigo As Long
    sDescricao As String
    sComentarios As String
    lProxCodValor As Long
    sValidaExclusao As String
End Type

Type typeCamposGenericosValores
    lCodCampo As Long
    lCodValor As Long
    iPadrao As Integer
    sValor As String
    sComplemento1 As String
    sComplemento2 As String
    sComplemento3 As String
    sComplemento4 As String
    sComplemento5 As String
End Type
'*********************************************************

'Incluido por Jorge Specian
'-----------------------------------------
Type typeProjeto
    lNumIntDoc As Long
    sNomeReduzido As String
    sDescricao As String
    lCodigo As Long
    lCodCliente As Long
    iCodFilial As Integer
    sResponsavel As String
    dtDataCriacao As Date
    dtDataValidade As Date
    sObservacao As String
End Type

Type typeProjetoItens
    lNumIntDoc As Long
    lNumIntDocProj As Long
    iSeq As Integer
    sProduto As String
    sVersao As String
    sUMedida As String
    dQuantidade As Double
    dtDataMaxTermino As Date
    dtDataInicioPrev As Date
    dtDataTerminoPrev As Date
    iDestino As Integer
    dCustoTotalItem As Double
    dPrecoTotalItem As Double
    lNumIntDocCusteio As Long
End Type

Type typeProjetoItensRegGerados
    lNumIntDoc As Long
    lNumIntDocItemProj As Long
    lNumIntDocDestino As Long
    iDestino As Integer
End Type
'-----------------------------------------

'########################
'Inserido por Wagner
'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeCustoEmbMPAux
    iFilialEmpresa As Integer
    sProduto As String
    sMnemonico As String
    sValor As String
End Type
'########################

'########################
'Inserido por Wagner 17/05/2006
'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeKitVenda
    sProduto As String
    sUM As String
    dQuantidade As Double
    dtData As Date
    sObservacao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeProdutoKitVenda
    sProdutoKit As String
    sProduto As String
    iSeq As Integer
    dQuantidade As Double
    sUM As String
End Type
'########################

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRPS
    lNumIntDoc As Long
    sSerie As String
    sTipo As String
    lNumero As Long
    dtDataEmissao As Date
    sSituacao As String
    dValorServico As Double
    dValorDeducao As Double
    lCodigoServico As Long
    dAliquota As Double
    iISSRetido As Integer
    iTipoCGC As Integer
    sCgc As String
    sInscricaoMunicipal As String
    sInscricaoEstadual As String
    sRazaoSocial As String
    sEndereco As String
    sEndNumero As String
    sEndComplemento As String
    sBairro As String
    sCidade As String
    sUF As String
    sCEP As String
    sEmail As String
    sDiscriminacao As String
    iFilialEmpresa As Integer
    lNumIntDocCab As Long
    lNumIntDocNF As Long
    iFilialCliente As Integer
    lCliente As Long
    dValorCofins As Double
    dValorCSLL As Double
    dValorINSS As Double
    dValorIRPJ As Double
    dValorPIS As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRPSCab
    lNumIntDoc As Long
    sNomeArquivo As String
    dtDataGeracao As Date
    dHoraGeracao As Double
    sUsuario As String
    sVersao As String
    lInscricaoMunicipal As Long
    dtDataInicio As Date
    dtDataFim As Date
End Type


'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeNFe
    lNumIntDoc As Long
    lNumIntDocNF As Long
    lNumNFe As Long
    dtData As Date
    dHora As Double
    sCodVerificacao As String
    sTipoRPS As String
    sSerieRPS As String
    lNumeroRPS As Long
    dtDataEmissaoRPS As Date
    sInscricaoMunicialPrest As String
    iTipoCGCPrest As Integer
    sCGCPrest As String
    sRazaoSocialPrest As String
    sTipoEnderecoPrest As String
    sEnderecoPrest As String
    sEndNumeroPrest As String
    sEndComplementoPrest As String
    sBairroPrest As String
    sCidadePrest As String
    sUFPrest As String
    sCEPPrest As String
    sEmailPrest As String
    iOPTSimples As Integer
    sSituacaoNF As String
    dtDataCancelamento As Date
    sNumGuia As String
    dtDataQuitacaoGuia As Date
    dValorServicos As Double
    dValorDeducoes As Double
    lCodServico As Long
    dAliquota As Double
    dValorISS As Double
    dValorCredito As Double
    sISSRetido As String
    iTipoCGCTom As Integer
    sCGCTom As String
    sInscricaoMunicipalTom As String
    sInscricaoEstadualTom As String
    sRazaoSocialTom As String
    sTipoEnderecoTom As String
    sEnderecoTom As String
    sEndNumeroTom As String
    sEndComplementoTom As String
    sBairroTom As String
    sCidadeTom As String
    sUFTom As String
    sCEPTom As String
    sEmailTom As String
    sDiscriminacao As String
    iFilialEmpresa As Integer
    dValorCofins As Double
    dValorCSLL As Double
    dValorINSS As Double
    dValorIRPJ As Double
    dValorPIS As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeNFeCab
    lNumIntDoc As Long
    sNomeArquivo As String
    dtDataImportacao As Date
    dHoraImportacao As Double
    sUsuario As String
    sVersao As String
    lInscricaoMunicipal As Long
    dtDataInicio As Date
    dtDataFim As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeMapeamentoBloqGen
    iTipoTelaBloqueio As Integer
    iClassePossuiFilEmp As Integer
    sNomeTelaEditaDocBloq As String
    sNomeClasseDocBloq As String
    sClasseNomeCampoChave As String
    sNomeBrowseChave As String
    sNomeTabelaBloqueios As String
    sNomeTelaTestaPermissao As String
    sNomeFuncLiberaCust As String
    sProjetoClasseDocBloq As String
    sNomeViewLeBloqueios As String
    sTabelaNomeCampoChave As String
    iTabelaBloqPossuiTipoTela As Integer
    sNomeFuncLeDoc As String
    sNomeColecaoBloqDoc As String
    sClasseDocNomeQTD As String
    sClasseDocNomeQTDReservada As String
    sClasseDocNomeUM As String
    sClasseDocNomeColItem As String
    iClasseDocQTDNoItem As Integer
    sClasseDocNomeProduto As String
    sClasseDocNomeValor As String
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeTiposDeBloqueioGen
    iTipoTelaBloqueio As Integer
    iCodigo As Integer
    sNomeReduzido As String
    sDescricao As String
    sNomeFuncTrataTipo As String
    iNaoApareceTelaLib As Integer
    sNomeFuncGeraTipo As String
    iTestaValorAlteracao As Integer
    iAlteracaoForcaInclusao As Integer
    iBloqueioTotal As Integer
    sNomeFuncTrataGravaReserva As String
    iBloqueioReserva As Integer
    iInterno As Integer
End Type

Type typeBloqueioGen
    iTipoTelaBloqueio As Integer
    iFilialEmpresa As Integer
    lCodigo As Long
    iSequencial As Integer
    iTipoDeBloqueio As Integer
    sCodUsuario As String
    sResponsavel As String
    dtData As Date
    sCodUsuarioLib As String
    sResponsavelLib As String
    dtDataLib As Date
    sObservacao As String
    dtDataEmissaoDoc As Date
    sNomeClienteDoc As String
    dValorDoc As Double
    lClienteDoc As Long
    sNomeTipoDeBloqueio As String
    dValorDocAnt As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeISSQN
    sCodigo As String
    sDescricao As String
    iTipo As Integer
    lCodServNFe As Long
End Type

Type typeProdutoGenero
    sCodigo As String
    sDescricao As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeEmailConfig
    sUsuario As String
    sSMTP As String
    sSMTPUsu As String
    sSMTPSenha As String
    lSMTPPorta As Long
    iSSL As Integer
    iConfirmacaoLeitura As Integer
    iPgmEmail As Integer
    sEmail As String
    sNome As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRotas
    lNumIntDoc As Long
    sCodigo As String
    lChave As Long
    sChaveValor As String
    sChaveNome As String
    iFilialEmpresa As Integer
    sDescricao As String
    iAtivo As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRotasVend
    lNumIntDoc As Long
    lNumIntDocRota As Long
    iSeq As Integer
    iVendedor As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeRotasPontos
    lNumIntDoc As Long
    lNumIntDocRota As Long
    iSeq As Integer
    lCliente As Long
    sNomeCliente As String
    sObservacao As String
    lMeio As Long
    dTempo As Double
    dDistancia As Double
    iSelecionado As Integer
    iFilialCliente As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeVeiculos
    lCodigo As Long
    sDescricao As String
    lTipo As Long
    iProprio As Integer
    sPlaca As String
    sPlacaUF As String
    dCapacidadeKg As Double
    dVolumeM3 As Double
    dCustoHora As Double
    dDispPadraoDe As Double
    dDispPadraoAte As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeMapaDeEntrega
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    dtData As Date
    iRegiao As Integer
    lVeiculo As Long
    dVolumeTotal As Double
    dPesoTotal As Double
    iNumViagens As Integer
    dHoraSaida As Double
    dHoraRetorno As Double
    sResponsavel As String
    iTipoDoc As Integer
    iTransportadora As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeMapaDeEntregaDoc
    lNumIntDocMapa As Long
    lNumIntDoc As Long
    lSeq As Long
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeImportTabelas
    lCodigo As Long
    sTabela As String
    sDescricao As String
    sFuncaoGrava As String
    sFuncaoValida As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeImportTabelasCampos
    lTabela As Long
    lCodigo As Long
    sCampo As String
    sNomeExibicao As String
    iTipo As Integer
    iChave As Integer
    iExibe As Integer
    sNomeIgual1 As String
    sNomeIgual2 As String
    sNomeIgual3 As String
    sNomeLike1 As String
    sNomeLike2 As String
    sNomeLike3 As String
    sValorPadrao As String
    iObrigatorio As Integer
    iTamMax As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeDEInfo
    lNumIntDoc As Long
    sNumero As String
    dtData As Date
    iFilialEmpresa As Integer
    sDescricao As String
    iTipoDoc As Integer
    iNatureza As Integer
    sNumConhEmbarque As String
    sUFEmbarque As String
    sLocalEmbarque As String
    dtDataConhEmbarque As Date
    iTipoConhEmbarque As Integer
    iCodPais As Integer
    dtDataAverbacao As Date
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeDERegistro
    lNumIntDE As Long
    sNumRegistro As String
    dtDataRegistro As Date
End Type

Public Function ColecaoDef_Trans_Collection(ByVal colColecaoDef As Variant) As Collection

Dim vVar As Variant
Dim colAux As New Collection

    For Each vVar In colColecaoDef
        colAux.Add vVar
    Next
    
    Set ColecaoDef_Trans_Collection = colAux

End Function

Public Function Collection_Trans_ColecaoDef(ByVal colCollection As Collection, ByVal colColecaoDef As Variant) As Long

Dim vVar As Object
Dim colAux As New Collection

    For Each vVar In colColecaoDef
        colColecaoDef.Remove 1
    Next
    
    If UCase(TypeName(colColecaoDef)) <> "COLLECTION" Then
        For Each vVar In colCollection
            colColecaoDef.AddObj vVar
        Next
    Else
        For Each vVar In colCollection
            colColecaoDef.Add vVar
        Next
    End If
    
    Collection_Trans_ColecaoDef = SUCESSO

End Function

Public Function Colecao_Altera_Ordem_Itens(ByVal colColecao As Collection, ByVal iLinhaAtual As Integer, ByVal iLinhaNova As Integer) As Long

Dim iIndice As Integer
Dim vVar As Object
Dim colAux As New Collection

    iIndice = 0
    For Each vVar In colColecao
        iIndice = iIndice + 1
        If iIndice <> iLinhaAtual And iIndice <> iLinhaNova Then
            colAux.Add vVar
        Else
            If iIndice = iLinhaNova Then
                colAux.Add colColecao.Item(iLinhaAtual)
            Else
                colAux.Add colColecao.Item(iLinhaNova)
            End If
        End If
    Next
    
    For iIndice = colColecao.Count To 1 Step -1
        colColecao.Remove iIndice
    Next
    
    For Each vVar In colAux
        colColecao.Add vVar
    Next
    
    Colecao_Altera_Ordem_Itens = SUCESSO

End Function

Public Function Compara_Telefone(ByVal sTelefone1 As String, ByVal sTelefone2 As String) As Boolean

Dim sTel1 As String
Dim sTel2 As String

    sTel1 = Replace(Replace(Replace(Replace(sTelefone1, ")", ""), "(", ""), "-", ""), " ", "")
    sTel2 = Replace(Replace(Replace(Replace(sTelefone2, ")", ""), "(", ""), "-", ""), " ", "")

    If sTel1 = sTel2 Then
        Compara_Telefone = True
    Else
        Compara_Telefone = False
    End If

End Function

Public Function Compara_Contato(ByVal sContato1 As String, ByVal sContato2 As String) As Boolean

    If UCase(sContato1) = UCase(sContato2) Or Len(Trim(sContato1)) = 0 Or Len(Trim(sContato2)) = 0 Or sContato1 = "Sem nome" Or sContato2 = "Sem nome" Then
'    If UCase(sContato1) = UCase(sContato2) Or (Len(Trim(sContato1)) = 0 And sContato2 = "Sem nome") Or (Len(Trim(sContato2)) = 0 And sContato1 = "Sem nome") Then
        Compara_Contato = True
    Else
        Compara_Contato = False
    End If

End Function
