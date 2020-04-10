Attribute VB_Name = "GlobalMAT"
Option Explicit

'###################################################
'Inserido por Wagner 12/06/2006
Public Const PRODUTOFORNECEDOR_ORIGEM_PRODUTO = 0
Public Const PRODUTOFORNECEDOR_ORIGEM_FORNECEDOR = 1

Public Const PRODUTOFORNECEDOR_ACAO_ADICIONAR = 0
Public Const PRODUTOFORNECEDOR_ACAO_SOBREPOR = 1
'###################################################

Public Const NUM_MAX_TESTES_PRODUTO = 50

Public Const BENEF_COMBO_DISP = 0
Public Const BENEF_COMBO_BENEF = 1
Public Const BENEF_COMBO_OUTROS = 2

Public Const FORMATO_LIMITE_TESTE = "###,###,##0.0####"

Public Const FORMATO_KIT_PESO_FATOR = "###,###,##0.0#####"

Public Const OP_TIPO_CALC_NECESS_USA_QUANTOP = 0
Public Const OP_TIPO_CALC_NECESS_NAO_USA_QUANTOP = 1
Public Const OP_TIPO_CALC_NECESS_QUANTOP_ABATE_EMPENHO = 2

Public Const STRING_TESTEQUALIDADE_NOMERED = 100
Public Const STRING_TESTEQUALIDADE_ESPECIFICACAO = 250
Public Const STRING_TESTEQUALIDADE_METODOUSADO = 50
Public Const STRING_TESTEQUALIDADE_OBSERVACAO = 250

Public Const STRING_RESULTADO_ANALISE_ID = 20
Public Const STRING_RESULTADO_VALOR = 250
Public Const STRING_RESULTADO_OBS = 250

Public Const NAO_ATUALIZAR_TABELA_PADRAO = -1

Public Const OP_TIPO_OP = 0 'indica se esta lidando com uma ordem de producao
Public Const OP_TIPO_OC = 1 'indica se esta lidando com uma ordem de corte

Public Const ESTCFG_VALIDA_PRODUTO_BASE_CARGA = "VALIDA_PRODUTO_BASE_CARGA"

Public Const KIT_SITUACAO_PADRAO = 1
Public Const KIT_SITUACAO_ATIVO = 2
Public Const KIT_SITUACAO_INATIVO = 3

Public Const KIT_SITUACAO_STRING_PADRAO = "Padrão"
Public Const KIT_SITUACAO_STRING_ATIVO = "Ativo"
Public Const KIT_SITUACAO_STRING_INATIVO = "Inativo"

'Constantes dos tipos de log produto -> implementacao para o loja
'incluidas por tulio em 5/6/02
Public Const INCLUSAO_PRODUTO_CAIXA_CENTRAL = 16
Public Const EXCLUSAO_PRODUTO_CAIXA_CENTRAL = 17
Public Const ALTERACAO_PRODUTO_CAIXA_CENTRAL = 18

'indica que deve exibir mensagem de erro caso a quantidade que deveria ser apropriada não coincida com a quantidade a apropriar
Public Const MOSTRA_MENSAGEM_ERRO_QUANTAPROPRIADA = 1

'Indica que o registro na tabela de ApropriacaoInsumosProd foi criado pelo sistema.
Public Const APROPINSUMOSPROD_AUTOMATICO = 1

Public Const CAMPOINICIAL_INEXISTENTE As String = "CampoInicial_Inexistente"

Public Const STRING_FORNECEDORPRODUTO_DESCRICAO = 150

Public Const INVENTARIO_ATUALIZA_LOTE_ESTOQUE = 0
Public Const INVENTARIO_ATUALIZA_SO_LOTE = 1

'Declaracao das constantes utilizadas em ProdutoEmbalagem
Public Const STRING_PRODUTOEMBALAGEM_PRODUTO = 20
Public Const STRING_PRODUTOEMBALAGEM_NOMEPRODEMB = 20
Public Const STRING_PRODUTOEMBALAGEM_UMEMBALAGEM = 5
Public Const STRING_PRODUTOEMBALAGEM_UMPESO = 5
Public Const PRODUTOEMBALAGEM_PADRAO = 1
Public Const UM_PESO_CLASSE = 1

'#####################
'Inserido por Wagner
Public Const STRING_MOVESTOQUE_OBSERVACAO = 255
'#####################

Public Const TIPO_FATURAMENTO = 1
Public Const TIPO_FATURAMENTO_DEVOLUCAO = 2

Public Const ESCANINHO_RASTRO_ESTOQUE_INICIAL = 1
Public Const ESCANINHO_SEM_RASTRO_ESTOQUE_INICIAL = 0
Public Const STRING_NOME_ESCANINHO = 50
Public Const ESCANINHO_DISPONIVEL = 1
Public Const ESCANINHO_CONSERTO_NOSSO = 2
Public Const ESCANINHO_CONSIG_NOSSO = 3
Public Const ESCANINHO_DEMO_NOSSO = 4
Public Const ESCANINHO_OUTROS_NOSSO = 5
Public Const ESCANINHO_BENEF_NOSSO = 6
Public Const ESCANINHO_CONSERTO_3 = 7
Public Const ESCANINHO_CONSIG_3 = 8
Public Const ESCANINHO_DEMO_3 = 9
Public Const ESCANINHO_OUTROS_3 = 10
Public Const ESCANINHO_BENEF_3 = 11

Public Const RASTRO_ADICAO = 0
Public Const RASTRO_SUBTRACAO = 1

'Reprocessamento
Public Const REPROCESSAMENTO_ORDENA_HORA = 1
Public Const REPROCESSAMENTO_ORDENA_ENTRADAS = 2
Public Const REPROCESSAMENTO_NAO_GERA_LOG = 0
Public Const REPROCESSAMENTO_GERA_LOG = 1
Public Const REPROCESSAMENTO_NORMAL = 1
Public Const REPROCESSAMENTO_PULA_DESFAZ = 2
Public Const REPROCESSAMENTO_ACERTA_ESTPROD = 1
Public Const REPROCESSAMENTO_NAO_ACERTA_ESTPROD = 0

'Constantes da modificação
Public Const STRING_EMBALAGEM_SIGLA = 20
Public Const STRING_EMBALAGEM_DESCRICAO = 50
Public Const STRING_FIGURA = 80
Public Const STRING_REFERENCIA = 20
Public Const STRING_TIPOSTRIBICMS_DESCRICAO = 50

'para a tabela ImportProdDesc
Public Const STRING_FFL_R = 20
'Para a tabela ImportProdAux
Public Const STRING_CODIGO_IMPORTPRODAUX = 5

Public Const INCLUSAO_MOVIMENTO = 1
Public Const EXCLUSAO_MOVIMENTO = 2
Public Const REPROCESSAMENTO_DESFAZ = 3
Public Const REPROCESSAMENTO_REFAZ = 4
Public Const ALTERACAO_MOVIMENTO_FASE_EXCLUSAO = 5
Public Const ALTERACAO_MOVIMENTO_FASE_INCLUSAO = 6
Public Const REPROCESSAMENTO_TESTA_INTEGRIDADE = 7
Public Const APURACAO_CUSTO_PRODUCAO = 8

'Rastreamento
Public Const PRODUTO_RASTRO_NENHUM = 0
Public Const PRODUTO_RASTRO_LOTE = 1
Public Const PRODUTO_RASTRO_OP = 2
Public Const PRODUTO_RASTRO_NUM_SERIE = 3

Public Const TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE = 0
Public Const TIPO_RASTREAMENTO_MOVTO_APROP_PRODUCAO_ENTRADA = 1

'Tipo de Origem para NumIntDocOrigem de Movimento Estoque
Public Const TIPO_ORIGEM_ITEMNF = 1

Public Const INTERVALO_INSUMO_INIC_PRODUCAO = "INTERVALO_INSUMO_INIC_PRODUCAO"

Public Const RATREAMENTOLOTE_STATUS_ABERTO = 0
Public Const RATREAMENTOLOTE_STATUS_BAIXADO = 1

'Limites versão light
Public Const LIMITE_ALMOX_VLIGHT = 5
Public Const LIMITE_FORN_PRODUTO_VLIGHT = 5
Public Const LIMITE_CLASSE_UM_VGLIGHT = 5
Public Const LIMITE_CCUSTO_VLIGHT = 5

Public Const POSSUI_PRODUTO = 1
Public Const NAO_POSSUI_PRODUTO = 0

'Constantes de Campo (tela de Comissoes)
Public Const BROWSER = 2
Public Const CANCELA_BROWSER = 3
Public Const ANTIGO_NOTA_FISCAL = 4
Public Const ANTIGO_PEDIDO_DE_VENDA = 5

Public Const RECALCULAR_PEDIDOVENDA = 1
Public Const NAO_RECALCULAR_PEDIDOVENDA = 0

Public Const FECHAR_SETAS = 1
Public Const NAO_FECHAR_SETAS = 0

'Lote Est Atualiza
Public Const LOTES_PENDENTES = 1

'Movimento Estoque
Public Const MOVESTOQUEINT = 1
Public Const MOVESTOQUEINTSRV = 2
Public Const TIPO_ENTRADA = "E"
Public Const TIPO_SAIDA = "S"

Public Const ESTOQUE_SAIDA = 0 'Indica saida do material do estoque
Public Const ESTOQUE_ENTRADA = 1 'Indica entrada do material no estoque

'KIT
Public Const NIVEL_MAXIMO_KIT = 1
Public Const COMPOSICAO_VARIAVEL = "Variável"

'Geração O.P.
Public Const S_MARCADO As String = "1"
Public Const S_DESMARCADO As String = "0"

Public Const STATUS_NAO_ATENDIDO = 0
Public Const STATUS_ATENDIDO = 1

'Falta Estoque
Public Const RESERVA_AUTO_RESP = "RESERVA AUTOMÁTICA"
Public Const TR_CANCELA = 1
Public Const TR_CANCELA_ULTR_DISP = 2
Public Const TR_ALOC_MANUAL = 3
Public Const TR_NAO_RESERVA = 4
Public Const TR_RESERVA_EST = 5
Public Const TR_SUBSTITUI = 6

Public Const FORN_PROD_NAO_PADRAO = 0
Public Const FORN_PROD_PADRAO = 1

Public Const Padrao = 1
Public Const NAO_PADRAO = 0

'Classificação ABC
Public Const ANO_VALIDO = 1900

'CurvaABC
'altera a granularidade da curva ABC
Public Const CURVA_ABC_MAX_PONTOS = 100

'ClassifABC atualiza ProdutosFilial
Public Const CLASSABC_ATUALIZA_PRODFILIAL = 1

'Tipo de Produto
Public Const TODOS_TIPOS = 0

'Indica o numero de lComandos do array de Movimentacao de Estoque
Public Const NUM_MAX_LCOMANDO_MOVESTOQUE = 107

'Tela ReqConsumo
'Public Const NUM_MAX_ITENS_MOV_ESTOQUE = 100
Public Const MOV_REQCONSUMO = 1
Public Const MOVIMENTO_NOVO = 1

'Tela RecebMaterialC e RecebMaterialF
Public Const NUM_MAX_ITENS_RECEB = 700

'Nota Fiscal
Public Const NUM_MAX_ITENS_NF = 700

'Public Const TIPO_ALMOXARIFADO_NORMAL = 0

Public Const TIPOMOV_EST_VALIDADATAULTMOV = 1 'Valida a data do movimento para que não seja menor do que a data do ultimo movimento cadastrado
Public Const TIPOMOV_EST_NAO_VALIDADATAULTMOVEST = 0 'Não Valida a data do movimento

'para tabela FornecedorProduto
Public Const STRING_PRODUTO_FORNECEDOR = 20 'tamanho máximo do codigo do produto para um fornecedor

Public Const STRING_FORNECEDOR_PRODUTO_DESCRICAO = 150

'Inventario
Public Const NUM_MAX_ITENS_INVENTARIO = 100
Public Const NUM_GRUPO_ITENS_INVENTARIO = 100 'incremento de linhas a serem criadas

'Constante que indica o numero maximo de codigo de barras
'que podem existir no grid
Public Const NUM_MAX_CODBARRAS_PRODUTO = 50

'Naturezas de Produto
Public Const NATUREZA_PROD_MATERIA_PRIMA = 1
Public Const NATUREZA_PROD_PRODUTO_INTERMEDIARIO = 2
Public Const NATUREZA_PROD_EMBALAGENS = 3
Public Const NATUREZA_PROD_PRODUTO_ACABADO = 4
Public Const NATUREZA_PROD_PRODUTO_REVENDA = 5
Public Const NATUREZA_PROD_PRODUTO_MANUTENCAO = 6
Public Const NATUREZA_PROD_OUTROS = 7
Public Const NATUREZA_PROD_SERVICO = 8
Public Const NATUREZA_PROD_PRODUTO_EM_PROCESSO = 9
Public Const NATUREZA_PROD_SUBPRODUTO = 10
Public Const NATUREZA_PROD_MATERIAL_DE_USO_E_CONSUMO = 11
Public Const NATUREZA_PROD_ATIVO_IMOBILIZADO = 12
Public Const NATUREZA_PROD_OUTROS_INSUMOS = 13

Public Const NATUREZA_PROD_MATERIA_PRIMA_COD_OFICIAL = 1
Public Const NATUREZA_PROD_PRODUTO_INTERMEDIARIO_COD_OFICIAL = 6
Public Const NATUREZA_PROD_EMBALAGENS_COD_OFICIAL = 2
Public Const NATUREZA_PROD_PRODUTO_ACABADO_COD_OFICIAL = 4
Public Const NATUREZA_PROD_PRODUTO_REVENDA_COD_OFICIAL = 0
Public Const NATUREZA_PROD_PRODUTO_MANUTENCAO_COD_OFICIAL = 99
Public Const NATUREZA_PROD_OUTROS_COD_OFICIAL = 99
Public Const NATUREZA_PROD_SERVICO_COD_OFICIAL = 9
Public Const NATUREZA_PROD_PRODUTO_EM_PROCESSO_COD_OFICIAL = 3
Public Const NATUREZA_PROD_SUBPRODUTO_COD_OFICIAL = 5
Public Const NATUREZA_PROD_MATERIAL_DE_USO_E_CONSUMO_COD_OFICIAL = 7
Public Const NATUREZA_PROD_ATIVO_IMOBILIZADO_COD_OFICIAL = 8
Public Const NATUREZA_PROD_OUTROS_INSUMOS_COD_OFICIAL = 10

'ProdutoKit
Public Const KIT_NIVEL_RAIZ = 0   'Nível do Produto Raiz no Kit
Public Const PRODUTOKIT_COMPOSICAO_FIXA = 0
Public Const PRODUTOKIT_COMPOSICAO_VARIAVEL = 1

'Nas telas de Movimento de Estoque determina se é ESTORNO
'Retorno da sub MovimentoEstorno(iMovimento)
Public Const MOVIMENTO_NORMAL = 0
Public Const MOVIMENTO_ESTORNO = 1

'ItemOP - Situacao
Public Const ITEMOP_SITUACAO_NORMAL = 0
Public Const ITEMOP_SITUACAO_DESAB = 1
Public Const ITEMOP_SITUACAO_SACR = 2
Public Const ITEMOP_SITUACAO_BAIXADA = 3
Public Const ITEMOP_SITUACAO_PLANEJADA = 4

'ItemOP - Destinação
Public Const ITEMOP_DESTINACAO_ESTOQUE = 0
Public Const ITEMOP_DESTINACAO_PV = 1
Public Const ITEMOP_DESTINACAO_CONSUMO = 2

'CAMPOS de TIPOSDOCINFO

'Campos Emitente, Destinatario, Origem
Public Const DOCINFO_EMPRESA = 0
Public Const DOCINFO_CLIENTE = 1
Public Const DOCINFO_FORNECEDOR = 2

'Janaina
Public Const DOCORIGEM_PV = 1
Public Const DOCORIGEM_NF = 2

Public Const NUM_MAX_EMBALAGENS = 500
'Janaina

'Tipo especifico de um codigo na tabela MATConfig
Public Const NUM_PROX_EMPENHO = "NUM_PROX_EMPENHO"
Public Const NUM_PROX_ITEM_MOV_ESTOQUE = "NUM_PROX_ITEM_MOV_ESTOQUE"
Public Const NUM_PROX_MOV_ESTOQUE = "NUM_PROX_MOV_ESTOQUE"
Public Const NUM_PROX_RESERVA = "NUM_PROX_RESERVA"
Public Const NUM_PROX_INT_RESERVA = "NUM_PROX_INT_RESERVA"
Public Const NUM_PROX_ITEM_INVENTARIO = "NUM_PROX_ITEM_INVENTARIO"
Public Const NUM_PROX_INVENTARIO = "NUM_PROX_INVENTARIO"
Public Const NUM_PROX_LOTE_INVENTARIO = "NUM_PROX_LOTE_INVENTARIO"
Public Const NUM_PROX_RECEBIMENTO = "NUM_PROX_RECEBIMENTO"
Public Const NUM_PROX_RASTREAMENTOMOVTO = "NUM_PROX_RASTREAMENTOMOVTO"
Public Const NUM_PROX_RASTREAMENTOLOTE = "NUM_PROX_RASTREAMENTOLOTE"
Public Const NUM_PROX_CODIGO_RASTREAMENTOLOTE = "NUM_PROX_CODIGO_RASTREAMENTOLOTE"
Public Const NUM_PROX_APROPRIACAOINSUMOSPROD = "NUM_PROX_APROPRIACAOINSUMOSPROD"
Public Const DATA_REPROCESSAMENTO = "DATA_REPROCESSAMENTO"
Public Const DATA_REPROCESSAMENTO_DESCR = "Data a partir da qual necessita realizar a rotina de reprocessamento"
Public Const HABILITA_FIFO_NFISCAIS = "HABILITA_FIFO_NFISCAIS"
Public Const DATAINICIO_ULTIMO_REPROC = "DATAINICIO_ULTIMO_REPROC"
Public Const DATAINICIO_ULTIMO_REPROC_DESCR = "Data a partir da qual iniciou-se o último reprocessamento feito"

Public Const HABILITA_FIFO_NF = 1
Public Const NAO_HABILITA_FIFO_NF = 0

'Declaração das constantes de String
Public Const STRING_RESPONSAVEL_RESERVA = 50 'Responsavel pela reserva

'tipos de movimento de estoque
Public Const MOV_EST_NF_VENDA = 1
Public Const MOV_EST_NF_DEV_VENDA = 2
Public Const MOV_EST_NF_FORNECEDOR = 3
Public Const MOV_EST_NF_ENTRADA = 4
Public Const MOV_EST_NF_DEV_COMPRA = 5
Public Const MOV_EST_REQ_PRODUCAO = 6
Public Const MOV_EST_PRODUCAO = 7
Public Const MOV_EST_CONSUMO = 8
Public Const MOV_EST_ENTRADA_TRANSF_DISP = 9
Public Const MOV_EST_SAIDA_TRANSF_DISP = 10
Public Const MOV_EST_ACRESCIMO_INVENT_DISPONIVEL_NOSSA = 11
Public Const MOV_EST_DECRESCIMO_INVENT_DISPONIVEL_NOSSA = 12
Public Const MOV_EST_NF_ENTRADA_OUTRAS_REMESSAS = 13
Public Const MOV_EST_DEV_CONSUMO = 14
Public Const MOV_EST_DEV_MATERIAL_PRODUCAO = 15
Public Const MOV_EST_ESTORNO_PRODUCAO = 16
Public Const MOV_EST_ESTORNO_CONSUMO = 17
Public Const MOV_EST_ESTORNO_DEV_CONSUMO = 18
Public Const MOV_EST_ESTORNO_REQ_PRODUCAO = 19
Public Const MOV_EST_ESTORNO_DEV_MATERIAL_PRODUCAO = 20
Public Const MOV_EST_PROCESSO = 21
Public Const MOV_EST_DEV_PROCESSO = 22
Public Const MOV_EST_ESTORNO_PROCESSO = 23
Public Const MOV_EST_ENTRADA_TRANSF_DEFEIT = 24
Public Const MOV_EST_SAIDA_TRANSF_DEFEIT = 25
Public Const MOV_EST_ENTRADA_TRANSF_INDISP = 26
Public Const MOV_EST_SAIDA_TRANSF_INDISP = 27
Public Const MOV_EST_NF_ENTRADA_CONSIG = 28
Public Const MOV_EST_NF_VENDA_MAT_CONSIG = 29
Public Const MOV_EST_ACRESCIMO_INVENT_RECEB_INDISP = 30
Public Const MOV_EST_NF_SAIDA_DEV_CONSIG = 31
Public Const MOV_EST_NF_SAIDA_REMESSA_CONSIG = 32
Public Const MOV_EST_NF_ENTRADA_DEMO = 33
Public Const MOV_EST_NF_ENTRADA_DEV_DEMO = 34
Public Const MOV_EST_NF_SAIDA_REMESSA_DEMO = 35
Public Const MOV_EST_NF_SAIDA_DEV_DEMO = 36
Public Const MOV_EST_NF_ENTRADA_CONSERTO = 37
Public Const MOV_EST_NF_ENTRADA_DEV_CONSERTO = 38
Public Const MOV_EST_NF_SAIDA_REMESSA_CONSERTO = 39
Public Const MOV_EST_NF_SAIDA_DEV_CONSERTO = 40
Public Const MOV_EST_ESTORNO_DEV_PROCESSO = 41
Public Const MOV_EST_NF_SAIDA_OUTRAS_REMESSAS_TERC = 42
Public Const MOV_EST_NF_SAIDA_OUTRAS_DEV_MAT_TERC = 43
Public Const MOV_EST_ACRESCIMO_INVENT_DEFEITUOSO = 44
Public Const MOV_EST_DECRESCIMO_INVENT_DEFEITUOSO = 45
Public Const MOV_EST_ACRESCIMO_INVENT_CONSERTO_TERC = 46
Public Const MOV_EST_DECRESCIMO_INVENT_CONSERTO_TERC = 47
Public Const MOV_EST_ACRESCIMO_INVENT_CONSIG_TERC = 48
Public Const MOV_EST_DECRESCIMO_INVENT_CONSIG_TERC = 49
Public Const MOV_EST_ACRESCIMO_INVENT_DEMO_TERC = 50
Public Const MOV_EST_DECRESCIMO_INVENT_DEMO_TERC = 51
Public Const MOV_EST_ACRESCIMO_INVENT_OUTROS_TERC = 52
Public Const MOV_EST_DECRESCIMO_INVENT_OUTROS_TERC = 53
Public Const MOV_EST_ACRESCIMO_INVENT_INDISP_OUTRAS = 54
Public Const MOV_EST_DECRESCIMO_INVENT_INDISP_OUTRAS = 55
Public Const MOV_EST_RECEBIMENTO_MATERIAL = 56
Public Const MOV_EST_ESTORNO_ENTRADA_TRANSF_DISPONIVEL = 57
Public Const MOV_EST_ESTORNO_SAIDA_TRANSF_DISPONIVEL = 58
Public Const MOV_EST_ESTORNO_ENTRADA_TRANSF_DEFEITUOSO = 59
Public Const MOV_EST_ESTORNO_SAIDA_TRANSF_DEFEITUOSO = 60
Public Const MOV_EST_ESTORNO_ENTRADA_TRANSF_INDISPONIVEL = 61
Public Const MOV_EST_ESTORNO_SAIDA_TRANSF_INDISPONIVEL = 62
Public Const MOV_EST_EXCLUSAO_RECEBIMENTO_MATERIAL = 63
Public Const MOV_EST_EXCLUSAO_NOTA_FISCAL_ENTRADA = 64
Public Const MOV_EST_BAIXA_RECEBIMENTO_MATERIAL = 65
Public Const MOV_EST_NF_ENTRADA_DEV_CONSIG = 73
Public Const MOV_EST_DECRESCIMO_INVENT_RECEB_INDISP = 77
Public Const MOV_EST_ESTORNO_ENTRADA_TRANSF_CONSIG_TERC = 79
Public Const MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_TERC = 80
Public Const MOV_EST_ENTRADA_TRANSF_CONSIG_TERC = 81
Public Const MOV_EST_SAIDA_TRANSF_CONSIG_TERC = 82
Public Const MOV_EST_UTILIZA_RESERVA = 83
Public Const MOV_EST_INCLUI_RESERVA = 84
Public Const MOV_EST_CANCELA_RESERVA = 85
Public Const MOV_EST_ENTRADA_MAT_BENEF_PARA_TERC = 86
Public Const MOV_EST_SAIDA_MAT_BENEF_PARA_TERC = 87
Public Const MOV_EST_MAT_NOSSO_PARA_BENEF_ENTRADA = 90
Public Const MOV_EST_MAT_NOSSO_PARA_BENEF_SAIDA = 91
Public Const MOV_EST_ACRESCIMO_INVENT_BENEF_TERC = 92
Public Const MOV_EST_DECRESCIMO_INVENT_BENEF_TERC = 93
Public Const MOV_EST_ENTRADA_TRANSF_CONSIG_NOSSO = 94
Public Const MOV_EST_SAIDA_TRANSF_CONSIG_NOSSO = 95
Public Const MOV_EST_ESTORNO_ENTRADA_TRANSF_CONSIG_NOSSO = 96
Public Const MOV_EST_ESTORNO_SAIDA_TRANSF_CONSIG_NOSSO = 97
Public Const MOV_EST_CANCELA_BAIXA_RECEBIMENTO_MATERIAL = 99
Public Const MOV_EST_ENTRADA_TRANSF_DISP1 = 100
Public Const MOV_EST_ENTRADA_TRANSF_DISP_CONSIG3 = 122
Public Const MOV_EST_ENTRADA_TRANSF_DISP_CONSIG = 123
Public Const MOV_EST_AJUSTE_CUSTO_STD_NOSSO = 141
Public Const MOV_EST_AJUSTE_CUSTO_STD_CONSIG_NOSSO = 142
Public Const MOV_EST_AJUSTE_CUSTO_STD_DEMO_NOSSO = 143
Public Const MOV_EST_AJUSTE_CUSTO_STD_CONSERTO_NOSSO = 144
Public Const MOV_EST_AJUSTE_CUSTO_STD_OUTROS_NOSSO = 145
Public Const MOV_EST_AJUSTE_CUSTO_STD_BENEF_NOSSO = 146
Public Const MOV_EST_NF_ENTRADA_BENEF3 = 150
Public Const MOV_EST_REQ_PRODUCAO_BENEF3 = 151
Public Const MOV_EST_ESTORNO_REQ_PRODUCAO_BENEF3 = 152
Public Const MOV_EST_PRODUCAO_BENEF3 = 154
Public Const MOV_EST_ESTORNO_PRODUCAO_BENEF3 = 155
Public Const MOV_EST_NF_SAIDA_BENEF3 = 156
Public Const MOV_EST_NF_SAIDA_DEV_BENEF3 = 158
Public Const MOV_EST_DECR_INVENT_BENEF_TERC_SOLOTE = 173
Public Const MOV_EST_DECR_INVENT_DEMO_TERC_SOLOTE = 174
Public Const MOV_EST_DECR_INVENT_DISP_NOSSA_SOLOTE = 175
Public Const MOV_EST_ACRES_INVENT_DISP_NOSSA_SOLOTE = 176
Public Const MOV_EST_ACRES_INVENT_DEMO_TERC_SOLOTE = 177
Public Const MOV_EST_DECR_INVENT_OUTROS_TERC_SOLOTE = 178
Public Const MOV_EST_ACRES_INVENT_CONS_TERC_SOLOTE = 179
Public Const MOV_EST_ACRES_INVENT_IND_OUTRAS_SOLOTE = 180
Public Const MOV_EST_ACRES_INVENT_BENEF_TERC_SOLOTE = 181
Public Const MOV_EST_DECR_INVENT_IND_OUTRAS_SOLOTE = 182
Public Const MOV_EST_ACRES_INVENT_OUTROS_TERC_SOLOTE = 183
Public Const MOV_EST_DECR_INVENT_CONSIG_TERC_SOLOTE = 184
Public Const MOV_EST_ACRES_INVENT_RECEB_IND_SOLOTE = 185
Public Const MOV_EST_DECR_INVENT_CONS_TERC_SOLOTE = 186
Public Const MOV_EST_DECR_INVENT_DEFEITUOSO_SOLOTE = 187
Public Const MOV_EST_ACRES_INVENT_DEFEITUOSO_SOLOTE = 188
Public Const MOV_EST_DECR_INVENT_RECEB_IND_SOLOTE = 189
Public Const MOV_EST_ACRES_INVENT_CONSIG_TERC_SOLOTE = 190
Public Const MOV_EST_SAIDA_TRANSF_BENEF_TERC = 193
Public Const MOV_EST_ESTORNO_SAIDA_TRANSF_BENEF_TERC = 194
Public Const MOV_EST_ENTRADA_TRANSF_DISP4 = 195
Public Const MOV_EST_NF_VENDA_MAT_BENEF3 = 196
Public Const MOV_EST_SAIDA_TRANSF_OUTRAS_TERC = 199
Public Const MOV_EST_ENTRADA_TRANSF_OUTRAS_TERC = 200
Public Const MOV_EST_CUPOM_FISCAL = 201
Public Const MOV_EST_EXCLUSAO_CUPOM_FISCAL = 202
Public Const MOV_EST_CUSTO_BENEF_DEV_SIMBOLICA = 205
Public Const MOV_EST_TROCA_VENDA_CUPOM_FISCAL = 206
Public Const MOV_EST_DESMEMBRAMENTO_SAIDA = 208
Public Const MOV_EST_DESMEMBRAMENTO_ENTRADA = 209
Public Const MOV_EST_COMPRA_ENTREGA_TERCEIRO = 211
Public Const MOV_EST_ACRESCIMO_INVENT_DISPONIVEL_NOSSA_CI = 214
Public Const MOV_EST_DECRESCIMO_INVENT_DISPONIVEL_NOSSA_CI = 215
Public Const MOV_EST_ACRESCIMO_INVENT_RECEB_INDISP_CI = 216
Public Const MOV_EST_DECRESCIMO_INVENT_RECEB_INDISP_CI = 217
Public Const MOV_EST_ACRESCIMO_INVENT_DEFEITUOSO_CI = 218
Public Const MOV_EST_DECRESCIMO_INVENT_DEFEITUOSO_CI = 219
Public Const MOV_EST_ACRESCIMO_INVENT_CONSIG_TERC_CI = 220
Public Const MOV_EST_DECRESCIMO_INVENT_CONSIG_TERC_CI = 221
Public Const MOV_EST_ACRESCIMO_INVENT_INDISP_OUTRAS_CI = 222
Public Const MOV_EST_DECRESCIMO_INVENT_INDISP_OUTRAS_CI = 223
Public Const MOV_EST_ACRES_INVENT_DISP_NOSSA_SOLOTE_CI = 224
Public Const MOV_EST_DECR_INVENT_DISP_NOSSA_SOLOTE_CI = 225
Public Const MOV_EST_ACRES_INVENT_RECEB_IND_SOLOTE_CI = 226
Public Const MOV_EST_DECR_INVENT_RECEB_IND_SOLOTE_CI = 227
Public Const MOV_EST_ACRES_INVENT_DEFEITUOSO_SOLOTE_CI = 228
Public Const MOV_EST_DECR_INVENT_DEFEITUOSO_SOLOTE_CI = 229
Public Const MOV_EST_ACRES_INVENT_CONSIG_TERC_SOLOTE_CI = 230
Public Const MOV_EST_DECR_INVENT_CONSIG_TERC_SOLOTE_CI = 231
Public Const MOV_EST_ACRES_INVENT_IND_OUTRAS_SOLOTE_CI = 232
Public Const MOV_EST_DECR_INVENT_IND_OUTRAS_SOLOTE_CI = 233

Public Const MOV_EST_REQ_PRODUCAO_OUTROS = 250
Public Const MOV_EST_ESTORNO_REQ_PRODUCAO_OUTROS = 251
Public Const MOV_EST_PRODUCAO_OUTROS = 252
Public Const MOV_EST_ESTORNO_PRODUCAO_OUTROS = 253
Public Const MOV_EST_ENTRADA_TRANSF_BENEF_TERC = 260
Public Const MOV_EST_ESTORNO_ENTRADA_TRANSF_BENEF_TERC = 261
Public Const MOV_EST_APONT_SRV = 262
Public Const MOV_EST_ESTORNO_APONT_SRV = 263
Public Const MOV_EST_INT2_ENT_CONSERTO_TERC = 264
Public Const MOV_EST_INT2_DEV_CONSERTO_TERC = 265
Public Const MOV_EST_INT2_SAIDA_CONSERTO = 266
Public Const MOV_EST_INT2_DEV_CONSERTO = 267
Public Const MOV_EST_ESTORNO_CRED_TRIBUTARIO = 304

Public Const MOV_EST_ACRES_INVENT_NOSSO_BENEF = 305
Public Const MOV_EST_DECR_INVENT_NOSSO_BENEF = 306
Public Const MOV_EST_ACRES_INVENT_NOSSO_DEMO = 307
Public Const MOV_EST_DECR_INVENT_NOSSO_DEMO = 308
Public Const MOV_EST_ACRES_INVENT_NOSSO_OUTROS = 309
Public Const MOV_EST_DECR_INVENT_NOSSO_OUTROS = 310
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSIG = 311
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSIG = 312
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSERTO = 313
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSERTO = 314

Public Const MOV_EST_ACRES_INVENT_NOSSO_BENEF_SOLOTE = 315
Public Const MOV_EST_DECR_INVENT_NOSSO_BENEF_SOLOTE = 316
Public Const MOV_EST_ACRES_INVENT_NOSSO_DEMO_SOLOTE = 317
Public Const MOV_EST_DECR_INVENT_NOSSO_DEMO_SOLOTE = 318
Public Const MOV_EST_ACRES_INVENT_NOSSO_OUTROS_SOLOTE = 319
Public Const MOV_EST_DECR_INVENT_NOSSO_OUTROS_SOLOTE = 320
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSIG_SOLOTE = 321
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSIG_SOLOTE = 322
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSERTO_SOLOTE = 323
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSERTO_SOLOTE = 324

Public Const MOV_EST_ACRES_INVENT_NOSSO_BENEF_CI = 325
Public Const MOV_EST_DECR_INVENT_NOSSO_BENEF_CI = 326
Public Const MOV_EST_ACRES_INVENT_NOSSO_DEMO_CI = 327
Public Const MOV_EST_DECR_INVENT_NOSSO_DEMO_CI = 328
Public Const MOV_EST_ACRES_INVENT_NOSSO_OUTROS_CI = 329
Public Const MOV_EST_DECR_INVENT_NOSSO_OUTROS_CI = 330
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSIG_CI = 331
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSIG_CI = 332
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSERTO_CI = 333
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSERTO_CI = 334

Public Const MOV_EST_ACRES_INVENT_NOSSO_BENEF_SOLOTE_CI = 335
Public Const MOV_EST_DECR_INVENT_NOSSO_BENEF_SOLOTE_CI = 336
Public Const MOV_EST_ACRES_INVENT_NOSSO_DEMO_SOLOTE_CI = 337
Public Const MOV_EST_DECR_INVENT_NOSSO_DEMO_SOLOTE_CI = 338
Public Const MOV_EST_ACRES_INVENT_NOSSO_OUTROS_SOLOTE_CI = 339
Public Const MOV_EST_DECR_INVENT_NOSSO_OUTROS_SOLOTE_CI = 340
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSIG_SOLOTE_CI = 341
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSIG_SOLOTE_CI = 342
Public Const MOV_EST_ACRES_INVENT_NOSSO_CONSERTO_SOLOTE_CI = 343
Public Const MOV_EST_DECR_INVENT_NOSSO_CONSERTO_SOLOTE_CI = 344

Public Const MOV_EST_ACRES_INVENT_DISPONIVEL_NOSSA_FX_ZERA = 345
Public Const MOV_EST_ACRES_INVENT_DISP_NOSSA_SOLOTE_FX_ZERA = 346
Public Const MOV_EST_DECR_INVENT_DISPONIVEL_NOSSA_FX_ZERA = 347
Public Const MOV_EST_DECR_INVENT_DISP_NOSSA_SOLOTE_FX_ZERA = 348
Public Const MOV_EST_INVENT_DISPONIVEL_NOSSA_FX_RECOLOCA = 349
Public Const MOV_EST_INVENT_DISP_NOSSA_SOLOTE_FX_RECOLOCA = 350


Public Const MOV_EST_ESTORNO_SAIDA_TRANSF_OUTRAS_TERC = 494
Public Const MOV_EST_ESTORNO_ENTRADA_TRANSF_OUTRAS_TERC = 495

Public Const MOV_EST_ESTORNO_OUTRAS_REQ = 496
Public Const MOV_EST_ESTORNO_OUTRAS_DEV = 497
Public Const MOV_EST_OUTRAS_DEV = 498
Public Const MOV_EST_OUTRAS_REQ = 499

Public Const MOV_EST_ENTRADA_OC = 500

Public Const MOV_EST_ESTORNO_TRANSF = 1000   'Para informar que o movimento é um estorno de transferência

Public Const MOV_TRANSFERENCIA = 1

'Ordem de Produção
Public Const STRING_NORMAL = "Normal"
Public Const STRING_DESABILITADA = "Desabilitada"
Public Const STRING_SACRAMENTADA = "Sacramentada"
Public Const STRING_BAIXADA = "Baixada"
Public Const STRING_PEDIDOVENDA = "Pedido de Venda"
Public Const STRING_ESTOQUE = "Estoque"
Public Const STRING_CONSUMO = "Consumo"


Public Const STRING_DESCRICAO_ITEM = 250
Public Const STRING_LOTE_ITEM = 10
Public Const STRING_ALMOXARIFADO_NOMEREDUZIDO = 20

Public Const STRING_CLASSEABC = 1

Public Const STRING_DESCRICAO_TIPOMOVEST = 100
Public Const STRING_SIGLA_TIPOMOVEST = 5
Public Const STRING_DESCR_NUMINTDOCORIG_TIPOMOVEST = 50

Public Const STRING_ENTRADAOUSAIDA = 1
Public Const STRING_MATCONFIG_CONTEUDO = 255
'Public Const STRING_OPCODIGO = 9
'Public Const STRING_ORDEM_DE_PRODUCAO = 9
Public Const STRING_CODIGO_ITEM = 6
Public Const STRING_TIPODEPRODUTO_DESCRICAO = 50
Public Const STRING_TIPODEPRODUTO_SIGLA = 5
Public Const STRING_CLASSEUM_NOME = 20
Public Const STRING_CLASSEUM_DESCRICAO = 50
Public Const STRING_LOCALIZACAO_FISICA = 20
Public Const STRING_MOVESTOQUE_DOCORIGEM = 25

'Tela Estoque Produto
Public Const STRING_ESTOQUEPRODUTO_LOCALIZACAOFISICA = 20

Public Const STRING_CATEGORIAPRODUTO_CATEGORIA = 20
Public Const STRING_CATEGORIAPRODUTO_DESCRICAO = 50
Public Const STRING_CATEGORIAPRODUTO_SIGLA = 4
Public Const STRING_CATEGORIAPRODUTOITEM_DESCRICAO = 50
Public Const STRING_PRODUTOFILIAL_CLASSEABC = 1
Public Const STRING_PRODUTOFILIAL_SITUACAOTRIBECF = 3
Public Const STRING_TIPOMOV_EST_DESCRICAO = 100
Public Const STRING_TIPOMOV_EST_ENTRADAOUSAIDA = 1
Public Const STRING_PRODUTO_ITEM = 20
Public Const STRING_PRODUTO_CATEGORIA = 20
Public Const STRING_PRODUTOFILIAL_ICMS = 2

Public Const STRING_SIGLA_TIPOPEDIDO = 4
Public Const STRING_TIPOPEDIDO_DESCRICAO = 50
Public Const STRING_RESERVA_USUARIO = 10

'Inventário
Public Const STRING_INVENTARIO_CODIGO = 20
Public Const STRING_INVENTARIO_ETIQUETA = 10
Public Const STRING_INVLOTE_DESCRICAO = 50
Public Const TRANSACAO_INVLOTE = "InventarioLote"

'Classificação ABC
Public Const STRING_CLASSABC_CODIGO = 15
Public Const STRING_CLASSABC_DESCRICAO = 50

'Transfer
Public Const STRING_TRANSF_CONSIG3 = "Consignada de Terceiros"
Public Const STRING_TRANSF_CONSIG = "Consignada"

'Livros Fiscais
Public Const STRING_CODIGO_NCM = 8

'Rastreamento
Public Const STRING_RASTRO_OBSERVACAO = 50
Public Const STRING_RASTRO_LOCALIZACAO = 50

'TipoMovEst
Public Const TIPOMOV_EST_ENTRADA As String = "E"
Public Const TIPOMOV_EST_SAIDA As String = "S"
Public Const TIPOMOV_EST_AJUSTE_CUSTO_STANDARD As String = "A"
Public Const TIPOMOV_EST_NAO_ENTRADASAIDA As String = "N"
Public Const TIPOMOV_EST_NAOEDITAVEL = 0 'tipos pre-definidos, nao ediatveis pelo usuario
Public Const TIPOMOV_EST_EDITAVEL = 1
Public Const TIPOMOV_EST_NAO_INF_CUSTO = 0 'o usuario nao pode informar custo manualmente
Public Const TIPOMOV_EST_INF_CUSTO = 1
Public Const TIPOMOV_EST_NAO_INTERNO = 0 'invalido para a tela de movtos internos
Public Const TIPOMOV_EST_INTERNO = 1
Public Const TIPOMOV_EST_NAOATUALIZACONSERTO = 0 'não atualiza o totalizador de produtos em conserto
Public Const TIPOMOV_EST_ADICIONACONSERTO = 1 'adiciona ao totalizador de produtos em conserto
Public Const TIPOMOV_EST_SUBTRAICONSERTO = 2 'subtrai do totalizador de produtos em conserto
Public Const TIPOMOV_EST_NAOATUALIZADEMO = 0 'não atualiza o totalizador de produtos em demonstração
Public Const TIPOMOV_EST_ADICIONADEMO = 1 'adiciona ao totalizador de produtos em demonstração
Public Const TIPOMOV_EST_SUBTRAIDEMO = 2 'subtrai do totalizador de produtos em demonstração
Public Const TIPOMOV_EST_NAOATUALIZACONSUMO = 0 'não atualiza o totalizador de consumo
Public Const TIPOMOV_EST_ADICIONACONSUMO = 1 'adiciona ao totalizador de consumo
Public Const TIPOMOV_EST_SUBTRAICONSUMO = 2 'subtrai do totalizador de consumo
Public Const TIPOMOV_EST_NAOATUALIZAVENDA = 0 'não atualiza o totalizador de venda
Public Const TIPOMOV_EST_ADICIONAVENDA = 1 'adiciona ao totalizador de venda
Public Const TIPOMOV_EST_SUBTRAIVENDA = 2 'subtrai do totalizador de venda
Public Const TIPOMOV_EST_NAOATUALIZAVENDACONSIG3 = 0 'não atualiza o totalizador de venda
Public Const TIPOMOV_EST_ADICIONAVENDACONSIG3 = 1 'adiciona ao totalizador de venda
Public Const TIPOMOV_EST_SUBTRAIVENDACONSIG3 = 2 'subtrai do totalizador de venda
Public Const TIPOMOV_EST_NAOATUALIZACONSIGNACAO = 0 'não atualiza o totalizador de produtos em consignação
Public Const TIPOMOV_EST_ADICIONACONSIGNACAO = 1 'adiciona ao totalizador de produtos em consignação
Public Const TIPOMOV_EST_SUBTRAICONSIGNACAO = 2 'subtrai do totalizador de produtos em consignação
Public Const TIPOMOV_EST_SUBTRAICONSIGNACAO1 = 3 'subtrai do totalizador de produtos em consignação o que restou depois de retirar materiar nosso disponivel
Public Const TIPOMOV_EST_PRODUTONOSSO = 0 'o produto é nosso
Public Const TIPOMOV_EST_PRODUTODETERCEIROS = 1 'o produto é de terceiros
Public Const TIPOMOV_EST_NAOATUALIZAOUTRAS = 0 'não atualiza totalizador de outras quantidades disponiveis
Public Const TIPOMOV_EST_ADICIONAOUTRAS = 1 'adiciona ao totalizador de outras quantidades disponiveis
Public Const TIPOMOV_EST_SUBTRAIOUTRAS = 2 'subtrai do totalizador de outras quantidades disponiveis
Public Const TIPOMOV_EST_NAOATUALIZAINDOUTRAS = 0 'não atualiza totalizador de outras quantidades indisponiveis
Public Const TIPOMOV_EST_ADICIONAINDOUTRAS = 1 'adiciona ao totalizador de outras quantidades indisponiveis
Public Const TIPOMOV_EST_SUBTRAIINDOUTRAS = 2 'subtrai do totalizador de outras quantidades indisponiveis
Public Const TIPOMOV_EST_NAOATUALIZANOSSADISP = 0 'não atualiza totalizador de quantidade nossa disponivel
Public Const TIPOMOV_EST_ADICIONANOSSADISP = 1 'adiciona ao totalizador de quantidade nossa disponivel
Public Const TIPOMOV_EST_SUBTRAINOSSADISP = 2 'subtrai do totalizador de quantidade nossa disponivel
Public Const TIPOMOV_EST_NAOATUALIZADEFEITUOSA = 0 'não atualiza totalizador de quantidade defeituosa
Public Const TIPOMOV_EST_ADICIONADEFEITUOSA = 1 'adiciona ao totalizador de quantidade defeituosa
Public Const TIPOMOV_EST_SUBTRAIDEFEITUOSA = 2 'subtrai do totalizador de quantidade defeituosa
Public Const TIPOMOV_EST_NAOATUALIZARECEBINDISP = 0 'não atualiza totalizador de quantidade recebida indisponivel
Public Const TIPOMOV_EST_ADICIONARECEBINDISP = 1 'adiciona ao totalizador de quantidade recebida indisponivel
Public Const TIPOMOV_EST_SUBTRAIRECEBINDISP = 2 'subtrai do totalizador de quantidade recebida indisponivel
Public Const TIPOMOV_EST_INSEREMOV = 0 'indicacao que este movimento deve ser inserido
Public Const TIPOMOV_EST_EXCLUIMOV = 1 'indicacao que este movimento deve ser excluido
Public Const TIPOMOV_EST_ESTORNOMOV = 2 'indicacao que este movimento deve ser estornado
Public Const TIPOMOV_EST_NAOTRANSFERENCIA = 0 'indicação de que este movimento não é de transferencia
Public Const TIPOMOV_EST_TRANSFERENCIA = 1 'indicação de que este movimento é de transferencia
Public Const TIPOMOV_EST_ADICIONARESERVANOSSACONFIG = 3 'adciona ao totalizador de reserva nossa disponivel e baixa do escaninho nosso disponivel. Se não for suficiente baixa do produto em consig3 e reserva em consig3.
Public Const TIPOMOV_EST_SUBTRAIRESERVANOSSACONFIG = 4 'subtrai do totalizador de reserva nossa disponivel e incrementa no escaninho nosso disponivel. Se não for suficiente subtrai a reserva do produto em consig3 e incrementa em consig3.
Public Const TIPOMOV_EST_SUBTRAIRESERVACONFIGNOSSA = 5 'subtrai do totalizador de reserva config3 e incrementa no config3. Se não for suficiente baixa da reserva do produto nosso disponivel e incrementa em produto nosso disponivel.
Public Const TIPOMOV_EST_NAOATUALIZACUSTO = 0 'não atualiza o custo medio
Public Const TIPOMOV_EST_ATUALIZACUSTOADICIONA = 1 'indica que este movimento atualiza o custo medio adicionando a quantidade em questao e seu valor
Public Const TIPOMOV_EST_ATUALIZACUSTOSUBTRAI = 2 'indica que este movimento atualiza o custo medio subtraindo a quantidade em questao e seu valor
Public Const TIPOMOV_EST_NAOATUALIZACMRP = 0 'não apropria pelo custo real de produção
Public Const TIPOMOV_EST_ATUALIZACMRPDICIONA = 1 'apropria pelo custo real de produção se o produto não tiver apropriação pelo custo standard
Public Const TIPOMOV_EST_CUSTOINFORMADO = 1 'Custo é informado
Public Const TIPOMOV_EST_CUSTONAOINFORMADO = 0 'Custo não é informado
Public Const TIPOMOV_EST_NAOATUALIZACOMPRA = 0 'não atualiza o totalizador de compra
Public Const TIPOMOV_EST_ADICIONACOMPRA = 1 'adiciona ao totalizador de compra
Public Const TIPOMOV_EST_SUBTRAICOMPRA = 2 'subtrai do totalizador de compra
Public Const TIPOMOV_EST_NAOATUALIZABENEF = 0 'não atualiza o totalizador de produtos nossos em benefiamento em terceiros
Public Const TIPOMOV_EST_ADICIONABENEF = 1 'adiciona ao totalizador de produtos nossos em benefiamento em terceiros
Public Const TIPOMOV_EST_SUBTRAIBENEF = 2 'subtrai do totalizador de produtos nossos em benefiamento em terceiros
Public Const TIPOMOV_EST_NAOATUALIZAOP = 0 'não atualiza o totalizador de produtos em ordem de producao
Public Const TIPOMOV_EST_ADICIONAOP = 1 'adiciona ao totalizador de produtos em ordem de producao
Public Const TIPOMOV_EST_SUBTRAIOP = 2 'subtrai do totalizador de produtos em ordem de producao
Public Const TIPOMOV_EST_NAOATUALIZASALDOCUSTO = 0 'não atualiza o saldo (quantidade e valor) que calcula o custo medio
Public Const TIPOMOV_EST_ATUALIZASALDOCUSTOADICIONA = 1 'indica que este movimento atualiza o saldo (quantidade e valor) que calcula o custo medio adicionando a quantidade em questao e seu valor
Public Const TIPOMOV_EST_ATUALIZASALDOCUSTOSUBTRAI = 2 'indica que este movimento atualiza o saldo (quantidade e valor) que calcula o custo medio subtraindo a quantidade em questao e seu valor
Public Const TIPOMOV_EST_CUSTOMEDIO = 0 'utiliza o custo medio de estoque
Public Const TIPOMOV_EST_CUSTOMEDIO_CONSIG3 = 1 'utiliza o custo médio de material em consignação de terceiros
Public Const TIPOMOV_EST_CUSTOMEDIO_CONSIG = 2 'utiliza o custo médio de material nosso em consignação
Public Const TIPOMOV_EST_NAOINVENTARIO = 0 'não é um registro associado a movimento de inventário
Public Const TIPOMOV_EST_INVENTARIO = 1 'é um registro associado a movimento de inventário

Public Const TIPOMOV_EST_ULTIMOCODIGORESERVADO = 499

'Status da tabela SldMesEst
Public Const SLDMESEST_STATUS_ABERTO_NAO_ALTERADO = 0
Public Const SLDMESEST_STATUS_ABERTO_ALTERADO = 1
Public Const SLDMESEST_STATUS_FECHADO = 2

'Status do campo fechamento da tabela EstoqueMes
Public Const ESTOQUEMES_FECHAMENTO_ABERTO = 0
Public Const ESTOQUEMES_FECHAMENTO_FECHADO = 1

'Kit
Public Const STRING_KIT_VERSAO = 10
Public Const STRING_KIT_OBSERVACAO = 255

'Tipos de Documento associdados a Reservas
Public Const TIPO_MANUTENCAO As String = "Esta Tela"
Public Const TIPO_MANUTENCAO_COD = 0
Public Const TIPO_PEDIDO As String = "Pedido de Venda"
Public Const TIPO_PEDIDO_COD = 1
Public Const TIPO_PEDIDO_GRADE = 2
Public Const TIPO_PEDIDO_SRV As String = "Pedido de Serviço"
Public Const TIPO_PEDIDO_SRV_COD = 3

'Número máximo de Bloqueios, Almoxarifados, Reservas
Public Const NUM_MAX_BLOQUEIOS = 1000
Public Const NUM_MAX_ALMOXARIFADOS = 9999
Public Const NUM_MAX_RESERVAS = 1500

''Produto vazio ou preenchido
'Public Const PRODUTO_VAZIO = 0
'Public Const PRODUTO_PREENCHIDO = 1
'

'Tipo de Bloqueio de Estoque
Public Const BLOQUEIO_TOTAL = 1
Public Const BLOQUEIO_PARCIAL = 2
Public Const BLOQUEIO_NAO_RESERVA = 3
Public Const BLOQUEIO_CREDITO = 4
Public Const BLOQUEIO_DIAS_ATRASO = 6
Public Const BLOQUEIO_TAB_PRECO = 7
Public Const BLOQUEIO_PRECO_BAIXO = 8
Public Const BLOQUEIO_PRECO_DEFASADO = 9
Public Const BLOQUEIO_CARGO_COND_PAGTO = 10
Public Const BLOQUEIO_CARGO_TAB_PRECO = 11
Public Const BLOQUEIO_MARGEM_BAIXA = 12

'Public Const BLOQUEIO_PARCIAL_NOME_RED As String = "Bloqueio Parcial"
'Public Const BLOQUEIO_NAO_RESERVA_NOME_RED As String = "Não Reserva"

'Tipo de Almoxarifado
Public Const ALMOXARIFADO_NORMAL = 0
Public Const ALMOXARIFADO_DEMONSTRACAO = 1  'Produtos em Demonstração
Public Const ALMOXARIFADO_CONSIGNACAO = 2   'Produtos em Consignação
Public Const ALMOXARIFADO_CONSERTO = 3      'Produtos em Conserto
Public Const ALMOXARIFADO_PODER_TERCEIROS = 4 'Produtos em Poder Terceiros / Outros

'Tipos e string de quantidade p/ telas de Inventário
Public Const TIPO_QUANT_DISPONIVEL_NOSSA = 1
Public Const TIPO_QUANT_RECEB_INDISP = 2
Public Const TIPO_QUANT_OUTRAS_INDISP = 3
Public Const TIPO_QUANT_DEFEIT = 4
Public Const TIPO_QUANT_3_CONSIG = 5
Public Const TIPO_QUANT_3_DEMO = 6
Public Const TIPO_QUANT_3_CONSERTO = 7
Public Const TIPO_QUANT_3_OUTRAS = 8
Public Const TIPO_QUANT_3_BENEF = 9
Public Const TIPO_QUANT_DISPONIVEL_NOSSA_CI = 10  'CI = custo informado
Public Const TIPO_QUANT_RECEB_INDISP_CI = 11      'CI = custo informado
Public Const TIPO_QUANT_OUTRAS_INDISP_CI = 12     'CI = custo informado
Public Const TIPO_QUANT_DEFEIT_CI = 13            'CI = custo informado
Public Const TIPO_QUANT_3_CONSIG_CI = 14          'CI = custo informado

Public Const TIPO_QUANT_NOSSO_CONSIG = 15
Public Const TIPO_QUANT_NOSSO_DEMO = 16
Public Const TIPO_QUANT_NOSSO_CONSERTO = 17
Public Const TIPO_QUANT_NOSSO_OUTRAS = 18
Public Const TIPO_QUANT_NOSSO_BENEF = 19

Public Const TIPO_QUANT_NOSSO_CONSIG_CI = 20
Public Const TIPO_QUANT_NOSSO_DEMO_CI = 21
Public Const TIPO_QUANT_NOSSO_CONSERTO_CI = 22
Public Const TIPO_QUANT_NOSSO_OUTRAS_CI = 23
Public Const TIPO_QUANT_NOSSO_BENEF_CI = 24

Public Const TIPO_QUANT_DISPONIVEL_NOSSA_CI2P = 25  'FX = Fixo

Public Const STRING_QUANT_DISPONIVEL_NOSSA = "Nossa Disponível"
Public Const STRING_QUANT_RECEB_INDISP = "Recebida  e Indisponível"
Public Const STRING_QUANT_OUTRAS_INDISP = "Outras Indisponíveis"
Public Const STRING_QUANT_DEFEIT = "Defeituosa"
Public Const STRING_QUANT_3_CONSIG = "De 3º em Consignação"
Public Const STRING_QUANT_3_DEMO = "De 3º em Demonstração"
Public Const STRING_QUANT_3_CONSERTO = "De 3º em Conserto"
Public Const STRING_QUANT_3_OUTRAS = "Outras de 3º"
Public Const STRING_QUANT_3_BENEF = "De 3º em Beneficiamento"

Public Const STRING_QUANT_DISPONIVEL_NOSSA_CI = "Nossa Disponível - CI"        'CI = custo informado
Public Const STRING_QUANT_RECEB_INDISP_CI = "Recebida  e Indisponível - CI"    'CI = custo informado
Public Const STRING_QUANT_OUTRAS_INDISP_CI = "Outras Indisponíveis - CI"       'CI = custo informado
Public Const STRING_QUANT_DEFEIT_CI = "Defeituosa - CI"                        'CI = custo informado
Public Const STRING_QUANT_3_CONSIG_CI = "De 3º em Consignação - CI"            'CI = custo informado

Public Const STRING_QUANT_NOSSO_CONSIG = "Nosso em Consignação"
Public Const STRING_QUANT_NOSSO_DEMO = "Nosso em Demonstração"
Public Const STRING_QUANT_NOSSO_CONSERTO = "Nosso em Conserto"
Public Const STRING_QUANT_NOSSO_OUTRAS = "Nosso Outras"
Public Const STRING_QUANT_NOSSO_BENEF = "Nosso em Beneficiamento"

Public Const STRING_QUANT_NOSSO_CONSIG_CI = "Nosso em Consignação - CI"
Public Const STRING_QUANT_NOSSO_DEMO_CI = "Nosso em Demonstração - CI"
Public Const STRING_QUANT_NOSSO_CONSERTO_CI = "Nosso em Conserto - CI"
Public Const STRING_QUANT_NOSSO_OUTRAS_CI = "Nosso Outras - CI"
Public Const STRING_QUANT_NOSSO_BENEF_CI = "Nosso em Beneficiamento - CI"

Public Const STRING_QUANT_DISPONIVEL_NOSSA_CI2P = "Nossa Disponível - CI2P"

'Tabela de Preço
Public Const STRING_TABELA_DESCRICAO = 50
Public Const STRING_TABELA_OBSERVACAO = 255
Public Const STRING_TABELA_TEXTOGRADE = 50

'Indica se um codigo/filial de um movimento de estoque já está cadastrado
Public Const MOVESTOQUE_CODIGO_JA_CADASTRADO = 1

'Situação Tributária
Public Const STRING_SITUACAOTRIB = 2

Public Const STRING_LOTE_RASTREAMENTO = 20

'*** Constantes incluídas para tratamento de Grade ****
'Como a Tela de Romaneio é genérica ela vai ter vário modos de funcionamento
Public Const ROMANEIOGRADE_FUNCIONAMENTO_PEDIDO = 1
Public Const ROMANEIOGRADE_FUNCIONAMENTO_ORCAMENTO = 2
Public Const ROMANEIOGRADE_FUNCIONAMENTO_NFFATPEDIDO = 3
Public Const ROMANEIOGRADE_FUNCIONAMENTO_NFISCAL = 4
Public Const ROMANEIOGRADE_FUNCIONAMENTO_NFISCALREM = 5
Public Const ROMANEIOGRADE_FUNCIONAMENTO_RECEBIMENTO = 6
Public Const ROMANEIOGRADE_FUNCIONAMENTO_OP = 7
Public Const ROMANEIOGRADE_FUNCIONAMENTO_PRODSAI = 8
Public Const ROMANEIOGRADE_FUNCIONAMENTO_PRODENT = 9
'******************************************************

Public Const STRING_PEDIDOCOTACAO_CONTATO = 50
Public Const STRING_COTACAO_DESCRICAO = 50
Public Const STRING_COTACAOPRODUTO_PRODUTO = 20
Public Const STRING_COTACAOPRODUTO_UM = 5

Public Const STRING_NOME_FILIAL_FORN = 50
Public Const STRING_NOME_FILIAL_CLI = 50

'Para saber se vai ser ou não calculado os Parametros  de Ponto Pedido
Public Const PRODUTOFILIAL_CALCULA_VALORES = 1
Public Const PRODUTOFILIAL_NAO_CALCULA_VALORES = 0


Type typeBloqueioPV
    iFilialEmpresa As Integer
    lPedidoDeVendas As Long
    iSequencial As Integer
    iTipoDeBloqueio As Integer
    sCodUsuario As String
    sResponsavel As String
    dtData As Date
    sCodUsuarioLib As String
    sResponsavelLib As String
    dtDataLib As Date
    sObservacao As String
End Type

Type typeMovEst
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    iTipoMov As Integer
    lNumIntDocOrigem As Long
    dtData As Date
    dHora As Double
    '#####################
    'Inserido por Wagner
    sObservacao As String
    lRequisitante As Long
    '#####################
End Type

Type typeItemMovEstoque
    iFilialEmpresa As Integer
    lCodigo As Long
    lNumIntDoc As Long
    dCusto As Double
    iApropriacao As Integer
    sProduto As String
    sSiglaUM As String
    dQuantidade As Double
    iAlmoxarifado As Integer
    iTipoMov As Integer
    lNumIntDocOrigem As Long
    iTipoNumIntDocOrigem As Integer
    dtData As Date
    sCcl As String
    lNumIntDocEst As Long
    sProdutoDesc As String
    sAlmoxarifadoNomeRed As String
    sOPCodigo As String
    sProdutoOP As String
    sContaContabilEst As String
    sContaContabilAplic As String
    sDocOrigem As String
    lCliente As Long
    lFornecedor As Long
    lHorasMaquina As Long
    dtDataInicioProducao As Date
    dtDataRegistro As Date
    dHora As Double
    iClasseUM As Integer
    iNaturezaProduto As Integer
    sSiglaUMEst As String
    lNumIntDocGrade As Long
    dQuantInsumos As Double
    iItemNF As Integer
    '#####################
    'Inserido por Wagner
    sObservacao As String
    lRequisitante As Long
    '#####################
    iFilialCli As Integer
    iFilialForn As Integer
End Type

' *** Incluído por Luiz G.F.Nogueira em 28/08/01**
Type typeProdutoEmbalagem
    sProduto As String
    iEmbalagem As Integer
    iSeqGrid As Integer
    iPadrao As Integer
    sNomeProdEmb As String
    sUMEmbalagem As String
    dCapacidade As Double
    sUMPeso As String
    dPesoLiqTotal As Double
    dPesoEmbalagem As Double
    dPesoBruto As Double
End Type
 '**************************************************

' *** Incluído por Luiz G.F.Nogueira **
' *** Usado para fazer um select dinâmico, onde as variáveis são passadas pra o Bind uma a uma
'e todas precisam ser do tipo Variant ***
Type typeItemMovEstoqueVariant
    viFilialEmpresa As Variant
    vlCodigo As Variant
    vlNumIntDoc As Variant
    vdCusto As Variant
    viApropriacao As Variant
    vsProduto As Variant
    vsSiglaUM As Variant
    vdQuantidade As Variant
    viAlmoxarifado As Variant
    viTipoMov As Variant
    vlNumIntDocOrigem As Variant
    viTipoNumIntDocOrigem As Variant
    vdtData As Variant
    vsCcl As Variant
    vlNumIntDocEst As Variant
    vsProdutoDesc As Variant
    vsAlmoxarifadoNomeRed As Variant
    vsOPCodigo As Variant
    vsProdutoOP As Variant
    vsContaContabilEst As Variant
    vsContaContabilAplic As Variant
    vsDocOrigem As Variant
    vlCliente As Variant
    vlFornecedor As Variant
    vlHorasMaquina As Variant
    vdtDataInicioProducao As Variant
    vdtDataRegistro As Variant
    vdHora As Variant
End Type
'******************************************

Type typeTipoMovEst
    iCodigo As Integer
    sDescricao As String
    sEntradaOuSaida As String
    iInativo As Integer
    iAtualizaConsumo As Integer
    iAtualizaVenda As Integer
    iAtualizaVendaConsig3 As Integer
    iEditavel As Integer
    iValidoMovInt As Integer
    iAtualizaReserva As Integer
    iTransferencia As Integer
    iAtualizaCusto As Integer
    iAtualizaConsig As Integer
    iAtualizaDemo As Integer
    iAtualizaConserto As Integer
    iProdutoDeTerc As Integer
    iAtualizaOutras As Integer
    iAtualizaIndOutras As Integer
    iAtualizaNossaDisp As Integer
    iAtualizaDefeituosa As Integer
    iAtualizaRecebIndisp As Integer
    iAtualizaMovEstoque As Integer
    iAtualizaCMRProd As Integer
    iCustoInformado As Integer
    iValidaDataUltMov As Integer
    iAtualizaCompra As Integer
    iAtualizaBenef As Integer
    iAtualizaOP As Integer
    sDescrNumIntDocOrigem As String
    iTipoNumIntDocOrigem As Integer
    iKardex As Integer
    sSigla As String
    sNomeTela As String
    iLivroMod3 As Integer
    iAtualizaSaldoCusto As Integer
    iCustoMedio As Integer
    sEntradaSaidaCMP As String
    iAtualizaCustoConsig As Integer
    iAtualizaCustoDemo As Integer
    iAtualizaCustoConserto As Integer
    iAtualizaCustoOutros As Integer
    iAtualizaCustoBenef As Integer
    iAtualizaCustoConsig3 As Integer
    iAtualizaCustoDemo3 As Integer
    iAtualizaCustoConserto3 As Integer
    iAtualizaCustoOutros3 As Integer
    iAtualizaCustoBenef3 As Integer
    iCodigoOrig As Integer
    iAtualizaSoLote As Integer
    iInventario As Integer
    iNFDevolucao As Integer
End Type

Type typeAlmoxarifado
    iFilialEmpresa As Integer
    iCodigo As Integer
    sNomeReduzido As String
    sDescricao As String
    lEndereco As Long
    sContaContabil As String
'    iTipo As Integer
End Type

Type typeEstoqueProduto
    sProduto As String
    iAlmoxarifado As Integer
    sAlmoxarifadoNomeReduzido As String
    dQuantDispNossa As Double
    dSaldo As Double
    sLocalizacaoFisica As String
    sContaContabil As String
    dQuantReservada As Double
    dQuantReservadaConsig As Double
    dtDataInventario As Date
    dQuantidadeInicial As Double
    dSaldoInicial As Double
    dtDataInicial As Date
    dQuantEmpenhada As Double
    dQuantPedido As Double
    dQuantRecIndl As Double
    dQuantInd As Double
    dQuantDefeituosa As Double
    dQuantConsig As Double
    dQuantConsig3 As Double
    dQuantDemo As Double
    dQuantDemo3 As Double
    dQuantConserto3 As Double
    dQuantConserto As Double
    dQuantOutras As Double
    dQuantOutras3 As Double
    dQuantBenef As Double
    dQuantBenef3 As Double
    dQuantOP As Double
    dValorConsig As Double
    dValorConsig3 As Double
    dValorDemo As Double
    dValorDemo3 As Double
    dValorConserto3 As Double
    dValorConserto As Double
    dValorOutras As Double
    dValorOutras3 As Double
    dValorBenef As Double
    dValorBenef3 As Double
    dQuantInicialConsig3 As Double
    dQuantInicialConsig As Double
    dQuantInicialDemo3 As Double
    dQuantInicialDemo As Double
    dQuantInicialConserto3 As Double
    dQuantInicialConserto As Double
    dQuantInicialOutras3 As Double
    dQuantInicialOutras As Double
    dQuantInicialBenef As Double
    dQuantInicialBenef3 As Double
    dValorInicialConsig3 As Double
    dValorInicialConsig As Double
    dValorInicialDemo3 As Double
    dValorInicialDemo As Double
    dValorInicialConserto3 As Double
    dValorInicialConserto As Double
    dValorInicialOutras3 As Double
    dValorInicialOutras As Double
    dValorInicialBenef As Double
    dValorInicialBenef3 As Double
    
End Type

Type typeProdutoKit
    sProdutoRaiz As String
    sVersao As String
    iNivel As Integer
    iSeq As Integer
    sProduto As String
    iSeqPai As Integer
    dQuantidade As Double
    sUnidadeMed As String
    iComposicao As Integer
    iPosicaoArvore As Integer
    dPercentualPerda As Double
    dCustoStandard As Double
    '##################
    'Inserido por Wagner
    sVersaoKitComp As String
    '##################
    
End Type

Type typeItemOP
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    sCodigo As String
    iItem As Integer
    sProduto As String
    iFilialPedido As Integer
    lCodPedido As Long
    lNumIntOrigem As Long
    iAlmoxarifado As Integer
    sCcl As String
    sSiglaUM As String
    dQuantidade As Double
    dtDataInicioProd As Date
    dtDataFimProd As Date
    dtDataEmissao As Date
    iPrioridade As Integer
    dQuantidadeProd As Double
    iSituacao As Integer
    iDestinacao As Integer
    iBeneficiamento As Integer
    lNumIntEquipamento As Long
    lNumIntEquipamento2 As Long
    sVersao As String
    sDescricao As String
    sSiglaUMEstoque As String
    iClasseUM As Integer
    iOrigemPedido As Integer
    lNumIntItemOP As Long
    iTipo As Integer
    
    '#############################################
    'INSERIDO POR WAGNER
    lNumIntDocPai As Long
    sJustificativaBloqueio As String
    iProduzLogo As Integer
    '#############################################
End Type


Type typeEmpenho
    iFilialEmpresa As Integer
    lCodigo As Long
    sCodigoOP As String
    iItemOP As Integer
    sProduto As String
    iAlmoxarifado As Integer
    dQuantidade As Double
    dQuantidadeRequisitada As Double
    lNumIntDocItemOP As Long
    dtData As Date
End Type

Type typeInventario
    iFilialEmpresa As Integer
    sCodigo As String
    dtData As Date
    iLote As Integer
    dHora As Double
End Type

Type typeItemInventario
    sProduto As String
    sProdutoDesc As String
    sSiglaUM As String
    dQuantidade As Double
    dQuantEst As Double
    dCusto As Double
    iAlmoxarifado As Integer
    sEtiqueta As String
    iTipo As Integer
    lNumIntDoc As Long
    sAlmoxarifadoNomeRed As String
    sContaContabilEst As String
    sContaContabilInv As String
    sLoteProduto As String
    iFilialOP As Integer
    iAtualizaSoLote As Integer
End Type

Type typeInvLote
    iFilialEmpresa As Integer
    iLote As Integer
    sDescricao As String
    iNumItensInf As Integer
    iNumItensAtual As Integer
    iIdAtualizacao As Integer
End Type

Type typeFornecedorProduto
    lFornecedor As Long
    sProduto As String
    sProdutoFornecedor As String
    dLoteMinimo As Double
    dLoteEconomico As Double
    dQuantPedAbertos As Double
    iTempoMedio As Integer
    dQuantPedida As Double
    dQuantRecebida As Double
    dValor As Double
    dtDataPedido As Date
    dtDataReceb As Date
    
End Type

Type typeTabelaPrecoItem
    iCodTabela As Integer
    sCodProduto As String
    dPreco As Double
    iFilialEmpresa As Integer
    dtDataVigencia As Date
    sDescricaoTabela As String
    dtDataLog As Date
    iAtivo As Integer
    sObservacao As String
    sTextoGrade As String
    dComissao As Double
    dPercDesconto As Double
    dPrecoComDesconto As Double
End Type

Type typeProdutoCusto
    sCodProduto As String
    dCusto As Double
    dCustoMesAnterior As Double
    sDescProduto As String
    sSiglaUMEstoque As String
End Type

Type typeItemClassifABC
    sCodProduto As String
    dDemanda As Double
    sDescProduto As String
    sClasseABC As String
    lClassifABC As Long
End Type

Type typeClassificacaoABC
    iFilialEmpresa As Integer
    lNumInt As Long
    sCodigo As String
    sDescricao As String
    dtData As Date
    iMesInicial As Integer
    iAnoInicial As Integer
    iMesFinal As Integer
    iAnoFinal As Integer
    iFaixaA As Integer
    iFaixaB As Integer
    iTipoProduto As Integer
    dDemandaTotal As Double
    iAtualizaProdutosFilial As Integer
End Type

Public Const KIT_NUM_FATORES = 6

Type typeKit
    sProdutoRaiz As String
    sVersao As String
    dtData As Date
    sObservacao As String
    iSituacao As Integer
    dPesoFator1 As Double
    dPesoFator2 As Double
    dPesoFator3 As Double
    dPesoFator4 As Double
    dPesoFator5 As Double
    dPesoFator6 As Double
    iVersaoFormPreco As Integer
End Type

Type typeApropricacaoInsumosProd

    lNumIntDoc As Long
    lNumIntDocOrigem As Long
    lNumIntReqProd As Long
    sProduto As String
    dQuantidade As Double
    iTipoInsumo As Integer
    iAutomatico As Integer
End Type

Type typeRastreamentoLote
    
    lNumIntDoc As Long
    sCodigo As String
    sProduto As String
    dtDataValidade As Date
    dtDataEntrada As Date
    dtDataFabricacao As Date
    sObservacao As String
    iFilialOP As Integer
    iStatus As Integer
    sLocalizacao As String
    lCliente As Long
    iFilialCli As Integer
End Type

Type typeRastreamentoLoteLoc
    sLocalizacao As String
End Type

Type typeRastreamentoLoteSaldo

    sLote As String
    sProduto As String
    iAlmoxarifado As Integer
    dQuantDispNossa As Double
    dQuantReservada As Double
    dQuantReservadaConsig As Double
    dQuantEmpenhada As Double
    dQuantPedida As Double
    dQuantRecIndl As Double
    dQuantIndOutras As Double
    dQuantDefeituosa As Double
    dQuantConsig3 As Double
    dQuantConsig As Double
    dQuantDemo3 As Double
    dQuantDemo As Double
    dQuantConserto3 As Double
    dQuantConserto As Double
    dQuantOutras3 As Double
    dQuantOutras As Double
    dQuantOP As Double
    dQuantBenef As Double
    dQuantBenef3 As Double
    iFilialOP As Integer
    lNumIntDocLote As Long
    
End Type

Type typeRastroEstIni

    sProduto As String
    iAlmoxarifado As Integer
    iEscaninho As Integer
    lNumIntDocLote As Long
    dQuantidade As Double
    dtDataEntrada As Date
    iFilialOP As Integer
    sLote As String

End Type

Type typeRastreamentoMovto

    lNumIntDoc As Long
    sProduto As String
    iTipoDocOrigem As Integer
    lNumIntDocOrigem As Long
    sLote As String
    dQuantidade As Double
    iFilialOP As Integer
    
End Type

Type typeTipoDeProduto
    sSiglaUMCompra As String
    sSiglaUMEstoque As String
    sSiglaUMVenda As String
    sDescricao As String
    sSigla As String
    sIPICodigo As String
    sIPICodDIPI As String
    sISSCodigo As String
    sContaContabil As String
    sContaProducao As String
    iTipo As Integer
    iClasseUM As Integer
    iCompras As Integer
    iControleEstoque As Integer
    iFaturamento As Integer
    iKitBasico As Integer
    iKitInt As Integer
    iPCP As Integer
    iPrazoValidade As Integer
    iIRIncide As Integer
    iICMSAgregaCusto As Integer
    iIPIAgregaCusto As Integer
    iFreteAgregaCusto As Integer
    iApropriacaoCusto As Integer
    iIntRessup As Integer
    iMesesConsumoMedio As Integer
    iConsideraQuantCotAnt As Integer
    iTemFaixaReceb As Integer
    iRecebFaixaFora As Integer
    iNatureza As Integer
    colCategoriaItem As Collection
    dIPIAliquota As Double
    dISSAliquota As Double
    dTempoRessupMax As Double
    dConsumoMedioMax As Double
    dResiduo As Double
    dPercentMaisQuantCotAnt As Double
    dPercentMenosQuantCotAnt As Double
    dPercentMaisReceb As Double
    dPercentMenosReceb As Double
    iKitVendaComp As Integer
    iExTIPI As Integer
    iProdutoEspecifico As Integer
    sGenero As String
    sISSQN As String
    sSiglaUMTrib As String
    iOrigem As Integer
    sCEST As String
End Type


Type typeCotacao

    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    sDescricao As String
    dtData As Date
    iTipoDestino As Integer
    lFornCliDestino As Long
    iFilialDestino As Integer
    iComprador As Integer

End Type

Type typePedidoCotacao
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    lFornecedor As Long
    iFilial As Integer
    sContato As String
    dtDataEmissao As Date
    dtData As Date
    dtDataValidade As Date
    iTipoFrete As Integer
    iStatus As Integer
    iCondPagtoPrazo As Integer
    dtDataBaixa As Date
End Type

Type typeCotacaoProduto
    lNumIntDoc As Long
    lCotacao As Long
    sProduto As String
    dQuantidade As Double
    sUM As String
    lFornecedor As Long
    iFilial As Integer
End Type

Type typeItemPedCotacao
    lNumIntDoc As Long
    sProduto As String
    dQuantidade As Double
    sUM As String
    lCotacaoProduto As Long
    sObservacao As String
End Type

Type typeItemCotacao
    lNumIntDoc As Long
    iCondPagto As Integer
    dtDataReferencia As Date
    dPrecoUnitario As Double
    dOutrasDespesas As Double
    dValorSeguro As Double
    dValorDesconto As Double
    dValorTotal As Double
    dValorIPI As Double
    dAliquotaIPI As Double
    dAliquotaICMS As Double
    iPrazoEntrega As Integer
    dQuantEntrega As Double
    lObservacao As Long
    dValorFrete As Double
    iMoeda As Integer
    dTaxa As Double
End Type

Type typeOrdemProducao
    iFilialEmpresa As Integer
    sCodigo As String
    dtDataEmissao As Date
    iNumItens As Integer
    iNumItensBaixados As Integer
    iGeraReqCompra As Integer
    iGeraOP As Integer
    lCodPrestador As Long
    iTipo As Integer
    iGeraOPsArvore As Integer
    sOPGeradora As String
    lNumIntDocOper As Long
    lCodigoNumerico As Long
    iTipoTerc As Integer
    iEscaninhoTerc As Integer
    iFilialTerc As Integer
    lCodTerc As Long
    iIgnoraEst As Integer
End Type

Type typeFornecedorProdutoFF
    iFilialEmpresa As Integer
    sProduto As String
    lFornecedor As Long
    iFilialForn As Integer
    sProdutoFornecedor As String
    dLoteMinimo As Double
    iNota As Integer
    dQuantPedAbertos As Double
    dtDataUltimaCompra As Date
    dTempoRessup As Double
    dQuantPedida As Double
    sUMQuantPedida As String
    dQuantRecebida As Double
    sUMQuantRecebida As String
    dtDataPedido As Date
    dtDataReceb As Date
    dPrecoTotal As Double
    dUltimaCotacao As Double
    dtDataUltimaCotacao As Date
    iTipoFreteUltimaCotacao As Integer
    dQuantUltimaCotacao As Double
    sUMQuantUltimaCotacao As String
    iPadrao As Integer
    iCondPagto As Integer
    sCondPagto As String
    sDescricao As String
End Type

Type typeEscaninho
    iCodigo As Integer
    sNome As String
    iRastroEstoqueInicial As Integer
End Type

Type typeEmbalagem
    iCodigo As Integer
    sDescricao As String
    sSigla As String
    dCapacidade As Double
    dPeso As Double
    sProduto As String
End Type

'Type usado para gravacao de historico de produto
Type typeProdutoHistorico
    lNumIntDoc As Long
    dtDataAtualizacao As Date
    sCodigoProd As String
    sDescProduto As String
End Type

'Incluído por Ivan em 04/04/03
Type typeCategoriaProdutoItem
    sCategoria As String
    sDescricao As String
    sItem As String
    iOrdem As Integer
    dvalor1 As Double
    dvalor2 As Double
    dvalor3 As Double
    dvalor4 As Double
    dvalor5 As Double
    dvalor6 As Double
    dvalor7 As Double
    dvalor8 As Double
End Type

'constantes de tipo de terceiros
Public Const TIPO_TERC_CLIENTE = 1
Public Const TIPO_TERC_FORNECEDOR = 2

Type typeInventarioTercProd
    iFilialEmpresa As Integer
    iTipoTerc As Integer
    lCodTerc As Long
    iFilialTerc As Integer
    dtData As Date
    sProduto As String
    iCodEscaninho As Integer
    dQuantTotal As Double
End Type


Type typeOrdemServicoProd
    iFilialEmpresa As Integer
    sCodigo As String
    dtDataEmissao As Date
    lCodPrestador As Long
    iStatus As Integer
End Type

Type typeItemOSProd
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    sCodigo As String
    iItem As Integer
    sProduto As String
    sSiglaUM As String
    dQuantidade As Double
    dQuantidadeProd As Double
    iStatus As Integer
End Type

Type typeItemOSProdCons
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    sCodigo As String
    iItem As Integer
    sProduto As String
    sSiglaUM As String
    dQuantidade As Double
    dtData As Date
End Type

Public Type typeItemRomaneioGrade
    lNumIntDoc As Long
    lNumIntItemPV As Long
    dQuantidade As Double
    dQuantCancelada As Double
    dQuantReservada As Double
    sProduto As String
    dQuantFaturada As Double
    dQuantAFaturar As Double
    dQuantOP As Double
    dQuantSC As Double
    sDescricao As String
    sSiglaUMEstoque As String
    iAlmoxarifado As Integer
    iControleEstoque As Integer
End Type


Public Type typePrecoCalculado
    sCodProduto As String
    dPrecoCalculado As Double
    dPrecoInformado As Double
    iFilialEmpresa As Integer
    dtDataVigencia As Date
    dtDataReferencia As Date
End Type

Public Type typeProdutoKitProdutos

    sProduto As String
    dQuantidade As Double
    sUnidadeMed As String
    iComposicao As Integer
    iControleEstoque As Integer
    sSiglaUMEstoque As String
    iClasseUM As Integer
    sProdutoRaiz As String
    sVersao As String
    dPercentualPerda As Double

End Type

Type typeProdutoTeste
    sProduto As String
    iTesteCodigo As Integer
    iSeqGrid As Integer
    sTesteEspecificacao As String
    iTesteTipoResultado As Integer
    dTesteLimiteDe As Double
    dTesteLimiteAte As Double
    sTesteMetodoUsado As String
    sTesteObservacao As String
    iTesteNoCertificado As Integer
End Type

Type typeRastreamentoLoteTeste
    lNumIntRastroLote As Long
    iTesteCodigo As Integer
    iSeqGrid As Integer
    sTesteEspecificacao As String
    iTesteTipoResultado As Integer
    dTesteLimiteDe As Double
    dTesteLimiteAte As Double
    sTesteMetodoUsado As String
    sTesteObservacao As String
    sRegistroAnaliseID As String
    dtRegistroAnaliseData As Date
    iResultadoNaoConforme As Integer
    sResultadoValor As String
    sResultadoObservacao As String
    iTesteNoCertificado As Integer
End Type

Type typeTestesQualidade
    iCodigo As Integer
    sNomeReduzido As String
    sEspecificacao As String
    iTipoResultado As Integer
    dLimiteDe As Double
    dLimiteAte As Double
    sMetodoUsado As String
    sObservacao As String
    iNoCertificado As Integer
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeInvCliForn
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    dtData As Date
    iEscaninho As Integer
    iTipoCliForn As Integer
    lCliForn As Long
    iFilial As Integer
    sUsuario As String
    dtDataGravacao As Date
    dHoraGravacao As Double
    sObs As String
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeInvCliFornItens
    lNumIntInvCliForn As Long
    iSeq As Integer
    sProduto As String
    dQtdData As Double
    dQtdEncontCliData As Double
    dQtdCliData As Double
    dQtdAcerto As Double
    sObs As String
End Type

