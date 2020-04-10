Attribute VB_Name = "GlobalCOM"
Option Explicit

Public Const STRING_ITEMPEDCOTACAO_OBSERVACAO = 255
Public Const PEDIDOCOMPRA_TIPODOCORIGEM_PV = 0
Public Const PEDIDOCOMPRA_TIPODOCORIGEM_PSRV = 1

Public Const STRING_OBSEMBALAGEM = 150

Public Const MOTIVO_EXCLUSIVO_DESCRICAO = "Exclusividade"
Public Const MOTIVO_MELHORPRECO_DESCRICAO = "Menor Preço"
Public Const MOTIVO_PRECO_PRAZO_DESCRICAO = "Preço e Prazo de Entrega"

Public Const FORNECEDOR_EXCLUSIVO = 1
Public Const FORNECEDOR_PREFERENCIAL = 0

Public Const PRODUTO_CONSIDERA_QUANT_COTACAO_ANTERIOR = 1

Public Const NOTA_FISCAL_COMPRAS As String = "NFiscalEntradaCom"
Public Const NOTA_FISCAL_FATURA_COMPRAS As String = "NFiscalFatEntradaCom"

Public Const TIPO_BAIXA_NORMAL = 0
Public Const NFISCAL_NAO_ACEITA_DIFERENCA_PC = 0
Public Const NFISCAL_ACEITA_DIFERENCA_PC = 1

'Strings
Public Const STRING_TIPOFRETE = 1
Public Const STRING_MOTIVO_BAIXA = 50
Public Const STRING_ALCADA_CODUSUARIO = 10
'Public Const STRING_OPCODIGO = 6

Public Const REQUISITANTE_AUTOMATICO_NOMEREDUZIDO = "AUTO"

Public Const REQUISITANTE_AUTOMATICO_CODIGO = 1

Public Const TIPO_ITEMPEDCOTACAO = 2
Public Const TIPO_COTACAOITEMCONCORRENCIA = 1

Public Const CONCORRENCIA_ASSOCIADA_RC = 1
Public Const CONCORRENCIA_NAO_ASSOCIADA_RC = 0

Public Const NUM_MAX_FORNFILIALFF = 100

Public Const NUM_MAX_PEDIDOS_SEL = 100
Public Const NUM_MAX_ITENSPED_SEL = 100
Public Const NUM_MAX_PRODUTOS_PONTOPEDIDO = 100

Public Const NUM_MAX_CONCORRENCIAS = 100
Public Const NUM_MAX_GERACAO = 100
'Public Const NUM_MAX_COTACOES = 100 'Alterado por Wagner
Public Const NUM_MAX_PEDCOTACOES = 100
'Public Const NUM_MAX_ITENS_GERACAO = 100 'Alterado por Wagner
'Public Const NUM_MAX_NFS_ITEMPED = 100 'Alterado por Wagner

Public Const TIPO_ORIGEM_PEDCOTACAO = 2
Public Const TIPO_ORIGEM_COTACAOITEMCONC = 1

Public Const TIPO_COMPRAS = 1

Public Const ALTERADO = 1
Public Const NAO_ALTERADO = 0

Public Const COMPRAS_CONFIG_DATA_CALCULO_PTO_PEDIDO = "DATA_CALCULO_PTO_PEDIDO"
Public Const COMPRAS_CONFIG_MESES_CONSUMO_MEDIO = "MESES_CONSUMO_MEDIO"
Public Const COMPRAS_CONFIG_MESES_MEDIA_TEMPO_RESSUP = "MESES_MEDIA_TEMPO_RESSUP"
Public Const COMPRAS_CONFIG_NUM_COMPRAS_TEMPO_RESSUP = "NUM_COMPRAS_TEMPO_RESSUP"
Public Const COMPRAS_CONFIG_CONSUMO_MEDIO_MAX = "CONSUMO_MEDIO_MAX"
Public Const COMPRAS_CONFIG_TEMPO_RESSUP_MAX = "TEMPO_RESSUP_MAX"


Public Const EXIBE_REQUISICOES_COTADAS = 1
Public Const NAO_EXIBE_REQUISICOES_COTADAS = 0

'constantes para Exclusivo(=1) e Preferencial(0)
Public Const ITEM_FILIALFORNECEDOR_EXCLUSIVO = 1
Public Const ITEM_FILIALFORNECEDOR_PREFERENCIAL = 0

Public Const BAIXA_MANUAL_PEDCOMPRA = 1

'Indica se a concorrência está associada a uma RequisiÇÃo.
Public Const CONC_NAO_ASSOICAIDA_RC = 0
Public Const CONC_ASSOCIADA_RC = 1

'Tipo de origem de Pedido de Compra (itens)
Public Const PC_TIPO_ORIGEM_CONCORRENCIA = 1
Public Const PC_TIPO_ORIGEM_PED_COTACAO = 2

'Motivo de Escolha de Fornecedor
Public Const MOTIVO_EXCLUSIVO = 4

'Código de Tipo de Bloqueio por Alçada
Public Const BLOQUEIO_ALCADA = 1

'Status de Item de Pedido de Compras
Public Const ITEM_PED_COMPRAS_ABERTO = 0
Public Const ITEM_PED_COMPRAS_RECEBIDO = 1

'Status de Item de Requisição
Public Const ITEM_REQ_ABERTO = 0
Public Const ITEM_REQ_PEDIDO = 1
Public Const ITEM_REQ_RECEBIDO = 2

'Variáveis que armazenam independente de instância da classe
Public ComGlob_objCOM As ClassCOM
Public ComGlob_Refs As Integer

'Número máximo de itens em uma requisição de compras.
'Public Const NUM_MAX_ITENS_REQUISICAO = 100 'Alterado por Wagner
'Número máximo de Bloqueios para liberar.
Public Const NUM_MAX_BLOQUEIOSPC_LIBERACAO = 1000
'Número máximo de Requisições a baixar.
Public Const NUM_MAX_REQUISICOES = 100
'Número máximo de itens de um pedido de cotação.
'Public Const NUM_MAX_ITENS_PEDIDO_COTACAO = 100 'Alterado por Wagner
'Número máximo de itens de um Pedido de Compra
'Public Const NUM_MAX_ITENS_PEDIDO_COMPRAS = 100 'Alterado por Wagner

'Public Const NUM_MAX_ITENS_DISTRIBUICAO = 100 'Alterado por Wagner

'Número máximo de produtos a serem cotados
'Public Const NUM_MAX_PRODUTOS_COTACAO = 100 'Alterado por Wagner

'Número máximo de fornecedores para cotação
'Public Const NUM_MAX_FORNECEDORES_COTACAO = 100 'Alterado por Wagner

Public Const NUM_MAX_PEDIDOS = 100

'Public Const NUM_MAX_NFS_ITEMREQ = 100 'Alterado por Wagner
'Public Const NUM_MAX_PEDIDOS_ITEMREQ = 100 'Alterado por Wagner

Public Const STRING_USUARIO_CONEXAO = 255
Public Const STRING_MOTIVO_ESCOLHA = 30
Public Const STRING_REQUISITANTE_NOMERED = 20
Public Const STRING_REQUISITANTE_CCL = 10
Public Const STRING_REQUISITANTE_NOME = 50

Public Const STRING_BLOQUEIOSPC_RESPONSAVEL = 50

Public Const STRING_TIPODEBLOQUEIOPC_DESCRICAO = 100
Public Const STRING_TIPODEBLOQUEIOPC_NOME_REDUZIDO = 20

Public Const STRING_PEDIDO_COTACAO_DESCRICAO = 50

Public Const STRING_ITENSREQCOMPRA_DESCPRODUTO = 50

Public Const STRING_OBSERVACAO = 255
Public Const STRING_DESCRICAO_TIPOTRIBUTACAO = 100

Public Const STRING_DESCRICAO_REQMODELO = 30

Public Const STRING_NOTASPC_NOTA = 150

'Status utilizados em pedido de cotação
Public Const STATUS_GERADO_NAO_ATUALIZADO = 0
Public Const STATUS_PARCIALMENTE_ATUALIZADO = 1
Public Const STATUS_ATUALIZADO = 2

Public Const CONDPAGTO_VISTA = 1
Public Const CONDPAGTO_PRAZO = 2

''Recebimento fora da faixa
'Public Const MENSAGEM_REJEITA_RECEBIMENTO = "Rejeita Recebimento"
'Public Const MENSAGEM_ACEITA_RECEBIMENTO = "Avisa e Aceita Recebimento"
'Public Const MENSAGEM_NAO_AVISA_ACEITA_RECEBIMENTO = "Não Avisa e Aceita Recebimento"

'Constantes Batch do Calculo de Parametros Pto Pedido
Public Const TITULO_TELA_BACH_CALCULO_PTOPEDIDO = "Cálculo dos Parametros para Ponto Pedido"
Public Const ROTINA_CALCULO_PTOPEDIDO = 1


Type typeRequisitante
   
    lCodigo As Long
    sNome As String
    sNomeReduzido As String
    sCcl As String
    sEmail As String
    sCodUsuario As String
End Type

Type typeTipoBloqueioPC
    
    iCodigo As Integer
    sNomeReduzido As String
    sDescricao As String

End Type

Type typeAlcada
    sCodUsuario As String
    dLimiteOperacao As Double
    dLimiteMensal As Double
End Type


Type typeRequisicaoModelo
    
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    sDescricao As String
    iUrgente As Integer
    lRequisitante As Long
    sCcl As String
    iFilialCompra As Integer
    iTipoDestido As Integer
    lFornCliDestino As Long
    iFilialDestino As Integer
    lObservacao As Long
    sObservacao As String
    iTipoTributacao As Integer

End Type

Type typeItemReqModelo
    lNumIntDoc As Long
    lReqModelo As Long
    sProduto As String
    sDescProduto As String
    dQuantidade As Double
    sUM As String
    sCcl As String
    iAlmoxarifado As Integer
    sContaContabil As String
    iCreditaICMS As Integer
    iCreditaIPI As Integer
    sObservacao As String
    lFornecedor As Long
    iFilial As Integer
    iExclusivo As Integer
    iTipoTributacao As Integer
End Type

Type typeBloqueioPC
    iFilialEmpresa As Integer
    lPedCompras As Long
    iSequencial As Integer
    iTipoDeBloqueio As Integer
    sCodUsuario As String
    sResponsavel As String
    dtData As Date
    sCodUsuarioLib As String
    dtDataLib As Date
End Type

Type typePedidoCompras
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    lFornecedor As Long
    iFilial As Integer
    iComprador As Integer
    sContato As String
    iTipoDestino As Integer
    lFornCliDestino As Long
    iFilialDestino As Integer
    dtData As Date
    dtDataEmissao As Date
    dtDataEnvio As Date
    dtDataAlteracao As Date
    dtDataBaixa As Date
    iCondicaoPagto As Integer
    dOutrasDespesas As Double
    dValorFrete As Double
    dValorSeguro As Double
    dValorDesconto As Double
    dValorTotal As Double
    dValorIPI As Double
    sObservacao As String
    lObservacao As Long
    sTipoFrete As String
    iTransportadora As Integer
    iProxSeqBloqueio As Integer
    iTipoBaixa As Integer
    sMotivoBaixa As String
    sAlcada As String
    dValorProdutos As Double
    iEmbalagem As Integer
    dTaxa As Double
    iMoeda As Integer
    sObsEmbalagem As String
    lCodigoPV As Long
    dtDataRefFluxo As Date
    iTabelaPreco As Integer
    sUsuReg As String
    sUsuRegAprov As String
    dtDataRegAprov As Date
    sUsuRegEnvio As String
End Type

Type typeItemPedCompra
    iMoeda As Integer
    dTaxa As Double
    lNumIntDoc As Long
    dtDataLimite As Date
    sProduto As String
    sDescProduto As String
    dQuantidade As Double
    dQuantRecebida As Double
    dQuantRecebimento As Double
    sUM As String
    dPrecoUnitario As Double
    dValorDesconto As Double
    iTipoOrigem As Integer
    lNumIntOrigem As Long
    dPercentMaisReceb As Double
    dPercentMenosReceb As Double
    iRebebForaFaixa As Integer
    iStatus As Integer
    lObservacao As Long
    sObservacao As String
    dValorIPI As Double
    dAliquotaIPI As Double
    dAliquotaICMS As Double
    lPedCompra As Long
    dtDeliveryDate As Date
    iTempoTransito As Integer
End Type

Type typeLocalizacaoItemPC
    iAlmoxarifado As Integer
    sCcl As String
    dQuantidade As Double
    sContaContabil As String
End Type


Type typeItemConcorrencia
    
    lNumIntDoc As Long
    sProduto As String
    lFornecedor As Long
    iFilial As Integer
    dQuantidade As Double
    sUM As String
    sDescricao As String
    dtDataNecessidade As Date

End Type

Type typeCotacaoItemConcorrencia
    lNumIntDoc As Long
    lItemCotacao As Long
    dValorPresente As Double
    iEscolhido As Integer
    sMotivoEscolha As String
    dQuantidadeComprar As Double
    dtDataEntrega As Date
    dPrecoAjustado As Double
    sFornecedor As String
    sFilial As String
    sCondPagto As String
    dPrecoUnitario As Double
    dCreditoICMS As Double
    dCreditoIPI As Double
    lPedCotacao As Long
    dtDataValidade As Date
    iPrazoEntrega As Integer
    dQuantEntrega As Double
    dPreferencia As Double
    dAliquotaIPI As Double
    dAliquotaICMS As Double
    dTaxa As Double
    iMoeda As Integer
End Type

Type typeItemReqCompra

    lNumIntDoc As Long
    sProduto As String
    sDescProduto As String
    iStatus As Integer
    dQuantidade As Double
    dQuantPedida As Double
    dQuantRecebida As Double
    dQuantCancelada As Double
    sUM As String
    sCcl As String
    iAlmoxarifado As Integer
    sContaContabil As String
    iCreditaICMS As Integer
    iCreditaIPI As Integer
    lObservacao As Long
    sObservacao As String
    lFornecedor As Long
    iFilial As Integer
    iExclusivo As Integer
    dQuantNaConcorrencia As Double
    dQuantNoPedido As Double
    dQuantNoPedidoRecebida As Double
    dQuantNaCotacao As Double
    lReqCompra As Long
    iTipoTributacao As Integer

End Type

Type typeQuantSuplementar
    iTipoDestino As Integer
    lFornCliDestino As Long
    iFilialDestino As Integer
    dQuantidade As Double
End Type

Type typeRequisicaoCompras
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    dtData As Date
    dtDataEnvio As Date
    dtDataLimite As Date
    dtDataBaixa As Date
    lUrgente As Long
    lRequisitante As Long
    sDigitador As String
    sCcl As String
    sOPCodigo As String
    lPVCodigo As Long
    iFilialCompra As Integer
    iTipoDestino As Integer
    lFornCliDestino As Long
    iFilialDestino As Integer
    lObservacao As Long
    iTipoTributacao As Integer
    lNumIntDocItemOP As Long
    sUsuReg As String
    sUsuRegAprov As String
    sUsuRegBaixa As String
    sUsuRegEnvio As String
End Type


Type typeItemPedidoCompraInfo

    lNumIntDoc As Long
    lPedCompra As Long
    dtDataLimite As Date
    sProduto As String
    sDescProduto As String
    dQuantidade As Double
    dQuantRecebida As Double
    dQuantRecebimento As Double
    sUM As String
    dAliquotaICMS As Double
    dAliquotaIPI As Double
    iMoeda As Integer
    dTaxa As Double

End Type

Type typeCotacaoItemConcAux

    iEscolhido As Integer
    sProduto As String
    sDescProduto As String
    iCondPagto As Integer
    dQuantComprarMax As Double
    sUM As String
    dPrecoUnitario As Double
    dPrecoAjustado As Double
    dValorPresente As Double
    iTipoTributacao As Integer
    lFornecedor As Long
    iFilialForn As Integer
    lPedidoCot As Long
    dtDataValidade As Date
    iPrazoEntrega As Integer
    dtDataEntrega As Date
    dtDataNecessidade As Date
    dtDataCotacao As Date
    dQuantidadeEntrega As Double
    iPreferencia As Integer
    dQuantComprar As Double
    sMotivoEscolha As String
    lItemCotacao As Long
    iMoeda As Integer
    dTaxa As Double
    
End Type

Type typeConcorrencia

    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    sDescricao As String
    dtData As Date
    iComprador As Integer
    dTaxaFinanceira As Double
    iTipoDestino As Integer
    lFornCliDestino As Long
    iFilialDestino As Integer

End Type

Type typeItemReqCompraInfo

    lPedCompra As Long
    lReqCompra As Long
    lNumIntDoc As Long
    sProduto As String
    sDescProduto As String
    sUM As String
    dQuantPedida As Double
    iFilialEmpresa As Integer
    lUrgente As Long
    dQuantidade As Double
    dQuantRecebida As Double
    dQuantCancelada As Double
    iTipoTributacao As Integer
    lNumIntDocItemPC As Long
    
End Type

Type typeItemMapaCotacao
    
    lPedCotacao As Long
    sProduto As String
    sDescricao As String
    dQuantidade As Double
    lMapaCotacao As Long
    dtData As Date
    dTaxaFinanceira As Double
    lNumIntItemPedCotacao As Long
    iFilialEmpresa As Integer
    sFornecedor As String
    sUM As String
    iFilialForn As Long
    sNomeFilialForn As String
    
End Type

'Incluído por Luiz Nogueira em 03/03/04
Type TypeRelABCComprasVar
    vdtDataDe As Variant
    vdtDataAte As Variant
    viFilialEmpresaDe As Variant
    viFilialEmpresaAte As Variant
    vsProdutoDe As Variant
    vsProdutoAte As Variant
    viTipoProduto As Variant
    vsCategoria As Variant
    avsItensCategoria() As Variant
    viNumIntRel As Variant
    vsProduto As Variant
    vlRanking As Variant
    vdQuantidade As Variant
    vdValor As Variant
    vdPercParticipacao As Variant
    vsItemCategoria As Variant
End Type

'Incluído por Luiz Nogueira em 03/03/04
Type TypeRelABCFornecedoresVar
    vdtDataDe As Variant
    vdtDataAte As Variant
    viFilialEmpresaDe As Variant
    viFilialEmpresaAte As Variant
    vsProdutoDe As Variant
    vsProdutoAte As Variant
    vlFornecedorDe As Variant
    vlFornecedorAte As Variant
    viTipoProduto As Variant
    vsCategoriaProdutos As Variant
    avsItensCategoriaProdutos() As Variant
    vsCategoriaFornecedores As Variant
    avsItensCategoriaFornecedores() As Variant
    viNumIntRel As Variant
    vlFornecedor As Variant
    viFilialFornecedor As Variant
    vlRanking As Variant
    vdValor As Variant
    vdPercParticipacao As Variant
    vsItemCategoria As Variant
End Type

