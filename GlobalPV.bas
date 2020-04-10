Attribute VB_Name = "GlobalPV"
Option Explicit


'Indica se está faturando tudo ou deixa os pedidos como estão (PedidoDeVenda_Baixar1)
Public Const NAO_FATURA_TUDO = 0
Public Const FATURA_TUDO = 1

Public Const STRING_IPICODIGOPRODUTO_TRIBUTACAO = 20

'para tabela de Pedido de Venda
'Public Const STRING_PED_VENDA_NATUREZA = 3
Public Const STRING_PEDIDOVENDA_VOLUME_ESPECIE = 20
Public Const STRING_PEDIDOVENDA_VOLUME_MARCA = 20
Public Const STRING_PEDIDOVENDA_PLACA = 10
Public Const STRING_PEDIDOVENDA_PLACA_UF = 2

Public Const PV_STATUS_ANDAMENTO_NAO_RESERVOU_TUDO = 1
Public Const PV_STATUS_ANDAMENTO_RESERVOU_TUDO = 2
Public Const PV_STATUS_ANDAMENTO_FAT_PARCIAL = 3
Public Const PV_STATUS_ANDAMENTO_FAT_TOTAL = 4

Public Const PV_STATUS_ANDAMENTO_ORD_NAO_RESERVOU_TUDO = 1
Public Const PV_STATUS_ANDAMENTO_ORD_RESERVOU_TUDO = 2
Public Const PV_STATUS_ANDAMENTO_ORD_FAT_PARCIAL = 3
Public Const PV_STATUS_ANDAMENTO_ORD_FAT_TOTAL = 4

Public Const PV_STATUS_ANDAMENTO_TIPO_PV_GRAVA = 0
Public Const PV_STATUS_ANDAMENTO_TIPO_NF_GRAVA = 1
Public Const PV_STATUS_ANDAMENTO_TIPO_RESERVA_GRAVA = 2

Public Const OV_VERSAO_NAO_GRAVA = 0
Public Const OV_VERSAO_PERGUNTA = 1
Public Const OV_VERSAO_GRAVA = 2

Public Const OV_DATA_ENTREGA_DATA = 1
Public Const OV_DATA_ENTREGA_PRAZO_DIAS_UTEIS = 0
Public Const OV_DATA_ENTREGA_PRAZO_DIAS_CORRIDOS = 2
Public Const OV_DATA_ENTREGA_PRAZO_SEMANAS = 3
Public Const OV_DATA_ENTREGA_PRAZO_MESES = 4
Public Const OV_DATA_ENTREGA_TEXTO = 5

'Indica que um bloqueio de crédito está liberado. Usado no teste de bloqueio de credito.
Public Const BLOQUEIO_CREDITO_LIBERADO = 1
'Indica que um bloqueio por Atraso de Pagamento está liberado.
Public Const BLOQUEIO_POR_ATRASO_LIBERADO = 1 'Incluido por Leo em 22/02/02
'Indica que um bloqueio está liberado.
Public Const BLOQUEIOPV_LIBERADO = 1 'Incluido por Leo em 22/02/02
Public Const BLOQUEIO_ATRASO_NAO_BLOQUEAR = 0 'Quando a empresa não controla dias de atraso de parcelas p/ bloqueio.

Public Const SIGLA_PV_NORMAL = "PVN"
Public Const CODIGO_PV_NORMAL = 50

Public Const SIGLA_PSRV_NORMAL = "PSRVN"
Public Const CODIGO_PSRV_NORMAL = 216

Public Const STRING_ITEM_PEDIDO_LOTE = 10
Public Const STRING_ITEM_PEDIDO_DESCRICAO = 250

Public Const STRING_RESERVA_RESPONSAVEL = 50
Public Const STRING_PEDIDOVENDA_MENSAGEM_NOTA = STRING_NFISCAL_MENSAGEM
Public Const STRING_PEDIDOVENDA_PEDIDO_CLIENTE = 20
Public Const STRING_PEDIDOVENDA_VOLUME_NUMERO = 20
Public Const STRING_BLOQUEIOSPV_COD_USUARIO = 10
Public Const STRING_BLOQUEIOSPV_RESPONSAVEL = 50
Public Const STRING_BLOQUEIOSPV_OBSERVACAO = 250
Public Const STRING_PEDIDOVENDA_PEDIDO_REPRESENTANTE = 20

Type typeItemPedido
    iFilialEmpresa As Integer
    lCodPedido As Long
    sProduto As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dPrecoTotal As Double
    sUnidadeMed As String
    dValorDesconto As Double
    dtDataEntrega As Date
    sProdutoDescricao As String
    sLote As String
    dValorAbatComissao As Double
    dQuantCancelada As Double
    dQuantReservada As Double
    colReservaItem As ColReserva
    sProdutoNomeReduzido As String
    sUMEstoque As String
    iClasseUM As Integer
    dQuantFaturada As Double
    dQuantAFaturar As Double
    dQuantOP As Double
    dQuantSC As Double
    sDescricao As String
    iStatus As Integer
    iControleEstoque As Integer
    lNumIntDoc As Long
    dPercDesc1 As Double
    iTipoDesc1 As Integer
    dPercDesc2 As Double
    iTipoDesc2 As Integer
    dPercDesc3 As Double
    iTipoDesc3 As Integer
    iPeca As Integer
    dPercentMaisReceb As Double
    dPercentMenosReceb As Double
    iRecebForaFaixa As Integer
    dQuantFaturadaAMais As Double
    iPrioridade As Integer
    iTabelaPreco As Integer
    dComissao As Double
End Type

Type typeReserva
    lNumIntDoc As Long
    iFilialEmpresa As Integer
    lCodigo As Long
    sProduto As String
    iAlmoxarifado As Integer
    iTipoDoc As Integer
    lDocOrigem As Long
    lNumIntOrigem As Long
    dQuantidade As Double
    dtDataReserva As Date
    dtDataValidade As Date
    sCodUsuario As String
    sResponsavel As String
End Type

Type typeTributacaoItemPV
'    iFilialEmpresa As Integer
'    lCodPedido As Long
'    iItem As Integer
    lNumIntDoc As Long
    sNaturezaOp As String
    iNaturezaOpManual As Integer
    iTipoTributacao As Integer
    iTipoTributacaoManual As Integer
    iIPITipo As Integer
    iIPITipoManual As Integer
    dIPIBaseCalculo As Double
    iIPIBaseManual As Integer
    dIPIPercRedBase As Double
    iIPIPercRedBaseManual As Integer
    dIPIAliquota As Double
    iIPIAliquotaManual As Integer
    dIPIValor As Double
    iIPIValorManual As Integer
    iICMSTipo As Integer
    iICMSTipoManual As Integer
    dICMSBase As Double
    iICMSBaseManual As Integer
    dICMSPercRedBase As Double
    iICMSPercRedBaseManual As Integer
    dICMSAliquota As Double
    iICMSAliquotaManual As Integer
    dICMSValor As Double
    iICMSValorManual As Integer
    dICMSSubstBase As Double
    iICMSSubstBaseManual As Integer
    dICMSSubstAliquota As Double
    iICMSSubstAliquotaManual As Integer
    dICMSSubstValor As Double
    iICMSSubstValorManual As Integer
    dICMSSubstPercRedBase As Double
    dICMSSubstPercMVA As Double
    dICMSCredito As Double
    dPISCredito As Double
    dCOFINSCredito As Double
    dIPICredito As Double
End Type

Type typeTributacaoComplPV
    iFilialEmpresa As Integer
    lCodPedido As Long
    iItem As Integer
    sNaturezaOp As String
    iNaturezaOpManual As Integer
    iTipoTributacao As Integer
    iTipoTributacaoManual As Integer
    iIPITipo As Integer
    iIPITipoManual As Integer
    dIPIBaseCalculo As Double
    iIPIBaseManual As Integer
    dIPIPercRedBase As Double
    iIPIPercRedBaseManual As Integer
    dIPIAliquota As Double
    iIPIAliquotaManual As Integer
    dIPIValor As Double
    iIPIValorManual As Integer
    iICMSTipo As Integer
    iICMSTipoManual As Integer
    dICMSBase As Double
    iICMSBaseManual As Integer
    dICMSPercRedBase As Double
    iICMSPercRedBaseManual As Integer
    dICMSAliquota As Double
    iICMSAliquotaManual As Integer
    dICMSValor As Double
    iICMSValorManual As Integer
    dICMSSubstBase As Double
    iICMSSubstBaseManual As Integer
    dICMSSubstAliquota As Double
    iICMSSubstAliquotaManual As Integer
    dICMSSubstValor As Double
    iICMSSubstValorManual As Integer
    dICMSCredito As Double
    dPISCredito As Double
    dCOFINSCredito As Double
    dIPICredito As Double
End Type

Type typePedidoVenda
    iFilialEmpresa As Integer
    lCodigo As Long
    iFilialEmpresaFaturamento As Integer
    lCliente As Long
    iFilial As Integer
    iFilialEntrega As Integer
    iTipoPedido As Integer
    iCodTransportadora As Integer
    iCondicaoPagto As Integer
    dPercAcrescFinanceiro As Double
    dValorProdutos As Double
    dtDataEmissao As Date
    sMensagemNota As String
    sNaturezaOp As String
    dValorTotal As Double
    dValorFrete As Double
    dValorDesconto As Double
    dValorSeguro As Double
    dValorOutrasDespesas As Double
    sPedidoCliente As String
    iCanalVenda As Integer
    iTabelaPreco As Integer
    iProxSeqBloqueio As Integer
    iFaturaIntegral As Integer
    iCobrancaAutomatica As Integer
    iComissaoAutomatica As Integer
    iFreteRespons As Integer
    dtDataReferencia As Date
    lNumIntDoc As Long
    dPesoLiq As Double
    dPesoBruto As Double
    lVolumeQuant As Long
    lVolumeEspecie As Long
    lVolumeMarca As Long
    sVolumeNumero As String
    sPlaca As String
    sPlacaUF As String
    dtDataEntrega As Date
    iCodTranspRedesp As Integer
    iDetPagFrete As Integer
    dVolumeTotal As Double
    iMoeda As Integer
    dTaxaMoeda As Double
    sPedidoRepresentante As String
    dtDataRefFluxo As Date
    iStatus As Integer
    lNumIntSolicSRV As Long
    iAndamento As Integer
    sOBS As String
    dValorDescontoTit As Double
    dValorDescontoItens As Double
    dValorItens As Double
    lCodigoBase As Long
    iParc As Integer
    sEmitente As String
    sUsuarioUltAlteracao As String
    dtDataInclusao As Date
    dtDataAlteracao As Date
    dHoraInclusao As Double
    dHoraAlteracao As Double
End Type

Type typeComissaoPedVenda
    iFilialEmpresa As Integer
    lPedidoDeVendas As Long
    iCodVendedor As Integer
    dValorBase As Double
    dPercentual As Double
    dValor As Double
    dPercentualEmissao As Double
    dValorEmissao As Double
    iIndireta As Integer
    iSeq As Integer
End Type

Type typeComissaoPorItem
    iTipoDoc As Integer
    lNumIntDocItem As Long
    iSeqComissao As Integer
    dValorBase As Double
    dPercentual As Double
    dValor As Double
    iLinha As Integer
    dPercentualEmissao As Double
    dValorEmissao As Double
End Type

Type typeParcelaPedidoVenda
    dValor As Double
    dtDataVencimento As Date
    iNumParcela As Integer
    iDesconto1Codigo As Integer
    dtDesconto1Ate As Date
    dDesconto1Valor As Double
    iDesconto2Codigo As Integer
    dtDesconto2Ate As Date
    dDesconto2Valor As Double
    dtDesconto3Ate As Date
    dDesconto3Valor As Double
    iDesconto3Codigo As Integer
    iTipoPagto As Integer
    iCodConta As Integer
    dtDataCredito As Date
    dtDataEmissaoCheque As Date
    iBancoCheque As Integer
    sAgenciaCheque As String
    sContaCorrenteCheque As String
    lNumeroCheque As Long
    dtDataDepositoCheque As Date
    iAdmMeioPagto As Integer
    iParcelamento As Integer
    sNumeroCartao As String
    dtValidadeCartao As Date
    sAprovacaoCartao As String
    dtDataTransacaoCartao As Date
    
End Type


Type typeSldMesFat
    iAno As Integer
    iFilialEmpresa As Integer
    sProduto As String
    adQuantFaturada(1 To 12) As Double
    adValorFaturado(1 To 12) As Double
    adQuantDevolvida(1 To 12) As Double
    adValorDevolvido(1 To 12) As Double
    adTotalDescontos(1 To 12) As Double
    adQuantPedida(1 To 12) As Double
    adQuantPedidaSRV(1 To 12) As Double
End Type

Type typeSldDiaFat
    iFilialEmpresa As Integer
    sProduto As String
    dtData As Date
    dValorFaturado As Double
    dTotalDescontos As Double
    dQuantPedida As Double
    dQuantFaturada As Double
    dValorDevolvido As Double
    dQuantDevolvida As Double
    dQuantPedidaSRV As Double
End Type


Type typeItemPedidoNF
    iFilialEmpresa As Integer
    lCodPedido As Long
    sProduto As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dPrecoTotal As Double
    dValorDesconto As Double
    iClasseUM As Integer
    dQuantFaturada As Double
    sDescricao As String
    lNumIntDoc As Long
    sUMItemNF As String
    sUMItemPV As String
    lNumNF As Long
    iItemNF As Integer
    iItemPV As Integer
    dPercDesconto As Double
    sSerie As String
    dtDataEmissao As Date
End Type

Type typeTributacaoPV
    iFilialEmpresa As Integer
    lCodPedido As Long
    iTaxacaoAutomatica As Integer
    iTipoTributacao As Integer
    iTipoTributacaoManual As Integer
    dICMSBase As Double
    iICMSBaseManual As Integer
    dICMSValor As Double
    iICMSValorManual As Integer
    dICMSSubstBase As Double
    iICMSSubstBaseManual As Integer
    dICMSSubstValor As Double
    iICMSSubstValorManual As Integer
    dIPIBase As Double
    iIPIBaseManual As Integer
    dIPIValor As Double
    iIPIValorManual As Integer
    dIRRFBase As Double
    dIRRFAliquota As Double
    iIRRFAliquotaManual As Integer
    dIRRFValor As Double
    iIRRFValorManual As Integer
    iISSIncluso As Integer
    dISSBase As Double
    dISSAliquota As Double
    iISSAliquotaManual As Integer
    dISSValor As Double
    iISSValorManual As Integer
    iPISRetidoManual As Integer
    iISSRetidoManual As Integer
    iCOFINSRetidoManual As Integer
    iCSLLRetidoManual As Integer
    dPISRetido As Double
    dISSRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
    dValorINSS As Double
    iINSSValorManual As Integer
    iINSSRetido As Integer
    iINSSRetidoManual As Integer
    dINSSBase As Double
    iINSSBaseManual As Integer
    dINSSDeducoes As Double
    iINSSDeducoesManual As Integer
    dPISCredito As Double
    iPISCreditoManual As Integer
    dCOFINSCredito As Double
    iCOFINSCreditoManual As Integer
    dICMSCredito As Double
    iICMSCreditoManual As Integer
    dIPICredito As Double
    iIPICreditoManual As Integer
End Type

Type typeTribComplOV
    lNumIntDoc As Long
    iTipo As Integer
    sNaturezaOp As String
    iTipoTributacao As Integer
    iIPITipo As Integer
    dIPIBaseCalculo As Double
    dIPIPercRedBase As Double
    dIPIAliquota As Double
    dIPIValor As Double
    dIPICredito As Double
    iICMSTipo As Integer
    dICMSBase As Double
    dICMSPercRedBase As Double
    dICMSAliquota As Double
    dICMSValor As Double
    dICMSCredito As Double
    dICMSSubstBase As Double
    dICMSSubstAliquota As Double
    dICMSSubstValor As Double
End Type

Type typeTributacaoOV

    iFilialEmpresa As Integer
    lCodOrcamento As Long
    iTaxacaoAutomatica As Integer
    iTipoTributacao As Integer
    iTipoTributacaoManual As Integer
    dIPIBase As Double
    iIPIBaseManual As Integer
    dIPIValor As Double
    iIPIValorManual As Integer
    dICMSBase As Double
    iICMSBaseManual As Integer
    dICMSValor As Double
    iICMSValorManual As Integer
    dICMSSubstBase As Double
    iICMSSubstBaseManual As Integer
    dICMSSubstValor As Double
    iICMSSubstValorManual As Integer
    iISSIncluso As Integer
    dISSBase As Double
    dISSAliquota As Double
    iISSAliquotaManual As Integer
    dISSValor As Double
    iISSValorManual As Integer
    dIRRFBase As Double
    dIRRFAliquota As Double
    iIRRFAliquotaManual As Integer
    dIRRFValor As Double
    iIRRFValorManual As Integer
    iPISRetidoManual As Integer
    iISSRetidoManual As Integer
    iCOFINSRetidoManual As Integer
    iCSLLRetidoManual As Integer
    dPISRetido As Double
    dISSRetido As Double
    dCOFINSRetido As Double
    dCSLLRetido As Double
    dValorINSS As Double
    iINSSValorManual As Integer
    iINSSRetido As Integer
    iINSSRetidoManual As Integer
    dINSSBase As Double
    iINSSBaseManual As Integer
    dINSSDeducoes As Double
    iINSSDeducoesManual As Integer
    dPISCredito As Double
    iPISCreditoManual As Integer
    dCOFINSCredito As Double
    iCOFINSCreditoManual As Integer
    dICMSCredito As Double
    iICMSCreditoManual As Integer
    dIPICredito As Double
    iIPICreditoManual As Integer
End Type

Type typeParcelaOV
    dValor As Double
    dtDataVencimento As Date
    iNumParcela As Integer
    iDesconto1Codigo As Integer
    dtDesconto1Ate As Date
    dDesconto1Valor As Double
    iDesconto2Codigo As Integer
    dtDesconto2Ate As Date
    dDesconto2Valor As Double
    dtDesconto3Ate As Date
    dDesconto3Valor As Double
    iDesconto3Codigo As Integer
End Type

Type typeOrcamentoVenda
    iFilialEmpresa As Integer
    lCodigo As Long
    lCliente As Long
    iFilial As Integer
    iCondicaoPagto As Integer
    dPercAcrescFinanceiro As Double
    dValorProdutos As Double
    dtDataEmissao As Date
    sNaturezaOp As String
    dValorTotal As Double
    dValorFrete As Double
    dValorDesconto As Double
    dValorSeguro As Double
    dValorOutrasDespesas As Double
    iTabelaPreco As Integer
    dtDataReferencia As Date
    lNumIntDoc As Long
    iVendedor As Integer
    iVendedor2 As Integer
    sNomeFilialCli As String
    sNomeCli As String
    lNumIntNFiscal As Long
    lNumIntPedVenda As Long
    iCobrancaAutomatica As Integer
    sUsuario As String
    lCodigoBase As Long
    lStatus As Long
    lMotivoPerda As Long
    iStatusComercial As Integer
    lNumIntSolicSRV As Long
    iVersao As Integer
    dtDataUltAlt As Date
    dHoraUltAlt As Double
    dValorDescontoTit As Double
    dValorDescontoItens As Double
    dValorItens As Double
    sContato As String
    sEmail As String
    iFilialEntrega As Integer
    iPrazoEntrega As Integer
    iCodTransportadora As Integer
    iCodTranspRedesp As Integer
    sMensagemNota As String
    sPedidoCliente As String
    sPedidoRepresentante As String
    iCanalVenda As Integer
    dPesoBruto As Double
    dPesoLiq As Double
    sPlaca As String
    sPlacaUF As String
    lVolumeQuant As Long
    lVolumeEspecie As Long
    lVolumeMarca As Long
    sVolumeNumero As String
    dVolumeTotal As Double
    dtDataEnvio As Date
    dtDataEntrega As Date
    iDetPagFrete As Integer
    iFreteRespons As Integer
    iDataEnt As Integer
    iMoeda As Integer
    dtDataPerda As Date
    sPrazoTexto As String
    dCotacao As Double
    dtDataPrevReceb As Date
    dtDataProxCobr As Date
    iIdioma As Integer
    dPercParticVend2 As Double
End Type

Type typeItemOV
    iFilialEmpresa As Integer
    lCodPedidoOV As Long
    sProduto As String
    dQuantidade As Double
    dPrecoUnitario As Double
    dPrecoTotal As Double
    sUnidadeMed As String
    dValorDesconto As Double
    dtDataEntrega As Date
    sProdutoDescricao As String
    sProdutoNomeReduzido As String
    sUMEstoque As String
    iClasseUM As Integer
    sDescricao As String
    iStatus As Integer
    iControleEstoque As Integer
    lNumIntDoc As Long
    sVersaoKit As String
    sVersaoKitBase As String
    lStatus As Long
    lMotivoPerda As Long
    sObservacao As String
    iPCSituacao As Integer
    dPCPrecoUnitCalc As Double
    iPrazoEntrega As Integer
    iMarcado As Integer
End Type

Type typeItemPVEntrega
    lNumIntDoc As Long
    lNumIntPV As Long
    lNumIntItemPV As Long
    dtDataEntrega As Date
    dQuantidade As Double
    sPedidoCliente As String
End Type

Public Const STRING_FILIAISEMPRESA_NOME = 50

Public Const ITEMOV_PCSITUACAO_NAO_COTAR = 0
Public Const ITEMOV_PCSITUACAO_COTAR = 1
Public Const ITEMOV_PCSITUACAO_EM_COTACAO = 2
Public Const ITEMOV_PCSITUACAO_COTADO = 3

Public Const ITEMOV_PCSITUACAO_STRING_NAO_COTAR = ""
Public Const ITEMOV_PCSITUACAO_STRING_COTAR = "Cotar"
Public Const ITEMOV_PCSITUACAO_STRING_EM_COTACAO = "Em Cotação"
Public Const ITEMOV_PCSITUACAO_STRING_COTADO = "Cotado"

Public Const OV_STATUS_COMERCIAL_NAO_COTAR = 0
Public Const OV_STATUS_COMERCIAL_EM_COTACAO = 1
Public Const OV_STATUS_COMERCIAL_COTADO = 2
Public Const OV_STATUS_COMERCIAL_REVISADO = 3
Public Const OV_STATUS_COMERCIAL_LIBERADO = 4

Public Const STRING_OV_STATUS_COMERCIAL_NAO_COTAR = ""
Public Const STRING_OV_STATUS_COMERCIAL_EM_COTACAO = "Em Cotação"
Public Const STRING_OV_STATUS_COMERCIAL_COTADO = "Cotado"
Public Const STRING_OV_STATUS_COMERCIAL_REVISADO = "Aguardando Liberação"
Public Const STRING_OV_STATUS_COMERCIAL_LIBERADO = "Liberado"

Public Const FPORIGEM_ITEMOV = 0

Public Const FPSITUACAO_NAO_COTAR = 0
Public Const FPSITUACAO_COTAR = 1
Public Const FPSITUACAO_EM_COTACAO = 2
Public Const FPSITUACAO_COTADO = 3

Public Const FPSITUACAO_STRING_NAO_COTAR = ""
Public Const FPSITUACAO_STRING_COTAR = "Cotar"
Public Const FPSITUACAO_STRING_EM_COTACAO = "Em Cotação"
Public Const FPSITUACAO_STRING_COTADO = "Cotado"

Type typeItemFormPreco
    lNumIntDoc As Long 'na tabela de ItensFormPreco
    iTipoDocOrigem As Integer '0:item de orcamento de venda
    lNumIntDocOrigem As Long 'correspondente ao tipodocorigem
    iSequencial As Integer 'para que possa ser listado sempre numa mesma sequencia
    sProduto As String  'é o insumo/componente/proprio item
    sUnidMed As String
    dQtde As Double 'deste produto que será usada no produto para o qual está sendo calculado o preço
    dCustoUnit As Double 'obtido da cotação
    dPercentMargem As Double 'margem a ser aplicada sobre o custo para obter o preço
    dPrecoUnit As Double 'calculado: custo unit * (1+margem)
    dPrecoTotal As Double 'calculado: qtde * precounit
    iSituacao As Integer '1:cotando, 2:cotado
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typeItemOPItemPV
    lNumIntDocItemOP As Long
    lNumIntDocItemPV As Long
    sCodigoOP As String
    lCodigoPV As Long
    iFilialEmpresa As Integer
    sProduto As String
    dQuantidade As Double
    sUM As String
    iPrioridade As Integer
    dQuantidadeProd As Double
End Type

'TYPE CRIADO AUTOMATICAMENTE PELA TELA BROWSECRIA
Type typePVAndamento
    iCodigo As Integer
    sDescricao As String
    iAuto As Integer
    iFatorAuto As Integer
End Type
