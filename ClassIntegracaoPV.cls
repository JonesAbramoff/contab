VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIntegracaoPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntGer As Long
Private mvarlNumIntArq As Long
Private mvarlSeqRegistro As Long
Private mvariTipoInt As Integer
Private mvarlCodPedidoExterno As Long
Private mvardtDataEmissao As Date
Private mvarlCodClienteExterno As Long
Private mvarlCodClienteCorporator As Long
Private mvariCodFilialCorporator As Integer
Private mvarsCGCCliente As String
Private mvarsNaturezaOP As String
Private mvariCodTabelaPreco As Integer
Private mvariCodCondPagto As Integer
Private mvarsNomeCondPagtoExterno As String
Private mvariCodCondPagtoExterno As Integer
Private mvariFilialEmpresaFat As Integer
Private mvardValorDescontoPedido As Double
Private mvariFrete As Integer
Private mvardValorFretePedido As Double
Private mvardValorSeguroPedido As Double
Private mvardValorOutrasDespesasPedido As Double
Private mvariFilialEmpresaEnt As Integer
Private mvariTrazerTranspAuto As Integer
Private mvariCodTransportadora As Integer
Private mvariCodTransportadoraExterno As Integer
Private mvarsNomeTransportadoraExterno As String
Private mvariTrazerMensagemAuto As Integer
Private mvarsMensagemPedido As String
Private mvariTrazerPesoAuto As Integer
Private mvardPesoBruto As Double
Private mvardPesoLiquido As Double
Private mvariTrazerCanalVendaAuto As Integer
Private mvariCanalVenda As Integer
Private mvariCodVendedor As Integer
Private mvariCodVendedorExterno As Integer
Private mvarsNomeVendedorExterno As String
Private mvariTrazerComissaoAuto As Integer
Private mvardValorComissao As Double
Private mvariTrazerReservaAuto As Integer
Private mvariCodAlmoxarifado As Integer
Private mvariCodAlmoxarifadoExterno As Integer
Private mvarsNomeAlmoxarifadoExterno As String
Private mvariItem As Integer
Private mvarsCodProduto As String
Private mvarsProdutoFormatado As String
Private mvarsCodProdutoExterno As String
Private mvariTrazerDescricaoAuto As Integer
Private mvarsDescricaoItem As String
Private mvardQuantidadePedida As Double
Private mvardQuantidadeCancelada As Double
Private mvarsUM As String
Private mvardPrecoUnitario As Double
Private mvardValorDescontoItem As Double
Private mvardtDataEntrega As Date
Private mvariComErro As Integer
Private mvardtDataAtualizacao As Date
Private mvariCodTabelaPrecoItem As Integer
Private mvarsPedRepr As String

Public Property Let sPedRepr(ByVal vData As String)
    mvarsPedRepr = vData
End Property

Public Property Get sPedRepr() As String
    sPedRepr = mvarsPedRepr
End Property

Public Property Let iCodTabelaPrecoItem(ByVal vData As Integer)
    mvariCodTabelaPrecoItem = vData
End Property

Public Property Get iCodTabelaPrecoItem() As Integer
    iCodTabelaPrecoItem = mvariCodTabelaPrecoItem
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntGer(ByVal vData As Long)
    mvarlNumIntGer = vData
End Property

Public Property Get lNumIntGer() As Long
    lNumIntGer = mvarlNumIntGer
End Property

Public Property Let lNumIntArq(ByVal vData As Long)
    mvarlNumIntArq = vData
End Property

Public Property Get lNumIntArq() As Long
    lNumIntArq = mvarlNumIntArq
End Property

Public Property Let lSeqRegistro(ByVal vData As Long)
    mvarlSeqRegistro = vData
End Property

Public Property Get lSeqRegistro() As Long
    lSeqRegistro = mvarlSeqRegistro
End Property

Public Property Let iTipoInt(ByVal vData As Integer)
    mvariTipoInt = vData
End Property

Public Property Get iTipoInt() As Integer
    iTipoInt = mvariTipoInt
End Property

Public Property Let lCodPedidoExterno(ByVal vData As Long)
    mvarlCodPedidoExterno = vData
End Property

Public Property Get lCodPedidoExterno() As Long
    lCodPedidoExterno = mvarlCodPedidoExterno
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let lCodClienteExterno(ByVal vData As Long)
    mvarlCodClienteExterno = vData
End Property

Public Property Get lCodClienteExterno() As Long
    lCodClienteExterno = mvarlCodClienteExterno
End Property

Public Property Let lCodClienteCorporator(ByVal vData As Long)
    mvarlCodClienteCorporator = vData
End Property

Public Property Get lCodClienteCorporator() As Long
    lCodClienteCorporator = mvarlCodClienteCorporator
End Property

Public Property Let iCodFilialCorporator(ByVal vData As Integer)
    mvariCodFilialCorporator = vData
End Property

Public Property Get iCodFilialCorporator() As Integer
    iCodFilialCorporator = mvariCodFilialCorporator
End Property

Public Property Let sCGCCliente(ByVal vData As String)
    mvarsCGCCliente = vData
End Property

Public Property Get sCGCCliente() As String
    sCGCCliente = mvarsCGCCliente
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOP = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOP
End Property

Public Property Let iCodTabelaPreco(ByVal vData As Integer)
    mvariCodTabelaPreco = vData
End Property

Public Property Get iCodTabelaPreco() As Integer
    iCodTabelaPreco = mvariCodTabelaPreco
End Property

Public Property Let iCodCondPagto(ByVal vData As Integer)
    mvariCodCondPagto = vData
End Property

Public Property Get iCodCondPagto() As Integer
    iCodCondPagto = mvariCodCondPagto
End Property

Public Property Let sNomeCondPagtoExterno(ByVal vData As String)
    mvarsNomeCondPagtoExterno = vData
End Property

Public Property Get sNomeCondPagtoExterno() As String
    sNomeCondPagtoExterno = mvarsNomeCondPagtoExterno
End Property

Public Property Let iCodCondPagtoExterno(ByVal vData As Integer)
    mvariCodCondPagtoExterno = vData
End Property

Public Property Get iCodCondPagtoExterno() As Integer
    iCodCondPagtoExterno = mvariCodCondPagtoExterno
End Property

Public Property Let iFilialEmpresaFat(ByVal vData As Integer)
    mvariFilialEmpresaFat = vData
End Property

Public Property Get iFilialEmpresaFat() As Integer
    iFilialEmpresaFat = mvariFilialEmpresaFat
End Property

Public Property Let dValorDescontoPedido(ByVal vData As Double)
    mvardValorDescontoPedido = vData
End Property

Public Property Get dValorDescontoPedido() As Double
    dValorDescontoPedido = mvardValorDescontoPedido
End Property

Public Property Let iFrete(ByVal vData As Integer)
    mvariFrete = vData
End Property

Public Property Get iFrete() As Integer
    iFrete = mvariFrete
End Property

Public Property Let dValorFretePedido(ByVal vData As Double)
    mvardValorFretePedido = vData
End Property

Public Property Get dValorFretePedido() As Double
    dValorFretePedido = mvardValorFretePedido
End Property

Public Property Let dValorSeguroPedido(ByVal vData As Double)
    mvardValorSeguroPedido = vData
End Property

Public Property Get dValorSeguroPedido() As Double
    dValorSeguroPedido = mvardValorSeguroPedido
End Property

Public Property Let dValorOutrasDespesasPedido(ByVal vData As Double)
    mvardValorOutrasDespesasPedido = vData
End Property

Public Property Get dValorOutrasDespesasPedido() As Double
    dValorOutrasDespesasPedido = mvardValorOutrasDespesasPedido
End Property

Public Property Let iFilialEmpresaEnt(ByVal vData As Integer)
    mvariFilialEmpresaEnt = vData
End Property

Public Property Get iFilialEmpresaEnt() As Integer
    iFilialEmpresaEnt = mvariFilialEmpresaEnt
End Property

Public Property Let iTrazerTranspAuto(ByVal vData As Integer)
    mvariTrazerTranspAuto = vData
End Property

Public Property Get iTrazerTranspAuto() As Integer
    iTrazerTranspAuto = mvariTrazerTranspAuto
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
End Property

Public Property Let iCodTransportadoraExterno(ByVal vData As Integer)
    mvariCodTransportadoraExterno = vData
End Property

Public Property Get iCodTransportadoraExterno() As Integer
    iCodTransportadoraExterno = mvariCodTransportadoraExterno
End Property

Public Property Let sNomeTransportadoraExterno(ByVal vData As String)
    mvarsNomeTransportadoraExterno = vData
End Property

Public Property Get sNomeTransportadoraExterno() As String
    sNomeTransportadoraExterno = mvarsNomeTransportadoraExterno
End Property

Public Property Let iTrazerMensagemAuto(ByVal vData As Integer)
    mvariTrazerMensagemAuto = vData
End Property

Public Property Get iTrazerMensagemAuto() As Integer
    iTrazerMensagemAuto = mvariTrazerMensagemAuto
End Property

Public Property Let sMensagemPedido(ByVal vData As String)
    mvarsMensagemPedido = vData
End Property

Public Property Get sMensagemPedido() As String
    sMensagemPedido = mvarsMensagemPedido
End Property

Public Property Let iTrazerPesoAuto(ByVal vData As Integer)
    mvariTrazerPesoAuto = vData
End Property

Public Property Get iTrazerPesoAuto() As Integer
    iTrazerPesoAuto = mvariTrazerPesoAuto
End Property

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property

Public Property Let dPesoLiquido(ByVal vData As Double)
    mvardPesoLiquido = vData
End Property

Public Property Get dPesoLiquido() As Double
    dPesoLiquido = mvardPesoLiquido
End Property

Public Property Let iTrazerCanalVendaAuto(ByVal vData As Integer)
    mvariTrazerCanalVendaAuto = vData
End Property

Public Property Get iTrazerCanalVendaAuto() As Integer
    iTrazerCanalVendaAuto = mvariTrazerCanalVendaAuto
End Property

Public Property Let iCanalVenda(ByVal vData As Integer)
    mvariCanalVenda = vData
End Property

Public Property Get iCanalVenda() As Integer
    iCanalVenda = mvariCanalVenda
End Property

Public Property Let iCodVendedor(ByVal vData As Integer)
    mvariCodVendedor = vData
End Property

Public Property Get iCodVendedor() As Integer
    iCodVendedor = mvariCodVendedor
End Property

Public Property Let iCodVendedorExterno(ByVal vData As Integer)
    mvariCodVendedorExterno = vData
End Property

Public Property Get iCodVendedorExterno() As Integer
    iCodVendedorExterno = mvariCodVendedorExterno
End Property

Public Property Let sNomeVendedorExterno(ByVal vData As String)
    mvarsNomeVendedorExterno = vData
End Property

Public Property Get sNomeVendedorExterno() As String
    sNomeVendedorExterno = mvarsNomeVendedorExterno
End Property

Public Property Let iTrazerComissaoAuto(ByVal vData As Integer)
    mvariTrazerComissaoAuto = vData
End Property

Public Property Get iTrazerComissaoAuto() As Integer
    iTrazerComissaoAuto = mvariTrazerComissaoAuto
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let iTrazerReservaAuto(ByVal vData As Integer)
    mvariTrazerReservaAuto = vData
End Property

Public Property Get iTrazerReservaAuto() As Integer
    iTrazerReservaAuto = mvariTrazerReservaAuto
End Property

Public Property Let iCodAlmoxarifado(ByVal vData As Integer)
    mvariCodAlmoxarifado = vData
End Property

Public Property Get iCodAlmoxarifado() As Integer
    iCodAlmoxarifado = mvariCodAlmoxarifado
End Property

Public Property Let iCodAlmoxarifadoExterno(ByVal vData As Integer)
    mvariCodAlmoxarifadoExterno = vData
End Property

Public Property Get iCodAlmoxarifadoExterno() As Integer
    iCodAlmoxarifadoExterno = mvariCodAlmoxarifadoExterno
End Property

Public Property Let sNomeAlmoxarifadoExterno(ByVal vData As String)
    mvarsNomeAlmoxarifadoExterno = vData
End Property

Public Property Get sNomeAlmoxarifadoExterno() As String
    sNomeAlmoxarifadoExterno = mvarsNomeAlmoxarifadoExterno
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let sProdutoFormatado(ByVal vData As String)
    mvarsProdutoFormatado = vData
End Property

Public Property Get sProdutoFormatado() As String
    sProdutoFormatado = mvarsProdutoFormatado
End Property

Public Property Let sCodProduto(ByVal vData As String)
    mvarsCodProduto = vData
End Property

Public Property Get sCodProduto() As String
    sCodProduto = mvarsCodProduto
End Property

Public Property Let sCodProdutoExterno(ByVal vData As String)
    mvarsCodProdutoExterno = vData
End Property

Public Property Get sCodProdutoExterno() As String
    sCodProdutoExterno = mvarsCodProdutoExterno
End Property

Public Property Let iTrazerDescricaoAuto(ByVal vData As Integer)
    mvariTrazerDescricaoAuto = vData
End Property

Public Property Get iTrazerDescricaoAuto() As Integer
    iTrazerDescricaoAuto = mvariTrazerDescricaoAuto
End Property

Public Property Let sDescricaoItem(ByVal vData As String)
    mvarsDescricaoItem = vData
End Property

Public Property Get sDescricaoItem() As String
    sDescricaoItem = mvarsDescricaoItem
End Property

Public Property Let dQuantidadePedida(ByVal vData As Double)
    mvardQuantidadePedida = vData
End Property

Public Property Get dQuantidadePedida() As Double
    dQuantidadePedida = mvardQuantidadePedida
End Property

Public Property Let dQuantidadeCancelada(ByVal vData As Double)
    mvardQuantidadeCancelada = vData
End Property

Public Property Get dQuantidadeCancelada() As Double
    dQuantidadeCancelada = mvardQuantidadeCancelada
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let dValorDescontoItem(ByVal vData As Double)
    mvardValorDescontoItem = vData
End Property

Public Property Get dValorDescontoItem() As Double
    dValorDescontoItem = mvardValorDescontoItem
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let iComErro(ByVal vData As Integer)
    mvariComErro = vData
End Property

Public Property Get iComErro() As Integer
    iComErro = mvariComErro
End Property

Public Property Let dtDataAtualizacao(ByVal vData As Date)
    mvardtDataAtualizacao = vData
End Property

Public Property Get dtDataAtualizacao() As Date
    dtDataAtualizacao = mvardtDataAtualizacao
End Property

Public Function Desmonta_Registro_PV(sRegistro As String)

Dim lErro As Long
Dim iPOS As Integer
Dim sData As String

On Error GoTo Erro_Desmonta_Registro_PV
    
    '00
    iPOS = 3
    
    lCodPedidoExterno = StrParaLong(Mid(sRegistro, iPOS, 6))
    iPOS = iPOS + 6
    
    sData = Mid(sRegistro, iPOS, 8)
    dtDataEmissao = StrParaDate(left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4))
    iPOS = iPOS + 8
    
    lCodClienteExterno = StrParaLong(Mid(sRegistro, iPOS, 6))
    iPOS = iPOS + 6
    
    lCodClienteCorporator = StrParaLong(Mid(sRegistro, iPOS, 6))
    iPOS = iPOS + 6
    
    iCodFilialCorporator = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    sCGCCliente = Trim(Mid(sRegistro, iPOS, STRING_CGC))
    iPOS = iPOS + STRING_CGC
    
    sNaturezaOp = Trim(Mid(sRegistro, iPOS, STRING_NATUREZAOP_CODIGO))
    iPOS = iPOS + STRING_NATUREZAOP_CODIGO
    
    iCodTabelaPreco = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4

    iCodCondPagto = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    sNomeCondPagtoExterno = Trim(Mid(sRegistro, iPOS, STRING_NOME_EXTERNO_INT))
    iPOS = iPOS + STRING_NOME_EXTERNO_INT

    iCodCondPagtoExterno = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    iFilialEmpresaFat = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    dValorDescontoPedido = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    iFrete = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    dValorFretePedido = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    dValorSeguroPedido = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    dValorOutrasDespesasPedido = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    iFilialEmpresaEnt = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    iTrazerTranspAuto = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    iCodTransportadora = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    iCodTransportadoraExterno = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    sNomeTransportadoraExterno = Trim(Mid(sRegistro, iPOS, STRING_NOME_EXTERNO_INT))
    iPOS = iPOS + STRING_NOME_EXTERNO_INT
    
    iTrazerMensagemAuto = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    sMensagemPedido = Trim(Mid(sRegistro, iPOS, STRING_PEDIDOVENDA_MENSAGEM_NOTA))
    iPOS = iPOS + STRING_PEDIDOVENDA_MENSAGEM_NOTA
    
    iTrazerPesoAuto = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    dPesoBruto = StrParaDbl(Mid(sRegistro, iPOS, 10)) / 100
    iPOS = iPOS + 10
    
    dPesoLiquido = StrParaDbl(Mid(sRegistro, iPOS, 10)) / 100
    iPOS = iPOS + 10
    
    iTrazerCanalVendaAuto = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    iCanalVenda = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    iCodVendedor = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    iCodVendedorExterno = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    sNomeVendedorExterno = Trim(Mid(sRegistro, iPOS, STRING_NOME_EXTERNO_INT))
    iPOS = iPOS + STRING_NOME_EXTERNO_INT
    
    iTrazerComissaoAuto = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    dValorComissao = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    iTrazerReservaAuto = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    iCodAlmoxarifado = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    iCodAlmoxarifadoExterno = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    sNomeAlmoxarifadoExterno = Trim(Mid(sRegistro, iPOS, STRING_NOME_EXTERNO_INT))
    iPOS = iPOS + STRING_NOME_EXTERNO_INT
    
    iItem = StrParaInt(Mid(sRegistro, iPOS, 4))
    iPOS = iPOS + 4
    
    sCodProduto = Trim(Mid(sRegistro, iPOS, STRING_PRODUTO))
    iPOS = iPOS + STRING_PRODUTO
    
    sCodProdutoExterno = Trim(Mid(sRegistro, iPOS, STRING_NOME_EXTERNO_INT))
    iPOS = iPOS + STRING_NOME_EXTERNO_INT
    
    iTrazerDescricaoAuto = StrParaInt(Mid(sRegistro, iPOS, 1))
    iPOS = iPOS + 1
    
    sDescricaoItem = Trim(Mid(sRegistro, iPOS, STRING_ITEM_PEDIDO_DESCRICAO))
    iPOS = iPOS + STRING_ITEM_PEDIDO_DESCRICAO

    dQuantidadePedida = StrParaDbl(Mid(sRegistro, iPOS, 10)) / 100
    iPOS = iPOS + 10
    
    dQuantidadeCancelada = StrParaDbl(Mid(sRegistro, iPOS, 10)) / 100
    iPOS = iPOS + 10
    
    sUM = Trim(Mid(sRegistro, iPOS, STRING_UM_SIGLA))
    iPOS = iPOS + STRING_UM_SIGLA
    
    dPrecoUnitario = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    dValorDescontoItem = StrParaDbl(Mid(sRegistro, iPOS, 12)) / 100
    iPOS = iPOS + 12
    
    sData = Mid(sRegistro, iPOS, 8)
    dtDataEntrega = StrParaDate(left(sData, 2) & "/" & Mid(sData, 3, 2) & "/" & right(sData, 4))
    iPOS = iPOS + 8

    Desmonta_Registro_PV = SUCESSO

    Exit Function

Erro_Desmonta_Registro_PV:

    Desmonta_Registro_PV = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190044)

    End Select

    Exit Function
    
End Function
