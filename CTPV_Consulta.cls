VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTPV_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjAnotacao As ClassAnotacoes 'Inserido por Wagner

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjTelaProjetoInfo As ClassTelaPRJInfo
Public gobjTribTab As New ClassTribTab
Public gobjPedidoVenda As ClassPedidoDeVenda 'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
    'todos os dados relevantes p/tributacao dentro de gobjPedidoVenda estarao sincronizados com a tela antes da atualizacao da tributacao

'Inicia objeto associado a GridComissoes
Dim objTabComissoes As New ClassTabComissoes
Dim gcolItemPedidoNF_Nota As Collection
Dim gcolItemPedidoNF_ItemPV As Collection

Public giFrameAtual As Integer
'Dim giFrameAtualTributacao As Integer
Dim giGridNotaAtual As Integer
Public giLinhaAnterior As Integer
Public iAlterado As Integer
Dim FrameL(1 To 9) As Frame

'Ordem das Colunas do Grid Itens
Public iGrid_Item_Col As Integer
Public iGrid_ProdutoAlmox_Col As Integer
Public iGrid_UMEstoque_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_QuantReservar_Col As Integer
Public iGrid_QuantReserv_Col As Integer
Public iGrid_Validade_Col As Integer
Public iGrid_Responsavel_Col As Integer
Public iGrid_ItemProduto_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescProduto_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_QuantPedida_Col As Integer
Public iGrid_QuantCancel_Col As Integer
Public iGrid_PrecoUnitario_Col As Integer
Public iGrid_PercDesc_Col As Integer
Public iGrid_Desconto_Col As Integer
Public iGrid_PrecoTotal_Col As Integer
Public iGrid_DataEntrega_Col As Integer
Public iGrid_QuantReservada_Col As Integer
Public iGrid_QuantFaturada_Col As Integer
Public iGrid_TabPrecoItem_Col As Integer
Public iGrid_ComissaoItem_Col As Integer

'Ordem das Colunas do Grid Parcelas
Public iGrid_Vencimento_Col  As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_Desc1Codigo_Col As Integer
Public iGrid_Desc1Ate_Col As Integer
Public iGrid_Desc1Valor_Col As Integer
Public iGrid_Desc1Perc_Col As Integer
Public iGrid_Desc2Codigo_Col As Integer
Public iGrid_Desc2Ate_Col As Integer
Public iGrid_Desc2Valor_Col As Integer
Public iGrid_Desc2Perc_Col As Integer
Public iGrid_Desc3Codigo_Col As Integer
Public iGrid_Desc3Ate_Col As Integer
Public iGrid_Desc3Valor_Col As Integer
Public iGrid_Desc3Perc_Col As Integer

'Ordem das Colunas do Grid Bloqueio
Public iGrid_TipoBloqueio_Col As Integer
Public iGrid_DataBloqueio_Col As Integer
Public iGrid_CodUsuario_Col As Integer
Public iGrid_ResponsavelBL_Col As Integer
Public iGrid_DataLiberacao_Col As Integer
Public iGrid_ResponsavelLib_Col As Integer
Public iGrid_Observacao_Col As Integer

'Ordem das Colunas do Grid de Itens PVNF
Public iGrid_ItemPV_Col As Integer
Public iGrid_ProdutoPV_Col As Integer
Public iGrid_DescProdutoPV_Col As Integer
Public iGrid_UnidadeMedPV_Col As Integer
Public iGrid_QuantFaturadaPV_Col As Integer
Public iGrid_SeriePV_Col As Integer
Public iGrid_NFiscalPV_Col As Integer
Public iGrid_ItemNFPV_Col As Integer
Public iGrid_QuantidadePV_Col As Integer
Public iGrid_PrecoUnitarioPV_Col As Integer
Public iGrid_PercDescPV_Col As Integer
Public iGrid_DescontoPV_Col As Integer
Public iGrid_PrecoTotalPV_Col As Integer

'Ordem das Colunas do Grid de Itens NFPV
Public iGrid_SerieNF_Col As Integer
Public iGrid_NFiscalNF_Col As Integer
Public iGrid_ItemNF_Col As Integer
Public iGrid_ItemPVNF_Col As Integer
Public iGrid_ProdutoNF_Col As Integer
Public iGrid_DescProdutoNF_Col As Integer
Public iGrid_UnidadeMedNF_Col As Integer
Public iGrid_QuantidadeNF_Col As Integer
Public iGrid_PrecoUnitarioNF_Col As Integer
Public iGrid_PercDescNF_Col As Integer
Public iGrid_DescontoNF_Col As Integer
Public iGrid_PrecoTotalNF_Col As Integer

Public objGridComissoes As AdmGrid

Dim objGridItens As AdmGrid
Dim objGridParcelas As AdmGrid
Dim objGridBloqueio As AdmGrid
Dim objGridReserva As AdmGrid
Dim objGridNFItemPV As AdmGrid
Dim objGridNFiscal As AdmGrid

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
'
'Dim giRecalculandoTributacao As Integer
'Dim gcolTiposTribICMS As New Collection
'Dim gcolTiposTribIPI As New Collection

'Constantes públicas dos tabs
Private Const TAB_DadosPrincipais = 1
Private Const TAB_Itens = 2
Private Const TAB_Complemento = 3
Private Const TAB_Cobranca = 4
Private Const TAB_Comissoes = 5
Private Const TAB_Bloqueio = 6
Private Const TAB_Almoxarifado = 7
Private Const TAB_NotasFiscais = 8
Dim TAB_Tributacao As Integer
Private Const TAB_Contabilizacao = 10

'--- inicio dos properties get dos controles da tela

Public Property Get PedidoRepr() As Object
     Set PedidoRepr = objUserControl.Controls("PedidoRepr")
End Property

Public Property Get TabPrecoItemPV() As Object
     Set TabPrecoItemPV = objUserControl.Controls("TabPrecoItemPV")
End Property

Public Property Get ComissaoItemPV() As Object
     Set ComissaoItemPV = objUserControl.Controls("ComissaoItemPV")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

'Public Property Get FrameTributacao() As Object
'     Set FrameTributacao = objUserControl.Controls("FrameTributacao")
'End Property
'
'Public Property Get Frame11() As Object
'     Set Frame11 = objUserControl.Controls("Frame11")
'End Property
'
'Public Property Get ISSIncluso() As Object
'     Set ISSIncluso = objUserControl.Controls("ISSIncluso")
'End Property
'
'Public Property Get ISSValor() As Object
'     Set ISSValor = objUserControl.Controls("ISSValor")
'End Property
'
'Public Property Get ISSValor1() As Object
'     Set ISSValor1 = objUserControl.Controls("ISSValor1")
'End Property
'
'Public Property Get ISSAliquota() As Object
'     Set ISSAliquota = objUserControl.Controls("ISSAliquota")
'End Property
'
'Public Property Get ISSBase() As Object
'     Set ISSBase = objUserControl.Controls("ISSBase")
'End Property
'
'Public Property Get ValorIRRF() As Object
'     Set ValorIRRF = objUserControl.Controls("ValorIRRF")
'End Property
'
'Public Property Get PISRetido() As Object
'     Set PISRetido = objUserControl.Controls("PISRetido")
'End Property
'
'Public Property Get COFINSRetido() As Object
'     Set COFINSRetido = objUserControl.Controls("COFINSRetido")
'End Property
'
'Public Property Get CSLLRetido() As Object
'     Set CSLLRetido = objUserControl.Controls("CSLLRetido")
'End Property
'
'Public Property Get ISSRetido() As Object
'     Set ISSRetido = objUserControl.Controls("ISSRetido")
'End Property
'
'Public Property Get IRAliquota() As Object
'     Set IRAliquota = objUserControl.Controls("IRAliquota")
'End Property
'
'Public Property Get IRBase() As Object
'     Set IRBase = objUserControl.Controls("IRBase")
'End Property
'
'Public Property Get IPIBase() As Object
'     Set IPIBase = objUserControl.Controls("IPIBase")
'End Property
'
'Public Property Get IPIValor() As Object
'     Set IPIValor = objUserControl.Controls("IPIValor")
'End Property
'
'Public Property Get Frame9() As Object
'     Set Frame9 = objUserControl.Controls("Frame9")
'End Property
'
'Public Property Get Frame10() As Object
'     Set Frame10 = objUserControl.Controls("Frame10")
'End Property

Public Property Get Cubagem() As Object
     Set Cubagem = objUserControl.Controls("Cubagem")
End Property

Public Property Get IPIValor1() As Object
     Set IPIValor1 = objUserControl.Controls("IPIValor1")
End Property

'Public Property Get ICMSSubstBase() As Object
'     Set ICMSSubstBase = objUserControl.Controls("ICMSSubstBase")
'End Property
'
'Public Property Get ICMSSubstValor() As Object
'     Set ICMSSubstValor = objUserControl.Controls("ICMSSubstValor")
'End Property
'
'Public Property Get ICMSBase() As Object
'     Set ICMSBase = objUserControl.Controls("ICMSBase")
'End Property
'
'Public Property Get ICMSValor() As Object
'     Set ICMSValor = objUserControl.Controls("ICMSValor")
'End Property
'
'Public Property Get TipoTributacao() As Object
'     Set TipoTributacao = objUserControl.Controls("TipoTributacao")
'End Property
'
'Public Property Get DescTipoTrib() As Object
'     Set DescTipoTrib = objUserControl.Controls("DescTipoTrib")
'End Property
'
'Public Property Get LblTipoTrib() As Object
'     Set LblTipoTrib = objUserControl.Controls("LblTipoTrib")
'End Property
'
'Public Property Get NatOpEspelho() As Object
'     Set NatOpEspelho = objUserControl.Controls("NatOpEspelho")
'End Property
'
'Public Property Get DescNatOp() As Object
'     Set DescNatOp = objUserControl.Controls("DescNatOp")
'End Property
'
'Public Property Get LblNatOpEspelho() As Object
'     Set LblNatOpEspelho = objUserControl.Controls("LblNatOpEspelho")
'End Property
'
'Public Property Get FrameItensTrib() As Object
'     Set FrameItensTrib = objUserControl.Controls("FrameItensTrib")
'End Property
'
'Public Property Get ComboItensTrib() As Object
'     Set ComboItensTrib = objUserControl.Controls("ComboItensTrib")
'End Property
'
'Public Property Get LabelValorItem() As Object
'     Set LabelValorItem = objUserControl.Controls("LabelValorItem")
'End Property
'
'Public Property Get LabelQtdeItem() As Object
'     Set LabelQtdeItem = objUserControl.Controls("LabelQtdeItem")
'End Property
'
'Public Property Get LabelUMItem() As Object
'     Set LabelUMItem = objUserControl.Controls("LabelUMItem")
'End Property
'
'Public Property Get FrameOutrosTrib() As Object
'     Set FrameOutrosTrib = objUserControl.Controls("FrameOutrosTrib")
'End Property
'
'Public Property Get LabelValorFrete() As Object
'     Set LabelValorFrete = objUserControl.Controls("LabelValorFrete")
'End Property
'
'Public Property Get LabelValorDesconto() As Object
'     Set LabelValorDesconto = objUserControl.Controls("LabelValorDesconto")
'End Property
'
'Public Property Get LabelValorSeguro() As Object
'     Set LabelValorSeguro = objUserControl.Controls("LabelValorSeguro")
'End Property
'
'Public Property Get LabelValorOutrasDespesas() As Object
'     Set LabelValorOutrasDespesas = objUserControl.Controls("LabelValorOutrasDespesas")
'End Property
'
'Public Property Get TribSobreItem() As Object
'     Set TribSobreItem = objUserControl.Controls("TribSobreItem")
'End Property
'
'Public Property Get TribSobreFrete() As Object
'     Set TribSobreFrete = objUserControl.Controls("TribSobreFrete")
'End Property
'
'Public Property Get TribSobreDesconto() As Object
'     Set TribSobreDesconto = objUserControl.Controls("TribSobreDesconto")
'End Property
'
'Public Property Get TribSobreSeguro() As Object
'     Set TribSobreSeguro = objUserControl.Controls("TribSobreSeguro")
'End Property
'
'Public Property Get TribSobreOutrasDesp() As Object
'     Set TribSobreOutrasDesp = objUserControl.Controls("TribSobreOutrasDesp")
'End Property
'
'Public Property Get Frame12() As Object
'     Set Frame12 = objUserControl.Controls("Frame12")
'End Property
'
'Public Property Get ICMSSubstBaseItem() As Object
'     Set ICMSSubstBaseItem = objUserControl.Controls("ICMSSubstBaseItem")
'End Property
'
'Public Property Get ICMSSubstAliquotaItem() As Object
'     Set ICMSSubstAliquotaItem = objUserControl.Controls("ICMSSubstAliquotaItem")
'End Property
'
'Public Property Get ICMSSubstValorItem() As Object
'     Set ICMSSubstValorItem = objUserControl.Controls("ICMSSubstValorItem")
'End Property
'
'Public Property Get ICMSValorItem() As Object
'     Set ICMSValorItem = objUserControl.Controls("ICMSValorItem")
'End Property
'
'Public Property Get ICMSAliquotaItem() As Object
'     Set ICMSAliquotaItem = objUserControl.Controls("ICMSAliquotaItem")
'End Property
'
'Public Property Get ICMSPercRedBaseItem() As Object
'     Set ICMSPercRedBaseItem = objUserControl.Controls("ICMSPercRedBaseItem")
'End Property
'
'Public Property Get ICMSBaseItem() As Object
'     Set ICMSBaseItem = objUserControl.Controls("ICMSBaseItem")
'End Property
'
'Public Property Get ComboICMSTipo() As Object
'     Set ComboICMSTipo = objUserControl.Controls("ComboICMSTipo")
'End Property
'
'Public Property Get IPIItemFrame() As Object
'     Set IPIItemFrame = objUserControl.Controls("IPIItemFrame")
'End Property
'
'Public Property Get IPIValorItem() As Object
'     Set IPIValorItem = objUserControl.Controls("IPIValorItem")
'End Property
'
'Public Property Get IPIAliquotaItem() As Object
'     Set IPIAliquotaItem = objUserControl.Controls("IPIAliquotaItem")
'End Property
'
'Public Property Get IPIPercRedBaseItem() As Object
'     Set IPIPercRedBaseItem = objUserControl.Controls("IPIPercRedBaseItem")
'End Property
'
'Public Property Get IPIBaseItem() As Object
'     Set IPIBaseItem = objUserControl.Controls("IPIBaseItem")
'End Property
'
'Public Property Get ComboIPITipo() As Object
'     Set ComboIPITipo = objUserControl.Controls("ComboIPITipo")
'End Property
'
'Public Property Get TipoTributacaoItem() As Object
'     Set TipoTributacaoItem = objUserControl.Controls("TipoTributacaoItem")
'End Property
'
'Public Property Get NaturezaOpItem() As Object
'     Set NaturezaOpItem = objUserControl.Controls("NaturezaOpItem")
'End Property
'
'Public Property Get LblTipoTribItem() As Object
'     Set LblTipoTribItem = objUserControl.Controls("LblTipoTribItem")
'End Property
'
'Public Property Get LabelDescrNatOpItem() As Object
'     Set LabelDescrNatOpItem = objUserControl.Controls("LabelDescrNatOpItem")
'End Property
'
'Public Property Get DescTipoTribItem() As Object
'     Set DescTipoTribItem = objUserControl.Controls("DescTipoTribItem")
'End Property
'
'Public Property Get NaturezaItemLabel() As Object
'     Set NaturezaItemLabel = objUserControl.Controls("NaturezaItemLabel")
'End Property
'
'Public Property Get OpcaoTributacao() As Object
'     Set OpcaoTributacao = objUserControl.Controls("OpcaoTributacao")
'End Property
'
'Public Property Get Frame8() As Object
'     Set Frame8 = objUserControl.Controls("Frame8")
'End Property

Public Property Get VolumeMarca() As Object
     Set VolumeMarca = objUserControl.Controls("VolumeMarca")
End Property

Public Property Get VolumeNumero() As Object
     Set VolumeNumero = objUserControl.Controls("VolumeNumero")
End Property

Public Property Get VolumeEspecie() As Object
     Set VolumeEspecie = objUserControl.Controls("VolumeEspecie")
End Property

Public Property Get VolumeQuant() As Object
     Set VolumeQuant = objUserControl.Controls("VolumeQuant")
End Property

Public Property Get PedidoCliente() As Object
     Set PedidoCliente = objUserControl.Controls("PedidoCliente")
End Property

Public Property Get PesoLiquido() As Object
     Set PesoLiquido = objUserControl.Controls("PesoLiquido")
End Property

Public Property Get CanalVenda() As Object
     Set CanalVenda = objUserControl.Controls("CanalVenda")
End Property

Public Property Get Mensagem() As Object
     Set Mensagem = objUserControl.Controls("Mensagem")
End Property

Public Property Get PesoBruto() As Object
     Set PesoBruto = objUserControl.Controls("PesoBruto")
End Property

Public Property Get CanalVendaLabel() As Object
     Set CanalVendaLabel = objUserControl.Controls("CanalVendaLabel")
End Property

Public Property Get MensagemLabel() As Object
     Set MensagemLabel = objUserControl.Controls("MensagemLabel")
End Property

'Public Property Get Frame16() As Object
'     Set Frame16 = objUserControl.Controls("Frame16")
'End Property

Public Property Get RedespachoCli() As Object
     Set RedespachoCli = objUserControl.Controls("RedespachoCli")
End Property

Public Property Get TranspRedespacho() As Object
     Set TranspRedespacho = objUserControl.Controls("TranspRedespacho")
End Property

Public Property Get TranspRedLabel() As Object
     Set TranspRedLabel = objUserControl.Controls("TranspRedLabel")
End Property

'Public Property Get Frame6() As Object
'     Set Frame6 = objUserControl.Controls("Frame6")
'End Property

Public Property Get TipoFrete() As Object
     Set TipoFrete = objUserControl.Controls("TipoFrete")
End Property

Public Property Get Destinatario() As Object
     Set Destinatario = objUserControl.Controls("Destinatario")
End Property

Public Property Get LabelVLight() As Object
     Set LabelVLight = objUserControl.Controls("LabelVLight")
End Property

Public Property Get DataEntregaPV() As Object
     Set DataEntregaPV = objUserControl.Controls("DataEntregaPV")
End Property

Public Property Get Placa() As Object
     Set Placa = objUserControl.Controls("Placa")
End Property

Public Property Get PlacaUF() As Object
     Set PlacaUF = objUserControl.Controls("PlacaUF")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
End Property

Public Property Get FilialEntrega() As Object
     Set FilialEntrega = objUserControl.Controls("FilialEntrega")
End Property

Public Property Get TransportadoraLabel() As Object
     Set TransportadoraLabel = objUserControl.Controls("TransportadoraLabel")
End Property

Public Property Get SSFrame1() As Object
     Set SSFrame1 = objUserControl.Controls("SSFrame1")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get TipoBloqueio() As Object
     Set TipoBloqueio = objUserControl.Controls("TipoBloqueio")
End Property

Public Property Get ResponsavelLib() As Object
     Set ResponsavelLib = objUserControl.Controls("ResponsavelLib")
End Property

Public Property Get DataLiberacao() As Object
     Set DataLiberacao = objUserControl.Controls("DataLiberacao")
End Property

Public Property Get CodUsuario() As Object
     Set CodUsuario = objUserControl.Controls("CodUsuario")
End Property

Public Property Get ResponsavelBL() As Object
     Set ResponsavelBL = objUserControl.Controls("ResponsavelBL")
End Property

Public Property Get DataBloqueio() As Object
     Set DataBloqueio = objUserControl.Controls("DataBloqueio")
End Property

Public Property Get GridBloqueio() As Object
     Set GridBloqueio = objUserControl.Controls("GridBloqueio")
End Property

Public Property Get Command1() As Object
     Set Command1 = objUserControl.Controls("Command1")
End Property

Public Property Get ValorDesconto() As Object
     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
End Property

Public Property Get ValorDespesas() As Object
     Set ValorDespesas = objUserControl.Controls("ValorDespesas")
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
End Property

Public Property Get ValorFrete() As Object
     Set ValorFrete = objUserControl.Controls("ValorFrete")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get ValorProdutos() As Object
     Set ValorProdutos = objUserControl.Controls("ValorProdutos")
End Property

Public Property Get ICMSBase1() As Object
     Set ICMSBase1 = objUserControl.Controls("ICMSBase1")
End Property

Public Property Get ICMSValor1() As Object
     Set ICMSValor1 = objUserControl.Controls("ICMSValor1")
End Property

Public Property Get ICMSSubstBase1() As Object
     Set ICMSSubstBase1 = objUserControl.Controls("ICMSSubstBase1")
End Property

Public Property Get ICMSSubstValor1() As Object
     Set ICMSSubstValor1 = objUserControl.Controls("ICMSSubstValor1")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get DescricaoProduto() As Object
     Set DescricaoProduto = objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get QuantCancelada() As Object
     Set QuantCancelada = objUserControl.Controls("QuantCancelada")
End Property

Public Property Get QuantFaturada() As Object
     Set QuantFaturada = objUserControl.Controls("QuantFaturada")
End Property

Public Property Get QuantReservadaPV() As Object
     Set QuantReservadaPV = objUserControl.Controls("QuantReservadaPV")
End Property

Public Property Get DataEntrega() As Object
     Set DataEntrega = objUserControl.Controls("DataEntrega")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get PercentDesc() As Object
     Set PercentDesc = objUserControl.Controls("PercentDesc")
End Property

Public Property Get PrecoUnitario() As Object
     Set PrecoUnitario = objUserControl.Controls("PrecoUnitario")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get PrecoTotal() As Object
     Set PrecoTotal = objUserControl.Controls("PrecoTotal")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get BotaoVendedores() As Object
     Set BotaoVendedores = objUserControl.Controls("BotaoVendedores")
End Property

Public Property Get ComissaoAutomatica() As Object
     Set ComissaoAutomatica = objUserControl.Controls("ComissaoAutomatica")
End Property

Public Property Get SSFrame4() As Object
     Set SSFrame4 = objUserControl.Controls("SSFrame4")
End Property

Public Property Get TotalPercentualComissao() As Object
     Set TotalPercentualComissao = objUserControl.Controls("TotalPercentualComissao")
End Property

Public Property Get TotalValorComissao() As Object
     Set TotalValorComissao = objUserControl.Controls("TotalValorComissao")
End Property

Public Property Get TotalValorBase() As Object
     Set TotalValorBase = objUserControl.Controls("TotalValorBase")
End Property

Public Property Get DiretoIndireto() As Object
     Set DiretoIndireto = objUserControl.Controls("DiretoIndireto")
End Property

Public Property Get ValorComissao() As Object
     Set ValorComissao = objUserControl.Controls("ValorComissao")
End Property

Public Property Get ValorBase() As Object
     Set ValorBase = objUserControl.Controls("ValorBase")
End Property

Public Property Get PercentualComissao() As Object
     Set PercentualComissao = objUserControl.Controls("PercentualComissao")
End Property

Public Property Get Vendedor() As Object
     Set Vendedor = objUserControl.Controls("Vendedor")
End Property

Public Property Get ValorEmissao() As Object
     Set ValorEmissao = objUserControl.Controls("ValorEmissao")
End Property

Public Property Get PercentualEmissao() As Object
     Set PercentualEmissao = objUserControl.Controls("PercentualEmissao")
End Property

Public Property Get PercentualBaixa() As Object
     Set PercentualBaixa = objUserControl.Controls("PercentualBaixa")
End Property

Public Property Get ValorBaixa() As Object
     Set ValorBaixa = objUserControl.Controls("ValorBaixa")
End Property

Public Property Get GridComissoes() As Object
     Set GridComissoes = objUserControl.Controls("GridComissoes")
End Property

Public Property Get SSFrame3() As Object
     Set SSFrame3 = objUserControl.Controls("SSFrame3")
End Property

Public Property Get TipoDesconto3() As Object
     Set TipoDesconto3 = objUserControl.Controls("TipoDesconto3")
End Property

Public Property Get TipoDesconto2() As Object
     Set TipoDesconto2 = objUserControl.Controls("TipoDesconto2")
End Property

Public Property Get TipoDesconto1() As Object
     Set TipoDesconto1 = objUserControl.Controls("TipoDesconto1")
End Property

Public Property Get Desconto1Percentual() As Object
     Set Desconto1Percentual = objUserControl.Controls("Desconto1Percentual")
End Property

Public Property Get Desconto3Valor() As Object
     Set Desconto3Valor = objUserControl.Controls("Desconto3Valor")
End Property

Public Property Get Desconto3Ate() As Object
     Set Desconto3Ate = objUserControl.Controls("Desconto3Ate")
End Property

Public Property Get Desconto2Valor() As Object
     Set Desconto2Valor = objUserControl.Controls("Desconto2Valor")
End Property

Public Property Get Desconto2Ate() As Object
     Set Desconto2Ate = objUserControl.Controls("Desconto2Ate")
End Property

Public Property Get Desconto1Valor() As Object
     Set Desconto1Valor = objUserControl.Controls("Desconto1Valor")
End Property

Public Property Get Desconto1Ate() As Object
     Set Desconto1Ate = objUserControl.Controls("Desconto1Ate")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get Desconto2Percentual() As Object
     Set Desconto2Percentual = objUserControl.Controls("Desconto2Percentual")
End Property

Public Property Get Desconto3Percentual() As Object
     Set Desconto3Percentual = objUserControl.Controls("Desconto3Percentual")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get FilialPedido() As Object
     Set FilialPedido = objUserControl.Controls("FilialPedido")
End Property

Public Property Get FilialFaturamento() As Object
     Set FilialFaturamento = objUserControl.Controls("FilialFaturamento")
End Property

Public Property Get NaturezaOp() As Object
     Set NaturezaOp = objUserControl.Controls("NaturezaOp")
End Property

Public Property Get NaturezaLabel() As Object
     Set NaturezaLabel = objUserControl.Controls("NaturezaLabel")
End Property

Public Property Get PercAcrescFin() As Object
     Set PercAcrescFin = objUserControl.Controls("PercAcrescFin")
End Property

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get TabelaPreco() As Object
     Set TabelaPreco = objUserControl.Controls("TabelaPreco")
End Property

Public Property Get CondPagtoLabel() As Object
     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
End Property

Public Property Get FaturaIntegral() As Object
     Set FaturaIntegral = objUserControl.Controls("FaturaIntegral")
End Property

Public Property Get FrameNFiscal() As Object
     Set FrameNFiscal = objUserControl.Controls("FrameNFiscal")
End Property

Public Property Get SerieNF() As Object
     Set SerieNF = objUserControl.Controls("SerieNF")
End Property

Public Property Get ItemNF() As Object
     Set ItemNF = objUserControl.Controls("ItemNF")
End Property

Public Property Get NFiscalNF() As Object
     Set NFiscalNF = objUserControl.Controls("NFiscalNF")
End Property

Public Property Get DescricaoProdutoNF() As Object
     Set DescricaoProdutoNF = objUserControl.Controls("DescricaoProdutoNF")
End Property

Public Property Get UMNF() As Object
     Set UMNF = objUserControl.Controls("UMNF")
End Property

Public Property Get ItemPVNF() As Object
     Set ItemPVNF = objUserControl.Controls("ItemPVNF")
End Property

Public Property Get PrecoTotalNF() As Object
     Set PrecoTotalNF = objUserControl.Controls("PrecoTotalNF")
End Property

Public Property Get DescontoNF() As Object
     Set DescontoNF = objUserControl.Controls("DescontoNF")
End Property

Public Property Get PercDescNF() As Object
     Set PercDescNF = objUserControl.Controls("PercDescNF")
End Property

Public Property Get PrecoUnitarioNF() As Object
     Set PrecoUnitarioNF = objUserControl.Controls("PrecoUnitarioNF")
End Property

Public Property Get QuantidadeNF() As Object
     Set QuantidadeNF = objUserControl.Controls("QuantidadeNF")
End Property

Public Property Get ProdutoNF() As Object
     Set ProdutoNF = objUserControl.Controls("ProdutoNF")
End Property

Public Property Get GridNFiscal() As Object
     Set GridNFiscal = objUserControl.Controls("GridNFiscal")
End Property

Public Property Get SeriePV() As Object
     Set SeriePV = objUserControl.Controls("SeriePV")
End Property

Public Property Get PrecoTotalPV() As Object
     Set PrecoTotalPV = objUserControl.Controls("PrecoTotalPV")
End Property

Public Property Get DescontoPV() As Object
     Set DescontoPV = objUserControl.Controls("DescontoPV")
End Property

Public Property Get PercDescPV() As Object
     Set PercDescPV = objUserControl.Controls("PercDescPV")
End Property

Public Property Get PrecoUnitarioPV() As Object
     Set PrecoUnitarioPV = objUserControl.Controls("PrecoUnitarioPV")
End Property

Public Property Get QuantidadePV() As Object
     Set QuantidadePV = objUserControl.Controls("QuantidadePV")
End Property

Public Property Get ItemNFPV() As Object
     Set ItemNFPV = objUserControl.Controls("ItemNFPV")
End Property

Public Property Get NFiscalPV() As Object
     Set NFiscalPV = objUserControl.Controls("NFiscalPV")
End Property

Public Property Get QuantFaturadaPV() As Object
     Set QuantFaturadaPV = objUserControl.Controls("QuantFaturadaPV")
End Property

Public Property Get DescricaoProdutoPV() As Object
     Set DescricaoProdutoPV = objUserControl.Controls("DescricaoProdutoPV")
End Property

Public Property Get UMPV() As Object
     Set UMPV = objUserControl.Controls("UMPV")
End Property

Public Property Get ItemPV() As Object
     Set ItemPV = objUserControl.Controls("ItemPV")
End Property

Public Property Get ProdutoPV() As Object
     Set ProdutoPV = objUserControl.Controls("ProdutoPV")
End Property

Public Property Get GridNFItemPV() As Object
     Set GridNFItemPV = objUserControl.Controls("GridNFItemPV")
End Property

Public Property Get BotaoNFiscal() As Object
     Set BotaoNFiscal = objUserControl.Controls("BotaoNFiscal")
End Property

Public Property Get ComboOrdenacao() As Object
     Set ComboOrdenacao = objUserControl.Controls("ComboOrdenacao")
End Property

Public Property Get ProdutoDescricao() As Object
     Set ProdutoDescricao = objUserControl.Controls("ProdutoDescricao")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get Responsavel() As Object
     Set Responsavel = objUserControl.Controls("Responsavel")
End Property

Public Property Get UnidadeMedEst() As Object
     Set UnidadeMedEst = objUserControl.Controls("UnidadeMedEst")
End Property

Public Property Get DataValidade() As Object
     Set DataValidade = objUserControl.Controls("DataValidade")
End Property

Public Property Get ProdutoAlmox() As Object
     Set ProdutoAlmox = objUserControl.Controls("ProdutoAlmox")
End Property

Public Property Get Almox() As Object
     Set Almox = objUserControl.Controls("Almox")
End Property

Public Property Get QuantReservar() As Object
     Set QuantReservar = objUserControl.Controls("QuantReservar")
End Property

Public Property Get QuantReservada() As Object
     Set QuantReservada = objUserControl.Controls("QuantReservada")
End Property

Public Property Get ItemPedido() As Object
     Set ItemPedido = objUserControl.Controls("ItemPedido")
End Property

Public Property Get GridReserva() As Object
     Set GridReserva = objUserControl.Controls("GridReserva")
End Property

Public Property Get BotaoConsulta() As Object
     Set BotaoConsulta = objUserControl.Controls("BotaoConsulta")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoEditar() As Object
     Set BotaoEditar = objUserControl.Controls("BotaoEditar")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get StatusPedido() As Object
     Set StatusPedido = objUserControl.Controls("StatusPedido")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get LabelCliente() As Object
     Set LabelCliente = objUserControl.Controls("LabelCliente")
End Property

Public Property Get Line1() As Object
     Set Line1 = objUserControl.Controls("Line1")
End Property

Public Property Get BotaoDesfazer() As Object
     Set BotaoDesfazer = objUserControl.Controls("BotaoDesfazer")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property
'--- fim dos properties get dos controles da tela

'@@@@@ para ser customizada
Public Function PVC_Inicializa_Grid_Itens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    
Dim iColunasEscondidas As Integer
 
 'Inicializa o Grid de Itens

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Pedida")
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Quant. Cancelada")
    End If
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Tab. Preço")
    End If
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")
    objGridInt.colColuna.Add ("Data Entrega")
    objGridInt.colColuna.Add ("Quant Reservada")
    objGridInt.colColuna.Add ("Quant Faturada")
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colColuna.Add ("Comissão")
    End If

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        objGridInt.colCampo.Add (QuantCancelada.Name)
    Else
        QuantCancelada.left = POSICAO_FORA_TELA
    End If
    
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        objGridInt.colCampo.Add (TabPrecoItemPV.Name)
    Else
        TabPrecoItemPV.left = POSICAO_FORA_TELA
    End If
    objGridInt.colCampo.Add (PrecoUnitario.Name)
    objGridInt.colCampo.Add (PercentDesc.Name)
    objGridInt.colCampo.Add (Desconto.Name)
    objGridInt.colCampo.Add (PrecoTotal.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (QuantReservadaPV.Name)
    objGridInt.colCampo.Add (QuantFaturada.Name)
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        objGridInt.colCampo.Add (ComissaoItemPV.Name)
    Else
        ComissaoItemPV.left = POSICAO_FORA_TELA
    End If

    'Colunas do Grid
    iGrid_ItemProduto_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescProduto_Col = 2
    iGrid_UnidadeMed_Col = 3
    iGrid_QuantPedida_Col = 4
    If gobjFAT.iExibeQtdeCanceladaItensPV = MARCADO Then
        iGrid_QuantCancel_Col = 5
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If
    
    If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
        iGrid_TabPrecoItem_Col = 6 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If
    
    iGrid_PrecoUnitario_Col = 7 - iColunasEscondidas
    iGrid_PercDesc_Col = 8 - iColunasEscondidas
    iGrid_Desconto_Col = 9 - iColunasEscondidas
    iGrid_PrecoTotal_Col = 10 - iColunasEscondidas
    iGrid_DataEntrega_Col = 11 - iColunasEscondidas
    iGrid_QuantReservada_Col = 12 - iColunasEscondidas
    iGrid_QuantFaturada_Col = 13 - iColunasEscondidas
    If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        iGrid_ComissaoItem_Col = 14 - iColunasEscondidas
    Else
        iColunasEscondidas = iColunasEscondidas + 1
    End If

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    PVC_Inicializa_Grid_Itens = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    TAB_Tributacao = 9
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        LabelVLight(0).Visible = False
'        LabelVLight(1).Visible = False
'        LabelVLight(2).Visible = False
'        LabelVLight(3).Visible = False
'        LabelVLight(4).Visible = False
'        CanalVendaLabel.Visible = False
'        FaturaIntegral.Visible = False
'        Filial.Visible = False
'        FilialFaturamento.Visible = False
'        TabelaPreco.Visible = False
'        FilialEntrega.Visible = False
'        CanalVenda.Visible = False
'        FilialPedido.Visible = False
'
'        'Inicializa FrameL para versão LIGHT
'        Set FrameL(1) = Frame1(1)
'        Set FrameL(2) = Frame1(2)
'        Set FrameL(3) = Frame1(3)
'        Set FrameL(4) = Frame1(4)
'        Set FrameL(5) = Frame1(5)
'        Set FrameL(6) = Frame1(7)
'        Set FrameL(7) = Frame1(8)
'        Set FrameL(8) = Frame1(9)
'        Set FrameL(9) = Frame1(10)
'
'        Opcao.Tabs.Remove TAB_Bloqueio
'        Frame1(TAB_Bloqueio).left = -20000
'
'        TAB_Tributacao = 8
'
'    End If
    
    Set objTabComissoes.objTela = Me
    
    giFrameAtual = 1
    'giFrameAtualTributacao = 1
        
    giGridNotaAtual = 1
    ComboOrdenacao.ListIndex = 0
    
'    If giTipoVersao = VERSAO_FULL Then
        'Carrega a combo de Tipos de Bloqueio
        lErro = Carrega_TipoBloqueio()
        If lErro <> SUCESSO Then Error 60200
'    End If
    
    'Carrega a combo de Tipo de desconto
    lErro = Carrega_TipoDesconto()
    If lErro <> SUCESSO Then Error 60201

    Quantidade.Format = FORMATO_ESTOQUE
    QuantCancelada.Format = FORMATO_ESTOQUE

    'Iniciliza as variáveis de browse
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento

    Set objGridItens = New AdmGrid
    Set objGridComissoes = New AdmGrid
    Set objGridParcelas = New AdmGrid
'    If giTipoVersao = VERSAO_FULL Then
        Set objGridBloqueio = New AdmGrid
'    endif
    Set objGridReserva = New AdmGrid
    Set objGridNFItemPV = New AdmGrid
    Set objGridNFiscal = New AdmGrid
    
    
    'Inserido por Wagner
    '#####################################
    'Faz as Inicializações dos Grids
     lErro = CF2(Me, "PVC_Inicializa_Grid_Itens", objGridItens)
     If lErro <> SUCESSO Then gError 129160
     '####################################

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 60203
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoNF)
    If lErro <> SUCESSO Then Error 60204
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoPV)
    If lErro <> SUCESSO Then Error 60205
    
    lErro = objTabComissoes.Inicializa_Grid_Comissoes(objGridComissoes)
    If lErro <> SUCESSO Then Error 60206

    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 60207

'    If giTipoVersao = VERSAO_FULL Then
        lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
        If lErro <> SUCESSO Then Error 60208
'    End If
    
    lErro = Inicializa_Grid_Reserva(objGridReserva)
    If lErro <> SUCESSO Then Error 60209
   
    lErro = Inicializa_Grid_NFItemPV(objGridNFItemPV)
    If lErro <> SUCESSO Then Error 60210
   
    lErro = Inicializa_Grid_NFiscal(objGridNFiscal)
    If lErro <> SUCESSO Then Error 60211
'
'    lErro = CarregaTiposTrib_Global()
'    If lErro <> SUCESSO Then Error 60212
'
'    lErro = TributacaoPV_Reset()
'    If lErro <> SUCESSO Then Error 60213

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjTribTab = New ClassTribTab
    lErro = gobjTribTab.Ativar(Me, , , gobjTribTab.TIPOTELA_PV, gobjTribTab.SUBTIPOTELA_CONSULTA)
    If lErro <> SUCESSO Then Error 60211
    
    Set gobjPedidoVenda = New ClassPedidoDeVenda
    lErro = gobjTribTab.TributacaoNF_Reset(gobjPedidoVenda)
    If lErro <> SUCESSO Then Error 60211
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 60201, 60202, 60203, 60204, 60205, 60206, 60207, 60208, 60209, _
            60210, 60211, 60212, 60213, 60214, 129160
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158199)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoConsulta_Click

    If Len(Trim(Codigo.Text)) = 0 Then Error 60215
    
    objPedidoDeVenda.lCodigo = CLng(Codigo.Text)
    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa
    
    lErro = Traz_Pedido_Tela(objPedidoDeVenda)
    If lErro <> SUCESSO Then Error 60216
    
    Exit Sub
    
Erro_BotaoConsulta_Click:

    Select Case Err
        
        Case 60215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", Err)
        
        Case 60216
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158200)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoEditar_Click()

Dim lErro As Long
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim iStatus As Integer

On Error GoTo Erro_BotaoConsulta_Click

    If Len(Trim(Codigo.Text)) = 0 Then Error 60217
    
    objPedidoDeVenda.lCodigo = CLng(Codigo.Text)
    objPedidoDeVenda.iFilialEmpresa = Codigo_Extrai(FilialPedido.Caption)
    
    If objPedidoDeVenda.iFilialEmpresa <> giFilialEmpresa Then Error 64468
            
    lErro = CF("PedidoDeVenda_Le_Todos_Completo", objPedidoDeVenda, iStatus)
    If lErro <> SUCESSO Then Error 60218
    
    If iStatus = STATUS_ABERTO Then
        Call Chama_Tela("PedidoVenda", objPedidoDeVenda)
    ElseIf iStatus = STATUS_BAIXADO Then
        Call Rotina_Aviso(vbOKOnly, "AVISO_PEDIDO_BAIXADO", objPedidoDeVenda.lCodigo)
    End If
        
    Exit Sub
    
Erro_BotaoConsulta_Click:

    Select Case Err
    
        Case 60217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_PEDIDO_NAO_PREENCHIDO", Err)
        
        Case 60218
    
        Case 64468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PV_NAOEDITAVEL_FILIALEMPRESA_DIFERENTE", Err, objPedidoDeVenda.lCodigo, objPedidoDeVenda.iFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158201)

    End Select

    Exit Sub

End Sub

Public Sub BotaoNFiscal_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim sTela As String

On Error GoTo Erro_BotaoNFiscal_Click

    If ComboOrdenacao.ItemData(ComboOrdenacao.ListIndex) = 1 Then
        
        If Len(Trim(GridNFItemPV.TextMatrix(GridNFItemPV.Row, iGrid_ProdutoPV_Col))) > 0 Then
        
        If gcolItemPedidoNF_ItemPV.Count = 0 Or GridNFItemPV.Row = 0 Then Exit Sub
            
            objNFiscal.sSerie = gcolItemPedidoNF_ItemPV.Item(GridNFItemPV.Row).sSerie
            objNFiscal.lNumNotaFiscal = gcolItemPedidoNF_ItemPV.Item(GridNFItemPV.Row).lNumNF
            objNFiscal.dtDataEmissao = gcolItemPedidoNF_ItemPV.Item(GridNFItemPV.Row).dtDataEmissao
            objNFiscal.iFilialEmpresa = giFilialEmpresa
            objNFiscal.lNumPedidoVenda = gcolItemPedidoNF_ItemPV.Item(GridNFItemPV.Row).lCodPedido
            objNFiscal.iFilialPedido = gcolItemPedidoNF_ItemPV.Item(GridNFItemPV.Row).iFilialEmpresa
                    
        End If
    
    ElseIf ComboOrdenacao.ItemData(ComboOrdenacao.ListIndex) = 2 Then
        
        If gcolItemPedidoNF_Nota.Count = 0 Or GridNFiscal.Row = 0 Then Exit Sub
        
        If Len(Trim(GridNFiscal.TextMatrix(GridNFiscal.Row, iGrid_ProdutoNF_Col))) > 0 Then
            
            objNFiscal.sSerie = gcolItemPedidoNF_Nota.Item(GridNFiscal.Row).sSerie
            objNFiscal.lNumNotaFiscal = gcolItemPedidoNF_Nota.Item(GridNFiscal.Row).lNumNF
            objNFiscal.dtDataEmissao = gcolItemPedidoNF_Nota.Item(GridNFiscal.Row).dtDataEmissao
            objNFiscal.iFilialEmpresa = giFilialEmpresa
            objNFiscal.lNumPedidoVenda = gcolItemPedidoNF_Nota.Item(GridNFiscal.Row).lCodPedido
            objNFiscal.iFilialPedido = gcolItemPedidoNF_Nota.Item(GridNFiscal.Row).iFilialEmpresa
        
        End If
    End If
                    
    lErro = CF("NFiscal_Le_Pedido_TipoDocInfo", objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 60222 Then Error 60223
    
    If lErro = 60222 Then Error 60224
            
'    If giTipoVersao = VERSAO_FULL Then
        Call Chama_Tela(sTela, objNFiscal)
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'        If objNFiscal.iTipoNFiscal = DOCINFO_NFISVPV Or objNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV Then
'            Call Chama_Tela("NFiscalFatura", objNFiscal)
'        Else
'            Call Chama_Tela(sTela, objNFiscal)
'        End If
'    End If
    
    Exit Sub
    
Erro_BotaoNFiscal_Click:

    Select Case Err
        
        Case 60223
        
        Case 60224
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA1", Err, objNFiscal.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158202)

    End Select

    Exit Sub
            
End Sub

Public Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo)

End Sub

'Public Sub ComboItensTrib_Click()
'
'Dim iIndice As Integer, objItemPedido As ClassItemPedido
'
'    iIndice = ComboItensTrib.ListIndex
'
'    If iIndice <> -1 Then
'
'        'preenche os campos da tela em funcao do item selecionado
'
'        Set objItemPedido = gobjPedidoVenda.colItensPedido.Item(iIndice + 1)
'
'        LabelValorItem.Caption = Format(objItemPedido.dPrecoTotal, "Standard")
'        LabelQtdeItem.Caption = CStr(objItemPedido.dQuantidade - objItemPedido.dQuantCancelada)
'        LabelUMItem.Caption = objItemPedido.sUnidadeMed
'
'        Call TributacaoItem_TrazerTela(objItemPedido.objTributacaoItemPV)
'
'    End If
'
'End Sub

Private Function Carrega_TipoBloqueio() As Long

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TipoBloqueio

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("Cod_Nomes_Le", "TiposDeBloqueio", "Codigo", "NomeReduzido", STRING_TIPO_BLOQUEIO_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 60225

    For Each objCodDescricao In colCodigoDescricao

        If objCodDescricao.iCodigo <> BLOQUEIO_PARCIAL And objCodDescricao.iCodigo <> BLOQUEIO_NAO_RESERVA And objCodDescricao.iCodigo <> BLOQUEIO_CREDITO Then
            'Adiciona o item na Lista de Tabela de Preços
            TipoBloqueio.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
            TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objCodDescricao.iCodigo
        End If
    Next

    Carrega_TipoBloqueio = SUCESSO

    Exit Function

Erro_Carrega_TipoBloqueio:

    Carrega_TipoBloqueio = Err

    Select Case Err

        Case 60225

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158203)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoDesconto() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    'Lê o código e a descrição de todos os Tipos de Desconto
    lErro = CF("Cod_Nomes_Le", "TiposDeDesconto", "Codigo", "DescReduzida", STRING_TIPOSDEDESCONTO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 60226

    For Each objCodDescricao In colCodigoDescricao
       
        'Adiciona o item nas List's das Combos de Tipos Desconto
        TipoDesconto1.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto1.ItemData(TipoDesconto1.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto2.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto2.ItemData(TipoDesconto2.NewIndex) = objCodDescricao.iCodigo
        TipoDesconto3.AddItem objCodDescricao.iCodigo & SEPARADOR & objCodDescricao.sNome
        TipoDesconto3.ItemData(TipoDesconto3.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = Err

    Select Case Err

        Case 60226

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158204)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_NFItemPV(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Notas Fiscais com Item PV

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Item P.V.")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Faturada")
    objGridInt.colColuna.Add ("Serie N.F.")
    objGridInt.colColuna.Add ("N. Fiscal")
    objGridInt.colColuna.Add ("Item N.F.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemPV.Name)
    objGridInt.colCampo.Add (ProdutoPV.Name)
    objGridInt.colCampo.Add (DescricaoProdutoPV.Name)
    objGridInt.colCampo.Add (UMPV.Name)
    objGridInt.colCampo.Add (QuantFaturadaPV.Name)
    objGridInt.colCampo.Add (SeriePV.Name)
    objGridInt.colCampo.Add (NFiscalPV.Name)
    objGridInt.colCampo.Add (ItemNFPV.Name)
    objGridInt.colCampo.Add (QuantidadePV.Name)
    objGridInt.colCampo.Add (PrecoUnitarioPV.Name)
    objGridInt.colCampo.Add (PercDescPV.Name)
    objGridInt.colCampo.Add (DescontoPV.Name)
    objGridInt.colCampo.Add (PrecoTotalPV.Name)

    'Colunas do Grid
    iGrid_ItemPV_Col = 1
    iGrid_ProdutoPV_Col = 2
    iGrid_DescProdutoPV_Col = 3
    iGrid_UnidadeMedPV_Col = 4
    iGrid_QuantFaturadaPV_Col = 5
    iGrid_SeriePV_Col = 6
    iGrid_NFiscalPV_Col = 7
    iGrid_ItemNFPV_Col = 8
    iGrid_QuantidadePV_Col = 9
    iGrid_PrecoUnitarioPV_Col = 10
    iGrid_PercDescPV_Col = 11
    iGrid_DescontoPV_Col = 12
    iGrid_PrecoTotalPV_Col = 13

    'Grid do GridInterno
    objGridInt.objGrid = GridNFItemPV

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridNFItemPV.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_NFItemPV = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_NFiscal(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Notas Fiscais com Item PV

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Serie N.F.")
    objGridInt.colColuna.Add ("N. Fiscal")
    objGridInt.colColuna.Add ("Item N.F.")
    objGridInt.colColuna.Add ("Item P.V.")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Preço Unitário")
    objGridInt.colColuna.Add ("% Desconto")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Preço Total")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (SerieNF.Name)
    objGridInt.colCampo.Add (NFiscalNF.Name)
    objGridInt.colCampo.Add (ItemNF.Name)
    objGridInt.colCampo.Add (ItemPVNF.Name)
    objGridInt.colCampo.Add (ProdutoNF.Name)
    objGridInt.colCampo.Add (DescricaoProdutoNF.Name)
    objGridInt.colCampo.Add (UMNF.Name)
    objGridInt.colCampo.Add (QuantidadeNF.Name)
    objGridInt.colCampo.Add (PrecoUnitarioNF.Name)
    objGridInt.colCampo.Add (PercDescNF.Name)
    objGridInt.colCampo.Add (DescontoNF.Name)
    objGridInt.colCampo.Add (PrecoTotalNF.Name)

    'Colunas do Grid
    iGrid_SerieNF_Col = 1
    iGrid_NFiscalNF_Col = 2
    iGrid_ItemNF_Col = 3
    iGrid_ItemPVNF_Col = 4
    iGrid_ProdutoNF_Col = 5
    iGrid_DescProdutoNF_Col = 6
    iGrid_UnidadeMedNF_Col = 7
    iGrid_QuantidadeNF_Col = 8
    iGrid_PrecoUnitarioNF_Col = 9
    iGrid_PercDescNF_Col = 10
    iGrid_DescontoNF_Col = 11
    iGrid_PrecoTotalNF_Col = 11

    'Grid do GridInterno
    objGridInt.objGrid = GridNFiscal

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridNFiscal.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_NFiscal = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Desconto 1 Tipo")
    objGridInt.colColuna.Add ("Desc. 1 Data")
    objGridInt.colColuna.Add ("Desc. 1 Valor")
    objGridInt.colColuna.Add ("Desc. 1 %")
    objGridInt.colColuna.Add ("Desconto 2 Tipo")
    objGridInt.colColuna.Add ("Desc. 2 Data")
    objGridInt.colColuna.Add ("Desc. 2 Valor")
    objGridInt.colColuna.Add ("Desc. 2 %")
    objGridInt.colColuna.Add ("Desconto 3 Tipo")
    objGridInt.colColuna.Add ("Desc. 3 Data")
    objGridInt.colColuna.Add ("Desc. 3 Valor")
    objGridInt.colColuna.Add ("Desc. 3 %")

    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoDesconto1.Name)
    objGridInt.colCampo.Add (Desconto1Ate.Name)
    objGridInt.colCampo.Add (Desconto1Valor.Name)
    objGridInt.colCampo.Add (Desconto1Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto2.Name)
    objGridInt.colCampo.Add (Desconto2Ate.Name)
    objGridInt.colCampo.Add (Desconto2Valor.Name)
    objGridInt.colCampo.Add (Desconto2Percentual.Name)
    objGridInt.colCampo.Add (TipoDesconto3.Name)
    objGridInt.colCampo.Add (Desconto3Ate.Name)
    objGridInt.colCampo.Add (Desconto3Valor.Name)
    objGridInt.colCampo.Add (Desconto3Percentual.Name)


    'Controles que participam do Grid
    iGrid_Vencimento_Col = 1
    iGrid_ValorParcela_Col = 2
    iGrid_Desc1Codigo_Col = 3
    iGrid_Desc1Ate_Col = 4
    iGrid_Desc1Valor_Col = 5
    iGrid_Desc1Perc_Col = 6
    iGrid_Desc2Codigo_Col = 7
    iGrid_Desc2Ate_Col = 8
    iGrid_Desc2Valor_Col = 9
    iGrid_Desc2Perc_Col = 10
    iGrid_Desc3Codigo_Col = 11
    iGrid_Desc3Ate_Col = 12
    iGrid_Desc3Valor_Col = 13
    iGrid_Desc3Perc_Col = 14

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 700

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Usuário")
    objGridInt.colColuna.Add ("Responsável")
    objGridInt.colColuna.Add ("Data Liberação")
    objGridInt.colColuna.Add ("Resp. Liberação")
    objGridInt.colColuna.Add ("Observação")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (CodUsuario.Name)
    objGridInt.colCampo.Add (ResponsavelBL.Name)
    objGridInt.colCampo.Add (DataLiberacao.Name)
    objGridInt.colCampo.Add (ResponsavelLib.Name)
    objGridInt.colCampo.Add (Observacao.Name)

    iGrid_TipoBloqueio_Col = 1
    iGrid_DataBloqueio_Col = 2
    iGrid_CodUsuario_Col = 3
    iGrid_ResponsavelBL_Col = 4
    iGrid_DataLiberacao_Col = 5
    iGrid_ResponsavelLib_Col = 6
    iGrid_Observacao_Col = 7

    'Grid do GridInterno
    objGridInt.objGrid = GridBloqueio

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_BLOQUEIOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridBloqueio.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Bloqueio = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Reserva(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Reserva

Dim lAlmoxarifado As Long

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Item")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Quant.a Reservar")
    objGridInt.colColuna.Add ("Quant. Reservada")
    objGridInt.colColuna.Add ("Validade")
    objGridInt.colColuna.Add ("Responsável")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ItemPedido.Name)
    objGridInt.colCampo.Add (ProdutoAlmox.Name)
    objGridInt.colCampo.Add (UnidadeMedEst.Name)
    objGridInt.colCampo.Add (Almox.Name)
    objGridInt.colCampo.Add (QuantReservar.Name)
    objGridInt.colCampo.Add (QuantReservada.Name)
    objGridInt.colCampo.Add (DataValidade.Name)
    objGridInt.colCampo.Add (Responsavel.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridReserva

    iGrid_Item_Col = 1
    iGrid_ProdutoAlmox_Col = 2
    iGrid_UMEstoque_Col = 3
    iGrid_Almoxarifado_Col = 4
    iGrid_QuantReservar_Col = 5
    iGrid_QuantReserv_Col = 6
    iGrid_Validade_Col = 7
    iGrid_Responsavel_Col = 8

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_RESERVAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridReserva.ColWidth(0) = 300

    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Reserva = SUCESSO

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colParcelas As ColParcelaPagar

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PedVenTodos"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60227

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objPedidoVenda.lCodigo, 0, "Codigo"
    colCampoValor.Add "Cliente", objPedidoVenda.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objPedidoVenda.iFilial, 0, "Filial"
    colCampoValor.Add "FilialEntrega", objPedidoVenda.iFilialEntrega, 0, "FilialEntrega"
    colCampoValor.Add "CodTransportadora", objPedidoVenda.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CondicaoPagto", objPedidoVenda.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "PercAcrescFinanceiro", objPedidoVenda.dPercAcrescFinanceiro, 0, "PercAcrescFinanceiro"
    colCampoValor.Add "DataEmissao", objPedidoVenda.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "MensagemNota", objPedidoVenda.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
    colCampoValor.Add "NaturezaOp", objPedidoVenda.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "ValorTotal", objPedidoVenda.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorFrete", objPedidoVenda.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorDesconto", objPedidoVenda.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorSeguro", objPedidoVenda.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "PedidoCliente", objPedidoVenda.sPedidoCliente, STRING_BUFFER_MAX_TEXTO, "PedidoCliente"
    colCampoValor.Add "CanalVenda", objPedidoVenda.iCanalVenda, 0, "CanalVenda"
    colCampoValor.Add "TabelaPreco", objPedidoVenda.iTabelaPreco, 0, "TabelaPreco"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 60227

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158205)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Tela_Preenche

    objPedidoVenda.lCodigo = colCampoValor.Item("Codigo").vValor
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    If objPedidoVenda.lCodigo <> 0 Then

        'Mostra os dados do Pedido de Venda na tela
        lErro = Traz_Pedido_Tela(objPedidoVenda)
        If lErro <> SUCESSO Then Error 60228

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 60228

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158206)

    End Select

    Exit Sub

End Sub

'@@@@@ para ser customizada
Function PVC_Traz_Pedido_Tela(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
    PVC_Traz_Pedido_Tela = SUCESSO
End Function

Private Function Traz_Pedido_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca na tela os dados do Pedido de Venda

Dim lErro As Long
Dim bCancel As Boolean
Dim iStatus As Integer
Dim objTabelaPreco As New ClassTabelaPreco
Dim objFilial As AdmFiliais
Dim objCamposGenericosValores As ClassCamposGenericosValores
Dim colTipoFrete As New Collection
Dim objTipoFrete As ClassTipoFrete

On Error GoTo Erro_Traz_Pedido_Tela

    Call Limpa_Tela_PedidoVenda_Consulta2

    'Lê os dados do Pedido de Venda
    lErro = CF("PedidoDeVenda_Le_Todos_Completo", objPedidoVenda, iStatus)
    If lErro <> SUCESSO Then Error 60229
    
    If iStatus = STATUS_BAIXADO Then
        StatusPedido.Caption = "Baixado"
        BotaoDesfazer.Enabled = True 'Inserido por Wagner
    ElseIf iStatus = STATUS_ABERTO Then
        StatusPedido.Caption = "Aberto"
        BotaoDesfazer.Enabled = False 'Inserido por Wagner
    End If
    
'    lErro = TributacaoPV_Reset(objPedidoVenda)
'    If lErro <> SUCESSO Then Error 60230

    lErro = gobjTribTab.Traz_NFiscal_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60241

    lErro = CF("ReservasPedido_Le", objPedidoVenda)
    If lErro <> SUCESSO Then Error 60231

    ValorTotal.Caption = Format(objPedidoVenda.dValorTotal, "Standard")
    ValorProdutos.Caption = Format(objPedidoVenda.dValorProdutos, "Standard")
    
    PercAcrescFin.Caption = ""
    'Coloca os dados do Pedido na tela
    
    Call Cliente_Formata(objPedidoVenda.lCliente)

    Filial = objPedidoVenda.iFilial
    Call Filial_Validate(bCancel)
    
    For Each objFilial In gcolFiliais
        
        If objFilial.iCodFilial = objPedidoVenda.iFilialEmpresaFaturamento Then
            
            FilialFaturamento.Caption = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
            Exit For
        End If
    
    Next
    
    For Each objFilial In gcolFiliais
        
        If objFilial.iCodFilial = objPedidoVenda.iFilialEmpresa Then
            
            FilialPedido.Caption = objFilial.iCodFilial & SEPARADOR & objFilial.sNome
            Exit For
        End If
    
    Next
    
    Codigo.Text = objPedidoVenda.lCodigo
    'NaturezaOp.Caption = objPedidoVenda.sNaturezaOp
    FaturaIntegral.Value = objPedidoVenda.iFaturaIntegral
    
    ValorFrete.Caption = Format(objPedidoVenda.dValorFrete, "Standard")
    ValorSeguro.Caption = Format(objPedidoVenda.dValorSeguro, "Standard")
    ValorDesconto.Caption = Format(objPedidoVenda.dValorDesconto, "Standard")
    ValorDespesas.Caption = Format(objPedidoVenda.dValorOutrasDespesas, "Standard")

    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(objPedidoVenda.dtDataEmissao, "dd/mm/yy")
    DataEmissao.PromptInclude = True

    'Se a tabela de preços estiver preenchida coloca na tela
    If objPedidoVenda.iTabelaPreco > 0 Then
                    
        objTabelaPreco.iCodigo = objPedidoVenda.iTabelaPreco
        
        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then Error 60232
        
        TabelaPreco.Caption = objTabelaPreco.iCodigo & SEPARADOR & objTabelaPreco.sDescricao
        
    Else
        TabelaPreco.Caption = ""
    End If

    Placa.Caption = objPedidoVenda.sPlaca
    PlacaUF.Caption = objPedidoVenda.sPlacaUF
    
    '*** Marca de volume (Incluído por Luiz Nogueira em 21/08/03)***
    If objPedidoVenda.lVolumeMarca <> 0 Then
    
        'Instancia o obj
        Set objCamposGenericosValores = New ClassCamposGenericosValores
        
        'Guarda no obj os dados necessários para leitura
        objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_VOLUMEMARCA
        objCamposGenericosValores.lCodValor = objPedidoVenda.lVolumeMarca
        
        'Lê os dados da marca em questão
        lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
        If lErro <> SUCESSO And lErro <> 102399 Then gError 102422
        
        'Se não encontrou a marca em questão => erro
        If lErro = 102399 Then gError 102423
        
        'Exibe a marca de volume na tela
        VolumeMarca.Caption = objCamposGenericosValores.lCodValor & SEPARADOR & objCamposGenericosValores.sValor
    
    Else
    
        VolumeMarca.Caption = ""
        
    End If
    
    '***************
    
    '*** Espécie de volume (Incluído por Luiz Nogueira em 21/08/03)***
    If objPedidoVenda.lVolumeEspecie <> 0 Then
    
        'Instancia o obj
        Set objCamposGenericosValores = New ClassCamposGenericosValores
        
        'Guarda no obj os dados necessários para leitura
        objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_VOLUMEESPECIE
        objCamposGenericosValores.lCodValor = objPedidoVenda.lVolumeEspecie
        
        'Lê os dados da espécie em questão
        lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
        If lErro <> SUCESSO And lErro <> 102399 Then gError 102424
        
        'Se não encontrou a espécie em questão => erro
        If lErro = 102399 Then gError 102425
        
        'Exibe a espécie de volume na tela
        VolumeEspecie.Caption = objCamposGenericosValores.lCodValor & SEPARADOR & objCamposGenericosValores.sValor
    
    Else
    
        VolumeEspecie.Caption = ""
    
    End If
    '***************
    
    VolumeNumero.Caption = objPedidoVenda.sVolumeNumero

    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207477

    For Each objTipoFrete In colTipoFrete

        If objTipoFrete.iCodigo = objPedidoVenda.iFreteRespons Then
            TipoFrete.Caption = objTipoFrete.sDescricao
            Exit For
        End If

    Next

    If objPedidoVenda.dPesoLiq > 0 Then PesoLiquido.Caption = Format(objPedidoVenda.dPesoLiq, "Fixed")
    If objPedidoVenda.dPesoBruto > 0 Then PesoBruto.Caption = Format(objPedidoVenda.dPesoBruto, "Fixed")

    'Inserido Por Wagner
    '############################
    If objPedidoVenda.dVolumeTotal > 0 Then Cubagem.Text = Format(objPedidoVenda.dVolumeTotal, "Fixed")

    lErro = CF2(Me, "PVC_Traz_Pedido_Tela", objPedidoVenda)
    If lErro <> SUCESSO Then gError 129161
    '############################

    'Carrega o Tab Complemento
    lErro = Carrega_Tab_Complemento(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60233
    
    'Carrega o Tab Cobrança
    lErro = Carrega_Tab_Cobranca(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60234
    
    'Carrega o Tab Comissões
    lErro = objTabComissoes.Carrega_Tab_Comissoes(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60235

'    If giTipoVersao = VERSAO_FULL Then
        'Carrega o Tab Bloqueio
        lErro = Carrega_Tab_Bloqueio(objPedidoVenda)
        If lErro <> SUCESSO Then Error 60236
'    End If

    'Carrega o Grid de itens
    lErro = Carrega_Grid_Itens(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60237
    
    'Carrega o Grid de Reservas
    Call Preenche_GridReserva(objPedidoVenda.colItensPedido)
''''    lErro = Carrega_Grid_Reserva(objPedidoVenda)
''''    If lErro <> SUCESSO Then Error 60238

    ValorTotal.Caption = Format(objPedidoVenda.dValorTotal, "Standard")

'    'Carrega o Tab de Tributação
'    lErro = Carrega_Tab_Tributacao(objPedidoVenda)
'    If lErro <> SUCESSO Then Error 60239
'
'    'Traz a Tributacao a nivel de item
'    lErro = TributacaoItem_InicializaTab
'    If lErro <> SUCESSO Then Error 60240

    'Carrega o Tab de Notas Fiscais Relacionada com o Pedido
    lErro = Carrega_Tab_NotaFiscal(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60241
    
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objPedidoVenda.objProjetoInfo, objPedidoVenda.lNumIntDoc, PRJ_CR_TIPO_PV)
    If lErro <> SUCESSO Then Error 60241
    
    Traz_Pedido_Tela = SUCESSO

    Exit Function

Erro_Traz_Pedido_Tela:

    Traz_Pedido_Tela = Err

    Select Case Err

        Case 60229, 60230, 60231, 60232, 60233, 60234, 60235, 60236, _
        60237, 60238, 60239, 60240, 60241, 60242, 129161, 207477

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158207)

    End Select

    Exit Function

End Function

Private Function Posicao_ItemPV(lNumIntItemPV As Long) As Integer

Dim objItemPV As ClassItemPedido
Dim iIndice As Integer

    For Each objItemPV In gobjPedidoVenda.colItensPedido
        
        iIndice = iIndice + 1
        If objItemPV.lNumIntDoc = lNumIntItemPV Then
            Posicao_ItemPV = iIndice
            Exit For
        End If
    Next
            
End Function

Function Carrega_Tab_NotaFiscal(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_Tab_NotaFiscal
    
    'Le as Notas Fiscais Relacionadas com o Pedido ordenado por Nota
    lErro = CF("NotaFiscalPedido_Le_Nota", objPedidoVenda.lCodigo, gcolItemPedidoNF_Nota)
    If lErro <> SUCESSO Then Error 60243
    
    'Le as Notas Fiscais Relacionada com o Pedido ordenado por ItemPV
    lErro = CF("NotaFiscalPedido_Le_ItemPV", objPedidoVenda.lCodigo, gcolItemPedidoNF_ItemPV)
    If lErro <> SUCESSO Then Error 60244
    
    'Carrega o Grid de Notas Fiscais ordenando por Nota Fiscal
    lErro = Carrega_GridNFiscal()
    If lErro <> SUCESSO Then Error 60245
    
    'Carrega o Grid de Notas Fiscais ordenando por Item PV
    lErro = Carrega_GridNFItemPV()
    If lErro <> SUCESSO Then Error 60246
    
    Carrega_Tab_NotaFiscal = SUCESSO
    
    Exit Function
    
Erro_Carrega_Tab_NotaFiscal:
    
    Carrega_Tab_NotaFiscal = Err
    
    Select Case Err
        
        Case 60243, 60244, 60245, 60246
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158208)

    End Select

    Exit Function
        
End Function

Private Function Carrega_GridNFiscal() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objItemPedidoNF As ClassItemPedidoNF

On Error GoTo Erro_Carrega_GridNFiscal

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridNFiscal)

    For Each objItemPedidoNF In gcolItemPedidoNF_Nota
        
        iIndice = iIndice + 1
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemPedidoNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 60247

        'Mascara o produto enxuto
        ProdutoNF.PromptInclude = False
        ProdutoNF.Text = sProdutoEnxuto
        ProdutoNF.PromptInclude = True

        'Coloca os dados dos itens na tela
        GridNFiscal.TextMatrix(iIndice, iGrid_NFiscalNF_Col) = objItemPedidoNF.lNumNF
        GridNFiscal.TextMatrix(iIndice, iGrid_SerieNF_Col) = objItemPedidoNF.sSerie
        GridNFiscal.TextMatrix(iIndice, iGrid_ItemNF_Col) = objItemPedidoNF.iItemNF
        GridNFiscal.TextMatrix(iIndice, iGrid_ItemPVNF_Col) = Posicao_ItemPV(objItemPedidoNF.lNumIntDoc)
        GridNFiscal.TextMatrix(iIndice, iGrid_ProdutoNF_Col) = ProdutoNF.Text
        GridNFiscal.TextMatrix(iIndice, iGrid_DescProdutoNF_Col) = objItemPedidoNF.sDescricao
        GridNFiscal.TextMatrix(iIndice, iGrid_UnidadeMedNF_Col) = objItemPedidoNF.sUMItemNF
        GridNFiscal.TextMatrix(iIndice, iGrid_QuantidadeNF_Col) = Formata_Estoque(objItemPedidoNF.dQuantidade)
        GridNFiscal.TextMatrix(iIndice, iGrid_PrecoUnitarioNF_Col) = Format(objItemPedidoNF.dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridNFiscal.TextMatrix(iIndice, iGrid_PercDescNF_Col) = Format(objItemPedidoNF.dPercDesconto, "Percent")
        GridNFiscal.TextMatrix(iIndice, iGrid_DescontoNF_Col) = Format(objItemPedidoNF.dValorDesconto, "Standard")
        GridNFiscal.TextMatrix(iIndice, iGrid_PrecoTotalNF_Col) = Format(objItemPedidoNF.dPrecoTotal, "Standard")

    Next

    'Atualiza o número de linhas existentes
    objGridNFiscal.iLinhasExistentes = gcolItemPedidoNF_Nota.Count

    Exit Function
    
    Carrega_GridNFiscal = SUCESSO
    
    Exit Function
    
Erro_Carrega_GridNFiscal:

    Carrega_GridNFiscal = Err

    Select Case Err

        Case 60247
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objItemPedidoNF.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158209)

    End Select

    Exit Function

End Function

Private Function Carrega_GridNFItemPV() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objItemPedidoNF As ClassItemPedidoNF
Dim dFator As Double

On Error GoTo Erro_Carrega_GridNFItemPV

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridNFItemPV)

    For Each objItemPedidoNF In gcolItemPedidoNF_ItemPV
        
        iIndice = iIndice + 1
        
        lErro = Mascara_RetornaProdutoEnxuto(objItemPedidoNF.sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 60248

        'Mascara o produto enxuto
        ProdutoPV.PromptInclude = False
        ProdutoPV.Text = sProdutoEnxuto
        ProdutoPV.PromptInclude = True
        
        lErro = CF("UM_Conversao", objItemPedidoNF.iClasseUM, objItemPedidoNF.sUMItemNF, objItemPedidoNF.sUMItemPV, dFator)
        If lErro <> SUCESSO Then Error 60350
        
        'Coloca os dados dos itens na tela
        GridNFItemPV.TextMatrix(iIndice, iGrid_ItemPV_Col) = Posicao_ItemPV(objItemPedidoNF.lNumIntDoc)
        GridNFItemPV.TextMatrix(iIndice, iGrid_ProdutoPV_Col) = ProdutoPV.Text
        GridNFItemPV.TextMatrix(iIndice, iGrid_DescProdutoPV_Col) = objItemPedidoNF.sDescricao
        GridNFItemPV.TextMatrix(iIndice, iGrid_UnidadeMedPV_Col) = objItemPedidoNF.sUMItemPV
        GridNFItemPV.TextMatrix(iIndice, iGrid_SeriePV_Col) = objItemPedidoNF.sSerie
        GridNFItemPV.TextMatrix(iIndice, iGrid_QuantFaturadaPV_Col) = Formata_Estoque(objItemPedidoNF.dQuantFaturada * dFator)
        GridNFItemPV.TextMatrix(iIndice, iGrid_NFiscalPV_Col) = objItemPedidoNF.lNumNF
        GridNFItemPV.TextMatrix(iIndice, iGrid_ItemNFPV_Col) = objItemPedidoNF.iItemNF
        GridNFItemPV.TextMatrix(iIndice, iGrid_QuantidadePV_Col) = Formata_Estoque(objItemPedidoNF.dQuantidade * dFator)
        GridNFItemPV.TextMatrix(iIndice, iGrid_PrecoUnitarioPV_Col) = Format(objItemPedidoNF.dPrecoUnitario / dFator, gobjFAT.sFormatoPrecoUnitario)
        GridNFItemPV.TextMatrix(iIndice, iGrid_PercDescPV_Col) = Format(objItemPedidoNF.dPercDesconto, "Percent")
        GridNFItemPV.TextMatrix(iIndice, iGrid_DescontoPV_Col) = Format(objItemPedidoNF.dValorDesconto, "Standard")
        GridNFItemPV.TextMatrix(iIndice, iGrid_PrecoTotalPV_Col) = Format(objItemPedidoNF.dPrecoTotal, "Standard")

    Next

    'Atualiza o número de linhas existentes
    objGridNFItemPV.iLinhasExistentes = gcolItemPedidoNF_ItemPV.Count

    Exit Function
    
    Carrega_GridNFItemPV = SUCESSO
    
    Exit Function
    
Erro_Carrega_GridNFItemPV:

    Carrega_GridNFItemPV = Err

    Select Case Err

        Case 60248
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objItemPedidoNF.sProduto)
        
        Case 60350
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158210)

    End Select

    Exit Function

End Function

'@@@@@ para ser customizada
Function PVC_Carrega_Grid_Itens(ByVal objCT As Object, ByVal objPedidoVenda As ClassPedidoDeVenda) As Long
    PVC_Carrega_Grid_Itens = SUCESSO
End Function

Private Function Carrega_Grid_Itens(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim dPercDesc As Double
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemPV As ClassItemPedido
Dim objTabelaPreco As New ClassTabelaPreco

On Error GoTo Erro_Carrega_Grid_Itens

    'Limpa o Grid antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridItens)

    For iIndice = 1 To objPedidoVenda.colItensPedido.Count
        
        Set objItemPV = objPedidoVenda.colItensPedido(iIndice)
        
        lErro = Mascara_RetornaProdutoEnxuto(objPedidoVenda.colItensPedido(iIndice).sProduto, sProdutoEnxuto)
        If lErro <> SUCESSO Then Error 60249

        'Mascara o produto enxuto
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True

        'Calcula o percentual de desconto
        If objPedidoVenda.colItensPedido(iIndice).dPrecoTotal > 0 Then
            dPercDesc = objPedidoVenda.colItensPedido(iIndice).dValorDesconto / (objPedidoVenda.colItensPedido(iIndice).dPrecoTotal + objPedidoVenda.colItensPedido(iIndice).dValorDesconto)
        Else
            dPercDesc = 0
        End If
        
        objPedidoVenda.colItensPedido(iIndice).iItem = iIndice
                
        'Coloca os dados dos itens na tela
        
        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objPedidoVenda.colItensPedido(iIndice).iPossuiGrade = MARCADO Then GridItens.TextMatrix(iIndice, 0) = "# " & GridItens.TextMatrix(iIndice, 0)
        
        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col) = objPedidoVenda.colItensPedido(iIndice).sDescricao
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objPedidoVenda.colItensPedido(iIndice).sUnidadeMed
        GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantCancelada)
        GridItens.TextMatrix(iIndice, iGrid_PrecoUnitario_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoUnitario, gobjFAT.sFormatoPrecoUnitario)
        GridItens.TextMatrix(iIndice, iGrid_PercDesc_Col) = Format(dPercDesc, "Percent")
        GridItens.TextMatrix(iIndice, iGrid_Desconto_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dValorDesconto, "Standard")
        GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dPrecoTotal, "Standard")
        If objPedidoVenda.colItensPedido(iIndice).dtDataEntrega <> DATA_NULA Then GridItens.TextMatrix(iIndice, iGrid_DataEntrega_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dtDataEntrega, "dd/mm/yyyy")
        GridItens.TextMatrix(iIndice, iGrid_QuantReservada_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantReservada)
        GridItens.TextMatrix(iIndice, iGrid_QuantFaturada_Col) = Formata_Estoque(objPedidoVenda.colItensPedido(iIndice).dQuantFaturada)

        If gobjFAT.iExibeTabPrecoItensPV = MARCADO Then
    
            'Se a tabela de preços estiver preenchida coloca na tela
            If objItemPV.iTabelaPreco > 0 Then
                            
                objTabelaPreco.iCodigo = objItemPV.iTabelaPreco
                
                lErro = CF("TabelaPreco_Le", objTabelaPreco)
                If lErro <> SUCESSO And lErro <> 28004 Then Error 60232
                
                GridItens.TextMatrix(iIndice, iGrid_TabPrecoItem_Col) = objTabelaPreco.iCodigo & SEPARADOR & objTabelaPreco.sDescricao
                
            End If
        
        End If
        
        If gobjFAT.iExibeComissaoItensPV = MARCADO Then
        
            If objPedidoVenda.colItensPedido(iIndice).dComissao <> -1 Then
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = Format(objPedidoVenda.colItensPedido(iIndice).dComissao, "Percent")
            Else
                GridItens.TextMatrix(iIndice, iGrid_ComissaoItem_Col) = ""
            End If
        
        End If
    
    Next
    
    'Atualiza o número de linhas existentes
    objGridItens.iLinhasExistentes = objPedidoVenda.colItensPedido.Count

    'Inserido por Wagner
    '########################################
    lErro = CF2(Me, "PVC_Carrega_Grid_Itens", objPedidoVenda)
    If lErro <> SUCESSO Then gError 124180
    '########################################

    Exit Function

Erro_Carrega_Grid_Itens:

    Carrega_Grid_Itens = Err

    Select Case Err

        Case 60232

        Case 60249
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objPedidoVenda.colItensPedido(iIndice).sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158211)

    End Select

    Exit Function

End Function

Public Function PVC_CargaPosFormLoad(ByVal objCT As Object) As Long
    PVC_CargaPosFormLoad = SUCESSO
End Function

Function Carrega_Tab_Complemento(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca os dados do Tab de Complementos na tela

Dim objCanalVenda As New ClassCanalVenda
Dim objTransportadora As New ClassTransportadora
Dim lErro As Long

On Error GoTo Erro_Carrega_Tab_Complemento

    PedidoCliente.Caption = objPedidoVenda.sPedidoCliente
    PedidoRepr.Caption = objPedidoVenda.sPedidoRepresentante
    Mensagem.Text = objPedidoVenda.sMensagemNota

    If objPedidoVenda.iCanalVenda > 0 Then
        
        objCanalVenda.iCodigo = objPedidoVenda.iCanalVenda
        
        lErro = CF("CanalVenda_Le", objCanalVenda)
        If lErro <> SUCESSO And lErro <> 23597 Then Error 60249
        
        CanalVenda.Caption = objCanalVenda.iCodigo & SEPARADOR & objCanalVenda.sNomeReduzido
    
    End If

    If objPedidoVenda.iFilialEntrega > 0 Then
        Call Filial_Formata(FilialEntrega, objPedidoVenda.iFilialEntrega)
    End If

    If objPedidoVenda.dtDataEntrega <> DATA_NULA Then DataEntregaPV.Caption = Format(objPedidoVenda.dtDataEntrega, "dd/mm/yyyy")
        
    If objPedidoVenda.iCodTransportadora > 0 Then
         
        objTransportadora.iCodigo = objPedidoVenda.iCodTransportadora
        
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 60250
        
        Transportadora.Caption = objTransportadora.iCodigo & SEPARADOR & objTransportadora.sNomeReduzido
        
    Else
    
        Transportadora.Caption = ""
        
    End If

    If objPedidoVenda.iCodTranspRedesp > 0 Then
         
        objTransportadora.iCodigo = objPedidoVenda.iCodTranspRedesp
        
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 60250
        
        TranspRedespacho.Caption = objTransportadora.iCodigo & SEPARADOR & objTransportadora.sNomeReduzido
        
    Else
        
        TranspRedespacho.Caption = ""
        
    End If

    RedespachoCli.Value = objPedidoVenda.iDetPagFrete
    
    If objPedidoVenda.lVolumeQuant > 0 Then VolumeQuant.Caption = objPedidoVenda.lVolumeQuant

    Carrega_Tab_Complemento = SUCESSO
    
    Exit Function
    
Erro_Carrega_Tab_Complemento:

    Carrega_Tab_Complemento = Err
    
    Select Case Err
        
        Case 60249, 60250
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158212)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Filial_Formata(objFilial As Object, iFilial As Integer)

Dim lErro As Long
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Formata

    sCliente = Cliente.Text
    objFilialCliente.iCodFilial = iFilial

    'Pesquisa se existe Filial com o código extraído
    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
    If lErro <> SUCESSO And lErro <> 17660 Then Error 60251

    If lErro = 17660 Then Error 60252

    'Coloca na tela a Filial lida
    objFilial.Caption = iFilial & SEPARADOR & objFilialCliente.sNome

    Exit Sub

Erro_Filial_Formata:

    Select Case Err

        Case 60251
        
        Case 60252
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, objFilial.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158213)

    End Select

    Exit Sub

End Sub

Function Carrega_Tab_Cobranca(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca os dados do tab de cobrança na tela

Dim objParcelaPV As ClassParcelaPedidoVenda
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim dValorDesconto As Double
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim lErro As Long

On Error GoTo Erro_Carrega_Tab_Cobranca

    PercAcrescFin.Caption = ""

    If objPedidoVenda.iCondicaoPagto > 0 Then
        
        objCondicaoPagto.iCodigo = objPedidoVenda.iCondicaoPagto
        
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 60253
        
        CondicaoPagamento.Caption = objCondicaoPagto.iCodigo & SEPARADOR & objCondicaoPagto.sDescReduzida
        
    Else
        CondicaoPagamento.Caption = ""
    End If

    PercAcrescFin.Caption = Format(objPedidoVenda.dPercAcrescFinanceiro * 100, "Fixed")

    'Limpa o Grid de Parcelas antes de preencher com dados da coleção
    Call Grid_Limpa(objGridParcelas)

    iIndice = 0

    For Each objParcelaPV In objPedidoVenda.colParcelas

        iIndice = iIndice + 1
        GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objParcelaPV.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objParcelaPV.dValor, "Standard")
        If objParcelaPV.dtDesconto1Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Ate_Col) = Format(objParcelaPV.dtDesconto1Ate, "dd/mm/yyyy")
        If objParcelaPV.dtDesconto2Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Ate_Col) = Format(objParcelaPV.dtDesconto2Ate, "dd/mm/yyyy")
        If objParcelaPV.dtDesconto3Ate <> DATA_NULA Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Ate_Col) = Format(objParcelaPV.dtDesconto3Ate, "dd/mm/yyyy")
        If objParcelaPV.iDesconto1Codigo = VALOR_FIXO Or objParcelaPV.iDesconto1Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto1Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(objParcelaPV.dDesconto1Valor, "Standard")
        ElseIf objParcelaPV.iDesconto1Codigo = Percentual Or objParcelaPV.iDesconto1Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto1Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc1Perc_Col) = Format(objParcelaPV.dDesconto1Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaPV.dValor > 0 Then
                dValorDesconto = objParcelaPV.dDesconto1Valor * objParcelaPV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc1Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaPV.iDesconto2Codigo = VALOR_FIXO Or objParcelaPV.iDesconto2Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto2Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(objParcelaPV.dDesconto2Valor, "Standard")
        ElseIf objParcelaPV.iDesconto2Codigo = Percentual Or objParcelaPV.iDesconto2Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto2Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc2Perc_Col) = Format(objParcelaPV.dDesconto2Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaPV.dValor > 0 Then
                dValorDesconto = objParcelaPV.dDesconto2Valor * objParcelaPV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc2Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        If objParcelaPV.iDesconto3Codigo = VALOR_FIXO Or objParcelaPV.iDesconto3Codigo = VALOR_ANT_DIA Or objParcelaPV.iDesconto3Codigo = VALOR_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(objParcelaPV.dDesconto3Valor, "Standard")
        ElseIf objParcelaPV.iDesconto3Codigo = Percentual Or objParcelaPV.iDesconto3Codigo = PERC_ANT_DIA Or objParcelaPV.iDesconto3Codigo = PERC_ANT_DIA_UTIL Then
            GridParcelas.TextMatrix(iIndice, iGrid_Desc3Perc_Col) = Format(objParcelaPV.dDesconto3Valor, "Percent")
            '*** Inicio código p/ colocar Valor Desconto
            If objParcelaPV.dValor > 0 Then
                dValorDesconto = objParcelaPV.dDesconto3Valor * objParcelaPV.dValor
                GridParcelas.TextMatrix(iIndice, iGrid_Desc3Valor_Col) = Format(dValorDesconto, "Standard")
            End If
            '*** Fim
        End If
        For iIndice2 = 0 To TipoDesconto1.ListCount - 1
            If TipoDesconto1.ItemData(iIndice2) = objParcelaPV.iDesconto1Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc1Codigo_Col) = TipoDesconto1.List(iIndice2)
            If TipoDesconto2.ItemData(iIndice2) = objParcelaPV.iDesconto2Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc2Codigo_Col) = TipoDesconto2.List(iIndice2)
            If TipoDesconto3.ItemData(iIndice2) = objParcelaPV.iDesconto3Codigo Then GridParcelas.TextMatrix(iIndice, iGrid_Desc3Codigo_Col) = TipoDesconto3.List(iIndice2)
        Next

    Next

    objGridParcelas.iLinhasExistentes = iIndice

    Carrega_Tab_Cobranca = SUCESSO
    
    Exit Function
    
Erro_Carrega_Tab_Cobranca:

    Carrega_Tab_Cobranca = Err
    
    Select Case Err
        
        Case 60253
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158214)

    End Select

    Exit Function

End Function

Function Carrega_Tab_Bloqueio(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca os dados do Tab de Bloqueio na tela

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objBloqueioPV As ClassBloqueioPV
Dim objTipoBloqueio As New ClassTipoDeBloqueio

On Error GoTo Erro_Carrega_Tab_Bloqueio

    'Limpa o Grid de Bloqueios antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV

        iIndice = iIndice + 1

        objTipoBloqueio.iCodigo = objBloqueioPV.iTipoDeBloqueio

        'Lê o Tipo de bloqueio
        lErro = CF("TipoDeBloqueio_Le", objTipoBloqueio)
        If lErro <> SUCESSO And lErro <> 23666 Then Error 60254
        
        If lErro = 23666 Then Error 60255
        
        'Coloca o bloqieo no Grid de bloqueios
        GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objTipoBloqueio.iCodigo & SEPARADOR & objTipoBloqueio.sNomeReduzido
        GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPV.dtData, "dd/mm/yy")
        GridBloqueio.TextMatrix(iIndice, iGrid_CodUsuario_Col) = objBloqueioPV.sCodUsuario
        GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelBL_Col) = objBloqueioPV.sResponsavel
        If objBloqueioPV.dtDataLib <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPV.dtDataLib, "dd/mm/yy")
        GridBloqueio.TextMatrix(iIndice, iGrid_ResponsavelLib_Col) = objBloqueioPV.sResponsavelLib
        GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col) = objBloqueioPV.sObservacao

    Next

    objGridBloqueio.iLinhasExistentes = iIndice

    Carrega_Tab_Bloqueio = SUCESSO

    Exit Function

Erro_Carrega_Tab_Bloqueio:

    Carrega_Tab_Bloqueio = Err

    Select Case Err

        Case 60254
        
        Case 60255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO", Err, objTipoBloqueio.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158215)

    End Select
    
    Exit Function
    
End Function

Function Carrega_Grid_Reserva(objPedidoVenda As ClassPedidoDeVenda) As Long
'Carrega na tela as reservas dos ítens do pedido passado

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim objItemPedido As ClassItemPedido
Dim objReserva As ClassReserva
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim dFator As Double
Dim objProduto As New ClassProduto

On Error GoTo Erro_Carrega_Grid_Reserva

    'Limpa o Grid de Reservas antes de prencher com os dados da coleção
    Call Grid_Limpa(objGridReserva)

    iIndice = 0
    iIndice1 = 0

    For Each objItemPedido In objPedidoVenda.colItensPedido

        iIndice1 = iIndice1 + 1
        
        objProduto.sCodigo = objItemPedido.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 60256
        
        If lErro <> SUCESSO Then Error 60257
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPedido.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then Error 60258
        
        For Each objReserva In objItemPedido.colReserva

            iIndice = iIndice + 1

            'Coloca os dados das reservas na tela
            GridReserva.TextMatrix(iIndice, iGrid_Item_Col) = iIndice1
            GridReserva.TextMatrix(iIndice, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(iIndice1, iGrid_Produto_Col)
            GridReserva.TextMatrix(iIndice, iGrid_QuantReservar_Col) = Formata_Estoque(objItemPedido.dQuantidade * dFator)
            GridReserva.TextMatrix(iIndice, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade)
            If objReserva.dtDataValidade <> DATA_NULA Then GridReserva.TextMatrix(iIndice, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
            GridReserva.TextMatrix(iIndice, iGrid_Responsavel_Col) = objReserva.sResponsavel
            GridReserva.TextMatrix(iIndice, iGrid_UMEstoque_Col) = objItemPedido.sUMEstoque

            objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado

            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then Error 60259
            If lErro = 25056 Then Error 60260

            GridReserva.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        Next

    Next

    objGridReserva.iLinhasExistentes = iIndice

    Carrega_Grid_Reserva = SUCESSO

    Exit Function

Erro_Carrega_Grid_Reserva:

    Carrega_Grid_Reserva = Err

    Select Case Err

        Case 60256, 60259, 60258
        
        Case 60257
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)

        Case 60260
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", Err, objAlmoxarifado.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158216)

    End Select

    Exit Function

End Function

Private Sub ClienteLabel_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)


End Sub

Public Sub ComboOrdenacao_Click()

    'Se frame selecionado não for o atual
    If ComboOrdenacao.ItemData(ComboOrdenacao.ListIndex) <> giGridNotaAtual Then

        'Esconde o frame atual, mostra o novo
        FrameNFiscal(ComboOrdenacao.ItemData(ComboOrdenacao.ListIndex)).Visible = True
        FrameNFiscal(giGridNotaAtual).Visible = False
        'Armazena novo valor de giFrameAtual
        giGridNotaAtual = ComboOrdenacao.ItemData(ComboOrdenacao.ListIndex)

    End If

End Sub

Public Sub Command1_Click()
    BotaoGrade_Click
End Sub

Public Sub DataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 60261

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then Error 60262

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 60263

        If lErro = 17660 Then Error 60264

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then Error 60265

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case Err

        Case 60261, 60263

        Case 60262
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 60264
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 60265
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158217)

    End Select

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    'Dispara o Validate de Cliente
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Public Sub NumeroLabel_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection
Dim sSelecao As String
Dim iPreenchido As Integer

On Error GoTo Erro_NumeroLabel_Click

    lErro = Move_Pedido_Memoria(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60266
    
    If Len(Trim(Cliente.Text)) > 0 Then
        sSelecao = "Cliente = ?"
        iPreenchido = 1
        'Adiciona o Filtro
        colSelecao.Add (objPedidoVenda.lCliente)
    End If
    
    If Len(Trim(Filial.Text)) > 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Filial = ?"
        Else
            iPreenchido = 1
            sSelecao = "Filial = ?"
        End If
        'Adiciona o Filtro
        colSelecao.Add (objPedidoVenda.iFilial)
    End If
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND DataEmissao = ?"
        Else
            iPreenchido = 1
            sSelecao = "DataEmissao = ?"
        End If
        'Adiciona o Filtro
        colSelecao.Add (objPedidoVenda.dtDataEmissao)
    End If
    
    Call Chama_Tela("PedidoVendaTodosLista", colSelecao, objPedidoVenda, objEventoNumero, sSelecao)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 60266

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158218)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim objPedidoVenda As ClassPedidoDeVenda
Dim lErro As Long

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objPedidoVenda = obj1

    lErro = Traz_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then Error 60267

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 60267

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158219)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim objProduto As ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click
    
    'verifica se tem alguma linha selecionada no grid
    If GridItens.Row = 0 Then Error 60268
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then
            
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
            If lErro <> SUCESSO Then Error 60269

            If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
        End If

        Set objProduto = New ClassProduto
        'preenche o codigo do produto
        objProduto.sCodigo = sProduto

    'Chama a tela de browse ProdutoVendaLista
    Call Chama_Tela("Produto", objProduto)

    Exit Sub
        
Erro_BotaoProdutos_Click:
    
    Select Case Err
        
        Case 60268
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
        
        Case 60269 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158220)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objPedidoVenda Is Nothing) Then


        lErro = CF2(Me, "PVC_CargaPosFormLoad")
        If lErro <> SUCESSO Then gError 124162
 
        'Coloca o Pedido de Venda na tela
        lErro = Traz_Pedido_Tela(objPedidoVenda)
        If lErro <> SUCESSO Then Error 60270
        
        If lErro <> SUCESSO Then  'Não encontrou no BD o código de Pedido

            'Limpa a tela e coloca o código na Tela
            Call Limpa_Tela_PedidoVenda_Consulta
            Codigo.Text = CStr(objPedidoVenda.lCodigo)

        End If

    Else

        Call Limpa_Tela_PedidoVenda_Consulta

    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 60270

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158221)

    End Select

    Exit Function

End Function

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim iStatus As Integer

On Error GoTo Erro_Codigo_Validate

    If Len(Trim(Codigo.ClipText)) = 0 Then Exit Sub

    lErro = Long_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 60271

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 60271

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158222)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long, sNatOp As String, iTipoTrib As Integer
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim objTipoCliente As New ClassTipoCliente
Dim objMensagem As New ClassMensagem

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then
        
        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then Error 60272

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then Error 60273

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)
            
        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear
        
        End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 60272, 60273

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158223)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 60274

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        'se houve erro de crítica, segura o foco
        Case 60274

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158224)

    End Select

    Exit Sub

End Sub

'Public Sub TribSobreDesconto_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoDesconto)
'
'End Sub
'
'Public Sub TribSobreFrete_Click()
'
'    'exibir o frame de "outros"
'    FrameOutrosTrib.Visible = True
'    FrameItensTrib.Visible = False
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoFrete)
'
'End Sub
'
'Public Sub TribSobreItem_Click()
'
'    'se houver itens na combo
'    If gobjPedidoVenda.colItensPedido.Count <> 0 Then
'        'mostra o frame de itens e esconde o de outros
'        FrameItensTrib.Visible = True
'        FrameOutrosTrib.Visible = False
'        'selecionar o 1o item
'        ComboItensTrib.ListIndex = 0
'        Call ComboItensTrib_Click
'    Else
'        'senao houver itens na combo selecionar Frete
'        TribSobreFrete.Value = True
'        Call TribSobreFrete_Click
'    End If
'
'End Sub
'
'Public Sub TribSobreOutrasDesp_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoOutras)
'
'End Sub
'
'Public Sub TribSobreSeguro_Click()
'
'    'se o frame atual for o de itens
'    If FrameItensTrib.Visible = True Then
'        'exibir o de outros
'        FrameOutrosTrib.Visible = True
'        FrameItensTrib.Visible = False
'    End If
'
'    'Call TributacaoItem_TrazerTela(gobjPedidoVenda.objTributacaoPV.objTributacaoSeguro)
'
'End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 60275

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 60275

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158225)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 60276

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 60276

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158226)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> giFrameAtual Then

        If TabStrip_PodeTrocarTab(giFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'se abriu o tab de detalhamento
        If Opcao.SelectedItem.Index = TAB_Tributacao Then
            lErro = gobjTribTab.TabClick
            If lErro <> SUCESSO Then Error 60277
        End If

'        If giTipoVersao = VERSAO_FULL Then
            'Esconde o frame atual, mostra o novo
            Frame1(Opcao.SelectedItem.Index).Visible = True
            Frame1(giFrameAtual).Visible = False
'        Else
'            'Esconde o frame atual, mostra o novo
'            FrameL(Opcao.SelectedItem.Index).Visible = True
'            FrameL(giFrameAtual).Visible = False
'        End If
        
        'Armazena novo valor de giFrameAtual
        giFrameAtual = Opcao.SelectedItem.Index

                
        Select Case giFrameAtual
        
            Case TAB_DadosPrincipais
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_DADOS_PRINCIPAIS
            
            Case TAB_Itens
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_ITENS
                
            Case TAB_Complemento
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_COMPLEMENTO
                
            Case TAB_Cobranca
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_COBRANCA
                
            Case TAB_Comissoes
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_COMISSOES
            
            Case TAB_Bloqueio
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_BLOQUEIO
                
            Case TAB_Almoxarifado
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_ALMOXARIFADO
                
            Case TAB_NotasFiscais
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_NOTAS_FISCAIS
                
            Case TAB_Tributacao
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_TRIBUTACAO
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_CONTABILIZACAO
                
        End Select
        
    End If

    Exit Sub

Erro_Opcao_Click:

    Select Case Err

        Case 60277

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158227)

    End Select

    Exit Sub

End Sub
'
'Public Sub OpcaoTributacao_Click()
'Dim lErro As Long
'On Error GoTo Erro_OpcaoTributacao_Click
'
'    'Se frame selecionado não for o atual
'    If OpcaoTributacao.SelectedItem.Index <> giFrameAtualTributacao Then
'
'        If TabStrip_PodeTrocarTab(giFrameAtualTributacao, OpcaoTributacao, Me) <> SUCESSO Then Exit Sub
'
'        'Esconde o frame atual, mostra o novo
'        FrameTributacao(OpcaoTributacao.SelectedItem.Index).Visible = True
'        FrameTributacao(giFrameAtualTributacao).Visible = False
'        'Armazena novo valor de giFrameAtualTributacao
'        giFrameAtualTributacao = OpcaoTributacao.SelectedItem.Index
'
'        'se abriu o tab de detalhamento
'        If OpcaoTributacao.SelectedItem.Index = 2 Then
'            lErro = TributacaoItem_InicializaTab
'            If lErro <> SUCESSO Then Error 60278
'        End If
'
'    End If
'
'    Exit Sub
'
'Erro_OpcaoTributacao_Click:
'
'    Select Case Err
'
'        Case 60278
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158228)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub Limpa_Tela_PedidoVenda_Consulta()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_PedidoVenda_Consulta

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Call Limpa_Tela_PedidoVenda_Consulta2
    
    Exit Sub

Erro_Limpa_Tela_PedidoVenda_Consulta:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158229)

    End Select

    Exit Sub

End Sub


'@@@@@ para ser customizada
Public Function PVC_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    PVC_Form_Unload = SUCESSO
End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set gobjInfoUsu = Nothing

    Call CF2(Me, "PVC_Form_Unload", Cancel)
    
    'Libera as variáveis globais
    Set objEventoCliente = Nothing
    Set objEventoNumero = Nothing

    Set objGridItens = Nothing
    Set objGridComissoes = Nothing
    Set objGridParcelas = Nothing
'    If giTipoVersao = VERSAO_FULL Then
        Set objGridBloqueio = Nothing
'    End If
    Set objGridReserva = Nothing

    'eventos associados a contabilidade
'    Set gcolTiposTribICMS = Nothing
'    Set gcolTiposTribIPI = Nothing

    'Libera referencia para essa tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

'    'Encerra tributacao
'    Call TributacaoPV_Terminar

    Set objTabComissoes = Nothing
    Set gcolItemPedidoNF_Nota = Nothing
    Set gcolItemPedidoNF_ItemPV = Nothing

    Set objGridNFItemPV = Nothing
    Set objGridNFiscal = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    Set gobjPedidoVenda = Nothing
    Set gobjTribTab = Nothing

End Sub
'
'Private Function TributacaoPV_Terminar() As Long
'
'Dim lErro As Long, objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoPV_Terminar
'
'    'Se gobjPedidoVenda já foi inicializado
'    If Not (gobjPedidoVenda Is Nothing) Then
'        Set objTributoDoc = gobjPedidoVenda
'
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then Error 60279
'
'        Set gobjPedidoVenda = Nothing
'    End If
'
'    TributacaoPV_Terminar = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Terminar:
'
'    TributacaoPV_Terminar = Err
'
'    Select Case Err
'
'        Case 60279
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158230)
'
'    End Select
'
'End Function

Public Sub BotaoVendedores_Click()

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_BotaoVendedores_Click
    
    objVendedor.sNomeReduzido = GridComissoes.TextMatrix(GridComissoes.Row, 1)
    
    lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
    If lErro <> SUCESSO And lErro <> 25006 Then Error 60280
    
    If lErro = 25006 Then Error 60281
    
    Call Chama_Tela("Vendedores", objVendedor)
    
    Exit Sub
    
Erro_BotaoVendedores_Click:

    Select Case Err
        
        Case 60280
        
        Case 60281
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158231)
        
    End Select
    
    Exit Sub

End Sub

Private Function Move_Pedido_Memoria(objPedidoVenda As ClassPedidoDeVenda) As Long
'Move os dados da tela para objPedidoVenda

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim dValorTotalParcelas As Double
Dim objParcelaPV As ClassParcelaPedidoVenda
Dim dValorIRRF As Double

On Error GoTo Erro_Move_Pedido_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objPedidoVenda.lCodigo = CLng(Codigo.Text)

    'Verifica se o Cliente foi preenchido
    If Len(Trim(Cliente.ClipText)) > 0 Then

        objCliente.sNomeReduzido = Cliente.Text

        'Lê o Cliente através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 60282

        'Se não achou o Cliente --> erro
        If lErro = 12348 Then Error 60283

        'Guarda código do Cliente em objPedidoVenda
        objPedidoVenda.lCliente = objCliente.lCodigo

    End If

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objPedidoVenda.iFilial = Codigo_Extrai(Filial.Text)
    End If
   
    'Preenche objPedidoVenda com dados da tela
    objPedidoVenda.iFilialEmpresaFaturamento = Codigo_Extrai(FilialFaturamento.Caption)
    objPedidoVenda.dtDataEmissao = MaskedParaDate(DataEmissao)
    objPedidoVenda.iTabelaPreco = Codigo_Extrai(TabelaPreco.Caption)
    objPedidoVenda.sNaturezaOp = gobjTribTab.sNatOpInterna
    objPedidoVenda.iFaturaIntegral = FaturaIntegral.Value
    objPedidoVenda.dValorFrete = StrParaDbl(ValorFrete.Caption)
    objPedidoVenda.dValorSeguro = StrParaDbl(ValorSeguro.Caption)
    objPedidoVenda.dValorDesconto = StrParaDbl(ValorDesconto.Caption)
    objPedidoVenda.dValorOutrasDespesas = StrParaDbl(ValorDespesas.Caption)
    objPedidoVenda.dValorProdutos = StrParaDbl(ValorProdutos.Caption)
    objPedidoVenda.dValorTotal = StrParaDbl(ValorTotal.Caption)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa
    objPedidoVenda.dVolumeTotal = StrParaDbl(Cubagem.Text) 'Inserido por Wagner
 
    Move_Pedido_Memoria = SUCESSO

    Exit Function

Erro_Move_Pedido_Memoria:

    Move_Pedido_Memoria = Err

    Select Case Err

        Case 60282

        Case 60283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158232)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_Tela_PedidoVenda_Consulta

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158233)

    End Select

    Exit Sub

End Sub
'
'Private Function TributacaoPV_Reset(Optional objPedidoDeVenda As ClassPedidoDeVenda) As Long
''cria ou atualiza gobjPedidoVenda, com dados correspondentes a objPedidoDeVenda (se este for passado) ou com dados "padrao"
'
'Dim lErro As Long, objTributoDoc As ClassTributoDoc
'
'On Error GoTo Erro_TributacaoPV_Reset
'
'    'se gobjPedidoVenda já foi inicializado
'    If Not (gobjPedidoVenda Is Nothing) Then
'        Set objTributoDoc = gobjPedidoVenda
'        lErro = objTributoDoc.Desativar
'        If lErro <> SUCESSO Then Error 60284
'        Set gobjPedidoVenda = Nothing
'    End If
'
'    'se o pedido de venda veio preenchido
'    If Not (objPedidoDeVenda Is Nothing) Then
'
'        Set gobjPedidoVenda = objPedidoDeVenda
'
'    Else
'
'        Set gobjPedidoVenda = New ClassPedidoDeVenda
'
'    End If
'
'    Set objTributoDoc = gobjPedidoVenda
'    lErro = objTributoDoc.Ativar
'    If lErro <> SUCESSO Then Error 60285
'
'    TributacaoPV_Reset = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoPV_Reset:
'
'    TributacaoPV_Reset = Err
'
'    Select Case Err
'
'        Case 60284, 60285
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158234)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function Carrega_Tab_Tributacao(objPedidoVenda As ClassPedidoDeVenda) As Long
'
'Dim lErro As Long, objTributacaoPV As ClassTributacaoPV
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'Dim objNaturezaOp As New ClassNaturezaOp
'
'On Error GoTo Erro_Carrega_Tab_Tributacao
'
'    Set objTributacaoPV = objPedidoVenda.objTributacaoPV
'
'    objNaturezaOp.sCodigo = objPedidoVenda.sNaturezaOp
'
'    'Le a Descricao da Natureza de Operacao
'    lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'    If lErro <> SUCESSO And lErro <> 17958 Then Error 60286
'
'    NatOpEspelho.Caption = objNaturezaOp.sCodigo
'    DescNatOp.Caption = objNaturezaOp.sDescricao
'
'    'no frame de "resumo"
'    objTributacaoTipo.iTipo = objTributacaoPV.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        TipoTributacao.Caption = CStr(objTributacaoPV.iTipoTributacao)
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then Error 60287
'
'        DescTipoTrib.Caption = objTributacaoTipo.sDescricao
'
'        'se nao incide ISS
'        If objTributacaoTipo.iISSIncide = 0 Then
'            ISSValor.Enabled = False
'            ISSAliquota.Enabled = False
'            ISSIncluso.Enabled = False
'        Else
'            ISSValor.Enabled = True
'            ISSAliquota.Enabled = True
'            ISSIncluso.Enabled = True
'        End If
'
'        'se nao incide IR
'        If objTributacaoTipo.iIRIncide = 0 Then
'            ValorIRRF.Enabled = False
'            IRAliquota.Enabled = False
'        Else
'            ValorIRRF.Enabled = True
'            IRAliquota.Enabled = True
'        End If
'
'    Else
'
'        TipoTributacao.Caption = ""
'        DescTipoTrib.Caption = ""
'
'    End If
'
'    IPIBase.Caption = Format(objTributacaoPV.dIPIBase, "Standard")
'    IPIValor.Caption = Format(objTributacaoPV.dIPIValor, "Standard")
'    ISSBase.Caption = Format(objTributacaoPV.dISSBase, "Standard")
'    ISSAliquota.Caption = CStr(objTributacaoPV.dISSAliquota * 100)
'    ISSValor.Caption = Format(objTributacaoPV.dISSValor, "Standard")
'    ISSValor1.Caption = Format(objTributacaoPV.dISSValor, "Standard")
'    ISSIncluso.Value = objTributacaoPV.iISSIncluso
'    IRBase.Caption = Format(objTributacaoPV.dIRRFBase, "Standard")
'    IRAliquota.Caption = Format(objTributacaoPV.dIRRFAliquota, "Percent")
'    ValorIRRF.Caption = Format(objTributacaoPV.dIRRFValor, "Standard")
'    ISSRetido.Caption = Format(objTributacaoPV.dISSRetido, "Standard")
'    PISRetido.Caption = Format(objTributacaoPV.dPISRetido, "Standard")
'    COFINSRetido.Caption = Format(objTributacaoPV.dCOFINSRetido, "Standard")
'    CSLLRetido.Caption = Format(objTributacaoPV.dCSLLRetido, "Standard")
'    ICMSBase.Caption = Format(objTributacaoPV.dICMSBase, "Standard")
'    ICMSValor.Caption = Format(objTributacaoPV.dICMSValor, "Standard")
'    ICMSSubstBase.Caption = Format(objTributacaoPV.dICMSSubstBase, "Standard")
'    ICMSSubstValor.Caption = Format(objTributacaoPV.dICMSSubstValor, "Standard")
'    ICMSBase1.Caption = Format(objTributacaoPV.dICMSBase, "Standard")
'    ICMSValor1.Caption = Format(objTributacaoPV.dICMSValor, "Standard")
'    ICMSSubstBase1.Caption = Format(objTributacaoPV.dICMSSubstBase, "Standard")
'    ICMSSubstValor1.Caption = Format(objTributacaoPV.dICMSSubstValor, "Standard")
'    IPIValor1.Caption = Format(objTributacaoPV.dIPIValor, "Standard")
'
'    'o frame de "detalhamento" vou deixar p/carregar qdo o usuario entrar nele
'
'    Carrega_Tab_Tributacao = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Tab_Tributacao:
'
'    Carrega_Tab_Tributacao = Err
'
'    Select Case Err
'
'        Case 60286, 60287
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158235)
'
'    End Select
'
'End Function
'
'Private Function TributacaoItem_InicializaTab() As Long
''deve ser chamada na entrada do tab de detalhamento dentro do tab de tributacao
'Dim lErro As Long, objItemPedido As ClassItemPedido, sItem As String
'
'On Error GoTo Erro_TributacaoItem_InicializaTab
'
'    'preencher o valor de frete, seguro, descontos e outras desp no frameOutros
'    LabelValorFrete.Caption = Format(gobjPedidoVenda.dValorFrete, "Standard")
'    LabelValorDesconto.Caption = Format(gobjPedidoVenda.dValorDesconto, "Standard")
'    LabelValorSeguro.Caption = Format(gobjPedidoVenda.dValorSeguro, "Standard")
'    LabelValorOutrasDespesas.Caption = Format(gobjPedidoVenda.dValorOutrasDespesas, "Standard")
'
'    'esvaziar a combo de itens
'    ComboItensTrib.Clear
'
'    'preencher a combo de itens: com "codigo do produto - descricao"
'    For Each objItemPedido In gobjPedidoVenda.colItensPedido
'
'        lErro = Mascara_MascararProduto(objItemPedido.sProduto, sItem)
'        If lErro <> SUCESSO Then Error 60288
'
'        sItem = sItem & " - " & objItemPedido.sDescricao
'        ComboItensTrib.AddItem sItem
'    Next
'
'    TribSobreItem.Value = True
'    Call TribSobreItem_Click
'
'    TributacaoItem_InicializaTab = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_InicializaTab:
'
'    TributacaoItem_InicializaTab = Err
'
'    Select Case Err
'
'        Case 60288
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158236)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Function TributacaoItem_TrazerTela(objTributacaoItemPV As ClassTributacaoItemPV) As Long
''Traz para a tela dados de tributacao de um item
'
'Dim iIndice As Integer, objItemPedido As ClassItemPedido
'Dim lErro As Long, objTipoTribIPI As New ClassTipoTribIPI, objTipoTribICMS As New ClassTipoTribICMS
'Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
'Dim objNaturezaOp As New ClassNaturezaOp, sNatOp As String
'Dim iIndiceTipo As Integer
'
'On Error GoTo Erro_TributacaoItem_TrazerTela
'
'    NaturezaOpItem.Caption = objTributacaoItemPV.sNaturezaOp
'
'    sNatOp = Trim(NaturezaOpItem.Caption)
'
'    If sNatOp <> "" Then
'
'        objNaturezaOp.sCodigo = sNatOp
'
'        'Lê a Natureza de Operação
'        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
'        If lErro <> SUCESSO And lErro <> 17958 Then Error 60289
'
'        'Se não achou a Natureza de Operação --> erro
'        If lErro <> SUCESSO Then Error 60290
'
'        LabelDescrNatOpItem.Caption = objNaturezaOp.sDescricao
'    Else
'        LabelDescrNatOpItem.Caption = ""
'    End If
'
'    objTributacaoTipo.iTipo = objTributacaoItemPV.iTipoTributacao
'    If objTributacaoTipo.iTipo <> 0 Then
'
'        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
'        If lErro <> SUCESSO Then Error 60291
'
'        TipoTributacaoItem.Caption = CStr(objTributacaoItemPV.iTipoTributacao)
'        DescTipoTribItem.Caption = objTributacaoTipo.sDescricao
'
'        'Se não incide IPI
'        If objTributacaoTipo.iIPIIncide = 0 Then
'            ComboIPITipo.Enabled = False
'            IPIBaseItem.Enabled = False
'        Else
'            ComboIPITipo.Enabled = True
'            IPIBaseItem.Enabled = True
'        End If
'
'        'Se não incide ICMS
'        If objTributacaoTipo.iICMSIncide = 0 Then
'            ComboICMSTipo.Enabled = False
'            ICMSBaseItem.Enabled = False
'        Else
'            ComboICMSTipo.Enabled = True
'            ICMSBaseItem.Enabled = True
'        End If
'
'    Else
'
'        TipoTributacaoItem.Caption = ""
'        DescTipoTribItem.Caption = ""
'
'    End If
'
'    For Each objTipoTribICMS In gcolTiposTribICMS
'        If objTipoTribICMS.iTipo = objTributacaoItemPV.iICMSTipo Then Exit For
'    Next
'
'    'Setar dados de ICMS
'    ComboICMSTipo.Caption = objTipoTribICMS.iTipo & SEPARADOR & objTipoTribICMS.sDescricao
'
'    ICMSBaseItem.Caption = Format(objTributacaoItemPV.dICMSBase, "Standard")
'    ICMSPercRedBaseItem.Caption = Format(objTributacaoItemPV.dICMSPercRedBase, "Percent")
'    ICMSAliquotaItem.Caption = Format(objTributacaoItemPV.dICMSAliquota, "Percent")
'    ICMSValorItem.Caption = Format(objTributacaoItemPV.dICMSValor, "Standard")
'    ICMSSubstBaseItem.Caption = Format(objTributacaoItemPV.iICMSSubstBaseManual, "Percent")
'    ICMSSubstAliquotaItem.Caption = Format(objTributacaoItemPV.dICMSSubstAliquota, "Percent")
'    ICMSSubstValorItem.Caption = Format(objTributacaoItemPV.dICMSSubstValor, "Standard")
'
'    'Se permite redução de base habilitar este campo
'    If objTipoTribICMS.iPermiteReducaoBase Then
'        ICMSPercRedBaseItem.Enabled = True
'    Else
'        'Desabilita-lo e limpa-lo em caso contrário
'        ICMSPercRedBaseItem.Enabled = False
'    End If
'
'    'Se permite aliquota habilitar este campo e valor.
'    If objTipoTribICMS.iPermiteAliquota Then
'        ICMSAliquotaItem.Enabled = True
'        ICMSValorItem.Enabled = True
'    Else
'        'Desabilitar os dois campos e coloca-los com zero
'        ICMSAliquotaItem.Enabled = False
'        ICMSValorItem.Enabled = False
'    End If
'
'    'Se permite margem de lucro habilitar campos do frame de substituicao
'    If objTipoTribICMS.iPermiteMargLucro Then
'        ICMSSubstBaseItem.Enabled = True
'        ICMSSubstAliquotaItem.Enabled = True
'        ICMSSubstValorItem.Enabled = True
'    Else
'        'Limpa-los e desabilita-los
'        ICMSSubstBaseItem.Enabled = False
'        ICMSSubstAliquotaItem.Enabled = False
'        ICMSSubstValorItem.Enabled = False
'    End If
'
'    For Each objTipoTribIPI In gcolTiposTribIPI
'        If objTipoTribIPI.iTipo = objTributacaoItemPV.iIPITipo Then Exit For
'    Next
'
'    'Setar dados de IPI
'    ComboIPITipo.Caption = objTipoTribIPI.iTipo & SEPARADOR & objTipoTribIPI.sDescricao
'
'    IPIBaseItem.Caption = Format(objTributacaoItemPV.dIPIBaseCalculo, "Standard")
'    IPIPercRedBaseItem.Caption = Format(objTributacaoItemPV.dIPIPercRedBase, "Percent")
'    IPIAliquotaItem.Caption = Format(objTributacaoItemPV.dIPIAliquota, "Percent")
'    IPIValorItem.Caption = Format(objTributacaoItemPV.dIPIValor, "Standard")
'
'    'Se permite redução de base habilitar este campo
'    If objTipoTribIPI.iPermiteReducaoBase Then
'        IPIPercRedBaseItem.Enabled = True
'    Else
'        'desabilita-lo e limpa-lo em caso contrário
'        IPIPercRedBaseItem.Enabled = False
'    End If
'
'    'Se permite alíquota habilitar este campo e valor.
'    If objTipoTribIPI.iPermiteAliquota Then
'        IPIAliquotaItem.Enabled = True
'        IPIValorItem.Enabled = True
'    Else
'        'Desabilitar os dois campos e coloca-los com zero
'        IPIAliquotaItem.Enabled = False
'        IPIValorItem.Enabled = False
'    End If
'
'    TributacaoItem_TrazerTela = SUCESSO
'
'    Exit Function
'
'Erro_TributacaoItem_TrazerTela:
'
'    TributacaoItem_TrazerTela = Err
'
'    Select Case Err
'
'        Case 60289, 60291
'
'        Case 60290
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", objNaturezaOp.sCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158237)
'
'    End Select
'
'    Exit Function
'
'End Function
'
'Private Function AtualizarTributacao() As Long
'
'Dim lErro As Long
'
'On Error GoTo Erro_AtualizarTributacao
'
'    If Not (gobjPedidoVenda Is Nothing) Then
'
'        'Atualiza os impostos
'        lErro = gobjTributacao.AtualizaImpostos(gobjPedidoVenda, giRecalculandoTributacao)
'        If lErro <> SUCESSO Then Error 60292
'
'        'joga dados do obj atualizado p/a tela
'        lErro = Carrega_Tab_Tributacao(gobjPedidoVenda)
'        If lErro <> SUCESSO Then Error 60293
'
'    End If
'
'    AtualizarTributacao = SUCESSO
'
'    Exit Function
'
'Erro_AtualizarTributacao:
'
'    AtualizarTributacao = Err
'
'    Select Case Err
'
'        Case 60292, 60293
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158238)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Preenche_GridReserva(colItensPV As colItemPedido) As Long
'Preenche o grid de reserva com as reservas dos itens passados em colItensPV

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim dFator As Double
Dim objReserva As ClassReserva
Dim objProduto As New ClassProduto
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Preenche_GridReserva

    iIndice1 = 0

    'Limpa o grid de reservas
    Call Grid_Limpa(objGridReserva)

    'Percorre a coleção de itens
    For iIndice = 1 To colItensPV.Count

        '******** COMPARAÇÃO INCLUÍDA PARA TRATAMENTO DA GRADE ********
        If colItensPV.Item(iIndice).iPossuiGrade = MARCADO Then
            
            Call Atualiza_Grid_Alocacao(colItensPV.Item(iIndice))
            
            objProduto.sCodigo = colItensPV.Item(iIndice).sProduto
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
            If lErro <> SUCESSO Then gError 42765

            If objProduto.iKitVendaComp <> MARCADO Then
                Call Atualiza_Grid_Itens(colItensPV(iIndice))
            End If
            
        Else
    
            objProduto.sCodigo = colItensPV.Item(iIndice).sProduto
            'Lê o produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
            If lErro = 28030 Then gError 42765 'Não encontrou
            'Faz a conversão da unidade do item para a unidade de estoque
            lErro = CF("UM_Conversao", objProduto.iClasseUM, colItensPV.Item(iIndice).sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 42766
            'Para cada reserva do item
            For Each objReserva In colItensPV.Item(iIndice).colReserva
                
                'Coloca os dados da reserva no grid
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Item_Col) = colItensPV.Item(iIndice).iItem
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = GridItens.TextMatrix(colItensPV.Item(iIndice).iItem, iGrid_Produto_Col)
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = colItensPV.Item(iIndice).sUMEstoque
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objReserva.sAlmoxarifado
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade)
                If objReserva.dtDataValidade <> DATA_NULA Then GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Responsavel_Col) = objReserva.sResponsavel
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque((colItensPV.Item(iIndice).dQuantidade - colItensPV.Item(iIndice).dQuantCancelada) * dFator)

                objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado
    
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
                If lErro = 25056 Then gError 46182
    
                GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

                objGridReserva.iLinhasExistentes = objGridReserva.iLinhasExistentes + 1

            Next
        End If
    Next

    Preenche_GridReserva = SUCESSO

    Exit Function

Erro_Preenche_GridReserva:

    Preenche_GridReserva = gErr

    Select Case gErr

        Case 42764, 42766, 46181

        Case 42765
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case 46182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158239)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_PedidoVenda_Consulta2()
'Limpa os campos da tela sem fechar o sistema de setas

Dim iIndice As Integer

    Call Limpa_Tela(Me)
    StatusPedido.Caption = ""
    Codigo.Text = ""
    Filial.Clear
    'NaturezaOp.Caption = ""
    TabelaPreco.Caption = ""
    FilialFaturamento.Caption = ""
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""
    ValorTotal.Caption = ""
    ValorProdutos.Caption = ""
    Transportadora.Caption = ""
    CanalVenda.Caption = ""
    CondicaoPagamento.Caption = ""
    PercAcrescFin.Caption = ""
    FaturaIntegral.Value = vbUnchecked
    PlacaUF.Caption = ""
    TotalPercentualComissao.Caption = Format(0, "Percent")
    TipoFrete.Caption = ""
    ProdutoDescricao.Caption = ""
    
    'Limpa Valores
    ICMSBase1.Caption = ""
    ICMSValor1.Caption = ""
    ICMSSubstBase1.Caption = ""
    ValorFrete.Caption = ""
    ValorSeguro.Caption = ""
    ValorDespesas.Caption = ""
    ICMSSubstValor1.Caption = ""
    ValorProdutos.Caption = ""
    IPIValor1.Caption = ""
    ValorDesconto.Caption = ""
    ValorTotal.Caption = ""
    
    'Limpa Tab Complemento
    FilialEntrega.Caption = ""
    DataEntregaPV.Caption = ""
    Transportadora.Caption = ""
    RedespachoCli.Value = False
    Placa.Caption = ""
    PlacaUF.Caption = ""
    VolumeQuant.Caption = ""
    VolumeEspecie.Caption = ""
    VolumeMarca.Caption = ""
    VolumeNumero.Caption = ""
    PedidoCliente.Caption = ""
    PesoLiquido.Caption = ""
    PesoBruto.Caption = ""
    CanalVenda.Caption = ""
    'Mensagem.Caption = ""
    
    'Inicializa as Coleções
    Set gcolItemPedidoNF_Nota = New Collection
    Set gcolItemPedidoNF_ItemPV = New Collection

'    'tab de tributacao resumo
'    TipoTributacao.Caption = ""
'    ComboItensTrib.Clear
'    NatOpEspelho.Caption = ""
'    DescNatOp.Caption = ""
'    ISSIncluso.Value = 0
'    IPIBase.Caption = ""
'    IPIValor.Caption = ""
'    ISSBase.Caption = ""
'    DescTipoTrib.Caption = ""
'    IRBase.Caption = ""
'    ICMSBase.Caption = ""
'    ICMSValor.Caption = ""
'    ICMSSubstBase.Caption = ""
'    ICMSSubstValor.Caption = ""
'
'    'tab de tributacao itens
'    LabelValorFrete.Caption = ""
'    LabelValorDesconto.Caption = ""
'    LabelValorSeguro.Caption = ""
'    LabelValorOutrasDespesas.Caption = ""
'    LabelValorItem.Caption = ""
'    LabelQtdeItem.Caption = ""
'    LabelUMItem.Caption = ""
'    LabelDescrNatOpItem.Caption = ""
'    DescTipoTribItem.Caption = ""
'    NaturezaOpItem.Caption = ""
'    TipoTributacaoItem.Caption = ""
'    ComboICMSTipo.Caption = ""
'    ICMSBaseItem.Caption = ""
'    ICMSPercRedBaseItem.Caption = ""
'    ICMSValorItem.Caption = ""
'    ICMSAliquotaItem.Caption = ""
'    ICMSSubstBaseItem.Caption = ""
'    ICMSSubstValorItem.Caption = ""
'    ICMSSubstAliquotaItem.Caption = ""
'    ComboIPITipo.Caption = ""
'    IPIBaseItem.Caption = ""
'    IPIPercRedBaseItem.Caption = ""
'    IPIAliquotaItem.Caption = ""
'    IPIValorItem.Caption = ""
    
    BotaoDesfazer.Enabled = False 'Inserido por Wagner
    
    '************** TRATAMENTO DE GRADE **************
    For iIndice = 1 To objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    '*************************************************
    
    Call Grid_Limpa(objGridItens)
    Call Grid_Limpa(objGridNFiscal)
    Call Grid_Limpa(objGridNFItemPV)
'    If giTipoVersao = VERSAO_FULL Then
        Call Grid_Limpa(objGridBloqueio)
'    End If
    Call Grid_Limpa(objGridReserva)
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridParcelas)

'    'Resseta tributação
'    Call TributacaoPV_Reset

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    Call gobjTribTab.Limpa_Tela
    Set gobjPedidoVenda = New ClassPedidoDeVenda
    
    Exit Sub

End Sub

Public Sub GridReserva_RowColChange()

Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_GridReserva_RowColChange

    Call Grid_RowColChange(objGridReserva)
    
    If objGridReserva.iExecutarFuncao = GRID_EXECUTAR_FUNCAO Then
    
        'se a linha anterior for diferente entao preenche a decricao do produto selecionado
        If (GridReserva.Row <> giLinhaAnterior) Then
            
            If GridReserva.Row > 0 Then
            
                lErro = CF("Produto_Formata", GridReserva.TextMatrix(GridReserva.Row, iGrid_ProdutoAlmox_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then Error 60297
    
                objProduto.sCodigo = sProduto
    
                'Lê o Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then Error 60298
    
                'Se não achou o Produto --> erro
                If lErro = 28030 Then Error 60299
                    
                ProdutoDescricao.Caption = objProduto.sDescricao
            
            End If
            
            'Guarda a Linha corrente
            giLinhaAnterior = GridReserva.Row
                
        End If

    End If
    
    Exit Sub

Erro_GridReserva_RowColChange:

    Select Case Err
        
        Case 60297, 60298
        
        Case 60299
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158240)

    End Select

    Exit Sub

End Sub
'
'Private Function CarregaTiposTrib_Global() As Long
'
'Dim lErro As Long, sCodigo As String
'Dim objTipoTribICMS As ClassTipoTribICMS
'Dim objTipoTribIPI As ClassTipoTribIPI
'
'On Error GoTo Erro_CarregaTiposTrib_Global
'
'    lErro = CF("TiposTribICMS_Le_Todos", gcolTiposTribICMS)
'    If lErro <> SUCESSO Then Error 60300
'
'    lErro = CF("TiposTribIPI_Le_Todos", gcolTiposTribIPI)
'    If lErro <> SUCESSO Then Error 60301
'
'    CarregaTiposTrib_Global = SUCESSO
'
'    Exit Function
'
'Erro_CarregaTiposTrib_Global:
'
'    CarregaTiposTrib_Global = Err
'
'    Select Case Err
'
'        Case 60300, 60301
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158241)
'
'    End Select
'
'    Exit Function
'
'End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_PEDIDO_VENDA_CONS_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Consulta de Pedido de Venda"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PedidoVenda_Consulta"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Private Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
        End If
    
    End If

End Sub

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente
    
    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then Error 60315

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then Error 60316

    'Preenche ComboBox de Filiais
    Call CF("Filial_Preenche", Filial, colCodigoNome)

    Exit Sub

Erro_Cliente_Formata:

    Select Case Err
    
        Case 60315, 60316
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158242)

    End Select

    Exit Sub

End Sub

'Public Sub OpcaoTributacao_BeforeClick(Cancel As Integer)
'    Call TabStrip_TrataBeforeClick(Cancel, OpcaoTributacao)
'End Sub

Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Sub Atualiza_Grid_Alocacao(objItemPV As ClassItemPedido)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objReserva As ClassReservaItem
Dim sProdutoMascarado As String
Dim lErro As Long
Dim dFator As Double
Dim dFator2 As Double
Dim objProduto As New ClassProduto
Dim dQuantReservada As Double
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_Atualiza_Grid_Alocacao

    For Each objItemRomaneio In objItemPV.colItensRomaneioGrade

        objProduto.sCodigo = objItemPV.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 42764
        If lErro = 28030 Then gError 42765 'Não encontrou
        'Faz a conversão da unidade do item para a unidade de estoque
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemRomaneio.sUMEstoque, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 42766
        
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPV.sUnidadeMed, objProduto.sSiglaUMEstoque, dFator2)
        If lErro <> SUCESSO Then gError 42766
        
        dQuantReservada = 0
        
        For Each objReserva In objItemRomaneio.colLocalizacao
        
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Item_Col) = objItemPV.iItem
            
            lErro = Mascara_MascararProduto(objItemRomaneio.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 86357
                        
            objAlmoxarifado.iCodigo = objReserva.iAlmoxarifado

            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 46181
            If lErro = 25056 Then gError 46182
                
            objReserva.sAlmoxarifado = objAlmoxarifado.sNomeReduzido
                        
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_ProdutoAlmox_Col) = sProdutoMascarado
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_UMEstoque_Col) = objItemRomaneio.sUMEstoque
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Almoxarifado_Col) = objReserva.sAlmoxarifado
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_QuantReservar_Col) = Formata_Estoque((objItemRomaneio.dQuantidade - objItemRomaneio.dQuantCancelada) * dFator2)
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_QuantReserv_Col) = Formata_Estoque(objReserva.dQuantidade * dFator)
            If objReserva.dtDataValidade <> DATA_NULA Then GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Validade_Col) = Format(objReserva.dtDataValidade, "dd/mm/yyyy")
            GridReserva.TextMatrix(objGridReserva.iLinhasExistentes + 1, iGrid_Responsavel_Col) = objReserva.sResponsavel
            
            objGridReserva.iLinhasExistentes = objGridReserva.iLinhasExistentes + 1
            
        Next
               
    Next

    Exit Sub
    
Erro_Atualiza_Grid_Alocacao:

    Select Case gErr
    
        Case 461821, 86357
        
        Case 46182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE2", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158243)
            
    End Select
    
    Exit Sub

End Sub

Sub Atualiza_Grid_Itens(objItemPV As ClassItemPedido)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim dQuantCancelada As Double
Dim dQuantReservada As Double
Dim dQuantFaturada As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
    
    For Each objItemRomaneioGrade In objItemPV.colItensRomaneioGrade
            
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
        dQuantCancelada = dQuantCancelada + objItemRomaneioGrade.dQuantCancelada
        dQuantReservada = dQuantReservada + objItemRomaneioGrade.dQuantReservada
        dQuantFaturada = dQuantFaturada + objItemRomaneioGrade.dQuantFaturada
        
    Next

    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantPedida_Col) = Formata_Estoque(dQuantidade)
    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantCancel_Col) = Formata_Estoque(dQuantCancelada)
    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantReservada_Col) = Formata_Estoque(dQuantReservada)
    GridItens.TextMatrix(objItemPV.iItem, iGrid_QuantFaturada_Col) = Formata_Estoque(dQuantFaturada)

    objItemPV.dQuantidade = dQuantidade
    objItemPV.dQuantCancelada = dQuantCancelada

    
    Exit Sub

End Sub

Public Sub BotaoGrade_Click()

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemPV As ClassItemPedido

On Error GoTo Erro_BotaoGrade_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemPV = gobjPedidoVenda.colItensPedido(GridItens.Row)
        
        If objItemPV.iPossuiGrade = MARCADO Then
            
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            objRomaneioGrade.iFilialFaturamento = Codigo_Extrai(FilialFaturamento.Caption)
            Set objRomaneioGrade.objObjetoTela = objItemPV
                        
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
            If lErro <> SUCESSO Then gError 86360
                        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 51037, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158244)
            
    End Select
    
    Exit Sub

End Sub

Function Transfere_Dados_ItensRomaneio(colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio existente do BD (Produtos Filhos do Produto passado)
    For Each objItemRomaneioGrade In colItensRomaneio
        'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
        For Each objItemRomaneioGradeTela In colItensRomaneioTela
            'Se encontrou o Item
            If objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto Then
                'Transfere as informações vindas da tela chamadora para essa tela
                objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
                objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
                objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
                objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
                objItemRomaneioGrade.dQuantFaturada = objItemRomaneioGradeTela.dQuantFaturada
                objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
                objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
                objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
                objItemRomaneioGrade.dQuantCancelada = objItemRomaneioGradeTela.dQuantCancelada
                
                'Transfere as informações de Localização
                Set objItemRomaneioGrade.colLocalizacao = New Collection
                    
                For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
                    
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                    objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                    objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                    objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                    objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                    objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                    objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                    
                    objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                    
                Next
                            
            End If
        
        Next
    Next

    Exit Function

End Function

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Public Sub Cliente_Change()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Change
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134010

    Exit Sub

Erro_Cliente_Change:

    Select Case gErr

        Case 134010

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158245)

    End Select
    
    Exit Sub

End Sub

'###################################################
'Inserido por Wagner
'Função copiada da versão da Martineli
Public Sub BotaoDesfazer_Click()

Dim objPedidoVenda As New ClassPedidoDeVenda
Dim lErro As Long

On Error GoTo Erro_BotaoDesfazer_Click

    If Len(Trim(Codigo.Text)) = 0 Then gError 86321

    objPedidoVenda.lCodigo = StrParaLong(Codigo.Text)
    objPedidoVenda.iFilialEmpresa = giFilialEmpresa

    lErro = CF("PedidoVenda_Desfaz_Baixa", objPedidoVenda)
    If lErro <> SUCESSO Then gError 86322
    
    Call Limpa_Tela_PedidoVenda_Consulta
    
    Exit Sub

Erro_BotaoDesfazer_Click:

    Select Case gErr
    
        Case 86321
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_INFORMADO1", gErr)
        
        Case 86322
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158246)
            
    End Select
    
    Exit Sub

End Sub
'###################################################

'#####################################################################
'Inserido por Wagner 13/09/05
Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_PEDIDOVENDA
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = CStr(giFilialEmpresa) & "," & Codigo.Text
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158247)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158248)
     
    End Select
     
    Exit Sub

End Sub
'#############################################################################

Public Function BotaoProjetos_Click() As Long
        Call gobjTelaProjetoInfo.BotaoProjetos_Click(True)
End Function

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objItensPRJCR = New ClassItensPRJCR
        
        objItensPRJCR.sItem = CStr(iIndice)
        objItensPRJCR.sDescricao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantPedida_Col)) - StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantCancel_Col))
        objItensPRJCR.dValorOriginal = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_PrecoTotal_Col))
        objItensPRJCR.sObservacao = GridItens.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridItens.TextMatrix(iIndice, iGrid_DescProduto_Col)

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_PV
        
        colItensPRJ.Add objItensPRJCR

    Next

    Call gobjTribTab.Preenche_Itens_Projeto_NF(colItensPRJ)
    
    Obter_ItensPRJCR = SUCESSO

    Exit Function

Erro_Obter_ItensPRJCR:

    Obter_ItensPRJCR = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
    
End Function

Public Function Preenche_Itens_Projeto_Aux(ByVal colItensPRJ As Collection, ByVal sItem As String, ByVal sDescricao As String, ByVal dValor As Double, ByVal iTipo As Integer, Optional ByVal iNegativo As Integer = 0, Optional ByVal sOBS As String = "") As Long

Dim objItensPRJCR As New ClassItensPRJCR

On Error GoTo Erro_Preenche_Itens_Projeto_Aux

    objItensPRJCR.sItem = sItem
    objItensPRJCR.sDescricao = sDescricao
    objItensPRJCR.dValorOriginal = dValor
    objItensPRJCR.iNegativo = iNegativo
    objItensPRJCR.iTipoOrigem = iTipo
    objItensPRJCR.sObservacao = sOBS

    colItensPRJ.Add objItensPRJCR

    Exit Function

Erro_Preenche_Itens_Projeto_Aux:

    Preenche_Itens_Projeto_Aux = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
    
End Function

Function ValorTotal_Calcula() As Long
    ValorTotal_Calcula = SUCESSO
End Function
