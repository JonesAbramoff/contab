VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTOrdemProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjAnotacao As ClassAnotacoes

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

Public iAlterado As Integer
Dim iLinhaAntiga As Integer
Dim iCodigoAlterado As Integer

Dim iFornecedorTercAlterado As Integer
Dim iClienteTercAlterado As Integer

Dim iPrestServAlterado As Integer

Public gcolItemOP As Collection
'criado por causa da forma como foi construido a tela romaneiograde
Public gobjOP As ClassOrdemDeProducao

Public objGrid As AdmGrid
Public iGrid_Produto_Col As Integer
Public iGrid_UnidadeMed_Col  As Integer
Public iGrid_Quantidade_Col  As Integer
Public iGrid_Almoxarifado_Col  As Integer
Public iGrid_Benef_Col  As Integer
Public iGrid_Ccl_Col  As Integer
Public iGrid_DescricaoItem_Col  As Integer
Public iGrid_DataPrevInicio_Col  As Integer
Public iGrid_DataPrevFim_Col  As Integer
Public iGrid_Situacao_Col  As Integer
Public iGrid_Destinacao_Col  As Integer
Public iGrid_PedidoDeVenda_Col  As Integer
Public iGrid_FilialPedido_Col  As Integer
Public iGrid_Cliente_Col As Integer
Public iGrid_FilialCliente_Col As Integer
Public iGrid_Prioridade_Col  As Integer
Public iGrid_Versao_Col  As Integer
Public iGrid_Maquina_Col  As Integer
Public iGrid_ProduzLogo_Col  As Integer

Public iFrameAtual As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoCclPadrao As AdmEvento
Attribute objEventoCclPadrao.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxPadrao As AdmEvento
Attribute objEventoAlmoxPadrao.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoPedidoDeVenda As AdmEvento
Attribute objEventoPedidoDeVenda.VB_VarHelpID = -1
Private WithEvents objEventoMaquina As AdmEvento
Attribute objEventoMaquina.VB_VarHelpID = -1
Private WithEvents objEventoPrestServ As AdmEvento
Attribute objEventoPrestServ.VB_VarHelpID = -1

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

'---------------------------------------
'Inserido por Jorge Specian - 29/04/2005
Dim gbOPValidaParaMRP As Boolean
Dim iFrameAtualOper As Integer

'Grid de OperacaoInsumos
Dim objGridOperacaoInsumos As AdmGrid
Dim iGrid_ProdutoInsumos_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_OrigemProduto_Col As Integer
Dim iGrid_QuantidadeProduto_Col As Integer
Dim iGrid_UMProduto_Col As Integer
Dim iGrid_VersaoKitComp_Col As Integer
Dim iGrid_PerdaInsumos_Col As Integer
Dim iGrid_ComposicaoInsumos_Col As Integer
Dim iGrid_CustoStandardInsumos_Col As Integer

'Grid de Maquinas
Dim objGridMaquinas As AdmGrid
Dim iGrid_NomeMaquina_Col As Integer
Dim iGrid_QuantidadeMaquina_Col As Integer
Dim iGrid_DataPO_Col As Integer
Dim iGrid_HorasPO_Col As Integer
Dim iGrid_TaxaProducao_Col As Integer

Dim colPMPItens As Collection
Dim colOPItens As Collection
Dim colComponentes As Collection
Dim iProxChave As Integer

'variaveis auxiliares para recalculo de nivel e sequencial
Dim aNivelSequencial(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o maior sequencial
Dim aSeqPai(NIVEL_MAXIMO_OPERACOES) As Integer 'para cada nivel guarda o SeqPai

Dim iUltimoNivel As Integer

Public iCompetenciaAlterada As Integer
Public iQtdeAlterada As Integer
Public giItemOP As Integer

Private WithEvents objEventoCompetencias As AdmEvento
Attribute objEventoCompetencias.VB_VarHelpID = -1
Private WithEvents objEventoCentroDeTrabalho As AdmEvento
Attribute objEventoCentroDeTrabalho.VB_VarHelpID = -1

Private Const TAB_Operacoes = 3
'---------------------------------------

'################################
'Inserido por Wagner 07/11/2005
'Grid de Maquinas
Dim objGridMaquinasReal As AdmGrid
Dim iGrid_NomeMaquinaReal_Col As Integer
Dim iGrid_QuantidadeMaquinaReal_Col As Integer
Dim iGrid_DataReal_Col As Integer
Dim iGrid_HorasReal_Col As Integer

Dim objGridMOReal As AdmGrid
Dim iGrid_NomeMaquinaRealMO_Col As Integer
Dim iGrid_QuantidadeMaquinaRealMO_Col As Integer
Dim iGrid_DataRealMO_Col As Integer
Dim iGrid_HorasRealMO_Col As Integer
Dim iGrid_TipoMaoDeObraReal_Col As Integer
Dim iGrid_DescricaoTipoMO_Col As Integer

Private WithEvents objEventoMaquinas As AdmEvento
Attribute objEventoMaquinas.VB_VarHelpID = -1
Private WithEvents objEventoMO As AdmEvento
Attribute objEventoMO.VB_VarHelpID = -1
'#################################

'--- inicio dos properties get dos controles da tela

Public Property Get IgnoraEst() As Object
     Set IgnoraEst = objUserControl.Controls("IgnoraEst")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get ClienteLabel() As Object
     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
End Property

Public Property Get FrameClienteFornec() As Object
     Set FrameClienteFornec = objUserControl.Controls("FrameClienteFornec")
End Property

Public Property Get FornecedorTerc() As Object
     Set FornecedorTerc = objUserControl.Controls("FornecedorTerc")
End Property

Public Property Get FilialTerc() As Object
     Set FilialTerc = objUserControl.Controls("FilialTerc")
End Property

Public Property Get ClienteTerc() As Object
     Set ClienteTerc = objUserControl.Controls("ClienteTerc")
End Property

Public Property Get EscaninhoTerc() As Object
     Set EscaninhoTerc = objUserControl.Controls("EscaninhoTerc")
End Property

Public Property Get OptionTipoTerc(ByVal iIndex As Integer) As Object
     Set OptionTipoTerc = objUserControl.Controls("OptionTipoTerc")(iIndex)
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get DestinacaoPadrao() As Object
     Set DestinacaoPadrao = objUserControl.Controls("DestinacaoPadrao")
End Property

Public Property Get CclPadrao() As Object
     Set CclPadrao = objUserControl.Controls("CclPadrao")
End Property

Public Property Get AlmoxPadrao() As Object
     Set AlmoxPadrao = objUserControl.Controls("AlmoxPadrao")
End Property

Public Property Get UpDownInicio() As Object
     Set UpDownInicio = objUserControl.Controls("UpDownInicio")
End Property

Public Property Get DataInicioPadrao() As Object
     Set DataInicioPadrao = objUserControl.Controls("DataInicioPadrao")
End Property

Public Property Get UpDownFim() As Object
     Set UpDownFim = objUserControl.Controls("UpDownFim")
End Property

Public Property Get DataFimPadrao() As Object
     Set DataFimPadrao = objUserControl.Controls("DataFimPadrao")
End Property

Public Property Get PrioridadePadrao() As Object
     Set PrioridadePadrao = objUserControl.Controls("PrioridadePadrao")
End Property

Public Property Get PrioridadePadraoLbl() As Object
     Set PrioridadePadraoLbl = objUserControl.Controls("PrioridadePadraoLbl")
End Property

Public Property Get DestPadraoLbl() As Object
     Set DestPadraoLbl = objUserControl.Controls("DestPadraoLbl")
End Property

Public Property Get DataPrevFimLbl() As Object
     Set DataPrevFimLbl = objUserControl.Controls("DataPrevFimLbl")
End Property

Public Property Get DataPrevIniLbl() As Object
     Set DataPrevIniLbl = objUserControl.Controls("DataPrevIniLbl")
End Property

Public Property Get AlmoxPadraoLabel() As Object
     Set AlmoxPadraoLabel = objUserControl.Controls("AlmoxPadraoLabel")
End Property

Public Property Get CclPadraoLabel() As Object
     Set CclPadraoLabel = objUserControl.Controls("CclPadraoLabel")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get BotaoProjetos() As Object
     Set BotaoProjetos = objUserControl.Controls("BotaoProjetos")
End Property

Public Property Get BotaoOPRC() As Object
     Set BotaoOPRC = objUserControl.Controls("BotaoOPRC")
End Property

Public Property Get botaoProxNum2() As Object
     Set botaoProxNum2 = objUserControl.Controls("botaoProxNum2")
End Property

Public Property Get BotaoProxNum() As Object
     Set BotaoProxNum = objUserControl.Controls("BotaoProxNum")
End Property

Public Property Get PrestadorServico() As Object
     Set PrestadorServico = objUserControl.Controls("PrestadorServico")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get GeraOP() As Object
     Set GeraOP = objUserControl.Controls("GeraOP")
End Property

Public Property Get GeraReqCompra() As Object
     Set GeraReqCompra = objUserControl.Controls("GeraReqCompra")
End Property

Public Property Get CodOPGerada() As Object
     Set CodOPGerada = objUserControl.Controls("CodOPGerada")
End Property

Public Property Get UpDownData() As Object
     Set UpDownData = objUserControl.Controls("UpDownData")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get GeraOPs() As Object
     Set GeraOPs = objUserControl.Controls("GeraOPs")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get LabelProjeto() As Object
     Set LabelProjeto = objUserControl.Controls("LabelProjeto")
End Property

Public Property Get LabelPrestador() As Object
     Set LabelPrestador = objUserControl.Controls("LabelPrestador")
End Property

Public Property Get CodigoOPLabel() As Object
     Set CodigoOPLabel = objUserControl.Controls("CodigoOPLabel")
End Property

Public Property Get LabelCodOPGerada() As Object
     Set LabelCodOPGerada = objUserControl.Controls("LabelCodOPGerada")
End Property

Public Property Get StatusOP() As Object
     Set StatusOP = objUserControl.Controls("StatusOP")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get OpcaoSimples() As Object
     Set OpcaoSimples = objUserControl.Controls("OpcaoSimples")
End Property

Public Property Get OpcaoDetalhado() As Object
     Set OpcaoDetalhado = objUserControl.Controls("OpcaoDetalhado")
End Property

Public Property Get OpcaoRelatorio() As Object
     Set OpcaoRelatorio = objUserControl.Controls("OpcaoRelatorio")
End Property

Public Property Get ImprimeAoGravar() As Object
     Set ImprimeAoGravar = objUserControl.Controls("ImprimeAoGravar")
End Property

Public Property Get BotaoKit() As Object
     Set BotaoKit = objUserControl.Controls("BotaoKit")
End Property

Public Property Get BotaoRoteiros() As Object
     Set BotaoRoteiros = objUserControl.Controls("BotaoRoteiros")
End Property

Public Property Get FrameConsideraAlgoritmo() As Object
     Set FrameConsideraAlgoritmo = objUserControl.Controls("FrameConsideraAlgoritmo")
End Property

Public Property Get OptionMRPII() As Object
     Set OptionMRPII = objUserControl.Controls("OptionMRPII")
End Property

Public Property Get OptionMRP() As Object
     Set OptionMRP = objUserControl.Controls("OptionMRP")
End Property

Public Property Get Observacao() As Object
     Set Observacao = objUserControl.Controls("Observacao")
End Property

Public Property Get CodigoCTPadrao() As Object
     Set CodigoCTPadrao = objUserControl.Controls("CodigoCTPadrao")
End Property

Public Property Get CodigoCompetencia() As Object
     Set CodigoCompetencia = objUserControl.Controls("CodigoCompetencia")
End Property

Public Property Get NumMaxMaqPorOper() As Object
     Set NumMaxMaqPorOper = objUserControl.Controls("NumMaxMaqPorOper")
End Property

Public Property Get Repeticao() As Object
     Set Repeticao = objUserControl.Controls("Repeticao")
End Property

Public Property Get LabelRepeticao() As Object
     Set LabelRepeticao = objUserControl.Controls("LabelRepeticao")
End Property

Public Property Get LabelNumMaxMaqPorOper() As Object
     Set LabelNumMaxMaqPorOper = objUserControl.Controls("LabelNumMaxMaqPorOper")
End Property

Public Property Get LabelDetUM() As Object
     Set LabelDetUM = objUserControl.Controls("LabelDetUM")
End Property

Public Property Get UMLabel() As Object
     Set UMLabel = objUserControl.Controls("UMLabel")
End Property

Public Property Get LabelDetQtde() As Object
     Set LabelDetQtde = objUserControl.Controls("LabelDetQtde")
End Property

Public Property Get QtdeLabel() As Object
     Set QtdeLabel = objUserControl.Controls("QtdeLabel")
End Property

Public Property Get LabelDetProduto() As Object
     Set LabelDetProduto = objUserControl.Controls("LabelDetProduto")
End Property

Public Property Get ProdutoLabel() As Object
     Set ProdutoLabel = objUserControl.Controls("ProdutoLabel")
End Property

Public Property Get LabelDetVersao() As Object
     Set LabelDetVersao = objUserControl.Controls("LabelDetVersao")
End Property

Public Property Get VersaoLabel() As Object
     Set VersaoLabel = objUserControl.Controls("VersaoLabel")
End Property

Public Property Get LabelCodigoCTPadrao() As Object
     Set LabelCodigoCTPadrao = objUserControl.Controls("LabelCodigoCTPadrao")
End Property

Public Property Get LabelCodigoCompetencia() As Object
     Set LabelCodigoCompetencia = objUserControl.Controls("LabelCodigoCompetencia")
End Property

Public Property Get CompetenciaLabel() As Object
     Set CompetenciaLabel = objUserControl.Controls("CompetenciaLabel")
End Property

Public Property Get CTLabel() As Object
     Set CTLabel = objUserControl.Controls("CTLabel")
End Property

Public Property Get LabelObservacao() As Object
     Set LabelObservacao = objUserControl.Controls("LabelObservacao")
End Property

Public Property Get LabelSeq() As Object
     Set LabelSeq = objUserControl.Controls("LabelSeq")
End Property

Public Property Get Nivel() As Object
     Set Nivel = objUserControl.Controls("Nivel")
End Property

Public Property Get Sequencial() As Object
     Set Sequencial = objUserControl.Controls("Sequencial")
End Property

Public Property Get LabelNivel() As Object
     Set LabelNivel = objUserControl.Controls("LabelNivel")
End Property

Public Property Get DescricaoCompetencia() As Object
     Set DescricaoCompetencia = objUserControl.Controls("DescricaoCompetencia")
End Property

Public Property Get DescricaoCTPadrao() As Object
     Set DescricaoCTPadrao = objUserControl.Controls("DescricaoCTPadrao")
End Property

Public Property Get NomeMaquinaRealMO() As Object
     Set NomeMaquinaRealMO = objUserControl.Controls("NomeMaquinaRealMO")
End Property

Public Property Get DescricaoTipoMO() As Object
     Set DescricaoTipoMO = objUserControl.Controls("DescricaoTipoMO")
End Property

Public Property Get TipoMaoDeObraReal() As Object
     Set TipoMaoDeObraReal = objUserControl.Controls("TipoMaoDeObraReal")
End Property

Public Property Get BotaoMO() As Object
     Set BotaoMO = objUserControl.Controls("BotaoMO")
End Property

Public Property Get HorasRealMO() As Object
     Set HorasRealMO = objUserControl.Controls("HorasRealMO")
End Property

Public Property Get DataRealMO() As Object
     Set DataRealMO = objUserControl.Controls("DataRealMO")
End Property

Public Property Get QuantidadeMaquinaRealMO() As Object
     Set QuantidadeMaquinaRealMO = objUserControl.Controls("QuantidadeMaquinaRealMO")
End Property

Public Property Get GridMOReal() As Object
     Set GridMOReal = objUserControl.Controls("GridMOReal")
End Property

Public Property Get BotaoTrazerMaquinasPrevistas() As Object
     Set BotaoTrazerMaquinasPrevistas = objUserControl.Controls("BotaoTrazerMaquinasPrevistas")
End Property

Public Property Get NomeMaquinaReal() As Object
     Set NomeMaquinaReal = objUserControl.Controls("NomeMaquinaReal")
End Property

Public Property Get BotaoMaquinasReal() As Object
     Set BotaoMaquinasReal = objUserControl.Controls("BotaoMaquinasReal")
End Property

Public Property Get HorasReal() As Object
     Set HorasReal = objUserControl.Controls("HorasReal")
End Property

Public Property Get DataReal() As Object
     Set DataReal = objUserControl.Controls("DataReal")
End Property

Public Property Get QuantidadeMaquinaReal() As Object
     Set QuantidadeMaquinaReal = objUserControl.Controls("QuantidadeMaquinaReal")
End Property

Public Property Get GridMaquinasReal() As Object
     Set GridMaquinasReal = objUserControl.Controls("GridMaquinasReal")
End Property

Public Property Get BotaoGrafico() As Object
     Set BotaoGrafico = objUserControl.Controls("BotaoGrafico")
End Property

Public Property Get BotaoCT() As Object
     Set BotaoCT = objUserControl.Controls("BotaoCT")
End Property

Public Property Get BotaoCompetencia() As Object
     Set BotaoCompetencia = objUserControl.Controls("BotaoCompetencia")
End Property

Public Property Get BotaoMaquinaPMP() As Object
     Set BotaoMaquinaPMP = objUserControl.Controls("BotaoMaquinaPMP")
End Property

Public Property Get BotaoAbrirOP() As Object
     Set BotaoAbrirOP = objUserControl.Controls("BotaoAbrirOP")
End Property

Public Property Get HorasPO() As Object
     Set HorasPO = objUserControl.Controls("HorasPO")
End Property

Public Property Get DataPO() As Object
     Set DataPO = objUserControl.Controls("DataPO")
End Property

Public Property Get FrameDatas() As Object
     Set FrameDatas = objUserControl.Controls("FrameDatas")
End Property

Public Property Get LabelDataInicio() As Object
     Set LabelDataInicio = objUserControl.Controls("LabelDataInicio")
End Property

Public Property Get DataInicio() As Object
     Set DataInicio = objUserControl.Controls("DataInicio")
End Property

Public Property Get DataFinal() As Object
     Set DataFinal = objUserControl.Controls("DataFinal")
End Property

Public Property Get LabelDataFinal() As Object
     Set LabelDataFinal = objUserControl.Controls("LabelDataFinal")
End Property

Public Property Get MRP() As Object
     Set MRP = objUserControl.Controls("MRP")
End Property

Public Property Get OPCodigoMRP() As Object
     Set OPCodigoMRP = objUserControl.Controls("OPCodigoMRP")
End Property

Public Property Get NomeMaquina() As Object
     Set NomeMaquina = objUserControl.Controls("NomeMaquina")
End Property

Public Property Get TaxaProducao() As Object
     Set TaxaProducao = objUserControl.Controls("TaxaProducao")
End Property

Public Property Get QuantidadeMaquina() As Object
     Set QuantidadeMaquina = objUserControl.Controls("QuantidadeMaquina")
End Property

Public Property Get GridMaquinas() As Object
     Set GridMaquinas = objUserControl.Controls("GridMaquinas")
End Property

Public Property Get LabelMRP() As Object
     Set LabelMRP = objUserControl.Controls("LabelMRP")
End Property

Public Property Get LabelOP() As Object
     Set LabelOP = objUserControl.Controls("LabelOP")
End Property

Public Property Get LabelMaquinas() As Object
     Set LabelMaquinas = objUserControl.Controls("LabelMaquinas")
End Property

Public Property Get OrigemProduto() As Object
     Set OrigemProduto = objUserControl.Controls("OrigemProduto")
End Property

Public Property Get DescricaoProduto() As Object
     Set DescricaoProduto = objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get UMProduto() As Object
     Set UMProduto = objUserControl.Controls("UMProduto")
End Property

Public Property Get ComposicaoInsumos() As Object
     Set ComposicaoInsumos = objUserControl.Controls("ComposicaoInsumos")
End Property

Public Property Get VersaoKitComp() As Object
     Set VersaoKitComp = objUserControl.Controls("VersaoKitComp")
End Property

Public Property Get ProdutoInsumos() As Object
     Set ProdutoInsumos = objUserControl.Controls("ProdutoInsumos")
End Property

Public Property Get PerdaInsumos() As Object
     Set PerdaInsumos = objUserControl.Controls("PerdaInsumos")
End Property

Public Property Get CustoStandardInsumos() As Object
     Set CustoStandardInsumos = objUserControl.Controls("CustoStandardInsumos")
End Property

Public Property Get QuantidadeProduto() As Object
     Set QuantidadeProduto = objUserControl.Controls("QuantidadeProduto")
End Property

Public Property Get GridOperacaoInsumos() As Object
     Set GridOperacaoInsumos = objUserControl.Controls("GridOperacaoInsumos")
End Property

Public Property Get BotaoImprimirOper() As Object
     Set BotaoImprimirOper = objUserControl.Controls("BotaoImprimirOper")
End Property

Public Property Get FrameItemOP() As Object
     Set FrameItemOP = objUserControl.Controls("FrameItemOP")
End Property

Public Property Get UpDownItemOP() As Object
     Set UpDownItemOP = objUserControl.Controls("UpDownItemOP")
End Property

Public Property Get ItemOP() As Object
     Set ItemOP = objUserControl.Controls("ItemOP")
End Property

Public Property Get FrameRoteiro() As Object
     Set FrameRoteiro = objUserControl.Controls("FrameRoteiro")
End Property

Public Property Get Roteiro() As Object
     Set Roteiro = objUserControl.Controls("Roteiro")
End Property

Public Property Get BotaoAlterar() As Object
     Set BotaoAlterar = objUserControl.Controls("BotaoAlterar")
End Property

Public Property Get TabStrip2() As Object
     Set TabStrip2 = objUserControl.Controls("TabStrip2")
End Property

Public Property Get ProduzLogo() As Object
     Set ProduzLogo = objUserControl.Controls("ProduzLogo")
End Property

Public Property Get BotaoImprimirRotulos() As Object
     Set BotaoImprimirRotulos = objUserControl.Controls("BotaoImprimirRotulos")
End Property

Public Property Get BotaoImprimirPrevia() As Object
     Set BotaoImprimirPrevia = objUserControl.Controls("BotaoImprimirPrevia")
End Property

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Public Property Get BotaoEstoque() As Object
     Set BotaoEstoque = objUserControl.Controls("BotaoEstoque")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get BotaoCcls() As Object
     Set BotaoCcls = objUserControl.Controls("BotaoCcls")
End Property

Public Property Get ComboFilialPedido() As Object
     Set ComboFilialPedido = objUserControl.Controls("ComboFilialPedido")
End Property

Public Property Get Situacao() As Object
     Set Situacao = objUserControl.Controls("Situacao")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get BotaoPedidoDeVenda() As Object
     Set BotaoPedidoDeVenda = objUserControl.Controls("BotaoPedidoDeVenda")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get FilialCliente() As Object
     Set FilialCliente = objUserControl.Controls("FilialCliente")
End Property

Public Property Get Benef() As Object
     Set Benef = objUserControl.Controls("Benef")
End Property

Public Property Get Destinacao() As Object
     Set Destinacao = objUserControl.Controls("Destinacao")
End Property

Public Property Get Versao() As Object
     Set Versao = objUserControl.Controls("Versao")
End Property

Public Property Get BotaoMaquinas() As Object
     Set BotaoMaquinas = objUserControl.Controls("BotaoMaquinas")
End Property

Public Property Get Maquina() As Object
     Set Maquina = objUserControl.Controls("Maquina")
End Property

Public Property Get PedidoDeVendaId() As Object
     Set PedidoDeVendaId = objUserControl.Controls("PedidoDeVendaId")
End Property

Public Property Get Prioridade() As Object
     Set Prioridade = objUserControl.Controls("Prioridade")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get Almoxarifado() As Object
     Set Almoxarifado = objUserControl.Controls("Almoxarifado")
End Property

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get DataPrevisaoFim() As Object
     Set DataPrevisaoFim = objUserControl.Controls("DataPrevisaoFim")
End Property

Public Property Get DataPrevisaoInicio() As Object
     Set DataPrevisaoInicio = objUserControl.Controls("DataPrevisaoInicio")
End Property

Public Property Get GridMovimentos() As Object
     Set GridMovimentos = objUserControl.Controls("GridMovimentos")
End Property

Public Property Get QuantDisponivel() As Object
     Set QuantDisponivel = objUserControl.Controls("QuantDisponivel")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoImprimir() As Object
     Set BotaoImprimir = objUserControl.Controls("BotaoImprimir")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

'--- fim dos properties get dos controles da tela

Function OP_Inicializa_GridMovimentos(ByVal objCT As Object) As Long

Dim iIndice As Integer

    Set objGrid = New AdmGrid

    'tela em questão
    Set objGrid.objForm = objCT

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("U.M.")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Almoxarifado")
    objGrid.colColuna.Add ("Escaninho")
    objGrid.colColuna.Add ("Ccl")
    objGrid.colColuna.Add ("Previsão Início")
    objGrid.colColuna.Add ("Previsão Fim")
    objGrid.colColuna.Add ("Situação")
    objGrid.colColuna.Add ("Destinação")
    objGrid.colColuna.Add ("Pedido de Venda")
    objGrid.colColuna.Add ("Filial do Pedido")
    objGrid.colColuna.Add ("Cliente")
    objGrid.colColuna.Add ("Filial do Cliente")
    objGrid.colColuna.Add ("Prioridade")
    objGrid.colColuna.Add ("Maquina")
    objGrid.colColuna.Add ("Produz Logo")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Produto.Name)
    objGrid.colCampo.Add (Versao.Name)
    objGrid.colCampo.Add (DescricaoItem.Name)
    objGrid.colCampo.Add (UnidadeMed.Name)
    objGrid.colCampo.Add (Quantidade.Name)
    objGrid.colCampo.Add (Almoxarifado.Name)
    objGrid.colCampo.Add (Benef.Name)
    objGrid.colCampo.Add (Ccl.Name)
    objGrid.colCampo.Add (DataPrevisaoInicio.Name)
    objGrid.colCampo.Add (DataPrevisaoFim.Name)
    objGrid.colCampo.Add (Situacao.Name)
    objGrid.colCampo.Add (Destinacao.Name)
    objGrid.colCampo.Add (PedidoDeVendaId.Name)
    objGrid.colCampo.Add (ComboFilialPedido.Name)
    objGrid.colCampo.Add (Cliente.Name)
    objGrid.colCampo.Add (FilialCliente.Name)
    objGrid.colCampo.Add (Prioridade.Name)
    objGrid.colCampo.Add (Maquina.Name)
    objGrid.colCampo.Add (ProduzLogo.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_Versao_Col = 2
    iGrid_DescricaoItem_Col = 3
    iGrid_UnidadeMed_Col = 4
    iGrid_Quantidade_Col = 5
    iGrid_Almoxarifado_Col = 6
    iGrid_Benef_Col = 7
    iGrid_Ccl_Col = 8
    iGrid_DataPrevInicio_Col = 9
    iGrid_DataPrevFim_Col = 10
    iGrid_Situacao_Col = 11
    iGrid_Destinacao_Col = 12
    iGrid_PedidoDeVenda_Col = 13
    iGrid_FilialPedido_Col = 14
    iGrid_Cliente_Col = 15
    iGrid_FilialCliente_Col = 16
    iGrid_Prioridade_Col = 17
    iGrid_Maquina_Col = 18
    iGrid_ProduzLogo_Col = 19

    objGrid.objGrid = GridMovimentos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    If m_objUserControl.Height > 9000 Then
        objGrid.iLinhasVisiveis = 20
    Else
        objGrid.iLinhasVisiveis = 9
    End If

    'Largura da primeira coluna
    GridMovimentos.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    objGrid.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    Call Grid_Inicializa(objGrid)

    OP_Inicializa_GridMovimentos = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iNum As Integer
'Inserido por Wagner
Dim sVersaoAnt As String
Dim objKit As New ClassKit
Dim colKits As New Collection
Dim sNomeMaqAnt As String
Dim lBenef As Long

On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada <> ROTINA_GRID_ABANDONA_CELULA Then

        'Verifica se produto está preenchido
        sCodProduto = GridMovimentos.TextMatrix(iLinha, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 21922

        iNum = 0
        
        If gcolItemOP.Count >= GridMovimentos.Row Then
            If gcolItemOP.Count > 0 And GridMovimentos.Row <> 0 Then
                iNum = gcolItemOP.Item(GridMovimentos.Row)
            End If
        End If

        'Pesquisa o controle da coluna em questão
        Select Case objControl.Name
    
            'Produto
            Case Produto.Name
            
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    Produto.Enabled = False
                Else
                    Produto.Enabled = True
                End If

            Case UnidadeMed.Name, DescricaoItem.Name, Cliente.Name, FilialCliente.Name
                'ficam sempre desabilitadas
    
            Case Versao.Name
                
                If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)) > 0 Then
                    'Impede a atualização dos demais dados da OP
                    If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 And left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                Else
                        objControl.Enabled = False
                End If
                
                '#############################################
                'Inserido por Wagner
                sVersaoAnt = Versao.Text
                
                Versao.Clear
    
                If iProdutoPreenchido <> PRODUTO_VAZIO Then
                    
                    'Armazena o Produto Raiz do kit
                    objKit.sProdutoRaiz = sProdutoFormatado
                    
                    'Le as Versoes Ativas e a Padrao
                    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
                    If lErro <> SUCESSO And lErro <> 106333 Then gError 131215
                    
                    'Carrega a Combo com os Dados da Colecao
                    For Each objKit In colKits
                    
                        Versao.AddItem (objKit.sVersao)
                                                    
                    Next
                    
                    'Tento selecionar na Combo a Unidade anterior
                    If Versao.ListCount <> 0 Then
        
                        For iIndice = 0 To Versao.ListCount - 1
        
                            If Versao.List(iIndice) = sVersaoAnt Then
                                Versao.ListIndex = iIndice
                                Exit For
                            End If
                        Next
                    End If
                End If
                '#############################################
            
            Case Situacao.Name, Maquina.Name
    
                'habilita atualização da situacao da OP
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case ComboFilialPedido.Name, PedidoDeVendaId.Name
                
                If Destinacao.ListIndex <> -1 Then
                    If Destinacao.ItemData(Destinacao.ListIndex) = ITEMOP_DESTINACAO_PV Then
                        'Impede a atualização dos demais dados da OP
                        If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 Then
                            
                            If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = "VÁRIOS" Then
                                objControl.Enabled = False
                            Else
                                objControl.Enabled = True
                            End If
                        Else
                            objControl.Enabled = False
                        End If
                    Else
                        objControl.Enabled = False
                    End If
                Else
                    objControl.Enabled = False
                   
                End If
            
            Case Destinacao.Name
            
                    Call Combo_Obtem_ItemData(Benef, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col), lBenef)
    
                    If lBenef <> BENEF_COMBO_DISP Then
                        objControl.Enabled = False
                    Else
                        If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 Then
                            objControl.Enabled = True
                        Else
                            objControl.Enabled = False
                        End If
                   
                   End If
    
            Case Almoxarifado.Name
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 And left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case Quantidade.Name
            
                'Alterado por Wagner
                'Se já fizer parte do plano mestre não pode alterar a quantidade
                If MRP.Value = vbChecked Then
                    objControl.Enabled = False
                Else
                    If iProdutoPreenchido = PRODUTO_PREENCHIDO And left(GridMovimentos.TextMatrix(GridMovimentos.Row, 0), 1) <> "#" Then
                        'Não tem código para redistribuir para vários PVs
                        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = "VÁRIOS" Then
                            objControl.Enabled = False
                        Else
                            objControl.Enabled = True
                        End If
                    Else
                        objControl.Enabled = False
                    End If
                End If
                
            '##########################################
            'Alterado por Wagner
            'se são campos do grid de insumos
             Case ProdutoInsumos.Name, DescricaoProduto.Name, OrigemProduto.Name, QuantidadeProduto.Name, _
                    UMProduto.Name, VersaoKitComp.Name, PerdaInsumos.Name, ComposicaoInsumos.Name, CustoStandardInsumos.Name

                    objControl.Enabled = False
                    
            'se são campos do grid de Maquinas
             Case NomeMaquina.Name, QuantidadeMaquina.Name, HorasPO.Name, _
                    DataPO.Name, TaxaProducao.Name
                    
                    objControl.Enabled = False
                    
            'se são campos do Grid de uso de máquinas real
            Case HorasReal.Name, QuantidadeMaquinaReal.Name
            
                If StrParaDate(GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_DataReal_Col)) = DATA_NULA Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
                
            Case NomeMaquinaReal.Name
                    
                'Se ainda não tiver data ou já tiver preechido a máquina então não deixa alterar
                If StrParaDate(GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_DataReal_Col)) = DATA_NULA Or Len(Trim(GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_NomeMaquinaReal_Col))) > 0 Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
                    
            'se são campos do Grid de uso de mão de obra real
            Case HorasRealMO.Name, QuantidadeMaquinaRealMO.Name, _
                    TipoMaoDeObraReal.Name
            
                If StrParaDate(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DataRealMO_Col)) = DATA_NULA Or Len(Trim(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_NomeMaquinaRealMO_Col))) = 0 Then
                    objControl.Enabled = False
                Else
                    objControl.Enabled = True
                End If
                
            Case DataReal.Name
                
                If MRP.Value = vbChecked Then
                    If StrParaDate(GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_DataReal_Col)) = DATA_NULA Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                Else
                    objControl.Enabled = False
                End If
                
            Case DataRealMO.Name
                
                If MRP.Value = vbChecked Then
                    If StrParaDate(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DataRealMO_Col)) = DATA_NULA Then
                        objControl.Enabled = True
                    Else
                        objControl.Enabled = False
                    End If
                Else
                    objControl.Enabled = False
                End If
                
            Case DescricaoTipoMO.Name
            
                objControl.Enabled = False
                
            Case NomeMaquinaRealMO.Name
            
                If MRP.Value = vbChecked Then
                    
                    sNomeMaqAnt = NomeMaquinaRealMO.Text
                    
                    NomeMaquinaRealMO.Clear
                    For iLinha = 1 To objGridMaquinasReal.iLinhasExistentes
                        If StrParaDate(GridMaquinasReal.TextMatrix(iLinha, iGrid_DataReal_Col)) = StrParaDate(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DataRealMO_Col)) Then
                            NomeMaquinaRealMO.AddItem GridMaquinasReal.TextMatrix(iLinha, iGrid_NomeMaquinaReal_Col)
                        End If
                    Next
                
                    'Tento selecionar na Combo a Máquina anterior
                    If NomeMaquinaRealMO.ListCount <> 0 Then
        
                        For iIndice = 0 To NomeMaquinaRealMO.ListCount - 1
        
                            If NomeMaquinaRealMO.List(iIndice) = sNomeMaqAnt Then
                                NomeMaquinaRealMO.ListIndex = iIndice
                                Exit For
                            End If
                        Next
                    End If
                    
                    If StrParaDate(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DataRealMO_Col)) = DATA_NULA Then
                        objControl.Enabled = False
                    Else
                        objControl.Enabled = True
                    End If
                Else
                    objControl.Enabled = False
                End If
                
            '##########################################
    
            Case Benef.Name, Ccl.Name, DataPrevisaoInicio.Name, DataPrevisaoFim.Name, Prioridade.Name, ProduzLogo.Name
                'Impede a atualização dos demais dados da OP
                If iProdutoPreenchido = PRODUTO_PREENCHIDO And iNum = 0 Then
                    objControl.Enabled = True
                Else
                    objControl.Enabled = False
                End If
    
            Case Else
                
                Call CF2(Me, "OP_Rotina_Grid_Enable", iLinha, objControl, iLocalChamada)
    
        End Select

    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 21922

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163800)

    End Select

    Exit Sub

End Sub

Public Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub AlmoxPadraoLabel_Click()

Dim colSelecao As New Collection
Dim objAlmoxarifado As New ClassAlmoxarifado

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxPadrao)

End Sub


Public Sub BotaoAbrirOP_Click()

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao
   
On Error GoTo Erro_BotaoVerEtapas_Click
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 139048
    
    objOrdemDeProducao.sCodigo = OPCodigoMRP
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa

    'traz OP para a tela
    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21966 Then gError 139049

    'Torna Frame atual invisível
    Frame1(TabStrip1.SelectedItem.Index).Visible = False
    iFrameAtual = 1
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    TabStrip1.Tabs.Item(iFrameAtual).Selected = True
    
    'Torna Frame Operacoes atual invisível
    Frame2(TabStrip2.SelectedItem.Index).Visible = False
    iFrameAtualOper = 1
    'Torna Frame Operacoes atual visível
    Frame2(iFrameAtualOper).Visible = True
    TabStrip2.Tabs.Item(iFrameAtualOper).Selected = True

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoVerEtapas_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 139048, 139049

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163801)

    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoCompetencia_Click()

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
   
On Error GoTo Erro_BotaoVerEtapas_Click
    
    GL_objMDIForm.MousePointer = vbHourglass

    objCompetencias.sNomeReduzido = LabelCodigoCompetencia.Caption
   
    'Lê a Competencia pelo NomeReduzido
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 139060
   
    'Chama a tela de ordem de produção
    Call Chama_Tela("Competencias", objCompetencias)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoVerEtapas_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 139060

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163802)

    End Select
    
    Exit Sub


End Sub

Public Sub BotaoCT_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
   
On Error GoTo Erro_BotaoVerEtapas_Click
    
    GL_objMDIForm.MousePointer = vbHourglass

    objCentrodeTrabalho.sNomeReduzido = LabelCodigoCTPadrao.Caption
   
    'Chama a tela de ordem de produção
    Call Chama_Tela("CentrodeTrabalho", objCentrodeTrabalho)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoVerEtapas_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163803)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoGrade_Click()

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemOP As ClassItemOP

On Error GoTo Erro_BotaoGrade_Click

    If GridMovimentos.Row > 0 And GridMovimentos.Row <= objGrid.iLinhasExistentes Then
    
        Set objItemOP = gobjOP.colItens(GridMovimentos.Row)
        
        If objItemOP.iPossuiGrade = DESMARCADO Then gError 126517
            
        objItemOP.sAlmoxarifadoNomeRed = AlmoxPadrao.Text
            
        Set objRomaneioGrade = New ClassRomaneioGrade
        
        objRomaneioGrade.sNomeTela = Me.Name
        Set objRomaneioGrade.objObjetoTela = objItemOP
        Set objRomaneioGrade.objTela = Me
                    
        Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
    
        Call Atualiza_Grid_Movimentos(objItemOP)
            
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
      
        Case 126517
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_NAO_GRADE", gErr, GridMovimentos.Row)
      
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163804)
            
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoGrafico_Click()

Dim lErro As Long
Dim objTelaGrafico As New ClassTelaGrafico

On Error GoTo Erro_BotaoGrafico_Click:

    lErro = Atualiza_Cronograma(objTelaGrafico)
    If lErro <> SUCESSO Then gError 139061
    
    Call Chama_Tela_Nova_Instancia("TelaGrafico", objTelaGrafico)

    Exit Sub

Erro_BotaoGrafico_Click:

    Select Case gErr
    
        Case 139061
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163805)

    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoMaquinaPMP_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
   
On Error GoTo Erro_BotaoVerEtapas_Click
    
    GL_objMDIForm.MousePointer = vbHourglass

    'Se não tiver linha selecionada => Erro
    If GridMaquinas.Row = 0 Then gError 139058
    
    'Verifica se a linha selecionada está preenchida
    If Len(GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_NomeMaquina_Col)) = 0 Then gError 139059
   
    objMaquinas.sNomeReduzido = GridMaquinas.TextMatrix(GridMaquinas.Row, iGrid_NomeMaquina_Col)
   
    'Chama a tela de ordem de produção
    Call Chama_Tela("Maquinas", objMaquinas)

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoVerEtapas_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 139058
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 139059
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163806)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoMaquinas_Click()
'Chama o Browser de Maquinas...

Dim lErro As Long
Dim objMaquina As New ClassMaquinas
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoMaquinas_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 106320

    'Se o equipamento foi preenchido => armazena no obj
    If Len(Trim(Maquina.Text)) > 0 Then
    
        If IsNumeric(Maquina.Text) Then
    
            objMaquina.iCodigo = StrParaInt(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col))
        
        Else
            
            objMaquina.sNomeReduzido = CStr(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col))
        
        End If
    
    End If

    'Lista de Equipamentos
    Call Chama_Tela("MaquinasLista", colSelecao, objMaquina, objEventoMaquina)

    Exit Sub

Erro_BotaoMaquinas_Click:

    Select Case gErr

        Case 106320
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 106433
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NOME_INEXISTENTE", gErr, Maquina.Text)
            
        Case 106434
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_CODIGO_INEXISTENTE", gErr, Maquina.Text, giFilialEmpresa)

        Case 55325, 106432

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163807)

    End Select

End Sub

Public Sub BotaoImprimir_Click()

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_BotaoImprimir_Click

    'Verifica se os campos obrigatórios foram preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 57722
    If Len(Trim(Data.ClipText)) = 0 Then gError 57723
    
    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.dtDataEmissao = StrParaDate(Data.Text)
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    
    'lErro = CF("OrdemDeProducao_TestaExistencia", objOrdemDeProducao)
    'If lErro <> SUCESSO And lErro <> 57721 Then gError 57724
    'If lErro = 57721 Then gError 57725
    
    'pesquisa a op nao baixada, e preenche seus itens
    lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 57724
    
    If lErro = 21960 Then
    
        'se nao achou antes, tenta ver se ja estava baixada
        lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 82797 Then gError 111823
    
        'se nao achou => erro de inexistencia de op
        If lErro <> SUCESSO Then gError 57725
    
    End If
    
    'Executa o(s) Relatorio(s) de acordo com a selecao
    lErro = CF2(Me, "OP_Executa_Relatorio", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 106409
    
    Exit Sub
    
Erro_BotaoImprimir_Click:

    Select Case gErr
    
        Case 57722
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr)
        
        Case 57723
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
    
        Case 57724, 106409, 111823
        
        Case 57725
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", gErr, objOrdemDeProducao.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163808)

    End Select
    
End Sub


Public Sub BotaoPedidoDeVenda_Click()

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim colSelecao As New Collection
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoPedidoDeVenda_Click

    If GridMovimentos.Row = 0 Then gError 52016

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52017

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        'verifica se a destinação é Pedido de Venda - - Se não for - - > Erro
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) <> STRING_PEDIDOVENDA Then gError 52018
        
        'Selecao
        colSelecao.Add sProdutoFormatado
        
        'chama a tela de lista de estoque do produto corrente
        Call Chama_Tela("ItemPVProdutoLista", colSelecao, objItemPedido, objEventoPedidoDeVenda)
    Else
        Error 52019
    End If

    Exit Sub

Erro_BotaoPedidoDeVenda_Click:

    Select Case gErr
        
        Case 52016
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 52017
                
        Case 52018
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINACAO_DEPENDENTE", gErr)
        
        Case 52019
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163809)

    End Select

    Exit Sub
End Sub

Public Sub Benef_Click()
    
Dim lErro As Long
Dim lBenef As Long

On Error GoTo Erro_Benef_Click

    iAlterado = REGISTRO_ALTERADO
    
    Call Combo_Obtem_ItemData(Benef, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col), lBenef)

    If lBenef <> BENEF_COMBO_DISP And GridMovimentos.Row <> 0 Then
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = "Estoque"
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) = ""
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = ""
    End If

    lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
    If lErro <> SUCESSO Then gError 91274

    Exit Sub
    
Erro_Benef_Click:

    Select Case gErr
    
        Case 91274
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163810)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Benef_GotFocus()
'trata o evento gotfocus associado ao campo Benef

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Benef_KeyPress(KeyAscii As Integer)
'trata o evento keypress associado ao campo Benef

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Benef_Validate(Cancel As Boolean)
'trata o evento validate associado ao campo Benef

Dim lErro As Long

    Set objGrid.objControle = Benef
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Ccl_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Ccl_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Ccl_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CclPadraoLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCclPadrao)

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO
    iCodigoAlterado = REGISTRO_ALTERADO

End Sub



Public Sub CodOPGerada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ComboFilialPedido_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub ComboFilialPedido_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Public Sub DataFimPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFimPadrao, iAlterado)

End Sub

Public Sub DataInicioPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicioPadrao, iAlterado)

End Sub

Public Sub DataPrevisaoFim_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub DataPrevisaoFim_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub DataPrevisaoFim_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub DataPrevisaoFim_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DataPrevisaoFim
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataPrevisaoInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataPrevisaoInicio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub DataPrevisaoInicio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub DataPrevisaoInicio_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = DataPrevisaoInicio
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Destinacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Destinacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Destinacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Destinacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Destinacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GeraOP_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If GeraOP.Value = vbChecked Then
        
        LabelCodOPGerada.Enabled = True
        CodOPGerada.Enabled = True
        botaoProxNum2.Enabled = True 'Alterado por Wagner
        '##############
        'INSERIDO POR WAGNER
        GeraOPs.Enabled = True
        '##############
    Else
        LabelCodOPGerada.Enabled = False
        CodOPGerada.Enabled = False
        CodOPGerada.Text = ""
        botaoProxNum2.Enabled = False 'Alterado por Wagner
        '##############
        'INSERIDO POR WAGNER
        GeraOPs.Enabled = False
        GeraOPs.Value = vbUnchecked
        '##############
    End If
    

End Sub

Public Sub GeraReqCompra_Click()

    iAlterado = REGISTRO_ALTERADO
    
    If gobjMAT.iOPDetalhamentoRCs = MARCADO Then
        If GeraReqCompra.Value = MARCADO Then
            BotaoOPRC.Enabled = True
        Else
            BotaoOPRC.Enabled = False
        End If
    End If

End Sub

Public Sub GridMovimentos_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer
Dim iNum As Integer
Dim lErro As Long
Dim iLinhasExistentes As Integer 'm

'Inserido por Jorge Specian - 14/05/2006
'---------------------------------------
Dim sProdutoRaiz As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objItOPcol As New ClassItemOP
'---------------------------------------

On Error GoTo Erro_GridMovimentos_KeyDown

    'Define se OP é nova ou existente
    If gcolItemOP.Count >= GridMovimentos.Row Then
        iNum = gcolItemOP.Item(GridMovimentos.Row)
    Else
        iNum = 0
    End If

    'Se for uma nova OP
    If iNum = 0 Then

        'Guarda iLinhasExistentes
        iLinhasExistentesAnterior = objGrid.iLinhasExistentes

        'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then

            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridMovimentos.Row
            
            'Inserido por Jorge Specian - 14/05/2006
            '---------------------------------------
            sProdutoRaiz = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
            '---------------------------------------
            
        End If

        Call Grid_Trata_Tecla1(KeyCode, objGrid)

        'Verifica se a Linha foi realmente excluída
        If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then

            gcolItemOP.Remove (iLinhaAnterior)
            gobjOP.colItens.Remove iLinhaAnterior  'm
            
           Call gobjTelaProjetoInfo.Remove_Linha(iLinhaAnterior)

            Call CF2(Me, "OP_GridMovimentos_KeyDown", iLinhaAnterior)
            
            'Inserido por Jorge Specian - 14/05/2006
            '---------------------------------------
            lErro = CF("Produto_Formata", sProdutoRaiz, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 137184
                        
            'ajusta a coleção
            For Each objItOPcol In colOPItens
                If objItOPcol.sProduto = sProdutoFormatado Then
                    colOPItens.Remove (sProdutoFormatado)
                    Exit For
                End If
            Next
            '---------------------------------------

            For iLinhasExistentes = 1 To objGrid.iLinhasExistentes 'm
                If gobjOP.colItens(iLinhasExistentes).iPossuiGrade = MARCADO Then
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = "# " & iLinhasExistentes
                Else
                    GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes
                End If
                
            Next

            GridMovimentos.TextMatrix(iLinhasExistentes, 0) = iLinhasExistentes

            lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
            If lErro <> SUCESSO Then gError 55328

        End If

    End If

    Exit Sub
    
Erro_GridMovimentos_KeyDown:

    Select Case gErr
    
        Case 55328, 137184 '(137184 - Inserido por Jorge Specian - 14/06/2005)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163811)
    
    End Select
    
    Exit Sub

End Sub

Public Sub ItemOP_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_ItemOP_Validate

    'Verifica se ItemOP está preenchida
    If Len(Trim(ItemOP.ClipText)) <> 0 Then

        'Critica o ItemOP
        lErro = Inteiro_Critica(ItemOP.Text)
        If lErro <> SUCESSO Then gError 137916
        
        iItem = StrParaInt(ItemOP.Text)
        
        If giItemOP <> iItem Then
        
            'Se o valor estiver fora do range do grid... Erro
            If iItem < 1 Or iItem > objGrid.iLinhasExistentes Then gError 137917
            
            GridMovimentos.Row = iItem
            
            'Então, mostra a nova arvore
            lErro = Mostra_Arvore(GridMovimentos.TextMatrix(iItem, iGrid_Produto_Col), StrParaDbl(GridMovimentos.TextMatrix(iItem, iGrid_Quantidade_Col)))
            If lErro <> SUCESSO Then gError 137918
            
            giItemOP = iItem
            
        End If

    End If

    Exit Sub

Erro_ItemOP_Validate:

    Cancel = True

    Select Case gErr

        Case 137916, 137918
            'erros tratados nas rotinas chamadas
            
        Case 137917
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_INEXISTENTE", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163812)

    End Select

    Exit Sub

End Sub

Public Sub LabelPrestador_Click()

Dim objPrestServ As New ClassPrestServ
Dim colSelecao As Collection

    'Preenche nomeReduzido com o fornecedor da tela
    objPrestServ.sNomeReduzido = PrestadorServico.Text

    Call Chama_Tela("PrestServLista", colSelecao, objPrestServ, objEventoPrestServ)

End Sub

Public Sub Maquina_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub objEventoCclPadrao_evSelecao(obj1 As Object)
'Preenche CclPadrao

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCclPadrao_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 22930

    CclPadrao.PromptInclude = False
    CclPadrao.Text = sCclMascarado
    CclPadrao.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCclPadrao_evSelecao:

    Select Case gErr

        Case 22930
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163813)

    End Select

    Exit Sub

End Sub

Private Sub objEventoAlmoxPadrao_evSelecao(obj1 As Object)

Dim objAlmoxarifado As ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    'Preenche AlmoxPadrao
    AlmoxPadrao.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

End Sub

Public Sub BotaoCcls_Click()
'chama tela de Lista de Ccl

Dim lErro As Long
Dim objCcls As New ClassCcl
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCcls_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 43742

    'Verifica se o Produto está preenchido
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then gError 43743

    Call Chama_Tela("CclLista", colSelecao, objCcls, objEventoCcl)
    
    Exit Sub
    
Erro_BotaoCcls_Click:

    Select Case gErr
    
        Case 43742
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 43743
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163814)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoEstoque_Click()
'Informa se produto é estocado em algum almoxarifado

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridMovimentos.Row = 0 Then gError 43719

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 21930

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado
        'chama a tela de lista de estoque do produto corrente
        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 43739
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 21930
        
        Case 43719
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43739
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163815)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    If GridMovimentos.Row <> 0 Then

        Set objEstoqueProduto = obj1

        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22941

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            'Calcula a Quantidade Disponível nesse Almoxarifado
            lErro = QuantDisponivel_Calcula(sCodProduto, objEstoqueProduto.sAlmoxarifadoNomeReduzido)
            If lErro <> SUCESSO Then gError 41304

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 22941, 41304

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163816)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()
'ativa a exclusão de uma OP

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'verifica se código está preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 21936

    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa

    'Pede ao usuário que confire a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_OP", objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'exclui a OP
    lErro = CF("OrdemDeProducao_Exclui", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21936 Then gError 21940

    'se OP não existir -> erro
    If lErro = 21936 Then gError 21950

    'Limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 21951

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 21936
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 21940, 21951

        Case 21950
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_INEXISTENTE", gErr, objOrdemDeProducao.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163817)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()
'implementa gravação de uma nova ou atualizacao de uma OP

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Rotina de gravação da OP
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 21750

    'limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 21951
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 21750, 21951

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163818)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 21952

    'limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 21953

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 21952, 21953

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163819)

    End Select

    Exit Sub

End Sub

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iPreenchido As Integer
Dim colSelecao As Collection

On Error GoTo Erro_BotaoProdutos_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridMovimentos.Row = 0 Then gError 43718

    'Verifica se o Produto está preenchido
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) > 0 Then
    
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 55325
        
        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""
        
    End If

    objProduto.sCodigo = sProduto

    'Lista de produtos produzíveis
    Call Chama_Tela("ProdutoProduzivelLista", colSelecao, objProduto, objEventoProduto)
        
    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 43718
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 55325
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163820)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoMaquina_evSelecao(obj1 As Object)

Dim objMaquinas As ClassMaquinas
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_objEventoMaquina_evSelecao

    Set objMaquinas = obj1
    
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 106426

    'verifica se o produto esta preenchido...
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        Maquina.Text = objMaquinas.sNomeReduzido
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Maquina_Col) = objMaquinas.sNomeReduzido
    
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoMaquina_evSelecao:

    Select Case gErr
        
        Case 106426
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163821)
    
    End Select

End Sub

Private Sub objEventoPedidoDeVenda_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objItemPedido As ClassItemPedido
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim iIndice As Integer

On Error GoTo Erro_objEventoPedidoDeVenda_evSelecao

    If GridMovimentos.Row <> 0 Then

        Set objItemPedido = obj1
        
        sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 52020

        'verifica se o produto esta preenchido e se a destinacao é pedido de Venda
        If iProdutoPreenchido = PRODUTO_PREENCHIDO And GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = STRING_PEDIDOVENDA Then
                               
            objPedidoDeVenda.lCodigo = objItemPedido.lCodPedido
            objPedidoDeVenda.iFilialEmpresa = objItemPedido.iFilialEmpresa
        
            lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
            If lErro <> SUCESSO And lErro <> 26509 Then gError 52051
        
            If lErro = 26509 Then gError 52052
            
            For iIndice = 0 To ComboFilialPedido.ListCount - 1
                If objItemPedido.iFilialEmpresa = ComboFilialPedido.ItemData(iIndice) Then
                    ComboFilialPedido.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = CStr(objItemPedido.lCodPedido)
            
            PedidoDeVendaId.Text = CStr(objItemPedido.lCodPedido)
            
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) = ComboFilialPedido.Text
            
            objCliente.lCodigo = objPedidoDeVenda.lCliente
            
            'le o nome reduzido do cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 12293 Then gError 52021
        
            If lErro <> SUCESSO Then gError 52022
                    
            'preenche com o nome reduzido do cliente
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                    
            objFilialCliente.lCodCliente = objCliente.lCodigo
            objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
        
            'le o nome da filial do cliente
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then gError 52023
        
            If lErro = 12567 Then gError 52024
        
            'preenche CODIGO - NOME
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
        
        End If
        
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoPedidoDeVenda_evSelecao:
    
    Select Case gErr
    
        Case 52020, 52021, 52023, 52051
        
        Case 52022
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
    
        Case 52024
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
        
        Case 52052
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163822)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPrestServ_evSelecao(obj1 As Object)

Dim objPrestServ As New ClassPrestServ

    If Not (obj1 Is Nothing) Then
    
        Set objPrestServ = obj1

        PrestadorServico.Text = objPrestServ.sNomeReduzido
        Call PrestadorServico_Validate(bSGECancelDummy)
        
        Me.Show
        
    End If
    
End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim objProduto As New ClassProduto
Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim objItemOP As New ClassItemOP
Dim bInseriuLinha As Boolean

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If GridMovimentos.Row <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22935

        'Se o produto não estiver preenchido
        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 126585

            If lErro = 28030 Then gError 126586

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            'mascara produto escolhido
            lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 22936

            'verifica sua existência na OP
            lErro = VerificaUso_Produto(objProduto)
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55327

            If lErro = 41316 Then gError 22958

            Call Carrega_ComboVersoes(objProduto.sCodigo)
            
            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 22937

            If lErro = 28030 Then gError 22939

            If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 55276

            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True
            
            If Not (Me.ActiveControl Is Produto) Then

                'preenche produto
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = sProdutoMascarado
    
                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
                If lErro <> SUCESSO Then gError 22938
    
                'calcula a qtd disponível
                lErro = QuantDisponivel_Calcula(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
                If lErro <> SUCESSO Then gError 41305

            End If
        
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 22935, 22937, 22938, 41305, 55327, 126585, 131084

        Case 22936
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
            
        Case 22939
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 22958
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, sProdutoMascarado, Codigo.Text)
            
        Case 55276
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, sProdutoMascarado)

        Case 126586
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163823)

    End Select

    Exit Sub

End Sub

Public Sub CodigoOPLabel_Click()

Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection
Dim sSelecao As String

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then objOrdemDeProducao.sCodigo = Codigo.Text
    
    sSelecao = "Tipo = 0"
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemDeProducao, objEventoCodigo, sSelecao)

End Sub

Public Sub ComboFilialPedido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboFilialPedido_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ComboFilialPedido
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function VerificaQuantidade_ItemPedido(objItemPedido As ClassItemPedido) As Long

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_VerificaQuantidade_ItemPedido

        lErro = CF("ItemPedido_Le", objItemPedido)
        If lErro <> SUCESSO And lErro <> 23971 Then gError 41308

        If lErro = 23971 Then gError 41309

        'avisa que qtd ordenada é diferente da qtd do item de pedido
        If Len(Quantidade.Text) > 0 Then
        
            If Trim(Quantidade.Text) <> CStr(objItemPedido.dQuantidade) Then
            
                vbMsg = Rotina_Aviso(vbOKOnly, "AVISO_QUANTIDADE_ITEMPEDIDO")
                
            End If
            
        End If

    VerificaQuantidade_ItemPedido = SUCESSO

    Exit Function

Erro_VerificaQuantidade_ItemPedido:

    VerificaQuantidade_ItemPedido = gErr

    Select Case gErr

        Case 41308

        Case 41309
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMPEDIDO_INEXISTENTE1", gErr, objItemPedido.lCodPedido, objItemPedido.iFilialEmpresa, objItemPedido.sProduto)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163824)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

    'Carrega índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub CargaCombo_Situacao(objSituacao As Object)
'Carga dos itens da combo Situação

    objSituacao.AddItem STRING_NORMAL
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_NORMAL
    objSituacao.AddItem STRING_DESABILITADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_DESAB
    objSituacao.AddItem STRING_SACRAMENTADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_SACR
    objSituacao.AddItem STRING_BAIXADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_BAIXADA

End Sub

Public Sub CargaCombo_Destinacao(objDestinacao As Object)
'Carga dos itens da combo Destinação

    objDestinacao.AddItem STRING_ESTOQUE
    objDestinacao.ItemData(objDestinacao.NewIndex) = ITEMOP_DESTINACAO_ESTOQUE
    objDestinacao.AddItem STRING_PEDIDOVENDA
    objDestinacao.ItemData(objDestinacao.NewIndex) = ITEMOP_DESTINACAO_PV
    objDestinacao.AddItem STRING_CONSUMO
    objDestinacao.ItemData(objDestinacao.NewIndex) = ITEMOP_DESTINACAO_CONSUMO

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim sItem As String
Dim sMascaraCclPadrao As String
Dim objFiliais As AdmFiliais
Dim colModulo As New AdmColModulo
Dim tModulo As typeModulo
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set gcolItemOP = New Collection

    Set objEventoCodigo = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoCclPadrao = New AdmEvento
    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoPedidoDeVenda = New AdmEvento
    Set objEventoMaquina = New AdmEvento
    Set objEventoPrestServ = New AdmEvento
    
    Set objEventoCliente = New AdmEvento
    Set objEventoFornecedor = New AdmEvento

    '---------------------------------------
    'Inserido por Jorge Specian - 29/04/2005
    Set objEventoCompetencias = New AdmEvento
    Set objEventoCentroDeTrabalho = New AdmEvento
    
    Set colPMPItens = New Collection
    Set colOPItens = New Collection
    Set colComponentes = New Collection
    '---------------------------------------
    
    '#######################################
    'Inserido por Wagner 07/11/2005
    Set objEventoMaquinas = New AdmEvento
    Set objEventoMO = New AdmEvento
    '#######################################
    
    For Each objFiliais In gcolFiliais
        If objFiliais.iCodFilial <> 0 Then
            sItem = CStr(objFiliais.iCodFilial) & SEPARADOR & CStr(objFiliais.sNome)
            ComboFilialPedido.AddItem sItem
            ComboFilialPedido.ItemData(ComboFilialPedido.NewIndex) = objFiliais.iCodFilial
        End If
    Next
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me

    'Carrega Ítens das Combos
    Call CargaCombo_Situacao(Situacao)
    Call CargaCombo_Destinacao(DestinacaoPadrao)
    Call CargaCombo_Destinacao(Destinacao)

    'Inicializa máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 22963

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInsumos)
    If lErro <> SUCESSO Then gError 22963

    'Inicializa Máscara de CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError 22964

    CclPadrao.Mask = sMascaraCclPadrao
    Ccl.Mask = sMascaraCclPadrao

    '---------------------------------------
    'Inserido por Jorge Specian - 29/04/2005
    
    iFrameAtualOper = 1
    
    'Grid de OperacaoInsumos
    Set objGridOperacaoInsumos = New AdmGrid
    
    'tela em questão
    Set objGridOperacaoInsumos.objForm = Me
    
    lErro = Inicializa_GridOperacaoInsumos(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then gError 137080
    
    'Grid de Maquinas
    Set objGridMaquinas = New AdmGrid
    
    'tela em questão
    Set objGridMaquinas.objForm = Me
    
    lErro = Inicializa_GridMaquinas(objGridMaquinas)
    If lErro <> SUCESSO Then gError 137081
    
    MRP.Value = vbUnchecked
    Call Habilita_MRP
    Call Habilita_Operacoes(Roteiro.Nodes.Count)
    
    EscaninhoTerc.ListIndex = 0
    Call EscaninhoTerc_Click
    
    iQtdeAlterada = 0
    
    'se a data de inicio de operacoes MRP for vazia
    If gobjEST.dtDataInicioMRP = DATA_NULA Then
        gbOPValidaParaMRP = False
    Else
        gbOPValidaParaMRP = True
    End If
    '---------------------------------------
    
    '#######################################
    'Inserido por Wagner 07/11/2005
    'Grid de uso de Maquinas real
    Set objGridMaquinasReal = New AdmGrid
    
    'tela em questão
    Set objGridMaquinasReal.objForm = Me
    
    lErro = Inicializa_GridMaquinasReal(objGridMaquinasReal)
    If lErro <> SUCESSO Then gError 140775
    
    'Grid de Mão de Obra
    Set objGridMOReal = New AdmGrid
    
    'tela em questão
    Set objGridMOReal.objForm = Me
    
    lErro = Inicializa_GridMOReal(objGridMOReal)
    If lErro <> SUCESSO Then gError 140776
    
    HorasPO.Format = FORMATO_ESTOQUE
    HorasReal.Format = FORMATO_ESTOQUE
    HorasRealMO.Format = FORMATO_ESTOQUE
    
    If gobjEST.iTemRepeticoesOper = MARCADO Then
        Repeticao.Visible = True
        LabelRepeticao.Visible = True
    Else
        Repeticao.Visible = False
        LabelRepeticao.Visible = False
    End If
    '#######################################
    
    Quantidade.Format = FORMATO_ESTOQUE

    'Preenche a combo Destinação com o padrão (Estoque)
    For iIndice = 0 To DestinacaoPadrao.ListCount - 1
        If DestinacaoPadrao.ItemData(iIndice) = ITEMOP_DESTINACAO_ESTOQUE Then
            DestinacaoPadrao.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Coloca a Data atual nas Datas de Previsão
    DataInicioPadrao.PromptInclude = False
    DataInicioPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataInicioPadrao.PromptInclude = True

    DataFimPadrao.PromptInclude = False
    DataFimPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataFimPadrao.PromptInclude = True

    lErro = CF2(Me, "OP_Form_Load")
    If lErro <> SUCESSO Then gError 21974
    
    'inicializa Grid
    lErro = CF2(Me, "OP_Inicializa_GridMovimentos")
    If lErro <> SUCESSO Then gError 21974

    'Verifica se módulo de Compras não está ativo
    If gcolModulo.Ativo(MODULO_COMPRAS) <> MODULO_ATIVO Then

        'Se Compras não estiver ativo desabilita a CheckBox GeraReqCompra
        GeraReqCompra.Enabled = False
    
    End If
        
    Set gobjOP = New ClassOrdemDeProducao

    If gobjMAT.iOPDetalhamentoRCs = MARCADO Then
        BotaoOPRC.Visible = True
    Else
        BotaoOPRC.Visible = False
    End If
    
    iFornecedorTercAlterado = 0
    iClienteTercAlterado = 0
    
    iAlterado = 0
    iCodigoAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 21974, 22963, 22964, 137080, 137081, 140775, 140776

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163825)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Call CF2(Me, "OP_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing
    
    Set objEventoCodigo = Nothing
    Set objEventoProduto = Nothing
    Set objEventoCcl = Nothing
    Set objEventoCclPadrao = Nothing
    Set objEventoAlmoxPadrao = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoPedidoDeVenda = Nothing
    Set objEventoMaquina = Nothing
    Set objEventoPrestServ = Nothing
    
    Set objGridMaquinasReal = Nothing
    Set objGridMOReal = Nothing
    Set objEventoMaquinas = Nothing
    Set objEventoMO = Nothing

    Set objEventoCliente = Nothing
    Set objEventoFornecedor = Nothing
    
    '---------------------------------------
    'Inserido por Jorge Specian - 29/04/2005
    Set objEventoCompetencias = Nothing
    Set objEventoCentroDeTrabalho = Nothing
    
    Set colPMPItens = Nothing
    Set colOPItens = Nothing
    Set colComponentes = Nothing
    
    Set objGridMaquinas = Nothing
    Set objGridOperacaoInsumos = Nothing
    '---------------------------------------

    Set gcolItemOP = Nothing
    Set gobjOP = Nothing
    
    Set objGrid = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    
    Set gobjAnotacao = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
    If lErro <> SUCESSO Then gError 21976

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case 21976

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163826)

    End Select

    Exit Sub

End Sub

Public Sub GridMovimentos_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGrid, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGrid, iAlterado)
        End If
End Sub

Public Sub GridMovimentos_GotFocus()
    Call Grid_Recebe_Foco(objGrid)
End Sub

Public Sub GridMovimentos_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Public Sub GridMovimentos_LeaveCell()
    Call Saida_Celula(objGrid)
End Sub

Public Sub GridMovimentos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Public Sub GridMovimentos_RowColChange()

Dim lErro As Long

On Error GoTo Erro_GridMovimentos_RowColChange

    Call Grid_RowColChange(objGrid)
    
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub
    
    If GridMovimentos.Row = 0 Then Exit Sub
    
    If (GridMovimentos.Row <> iLinhaAntiga) Then

        lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))
        If lErro <> SUCESSO Then gError 41310

        'Guarda a Linha corrente
        iLinhaAntiga = GridMovimentos.Row

    End If

    Exit Sub

Erro_GridMovimentos_RowColChange:

    Select Case gErr

        Case 41310

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163827)

    End Select

    Exit Sub

End Sub

Public Sub GridMovimentos_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica se é o GridMovimentos
        If objGridInt.objGrid.Name = GridMovimentos.Name Then

            Select Case GridMovimentos.Col
    
                Case iGrid_Produto_Col
    
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 21977
                    
                Case iGrid_Versao_Col
    
                    lErro = Saida_Celula_Versao(objGridInt)
                    If lErro <> SUCESSO Then gError 106337
                
                Case iGrid_Maquina_Col
    
                    lErro = Saida_Celula_Maquina(objGridInt)
                    If lErro <> SUCESSO Then gError 106361
                
                Case iGrid_Quantidade_Col
    
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 21979
    
                Case iGrid_Almoxarifado_Col
    
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 21980
                
    
                Case iGrid_Benef_Col
                    lErro = Saida_Celula_Benef(objGridInt)
                    If lErro <> SUCESSO Then gError 91272
                
                Case iGrid_Ccl_Col
    
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError 21981
    
                Case iGrid_DataPrevInicio_Col
    
                    lErro = Saida_Celula_DataPrevInicio(objGridInt)
                    If lErro <> SUCESSO Then gError 21982
    
                Case iGrid_DataPrevFim_Col
    
                    lErro = Saida_Celula_DataPrevFim(objGridInt)
                    If lErro <> SUCESSO Then gError 21983
    
                Case iGrid_Destinacao_Col
    
                    lErro = Saida_Celula_Destinacao(objGridInt)
                    If lErro <> SUCESSO Then gError 21984
    
                Case iGrid_FilialPedido_Col
    
                    lErro = Saida_Celula_FilialPedido(objGridInt)
                    If lErro <> SUCESSO Then gError 41311
    
                Case iGrid_Situacao_Col
    
                    lErro = Saida_Celula_Situacao(objGridInt)
                    If lErro <> SUCESSO Then gError 41312
    
                Case iGrid_PedidoDeVenda_Col
    
                    lErro = Saida_Celula_PedidoDeVenda(objGridInt)
                    If lErro <> SUCESSO Then gError 41313
    
                Case iGrid_Prioridade_Col
    
                    lErro = Saida_Celula_Prioridade(objGridInt)
                    If lErro <> SUCESSO Then gError 41315
    
            End Select
                   
        End If
        
        '######################################################
        'Inserido por Wagner
        'Verifica se é o GridMovimentos
        If objGridInt.objGrid.Name = GridMaquinasReal.Name Then

            Select Case GridMaquinasReal.Col
    
                Case iGrid_NomeMaquinaReal_Col
                
                    lErro = Saida_Celula_NomeMaquinaReal(objGridInt)
                    If lErro <> SUCESSO Then gError 140783
    
                Case iGrid_QuantidadeMaquinaReal_Col
                
                    lErro = Saida_Celula_QuantidadeMaquinaReal(objGridInt)
                    If lErro <> SUCESSO Then gError 140784
    
                Case iGrid_HorasReal_Col
                
                    lErro = Saida_Celula_HorasReal(objGridInt)
                    If lErro <> SUCESSO Then gError 140785
    
                Case iGrid_DataReal_Col
                
                    lErro = Saida_Celula_DataReal(objGridInt)
                    If lErro <> SUCESSO Then gError 140786
            
            End Select
        
        End If
        
        If objGridInt.objGrid.Name = GridMOReal.Name Then

            Select Case GridMOReal.Col
    
                Case iGrid_NomeMaquinaRealMO_Col
                
                    lErro = Saida_Celula_NomeMaquinaRealMO(objGridInt)
                    If lErro <> SUCESSO Then gError 140787
    
                Case iGrid_QuantidadeMaquinaRealMO_Col
                
                    lErro = Saida_Celula_QuantidadeMaquinaRealMO(objGridInt)
                    If lErro <> SUCESSO Then gError 140788
    
                Case iGrid_HorasRealMO_Col
                
                    lErro = Saida_Celula_HorasRealMO(objGridInt)
                    If lErro <> SUCESSO Then gError 140789
    
                Case iGrid_DataRealMO_Col
                
                    lErro = Saida_Celula_DataRealMO(objGridInt)
                    If lErro <> SUCESSO Then gError 140790
                    
                Case iGrid_TipoMaoDeObraReal_Col
                
                    lErro = Saida_Celula_TipoMaodeObraReal(objGridInt)
                    If lErro <> SUCESSO Then gError 140791
            
            End Select
        
        End If
        '######################################################
        
        lErro = CF2(Me, "OP_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 140791

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 21986

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 21977, 21978, 21979, 21980, 21981, 21982, 21983, 21984, 41311 To 41315, 91272, 106337, 106361, 140783 To 140791

        Case 21986
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163828)

    End Select

    Exit Function

End Function

Private Function VerificaUso_Produto(ByVal objProduto As ClassProduto) As Long
'Verifica se existem produtos repetidos na OP

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProdutoPai As String
Dim sProdutoFormatadoPai As String

On Error GoTo Erro_VerificaUso_Produto

    If objGrid.iLinhasExistentes > 0 Then

        For iIndice = 1 To objGrid.iLinhasExistentes

            If GridMovimentos.Row <> iIndice Then

                lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 22947

                If sProdutoFormatado = objProduto.sCodigo Then gError 41316
                
            End If

        Next

        'Se existir um produto pai de grade no grid ==> verificar se o produto em questao é filho deste pai de grade, se for ==> erro
        If Grid_Possui_Grade Then
            
            'Busca, caso exista, o produto pai de grade do prod em questão
            lErro = CF("Produto_Le_PaiGrade", objProduto, sProdutoPai)
            If lErro <> SUCESSO Then gError 126493
            
            'Se o produto tem um pai de grade
            If Len(Trim(sProdutoPai)) > 0 Then
                
                'Verifica se seu pai aparece no grid
                For iIndice = 1 To objGrid.iLinhasExistentes
                    
                    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatadoPai, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 126492
                    
                    'Se aparecer ==> erro
                    If sProdutoFormatadoPai = sProdutoPai Then gError 126494
                
                Next
            
            End If
            
        End If

    End If

    VerificaUso_Produto = SUCESSO

    Exit Function

Erro_VerificaUso_Produto:

    VerificaUso_Produto = gErr

    Select Case gErr

        Case 22947, 41316, 126492, 126493

        Case 126494
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_GRID", gErr, sProdutoPai, objProduto.sCodigo)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163829)

    End Select

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'faz a critica da celula de proddduto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim objMaquina As New ClassMaquinas
Dim objKit As New ClassKit
Dim objProdutoFilial As New ClassProdutoFilial
Dim iPossuiGrade As Integer
Dim iIndice As Integer
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim objItemOP As New ClassItemOP
Dim objRomaneioGrade As New ClassRomaneioGrade

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 106433
    
    'se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 126489

        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'se é um produto gerencial e não é pai de grade ==> erro
            If lErro = 25043 And Len(Trim(objProduto.sGrade)) = 0 Then gError 126490

            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
            If lErro <> SUCESSO And lErro <> 25043 Then gError 21988

            'verifica se este produto já foi usado na OP
            lErro = VerificaUso_Produto(objProduto)
            If lErro <> SUCESSO And lErro <> 41316 Then gError 55326

            If lErro = 41316 Then gError 41317

            iPossuiGrade = DESMARCADO

            If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO

            If iPossuiGrade = DESMARCADO Then

                'se o produto nao controla estoque ==> erro
                If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 126491

                If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 22946

                 'Preenche a linha do grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
                If lErro <> SUCESSO Then gError 22945

                'Calcula a Quantidade Disponível
                lErro = QuantDisponivel_Calcula(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
                If lErro <> SUCESSO Then gError 41318

                'Verifica se é um kit
                objKit.sProdutoRaiz = sProdutoFormatado
                lErro = CF("Kit_Le_Padrao", objKit)
                If lErro <> SUCESSO And lErro <> 106304 Then gError 106430

                'Se encontrou => É UM KIT => Carrega a Combo com as Versoes
                If lErro <> 106304 Then Call Carrega_ComboVersoes(objProduto.sCodigo)

                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) = 0 Then

                    objProdutoFilial.sProduto = sProdutoFormatado
                    objProdutoFilial.iFilialEmpresa = giFilialEmpresa

                    'Busca o Lote Economico do Produto/FilialEmpresa
                    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
                    If lErro <> SUCESSO And lErro <> 28261 Then gError 106402

                    'preenche com o lote econômico (caso exista)
                    If objProdutoFilial.dLoteEconomico > 0 Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = objProdutoFilial.dLoteEconomico

                End If

            'se é um produto pai de grade
            Else

                'Verifica se há filhos válidos da grade pai
                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
                If lErro <> SUCESSO Then gError 126495

                'Se nao existir, erro
                If colItensRomaneioGrade.Count = 0 Then gError 126496

                'Para cada filho de grade do produto
                For Each objItensRomaneio In colItensRomaneioGrade

                    'Verifica se ele já aparece no grid
                    For iIndice = 1 To objGridInt.iLinhasExistentes

                        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                        If lErro <> SUCESSO Then gError 126498

                        'Se aparecer ==> Erro
                        If sProdutoFormatado = objItensRomaneio.sProduto Then gError 126497

                    Next

                Next

                objItemOP.sProduto = objProduto.sCodigo
                objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
                objItemOP.sCodigo = Codigo.Text
                objItemOP.iItem = GridMovimentos.Row
                objItemOP.sDescricao = objProduto.sDescricao
                objItemOP.sAlmoxarifadoNomeRed = AlmoxPadrao.Text

                Set objRomaneioGrade = New ClassRomaneioGrade
                Set objRomaneioGrade.objTela = Me

                objRomaneioGrade.sNomeTela = Me.Name

                Set objRomaneioGrade.objObjetoTela = objItemOP

                Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
                If giRetornoTela <> vbOK Then gError 126499

                 'Preenche a linha do grid
                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
                If lErro <> SUCESSO Then gError 126552

                Call Atualiza_Grid_Movimentos(objItemOP)

            End If

        End If

        'Incluido por Jorge Specian - 15/06/2005
        '---------------------------------------
        If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) <> 0 Then

            lErro = Trata_Arvore(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)))
            If lErro <> SUCESSO Then gError 137123

        End If
        '---------------------------------------

    End If
    
    Call CF2(Me, "OP_GridMov_AtualizaLinha")
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21985

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 21987, 22945, 41318, 21985, 55326, 106430, 106433, 106402, 126495, 126498, 126499, 126552, 137123
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 21988
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
            
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 22946
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41317
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, Produto.Text, Codigo.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126490
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126491
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126496
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 126497
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, objProduto.sCodigo, GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163830)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantTotal As Double
Dim objProdutoFilial As New ClassProdutoFilial
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade
    
    'se a quantidade foi preenchida
    If Len(Quantidade.ClipText) > 0 Then

        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 41319
    
        '******* Dia 30/10/2002 Sergio: Alteração necessária para verificar se a quantidade solicitada para a produção(Produto), é maior que a Quantidade mínima que pode ser produzida *********************
        
        'Coloca o Produto no Formato do Banco de Dados
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 111429
        
        objProdutoFilial.sProduto = sProdutoFormatado
        objProdutoFilial.iFilialEmpresa = giFilialEmpresa
        
        'Busca o Lote Mínino do Produto/FilialEmpresa
        lErro = CF("ProdutoFilial_Le", objProdutoFilial)
        If lErro <> SUCESSO And lErro <> 28261 Then gError 111430
        
        'preenche Verifica se Lote mínimo esta preenchdo (caso exista)
        If objProdutoFilial.dLoteMinimo > 0 Then
        
            If StrParaDbl(Quantidade.Text) < objProdutoFilial.dLoteMinimo Then
                If gobjCRFAT.iNFImportacaoTribFlag02 <> DESMARCADO Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "O lote mínimo é de " & Formata_Estoque(objProdutoFilial.dLoteMinimo & ". Vai produzir menos ?"))
                    If vbMsgRes = vbNo Then gError 111432
                Else
                    gError 111432
                End If
            End If
        
        End If
                
        '**************** Sergio  *********************
    
        Quantidade.Text = Formata_Estoque(Quantidade.Text)
        
        'Incluido por Jorge Specian - 15/06/2005
        '---------------------------------------
        lErro = Trata_Arvore(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col), StrParaDbl(Quantidade.Text))
        If lErro <> SUCESSO Then gError 137125
        '---------------------------------------
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 21994

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 21994
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41319
            Quantidade.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111429, 111430, 137125
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 111432
            Call Rotina_Erro(vbOKOnly, "ERRO_QDTPRODUTO_MENOR_LOTEMININO", gErr, objProdutoFilial.dLoteMinimo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163831)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.ClipText)) <> 0 Then

        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 22947

        'se produto estiver preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'verifica almoxarifado
            lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 22948

            If lErro = 25157 Then gError 22949

            If lErro = 25162 Then gError 22950

            'calcula qtd disponivel
            lErro = QuantDisponivel_Calcula(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), Almoxarifado.Text)
            If lErro <> SUCESSO Then gError 41322

        End If

    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31501

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 22947, 22948, 31501, 41322
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 22949
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_ALMOXARIFADO2", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 22950
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163832)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Ccl.ClipText) > 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 22951

        If lErro = 5703 Then gError 22952

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31503

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 22951, 31503
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 22952
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            
            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163833)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevInicio(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevInicio

    Set objGridInt.objControle = DataPrevisaoInicio

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoInicio.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoInicio.Text)
        If lErro <> SUCESSO Then gError 31504

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31505

    Saida_Celula_DataPrevInicio = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevInicio:

    Saida_Celula_DataPrevInicio = gErr

    Select Case gErr

        Case 31504, 31505
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163834)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataPrevFim(objGridInt As AdmGrid) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataPrevFim

    Set objGridInt.objControle = DataPrevisaoFim

    'verifica se a data está preenchida
    If Len(Trim(DataPrevisaoFim.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataPrevisaoFim.Text)
        If lErro <> SUCESSO Then gError 31506

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31507

    Saida_Celula_DataPrevFim = SUCESSO

    Exit Function

Erro_Saida_Celula_DataPrevFim:

    Saida_Celula_DataPrevFim = gErr

    Select Case gErr

        Case 31506, 31507
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163835)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Prioridade(objGridInt As AdmGrid) As Long
'faz a critica da celula de Prioridade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Prioridade

    Set objGridInt.objControle = Prioridade

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41323

    Saida_Celula_Prioridade = SUCESSO

    Exit Function

Erro_Saida_Celula_Prioridade:

    Saida_Celula_Prioridade = gErr

    Select Case gErr

        Case 41323
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163836)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PedidoDeVenda(objGridInt As AdmGrid) As Long
'faz a critica da celula de PedidoDeVenda do grid que está deixando de ser a corrente

Dim lErro As Long, iIndice As Integer
Dim objItemPedido As New ClassItemPedido
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Saida_Celula_PedidoDeVenda

    Set objGridInt.objControle = PedidoDeVendaId

    If Len(Trim(PedidoDeVendaId.Text)) > 0 Then
    
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) <> STRING_PEDIDOVENDA Then gError 41645

        If Trim(PedidoDeVendaId.Text) <> Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col)) And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col))) <> 0 Then

            For iIndice = 0 To ComboFilialPedido.ListCount - 1
                If ComboFilialPedido.List(iIndice) = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) Then Exit For
            Next
            
            objItemPedido.iFilialEmpresa = ComboFilialPedido.ItemData(iIndice)
            objItemPedido.lCodPedido = CLng(PedidoDeVendaId.Text)

            sProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 41430

            objItemPedido.sProduto = sProdutoFormatado

            'Verifica se existe um ítem com os dados passados objItemPedido
            lErro = VerificaQuantidade_ItemPedido(objItemPedido)
            If lErro <> SUCESSO And lErro <> 41309 Then gError 41433

            If lErro = 41309 Then gError 41434
                        
            If Trim(PedidoDeVendaId.Text) <> "" And ComboFilialPedido.ListIndex <> -1 Then
                        
                objPedidoDeVenda.lCodigo = CLng(PedidoDeVendaId.Text)
                objPedidoDeVenda.iFilialEmpresa = ComboFilialPedido.ItemData(ComboFilialPedido.ListIndex)
            
                lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
                If lErro <> SUCESSO And lErro <> 26509 Then gError 52029
            
                If lErro = 26509 Then gError 52030
                        
                objCliente.lCodigo = objPedidoDeVenda.lCliente
                
                'le o nome reduzido do cliente
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 52025
            
                If lErro <> SUCESSO Then gError 52026
                        
                'preenche com o nome reduzido do cliente
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                        
                objFilialCliente.lCodCliente = objCliente.lCodigo
                objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
            
                'le o nome da filial do cliente
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 52027
            
                If lErro = 12567 Then gError 52028
            
                'preenche CODIGO - NOME
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
            End If
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41325

    Saida_Celula_PedidoDeVenda = SUCESSO

    Exit Function

Erro_Saida_Celula_PedidoDeVenda:

    Saida_Celula_PedidoDeVenda = gErr

    Select Case gErr
        
        Case 41325, 41430, 41433, 52025, 52027, 52029
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 41434
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41645
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINACAO_DEPENDENTE", gErr)
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52026
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52028
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52030
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163837)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Situacao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Situacao do grid que está deixando de ser a corrente
Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Situacao

    Set objGridInt.objControle = Situacao

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41326

    Saida_Celula_Situacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Situacao:

    Saida_Celula_Situacao = gErr

    Select Case gErr

        Case 41326
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163838)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Benef(objGridInt As AdmGrid) As Long
'faz a critica da celula de Benef do grid que está deixando de ser a corrente


Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Benef

    Set objGridInt.objControle = Benef

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91273
    
    Saida_Celula_Benef = SUCESSO

    Exit Function

Erro_Saida_Celula_Benef:

   Saida_Celula_Benef = gErr

    Select Case gErr

        Case 91273
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163839)

    End Select

    Exit Function

End Function
Private Function Saida_Celula_Destinacao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Destinacao do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Destinacao

    Set objGridInt.objControle = Destinacao

    If Destinacao.ListIndex >= 0 Then
        'limpa celulas se destinacao for diferente de Pedido de Venda e desabilita
        If Destinacao.List(Destinacao.ListIndex) <> STRING_PEDIDOVENDA Then
        
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) = ""
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) = ""
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = ""
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = ""
        
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31508

    Saida_Celula_Destinacao = SUCESSO

    Exit Function

Erro_Saida_Celula_Destinacao:

    Saida_Celula_Destinacao = gErr

    Select Case gErr

        Case 31508
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163840)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialPedido(objGridInt As AdmGrid) As Long
'faz a critica da celula FilialPedido do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objItemPedido As New ClassItemPedido
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sProduto As String
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Saida_Celula_FilialPedido

    Set objGridInt.objControle = ComboFilialPedido

    If ComboFilialPedido.ListIndex >= 0 Then
    
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) <> STRING_PEDIDOVENDA Then gError 41644

        If ComboFilialPedido.List(ComboFilialPedido.ListIndex) <> GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col) And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col))) <> 0 Then

            objItemPedido.iFilialEmpresa = ComboFilialPedido.ItemData(ComboFilialPedido.ListIndex)
            objItemPedido.lCodPedido = CLng(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col))

            sProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 41431

            objItemPedido.sProduto = sProdutoFormatado

            'Verifica se existe um ítem com os dados passados objItemPedido
            lErro = VerificaQuantidade_ItemPedido(objItemPedido)
            If lErro <> SUCESSO And lErro <> 41309 Then gError 41432

            If lErro = 41309 Then gError 41435
                
            If Trim(PedidoDeVendaId.Text) <> "" And ComboFilialPedido.ListIndex <> -1 Then
                        
                objPedidoDeVenda.lCodigo = CLng(PedidoDeVendaId.Text)
                objPedidoDeVenda.iFilialEmpresa = ComboFilialPedido.ItemData(ComboFilialPedido.ListIndex)
            
                lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
                If lErro <> SUCESSO And lErro <> 26509 Then gError 52035
            
                If lErro = 26509 Then gError 52036
                        
                objCliente.lCodigo = objPedidoDeVenda.lCliente
                
                'le o nome reduzido do cliente
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 52031
            
                If lErro <> SUCESSO Then gError 52032
                        
                'preenche com o nome reduzido do cliente
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                        
                objFilialCliente.lCodCliente = objCliente.lCodigo
                objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
            
                'le o nome da filial do cliente
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 52033
            
                If lErro = 12567 Then gError 52034
            
                'preenche CODIGO - NOME
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
            End If
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 31510

    Saida_Celula_FilialPedido = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialPedido:

    Saida_Celula_FilialPedido = gErr

    Select Case gErr

        Case 31510, 41431, 41432, 52031, 52033, 52035
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41435
            ComboFilialPedido.ListIndex = -1
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 41644
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINACAO_DEPENDENTE", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 52032
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52034
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 52036
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
            PedidoDeVendaId.SetFocus
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163841)

    End Select

    Exit Function

End Function

Sub Limpa_GridMovimentos()

    Call Grid_Limpa(objGrid)

End Sub

Function Preenche_GridMovimentos(colItensOP As Collection) As Long
'preenche o grid com os dados contidos na coleção colItensOP

Dim lErro As Long, sCclMascarado As String
Dim iIndice As Integer, iIndice1 As Integer, sProdutoMascarado As String
Dim objItemOP As New ClassItemOP, objProduto As New ClassProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_Preenche_GridMovimentos

    'Remove os ítens de gcolItemOP
    Set gcolItemOP = New Collection
    
    'Incluido por Jorge Specian - 10/05/2005
    '---------------------------------------
    'Remove os itens de colOPItens
    Set colOPItens = New Collection
    '---------------------------------------

    iIndice = 1

    'preenche o grid com os dados retornados na coleção colItensOP
    For Each objItemOP In colItensOP
    
        lErro = CF("ItemOPItemPV_Le2", objItemOP)
        If lErro <> SUCESSO Then gError 22927

        '****** IF INCLUÍDO PARA TRATAMENTO DE GRADE ***************
        If objItemOP.iPossuiGrade = MARCADO Then GridMovimentos.TextMatrix(iIndice, 0) = "# " & GridMovimentos.TextMatrix(iIndice, 0)

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Mascara produto
        lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 22927

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado

        'le o produto para obter sua descricao
        objProduto.sCodigo = objItemOP.sProduto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO Then gError 22928

        GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objProduto.sDescricao
        GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemOP.sSiglaUM
        GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemOP.dQuantidade)
        GridMovimentos.TextMatrix(iIndice, iGrid_ProduzLogo_Col) = CStr(objItemOP.iProduzLogo)

        If objItemOP.iPossuiGrade = DESMARCADO Then
        
            'Tenta ler almoxarifado
            objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
    
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 22235 Then gError 21970
    
            If lErro = 22235 Then gError 21971

            GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        End If

        'Preenche Benef
'        GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col) = objItemOP.iBeneficiamento
        
        Call Combo_Seleciona_ItemData(Benef, objItemOP.iBeneficiamento)
        
        GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col) = Benef.Text
        
        'mascara Ccl , se estiver informada
        If objItemOP.sCcl <> "" Then

            sCclMascarado = String(STRING_CCL, 0)

            lErro = Mascara_MascararCcl(objItemOP.sCcl, sCclMascarado)
            If lErro <> SUCESSO Then gError 22929

        Else
            sCclMascarado = ""
        End If

        GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col) = sCclMascarado

        'preenche datas
        If objItemOP.dtDataInicioProd <> DATA_NULA Then
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col) = Format(objItemOP.dtDataInicioProd, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col) = ""
        End If

        If objItemOP.dtDataFimProd <> DATA_NULA Then
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col) = Format(objItemOP.dtDataFimProd, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col) = ""
        End If

        'preenche Situação
        
        For iIndice1 = 0 To Situacao.ListCount - 1
            If Situacao.ItemData(iIndice1) = objItemOP.iSituacao Then
                Situacao.ListIndex = iIndice1
                Exit For
            End If
        Next
            
        GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col) = Situacao.Text

        'preenche Destinação
        For iIndice1 = 0 To Destinacao.ListCount - 1
            If Destinacao.ItemData(iIndice1) = objItemOP.iDestinacao Then
                Destinacao.ListIndex = iIndice1
                Exit For
            End If
        Next

        GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col) = Destinacao.Text
        
        If GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col) = STRING_PEDIDOVENDA Then
        
            If objItemOP.lCodPedido <> 0 Then
                
                'le o pedido para pegar o Cliente e a filial do cliente
                objPedidoDeVenda.lCodigo = objItemOP.lCodPedido
                objPedidoDeVenda.iFilialEmpresa = objItemOP.iFilialEmpresa
                
                lErro = CF("PedidoDeVenda_Le", objPedidoDeVenda)
                If lErro <> SUCESSO And lErro <> 26509 Then gError 52081
                
                'se nao encontrou  - - - > Erro
                If lErro = 26509 Then
                    'Verifica se o Pedido de Venda está baixado
                    lErro = CF("PedidoVendaBaixado_Le", objPedidoDeVenda)
                    If lErro <> SUCESSO And lErro <> 46135 Then gError 76085
                    
                    'se nao encontrou ==> Erro
                    If lErro = 46135 Then gError 52082
                
                End If
                
                'preenche Filial do pedido
                For iIndice1 = 0 To ComboFilialPedido.ListCount - 1
                    If ComboFilialPedido.ItemData(iIndice1) = objItemOP.iFilialPedido Then
                        ComboFilialPedido.ListIndex = iIndice1
                        Exit For
                    End If
                Next
                
                GridMovimentos.TextMatrix(iIndice, iGrid_FilialPedido_Col) = ComboFilialPedido.Text
            
                'preenche Pedido de Venda e ItemPV
                GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col) = CStr(objItemOP.lCodPedido)
                                
                objCliente.lCodigo = objPedidoDeVenda.lCliente
                    
                'le o nome reduzido do cliente
                lErro = CF("Cliente_Le", objCliente)
                If lErro <> SUCESSO And lErro <> 12293 Then gError 52083
                
                If lErro <> SUCESSO Then gError 52084
                            
                'preenche com o nome reduzido do cliente
                GridMovimentos.TextMatrix(iIndice, iGrid_Cliente_Col) = objCliente.sNomeReduzido
                            
                objFilialCliente.lCodCliente = objCliente.lCodigo
                objFilialCliente.iCodFilial = objPedidoDeVenda.iFilial
                
                'le o nome da filial do cliente
                lErro = CF("FilialCliente_Le", objFilialCliente)
                If lErro <> SUCESSO And lErro <> 12567 Then gError 52085
                
                If lErro = 12567 Then gError 52086
                
                'preenche CODIGO - NOME
                GridMovimentos.TextMatrix(iIndice, iGrid_FilialCliente_Col) = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
            
            Else
                GridMovimentos.TextMatrix(iIndice, iGrid_FilialPedido_Col) = ""
                GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col) = "VÁRIOS"
                GridMovimentos.TextMatrix(iIndice, iGrid_Cliente_Col) = ""
                GridMovimentos.TextMatrix(iIndice, iGrid_FilialCliente_Col) = ""
            
            End If
            
        End If
        
        'preenche prioridade
        GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col) = CStr(objItemOP.iPrioridade)
        
        'Preenche com a Versão do Kit
        GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col) = objItemOP.sVersao
        
        If objItemOP.lNumIntEquipamento <> 0 Then
        
            objMaquina.lNumIntDoc = objItemOP.lNumIntEquipamento
            
            'Le a Máquina atraves do NumIntDoc
            lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
            If lErro <> SUCESSO And lErro <> 106353 Then gError 106355
            
            'Se nao encontrou => Erro
            If lErro = 106353 Then gError 106356
            
            'Preenche a máquina
            GridMovimentos.TextMatrix(iIndice, iGrid_Maquina_Col) = objMaquina.sNomeReduzido
        
        End If
        
        'adiciona item à coleção
        gcolItemOP.Add objItemOP.iItem
        gobjOP.colItens.Add objItemOP

        'Incluido por Jorge Specian - 10/05/2005
        '---------------------------------------
        
        If gbOPValidaParaMRP = True Then
        
            'Inclui itens de colOPItens
            colOPItens.Add objItemOP, objItemOP.sProduto
        
        End If
        '---------------------------------------

        lErro = CF2(Me, "OP_Item_PreencheGridMov", objItemOP, iIndice)
        If lErro <> SUCESSO Then gError 22929
        
        iIndice = iIndice + 1

    Next

    objGrid.iLinhasExistentes = colItensOP.Count
    
    lErro = CF2(Me, "OP_PreencheGridMov")
    If lErro <> SUCESSO Then gError 22929
    
    Call Grid_Refresh_Checkbox(objGrid)
    
    Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_Preenche_GridMovimentos:

    Preenche_GridMovimentos = gErr

    Select Case gErr

        Case 21968, 21969, 21970, 22929, 52081, 52083, 52085, 76085

        Case 21971
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
        
        Case 22927
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)
        
        Case 22928
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 52082
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoDeVenda.lCodigo)
        
        Case 52084
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
    
        Case 52086
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_SEM_FILIAL", gErr, objFilialCliente.lCodCliente)
            
        Case 106355
        
        Case 106356
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.sNomeReduzido)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163842)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim sPedidoDeVenda As String
Dim vbMsg As VbMsgBoxResult
Dim sSituacao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim bAchou As Boolean
'Incluido por Jorge Specian
'------------------------------
Dim bEstaNoPMP As Boolean
Dim objPMPItem As ClassPMPItens
'------------------------------
Dim objItemOP As ClassItemOP

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    bAchou = False
    bEstaNoPMP = False   'by Jorge Specian
    
    If Len(Trim(Codigo.Text)) = 0 Then gError 31513

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 22954

    If objGrid.iLinhasExistentes = 0 Then gError 22953
    
    If GeraOP.Value = vbChecked Then
        If Len(Trim(CodOPGerada.Text)) = 0 Then gError 62562
    End If
    
    'Loop de Validação dos dados do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a quantidade foi digitada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 31515

        'Verifica se o almoxarifado foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 And left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then gError 31516

        'Verifica se a data de previsão de inicio da OP foi informada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col))) = 0 Then gError 41358

        'Verifica se a data de previsão de fim da OP foi informada e se é menor que a data de início
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col))) > 0 Then
            If CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col)) < CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col)) Then gError 41336
        Else
            gError 41337
        End If

        'Verifica se a Situação é Baixada para ítens novos
        sSituacao = GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col)
        If Len(Trim(sSituacao)) > 0 Then
            If sSituacao = STRING_BAIXADA And gcolItemOP.Item(iIndice) = 0 Then gError 41327
        End If
        
        'Verifica se é uma ordem de producao baixada
        If StatusOP.Caption = STRING_STATUS_BAIXADO Then
            
            'Verifica se a Situação é Normal para o item
            If Len(Trim(sSituacao)) > 0 Then
                
                If sSituacao = STRING_NORMAL Then bAchou = True
                
            End If
            
        End If
        
        ' se destinacao for PV ==> filialPV e COdPV tem que estar preenchidos
        sPedidoDeVenda = GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col)
        If Len(Trim(sPedidoDeVenda)) > 0 Then

            If sPedidoDeVenda = STRING_PEDIDOVENDA Then

                If GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col) <> "VÁRIOS" Then
                    If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col))) = 0 Then gError 41328
                    If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_FilialPedido_Col))) = 0 Then gError 41329
                End If
                
            End If

        End If
        
        'Incluido por Jorge Specian
        '--------------------------------------------------
        'Verifica se está no Plano Mestre de Producao
        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 139068

        'Para cada item do Plano Mestre
        For Each objPMPItem In colPMPItens
        
            If objPMPItem.sProduto = sProdutoFormatado Then
                bEstaNoPMP = True
                Exit For
            End If
            
        Next
        '--------------------------------------------------

    Next
    
    'If bEstaNoPMP = True Then gError 139069   'by Jorge Specian
    
    'se a Op está baixada e existe item com situacao='normal'
    If bAchou = True Then
    
        vbMsg = Rotina_Aviso(vbYesNo, "AVISO_REATIVACAO_OP", Codigo.Text)
        'se não for reativar a OP sai da gravação
        If vbMsg = vbNo Then gError 82804
    
    ElseIf bAchou = False And StatusOP.Caption = STRING_STATUS_BAIXADO Then
        gError 82803
    End If
    
    If StatusOP.Caption = STRING_STATUS_BAIXADO Then
        objOrdemDeProducao.iStatusOP = ITEMOP_SITUACAO_BAIXADA
    
    ElseIf StatusOP.Caption = STRING_NORMAL Then
        objOrdemDeProducao.iStatusOP = ITEMOP_SITUACAO_NORMAL
    End If
    
    lErro = CF2(Me, "OP_GravarRegistro")
    If lErro <> SUCESSO Then gError 31518
    
    lErro = Move_Tela_Memoria(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 31518
    
    '###############################################
    'Inserido por Wagner 10/08/2006
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 185836
    '###############################################
    
    If gobjMAT.iTrataEstTercCliForn = MARCADO Then
        If objOrdemDeProducao.lCodTerc = 0 Or objOrdemDeProducao.iFilialTerc = 0 Then
            For Each objItemOP In objOrdemDeProducao.colItens
                If objItemOP.iBeneficiamento <> BENEF_COMBO_DISP Then gError 209644
            Next
        End If
    End If

    lErro = CF("OrdemDeProducao_Grava", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 31520

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_OP
            gobjAnotacao.sID = CStr(objOrdemDeProducao.iFilialEmpresa) & "," & objOrdemDeProducao.sCodigo
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 31520
            
        End If
        
    End If
    
    'Se a opcao de imprimir o Relatorio estiver marcada
    If ImprimeAoGravar.Value = MARCADO Then
        
        'Gera o(s) Relatorio(s)
        lErro = CF2(Me, "OP_Executa_Relatorio", objOrdemDeProducao)
        If lErro <> SUCESSO Then gError 106405
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 22953
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEMOP_INFORMADO", gErr)

        Case 22954, 41358, 41337
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 31513
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr)

        Case 31515
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 31516
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 31518, 31520, 82804, 106405, 139068, 185836

        Case 41327
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAR_ITEMNOVO", gErr)
        
        Case 41328
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDAID_NAO_PREENCHIDO", gErr, iIndice)

        Case 41329
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_PREENCHIDA", gErr, iIndice)

        Case 41336
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_FIM_MENOR", gErr)

        Case 62562
''''            Call Rotina_Erro(vbOKOnly, "ERRO_CODOP_GERAR_NAO_INFORMADO", gErr)
        
        Case 82803
            Call Rotina_Erro(vbOKOnly, "ERRO_OPBAIXADA_NAO_REATIVADA", gErr)
            
        'Incluido por Jorge Specian
        '--------------------------------------------------
        Case 139069
            Call Rotina_Erro(vbOKOnly, "ERRO_OPINALTERAVEL_ESTA_NO_PMP", gErr)
        '--------------------------------------------------
        
        Case 209644
            Call Rotina_Erro(vbOKOnly, "ERRO_OP_TERC_SEM_CODIGO_TERC", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163843)

    End Select

    Exit Function

End Function

Function Move_Grid_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'move itens do Grid para objOrdemDeProducao

Dim lErro As Long
Dim iIndice As Integer, iCount As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCcl As String, sCclFormatada As String, iCclPreenchida As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objItemOP As ClassItemOP
Dim sSituacao As String, sDestinacao As String, sFilialPedido As String
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As AdmCodigoNome
Dim objMaquina As New ClassMaquinas
'Incluido por Jorge Specian - 10/05/2005
'---------------------------------------
Dim objItemOPOperacoes As New ClassItemOP
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
'---------------------------------------
Dim lBenef As Long

On Error GoTo Erro_Move_Grid_Memoria

    objOrdemDeProducao.iNumItens = 0
    objOrdemDeProducao.iNumItensBaixados = 0

    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objItemOP = New ClassItemOP

        If gobjOP.colItens.Count >= iIndice Then

            '#####################################
            'Alterado por Wagner - Estava gravando OPs sem itens quando
            'se trazia uma OP para tela, mudava-se o código e regravava
            'porque estava recebendo um NunIntDoc de um Item que já estava no BD
            'objItemOP.lNumIntDoc = gobjOP.colItens(iIndice).lNumIntDoc
            If gobjOP.colItens(iIndice).sCodigo = objOrdemDeProducao.sCodigo Then
                objItemOP.lNumIntDoc = gobjOP.colItens(iIndice).lNumIntDoc
            End If
            '#####################################
        
        End If

        objItemOP.sCodigo = objOrdemDeProducao.sCodigo
        objItemOP.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa

        sProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 31563

        objItemOP.sProduto = sProdutoFormatado

        objItemOP.sSiglaUM = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
            objItemOP.dQuantidade = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
        Else
            objItemOP.dQuantidade = 0
        End If

        'se nao for pai de grade
        If left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then

            objAlmoxarifado.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
            objItemOP.iPossuiGrade = DESMARCADO
    
            If colCodigoNome.Count > 0 Then
                
                For Each objCodigoNome In colCodigoNome
                    If objCodigoNome.sNome = objAlmoxarifado.sNomeReduzido Then
                        objItemOP.iAlmoxarifado = objCodigoNome.iCodigo
                        Exit For
                    End If
                Next
            
            End If
                    
            If objItemOP.iAlmoxarifado = 0 Then
    
                'lê nome reduzido do almoxarifado
                lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25060 Then gError 19307
        
                If lErro = 25060 Then gError 19308
        
                objItemOP.iAlmoxarifado = objAlmoxarifado.iCodigo
        
                colCodigoNome.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido
                
            End If
        
        Else
        
            objItemOP.iPossuiGrade = MARCADO
        
        End If
        
        Call Combo_Obtem_ItemData(Benef, GridMovimentos.TextMatrix(iIndice, iGrid_Benef_Col), lBenef)
        
        objItemOP.iBeneficiamento = lBenef

        sCcl = GridMovimentos.TextMatrix(iIndice, iGrid_Ccl_Col)

        If Len(Trim(sCcl)) <> 0 Then

            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError 19309

        Else
            sCclFormatada = ""
        End If

        objItemOP.sCcl = sCclFormatada

        If Len(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col)) > 0 Then
            objItemOP.dtDataInicioProd = CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevInicio_Col))
        Else
            objItemOP.dtDataInicioProd = DATA_NULA
        End If

        If Len(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col)) > 0 Then
            objItemOP.dtDataFimProd = CDate(GridMovimentos.TextMatrix(iIndice, iGrid_DataPrevFim_Col))
        Else
            objItemOP.dtDataFimProd = DATA_NULA
        End If

        'Seleciona a situação
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col))) > 0 Then
            sSituacao = GridMovimentos.TextMatrix(iIndice, iGrid_Situacao_Col)
            For iCount = 0 To Situacao.ListCount - 1
                If Situacao.List(iCount) = sSituacao Then
                    objItemOP.iSituacao = Situacao.ItemData(iCount)
                    Exit For
                End If
            Next
        End If

        'seleciona a destinação , junto com FilialPedido,PedidoDeVenda e ItemPV
        sDestinacao = GridMovimentos.TextMatrix(iIndice, iGrid_Destinacao_Col)
        If Len(Trim(sDestinacao)) > 0 Then

            If sDestinacao = STRING_PEDIDOVENDA Then
                sFilialPedido = GridMovimentos.TextMatrix(iIndice, iGrid_FilialPedido_Col)
                If Len(Trim(sFilialPedido)) > 0 Then
                    For iCount = 0 To ComboFilialPedido.ListCount - 1
                        If ComboFilialPedido.List(iCount) = sFilialPedido Then
                            objItemOP.iFilialPedido = ComboFilialPedido.ItemData(iCount)
                            Exit For
                        End If
                    Next
                End If
                If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col))) > 0 And GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col) <> "VÁRIOS" Then objItemOP.lCodPedido = CLng(GridMovimentos.TextMatrix(iIndice, iGrid_PedidoDeVenda_Col))
            End If

            'seleciona a Destinação
            For iCount = 0 To Destinacao.ListCount - 1
                If Destinacao.List(iCount) = sDestinacao Then
                    objItemOP.iDestinacao = Destinacao.ItemData(iCount)
                    Exit For
                End If
            Next

        End If

        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col))) > 0 Then objItemOP.iPrioridade = CInt(GridMovimentos.TextMatrix(iIndice, iGrid_Prioridade_Col))

        objItemOP.iItem = iIndice
        
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Maquina_Col))) > 0 Then
            
            'Le a Maquina a partir do nome reduzido
            objMaquina.sNomeReduzido = GridMovimentos.TextMatrix(iIndice, iGrid_Maquina_Col)
            
            lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
            If lErro <> SUCESSO And lErro <> 103100 Then gError 106345
            
            'Se nao encontrou => Erro
            If lErro = 103100 Then gError 106346
            
            objItemOP.lNumIntEquipamento = objMaquina.lNumIntDoc
            
        End If
        
        objItemOP.sVersao = GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col)
        objItemOP.iProduzLogo = StrParaInt(GridMovimentos.TextMatrix(iIndice, iGrid_ProduzLogo_Col))

        'Incluido por Jorge Specian - 10/05/2005
        '---------------------------------------
        
        'Se a OP é válida para MRP (Data emissão >= Data Inicio Operações MRP)
        If gbOPValidaParaMRP = True Then

            'Para cada item da coleção
            For Each objItemOPOperacoes In colOPItens
                
                'Se é o produto da linha do grid
                If objItemOPOperacoes.sProduto = objItemOP.sProduto Then
                
                    'Para cada Operação deste Produto
                    For Each objOrdemProducaoOperacoes In objItemOPOperacoes.colOrdemProducaoOperacoes
                    
                        'Verifica se a Competência está preenchida
                        If objOrdemProducaoOperacoes.lNumIntDocCompet = 0 Then gError 137704
                        
                        'Verifica se o Centro de Trabalho está preenchido
                        If objOrdemProducaoOperacoes.lNumIntDocCT = 0 Then gError 137705
                    
                        'Se tudo Ok... Alimenta a coleção do Item da OP
                        objItemOP.colOrdemProducaoOperacoes.Add objOrdemProducaoOperacoes
                        
                    Next
                    
                    'Finda a busca pelo produto na coleção
                    Exit For
                    
                End If
            
            Next
        
        End If
        '---------------------------------------

        objOrdemDeProducao.colItens.Add objItemOP

        If objItemOP.iPossuiGrade = DESMARCADO Then

            objOrdemDeProducao.iNumItens = objOrdemDeProducao.iNumItens + 1
            If objItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then objOrdemDeProducao.iNumItensBaixados = objOrdemDeProducao.iNumItensBaixados + 1

        End If

        Set objItemOP.colProdutoKitInfo = gobjOP.colItens.Item(iIndice).colProdutoKitInfo

        '********************* TRATAMENTO DE GRADE *****************
        Call Move_ItensGrade_Tela(objOrdemDeProducao, objItemOP, objItemOP.colItensRomaneioGrade, gobjOP.colItens(iIndice).colItensRomaneioGrade)
    
        lErro = CF2(Me, "OP_Move_Grid_Memoria", iIndice, objItemOP)
        If lErro <> SUCESSO Then gError 19309
        
    Next

    Move_Grid_Memoria = SUCESSO

    Exit Function

Erro_Move_Grid_Memoria:

    Move_Grid_Memoria = gErr

    Select Case gErr

        Case 19307, 19309, 31563

        Case 19308
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 106345
        
        Case 106346
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.sNomeReduzido)
            'ERRO_MAQUINA_NAO_CADASTRADA = "A máquina %s não está cadastrada."
        
        'Incluido por Jorge Specian - 07/07/2005
        '---------------------------------------
        Case 137704
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_NAO_INFORMADA", gErr, objOrdemProducaoOperacoes.iNivel, objOrdemProducaoOperacoes.iSeqArvore, objItemOP.sProduto)
        
        Case 137705
            Call Rotina_Erro(vbOKOnly, "ERRO_CENTRODETRABALHO_NAO_INFORMADO", gErr, objOrdemProducaoOperacoes.iNivel, objOrdemProducaoOperacoes.iSeqArvore, objItemOP.sProduto)
        '---------------------------------------
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163844)

    End Select

    Exit Function

End Function

Function Move_ItensGrade_Tela(objOrdemProducao As ClassOrdemDeProducao, objItemOP As ClassItemOP, colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeTela In colItensRomaneioTela
                    
        Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
        objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
        objItemRomaneioGrade.iAlmoxarifado = objItemRomaneioGradeTela.iAlmoxarifado
        objItemRomaneioGrade.sAlmoxarifado = objItemRomaneioGradeTela.sAlmoxarifado
        objItemRomaneioGrade.sVersao = objItemRomaneioGradeTela.sVersao
        objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
                    
        colItensRomaneio.Add objItemRomaneioGrade
        
        objOrdemProducao.iNumItens = objOrdemProducao.iNumItens + 1
        If objItemOP.iSituacao = ITEMOP_SITUACAO_BAIXADA Then objOrdemProducao.iNumItensBaixados = objOrdemProducao.iNumItensBaixados + 1
        
    Next

End Function

Function Limpa_Tela_OrdemDeProducao(Optional iFechaSetas As Integer = FECHAR_SETAS) As Long
'Limpa a Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_OrdemDeProducao

    If iFechaSetas = FECHAR_SETAS Then
    'Fecha o comando das setas se estiver aberto
     lErro = ComandoSeta_Fechar(Me.Name)
     If lErro <> SUCESSO Then gError 21801
    End If
    Call Limpa_Tela(Me)

    QuantDisponivel.Caption = ""
    StatusOP.Caption = ""
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    For iIndice = 1 To objGrid.iLinhasExistentes 'm
        GridMovimentos.TextMatrix(iIndice, 0) = iIndice
    Next
    
    Call Limpa_GridMovimentos

    Set gcolItemOP = New Collection

    Set gobjOP = New ClassOrdemDeProducao

    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Coloca a Data atual nas Datas de Previsão
    DataInicioPadrao.PromptInclude = False
    DataInicioPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataInicioPadrao.PromptInclude = True

    DataFimPadrao.PromptInclude = False
    DataFimPadrao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataFimPadrao.PromptInclude = True

    GeraOP.Value = vbUnchecked
    GeraReqCompra.Value = vbUnchecked
    GeraOPs.Value = vbUnchecked
    IgnoraEst.Value = vbUnchecked

    ImprimeAoGravar.Value = DESMARCADO
    OpcaoRelatorio(2).Value = True
    
    '---------------------------------------
    'Incluido por Jorge Specian - 02/05/2005
    Set colOPItens = New Collection
    Set colPMPItens = New Collection
    
    Call Limpa_Operacoes
        
    Call Limpa_Arvore_Roteiro
    
    MRP.Value = vbUnchecked
    Call Habilita_MRP
    Call Habilita_Operacoes(Roteiro.Nodes.Count)
    
    iQtdeAlterada = 0
    
    'se a data de inicio de operacoes MRP for vazia
    If gobjEST.dtDataInicioMRP = DATA_NULA Then
        gbOPValidaParaMRP = False
    Else
        gbOPValidaParaMRP = True
    End If
    '---------------------------------------
    
    FilialTerc.Clear
    EscaninhoTerc.ListIndex = 0
    Call EscaninhoTerc_Click
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    Call CF2(Me, "OP_Limpa_Tela_OP")
    
    iAlterado = 0
    iCodigoAlterado = 0
    iFornecedorTercAlterado = 0
    iClienteTercAlterado = 0
    
    Set gobjAnotacao = Nothing

    Limpa_Tela_OrdemDeProducao = SUCESSO

    Exit Function

Erro_Limpa_Tela_OrdemDeProducao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163845)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Trata_Parametros

    If Not (objOrdemDeProducao Is Nothing) Then

        'traz OP para a tela
        lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 21966 Then gError 31556

        If lErro = 21966 Then

            'Se não existe exibe apenas o código
            Codigo.Text = objOrdemDeProducao.sCodigo

        End If
                
        Call ComandoSeta_Fechar(Me.Name)
                
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 31556

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163846)

    End Select

    iAlterado = 0

    Exit Function

End Function


'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OrdemProducaoOP"

    objOrdemDeProducao.sCodigo = Codigo.Text
    objOrdemDeProducao.dtDataEmissao = CDate(Data.Text)
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    
    If IsNumeric(Codigo.Text) Then
        objOrdemDeProducao.lCodigoNumerico = StrParaLong(Codigo.Text)
    End If
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objOrdemDeProducao.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "CodigoNumerico", objOrdemDeProducao.lCodigoNumerico, 0, "CodigoNumerico"
    colCampoValor.Add "Codigo", objOrdemDeProducao.sCodigo, STRING_ORDEM_DE_PRODUCAO, "Codigo"
    colCampoValor.Add "DataEmissao", objOrdemDeProducao.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "OPGeradora", objOrdemDeProducao.sOPGeradora, STRING_ORDEM_DE_PRODUCAO, "OPGeradora"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        'Erro já tratado
        Case 31557

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163847)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long, objPrestServ As New ClassPrestServ
Dim objProjetoInfo As Object
Dim objCliente As ClassCliente
Dim objFornecedor As ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    objOrdemDeProducao.sCodigo = Trim(Codigo.Text)
    objOrdemDeProducao.dtDataEmissao = CDate(Data.Text)
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    objOrdemDeProducao.iGeraOP = GeraOP.Value
    objOrdemDeProducao.iGeraReqCompra = GeraReqCompra.Value
    
    '#######################
    'INSERIDO POR WAGNER
    objOrdemDeProducao.iGeraOPsArvore = GeraOPs.Value
    '#######################
    
    If IgnoraEst.Value = vbChecked Then
        objOrdemDeProducao.iIgnoraEst = MARCADO
    Else
        objOrdemDeProducao.iIgnoraEst = DESMARCADO
    End If
    
    'Verifica preenchimento de Prestador de servico
    If Len(Trim(PrestadorServico.Text)) <> 0 Then

        objPrestServ.sNomeReduzido = PrestadorServico.Text

        'Lê no BD
        lErro = CF("PrestServ_Le_NomeReduzido", objPrestServ)
        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError 124249

        'Se não achou --> erro
        If lErro <> SUCESSO Then gError 124250

        objOrdemDeProducao.lCodPrestador = objPrestServ.lCodigo

    End If
    
    If objOrdemDeProducao.iGeraOP = vbChecked Then
        objOrdemDeProducao.sCodOPGerar = CodOPGerada
    End If

    lErro = Move_Grid_Memoria(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 31519
    
    '##################################################
    'Inserido por Wagner 09/08/2006
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_OP)
    If lErro <> SUCESSO Then gError 185837
    
    Set objOrdemDeProducao.objProjetoInfo = objProjetoInfo
    '##################################################

    lErro = CF2(Me, "OP_Move_Tela_Memoria", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 185837
    
    
    'se for cliente
    If OptionTipoTerc(1).Value = True Then
        
        'instancia o obj cliente
        Set objCliente = New ClassCliente
        
        'carrega o obj c/ o tipo de terceiro (cliente)
        objOrdemDeProducao.iTipoTerc = TIPO_TERC_CLIENTE
        
        If Len(Trim(ClienteTerc.Text)) > 0 Then
        
            'preenche o objcliente c/ o nomered do cliente na tela
            objCliente.sNomeReduzido = Trim(ClienteTerc.Text)
            
            'busca o cód. do cliente a apartir do nomereduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError ERRO_SEM_MENSAGEM
            
            'cliente não cadastrado
            If lErro = 12348 Then gError 209631
            
            'preenche o obj c/ o cód do cliente
            objOrdemDeProducao.lCodTerc = objCliente.lCodigo
            
        End If
        
    'senão, é fornecedor
    Else
    
        'instancia o obj fornecedor
        Set objFornecedor = New ClassFornecedor
        
        'carrega o obj c/ o tipo de terceiro (fornecedor)
        objOrdemDeProducao.iTipoTerc = TIPO_TERC_FORNECEDOR
        
        If Len(Trim(FornecedorTerc.Text)) > 0 Then
        
            'carrega o objforncedor com o nomered do fornecedor da tela
            objFornecedor.sNomeReduzido = Trim(FornecedorTerc.Text)
            
            'busca o cód. do fornecedor a apartir do nomereduzido
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError ERRO_SEM_MENSAGEM
        
            'fornecedor não cadastrado
            If lErro = 6681 Then gError 209632
        
            'preenche o obj c/ o cód. do fornecedor
            objOrdemDeProducao.lCodTerc = objFornecedor.lCodigo
            
        End If
    
    End If

    'passa o código da filial do cliente / fornecedor p/ o obj
    objOrdemDeProducao.iFilialTerc = Codigo_Extrai(FilialTerc.Text)
    objOrdemDeProducao.iEscaninhoTerc = EscaninhoTerc.ItemData(EscaninhoTerc.ListIndex)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 31519, 124249, 185837
        
        Case 124250
            Call Rotina_Erro(vbOKOnly, "ERRO_PRESTSERV_NAO_CADASTRADO1", Err, objPrestServ.sNomeReduzido)

        Case 209631
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, FornecedorTerc.Text)
                    
        Case 209632
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, ClienteTerc.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163848)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Tela_Preenche

    objOrdemDeProducao.sCodigo = colCampoValor.Item("Codigo").vValor
    objOrdemDeProducao.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objOrdemDeProducao.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
    objOrdemDeProducao.sOPGeradora = colCampoValor.Item("OPGeradora").vValor

    'Traz dados da Ordem de Produção para a Tela
    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21966 Then gError 31558

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 31558

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163849)

    End Select

    Exit Sub

End Sub

Function Traz_Tela_OrdemDeProducao(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'preenche a tela com os dados da OP

Dim lErro As Long
Dim objPMP As ClassPMP
Dim objPMPItens As New ClassPMPItens

On Error GoTo Erro_Traz_Tela_OrdemDeProducao

    lErro = Limpa_Tela_OrdemDeProducao(NAO_FECHAR_SETAS)
    If lErro <> SUCESSO Then gError 41330

    Codigo.Text = CStr(objOrdemDeProducao.sCodigo)

    lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 21963

    If lErro = 21960 Then

        lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 82797 Then gError 82801
        
        If lErro = 82797 Then gError 21966
        
        StatusOP.Caption = STRING_STATUS_BAIXADO
        
    End If
    
    If objOrdemDeProducao.iTipo = OP_TIPO_OC Then gError 117662
    
    If objOrdemDeProducao.lCodPrestador <> 0 Then
        PrestadorServico.Text = CStr(objOrdemDeProducao.lCodPrestador)
    Else
        PrestadorServico.Text = ""
    End If
    Call PrestadorServico_Validate(bSGECancelDummy)
    
    Call DateParaMasked(Data, objOrdemDeProducao.dtDataEmissao)
    
    If Len(Trim(StatusOP.Caption)) = 0 Then
        StatusOP.Caption = "ATIVO"
    End If
    
    If objOrdemDeProducao.iIgnoraEst = MARCADO Then
        IgnoraEst.Value = vbChecked
    Else
        IgnoraEst.Value = vbUnchecked
    End If
    
    '---------------------------------------
    'Incluido por Jorge Specian - 18/05/2005
    iQtdeAlterada = 0
    
    'se a data de emissão da OP for maior ou igual que a data de inicio de operacoes MRP
    If objOrdemDeProducao.dtDataEmissao >= gobjEST.dtDataInicioMRP And gobjEST.dtDataInicioMRP <> DATA_NULA Then
    
        Set objPMP = New ClassPMP
        
        'Lê o Plano Mestre de Produção para preencher a coleção
        lErro = CF("PlanoMestreProducao_Le_ItensOP", objPMP, objOrdemDeProducao)
        If lErro <> SUCESSO Then gError 137144
        
        'Se foi gerado o MRP ...
        If objPMP.lCodGeracao > 0 Then
            
            'percorre os itens preenchendo a coleção
            For Each objPMPItens In objPMP.colItens
            
                colPMPItens.Add objPMPItens
            
            Next
        
        End If
        
    Else
    
        'Esta OP não será valida para o MRP
        gbOPValidaParaMRP = False
    
    End If
    '---------------------------------------
    
    'preenche o grid
    lErro = Preenche_GridMovimentos(objOrdemDeProducao.colItens)
    If lErro <> SUCESSO Then gError 21972
    
    lErro = CF("OPGeradora_Le_OPGerada", objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 62637 Then gError 62633
    
    If lErro = SUCESSO Then
        CodOPGerada.Text = objOrdemDeProducao.sCodOPGerar
    End If
    
    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objOrdemDeProducao.objProjetoInfo, 0, PRJ_CR_TIPO_OP, objOrdemDeProducao.sCodigo, objOrdemDeProducao.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 181676
    '###################################################
        
    lErro = CF2(Me, "OP_Traz_Tela_OP", objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 181676
    
    'se for cliente
    If objOrdemDeProducao.iTipoTerc = TIPO_TERC_CLIENTE Then
        
        'coloca a opt. cliente como marcada
        OptionTipoTerc(1).Value = True
        Call OptionTipoTerc_Trata
        
        If objOrdemDeProducao.lCodTerc > 0 Then
        
            'coloca o cód. do cliente na tela e chama o validate
            ClienteTerc.Text = objOrdemDeProducao.lCodTerc
            Call ClienteTerc_Validate(bSGECancelDummy)
            
            'coloca a filial do cliente na tela
            FilialTerc.Text = objOrdemDeProducao.iFilialTerc
            Call FilialTerc_Validate(bSGECancelDummy)
            
        End If
               
    'senão, é o fornecedor
    ElseIf objOrdemDeProducao.iTipoTerc = TIPO_TERC_FORNECEDOR Then
    
        'coloca a opt. do fornecedor marcada
        OptionTipoTerc(2).Value = True
        Call OptionTipoTerc_Trata
    
        If objOrdemDeProducao.lCodTerc > 0 Then
    
            'coloca o cód. do fornecedor na tela e chama o validate
            FornecedorTerc.Text = objOrdemDeProducao.lCodTerc
            Call FornecedorTerc_Validate(bSGECancelDummy)
            
            'coloca a filial do fornecedor na tela
            FilialTerc.Text = objOrdemDeProducao.iFilialTerc
            Call FilialTerc_Validate(bSGECancelDummy)
            
        End If
        
    End If
    
    Call Combo_Seleciona_ItemData(EscaninhoTerc, objOrdemDeProducao.iEscaninhoTerc)
    
    iAlterado = 0
    iCodigoAlterado = 0
    
    Traz_Tela_OrdemDeProducao = SUCESSO

    Exit Function

Erro_Traz_Tela_OrdemDeProducao:

    Traz_Tela_OrdemDeProducao = gErr

    Select Case gErr

        Case 21963, 21972, 21966, 41330, 62633, 82801, 137144, 181676
            'erros tratados nas rotinas chamadas

        Case 117662
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDECORTE", gErr, objOrdemDeProducao.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163850)

    End Select

    Exit Function

End Function

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    'Se o produto da linha corrente estiver preenchido e Linha corrente diferente da Linha fixa
    If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) <> 0 And GridMovimentos.Row <> 0 Then

        sCclMascarado = String(STRING_CCL, 0)

        lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 22934

        'Coloca o valor do Ccl na coluna correspondente
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = sCclMascarado

        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True

    End If

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 22934

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163851)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrdemDeProducao As ClassOrdemDeProducao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objOrdemDeProducao = obj1

    'traz OP para a tela
    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 34675

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 34675

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163852)
    End Select

    Exit Sub

End Sub

Public Sub PedidoDeVendaId_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PedidoDeVendaId_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)
    
End Sub

Public Sub PedidoDeVendaId_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub PedidoDeVendaId_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = PedidoDeVendaId
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrestadorServico_Change()

    iPrestServAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PrestadorServico_Validate(Cancel As Boolean)
    
Dim lErro As Long, objPrestServ As New ClassPrestServ

On Error GoTo Erro_PrestadorServico_Validate

    If iPrestServAlterado = 1 Then

        'Verifica o preenchimento
        If Len(Trim(PrestadorServico.Text)) > 0 Then

            'Tenta ler (NomeReduzido ou Código)
            lErro = TP_PrestServ_Le(PrestadorServico, objPrestServ)
            If lErro <> SUCESSO Then gError 124246

        End If
        
        iPrestServAlterado = 0
        
    End If
    
    Exit Sub
     
Erro_PrestadorServico_Validate:

    Select Case gErr
          
        Case 124246
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163853)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Prioridade_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Prioridade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Prioridade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Prioridade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Prioridade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PrioridadePadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(PrioridadePadrao, iAlterado)

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_GotFocus()

Dim lErro As Long

    Call Grid_Campo_Recebe_Foco(objGrid)

    If gobjEST.iInventarioCodBarrAuto = 1 Then

        If objGrid.lErroSaidaCelula = 0 Then

            lErro = Trata_CodigoBarras1

            objGrid.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

            Call Grid_Entrada_Celula(objGrid, iAlterado)

            objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

            If lErro <> SUCESSO Then
    
                objGrid.lErroSaidaCelula = 1
            End If

        Else
    
            objGrid.lErroSaidaCelula = 0
    
        End If
        
    End If
    
End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO
    iQtdeAlterada = REGISTRO_ALTERADO
    
End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Situacao_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Situacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Situacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Situacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Situacao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub



Public Sub UpDownData_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 22931

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 22931

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163854)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 22932

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 22932

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163855)

    End Select

    Exit Sub

End Sub

Public Sub UpDownInicio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownInicio_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownInicio_DownClick

    If Len(Trim(DataInicioPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataInicioPadrao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 41298

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownInicio_DownClick:

    Select Case gErr

        Case 41298

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163856)

    End Select

    Exit Sub

End Sub

Public Sub UpDownInicio_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownInicio_UpClick

    If Len(Trim(DataInicioPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataInicioPadrao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 41299

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownInicio_UpClick:

    Select Case gErr

        Case 41299

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163857)

    End Select

    Exit Sub

End Sub

Public Sub UpDownFim_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownFim_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFim_DownClick

    If Len(Trim(DataFimPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataFimPadrao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 41300

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownFim_DownClick:

    Select Case gErr

        Case 41300

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163858)

    End Select

    Exit Sub

End Sub

Public Sub UpDownFim_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownFim_UpClick

    If Len(Trim(DataFimPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Up_Down_Click(DataFimPadrao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 41301

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_UpDownFim_UpClick:

    Select Case gErr

        Case 41301

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163859)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_gcolItemOP(gcolItemOP As Collection)
'limpa gcolItemOP e a coluna de qtde das linhas que estavam na tela apos a troca do codigo da OP

Dim lErro As Long
Dim iCount As Integer
Dim iIndice As Integer

On Error GoTo Erro_Limpa_gcolItemOp

    iCount = gcolItemOP.Count
    Set gcolItemOP = New Collection

    For iIndice = 1 To iCount

        gcolItemOP.Add 0
        Call CF2(Me, "OP_Item_IncluirGrid")

    Next

    Exit Sub

Erro_Limpa_gcolItemOp:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163860)

    End Select

    Exit Sub

End Sub

Private Function ProdutoLinha_Preenche(objProduto As ClassProduto, objItemOP As ClassItemOP) As Long

Dim iIndice As Integer
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sAlmoxarifadoPadrao As String
Dim bInseriuLinha As Boolean

On Error GoTo Erro_ProdutoLinha_Preenche

    bInseriuLinha = False
    
    'Unidade de Medida
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao

    If Len(Trim(objProduto.sGrade)) = 0 Then

        'Almoxarifado
        '(Utiliza Almoxarifado Padrão caso esteja preenchido)
        If Len(Trim(AlmoxPadrao.ClipText)) > 0 And Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col))) = 0 Then
        
            lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadrao)
            
            If lErro = SUCESSO Then
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
            Else
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = ""
            End If
        Else
    
            'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
            lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
            If lErro <> SUCESSO Then gError 52282
    
            'preenche o grid
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
    
        End If

    End If

    'Ccl
    lErro = CF("Ccl_Formata", CclPadrao.Text, sCclFormata, iCclPreenchida)
    If lErro <> SUCESSO Then gError 22940

    If iCclPreenchida = CCL_PREENCHIDA Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Ccl_Col) = CclPadrao.Text

    If Len(Trim(DataInicioPadrao.ClipText)) > 0 Then
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataPrevInicio_Col) = Format(DataInicioPadrao.Text, "dd/mm/yyyy")
    End If

    If Len(Trim(DataFimPadrao.ClipText)) > 0 Then
        If CDate(DataFimPadrao.Text) >= CDate(DataInicioPadrao.Text) Then
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataPrevFim_Col) = Format(DataFimPadrao.Text, "dd/mm/yyyy")
        Else
            GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_DataPrevFim_Col) = Format(DataInicioPadrao.Text, "dd/mm/yyyy")
        End If
    End If

    Situacao.ListIndex = ITEMOP_SITUACAO_NORMAL
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Situacao_Col) = Situacao.Text

    If DestinacaoPadrao.ListIndex <> -1 Then
        For iIndice = 0 To Destinacao.ListCount - 1
            If Destinacao.List(iIndice) = DestinacaoPadrao.Text Then
                Destinacao.ListIndex = iIndice
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = Destinacao.Text
                Exit For
            End If
        Next
    Else
        
        For iIndice = 0 To Destinacao.ListCount - 1
            If Destinacao.ItemData(iIndice) = ITEMOP_DESTINACAO_ESTOQUE Then
                Destinacao.ListIndex = iIndice
                Exit For
            End If
        Next

        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) = Destinacao.Text
    End If

    If Len(Trim(PrioridadePadrao.ClipText)) > 0 Then
        Prioridade.PromptInclude = False
        Prioridade.Text = PrioridadePadrao.Text
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Prioridade_Col) = PrioridadePadrao.Text
        Prioridade.PromptInclude = True
    End If
    
    'Call Combo_Seleciona_ItemData(Benef, BENEF_COMBO_DISP)
    Call Combo_Seleciona_ItemData(Benef, EscaninhoTerc.ItemData(EscaninhoTerc.ListIndex))
    
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col) = Benef.Text

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridMovimentos.Row - GridMovimentos.FixedRows) = objGrid.iLinhasExistentes Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        bInseriuLinha = True
        gcolItemOP.Add 0
        gobjOP.colItens.Add objItemOP
        Call CF2(Me, "OP_Item_IncluirGrid")
        Call gobjTelaProjetoInfo.Adiciona_Linha(GridMovimentos.Row)
        
        If Len(Trim(objProduto.sGrade)) = 0 Then

            gobjOP.colItens(GridMovimentos.Row).iPossuiGrade = DESMARCADO
            
        Else
        
            gobjOP.colItens(GridMovimentos.Row).iPossuiGrade = MARCADO
'            Set gobjOP.colItens(GridMovimentos.Row).colItensRomaneioGrade = objItemOP.colItensRomaneioGrade
            GridMovimentos.TextMatrix(GridMovimentos.Row, 0) = "# " & GridMovimentos.TextMatrix(GridMovimentos.Row, 0)
        
        End If
        
        gobjOP.colItens(GridMovimentos.Row).sSiglaUMEstoque = objProduto.sSiglaUMEstoque
        gobjOP.colItens(GridMovimentos.Row).iItem = GridMovimentos.Row
        gobjOP.colItens(GridMovimentos.Row).sProduto = objProduto.sCodigo
    
    End If

    Call CF2(Me, "OP_GridMov_AtualizaLinha", True)
    
    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 22940, 52282

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163861)

    End Select

    '#########################################
    'Inserido por Wagner - CROMATON 10/11/04
    'Limpa o lixo deixado quando ocorre um erro nessa função
    Call Limpa_Linha_GridMovimentos(GridMovimentos.Row, bInseriuLinha)
    '#########################################
    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = ""
    
    Exit Function

End Function

Private Function QuantDisponivel_Calcula(sProduto As String, sAlmoxarifado As String, Optional objProduto As ClassProduto) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim dQuantTotal As Double
Dim dQuantidade As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim lBenef As Long

On Error GoTo Erro_QuantDisponivel_Calcula

    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41350

    If GridMovimentos.Row >= GridMovimentos.FixedRows And Len(Trim(sAlmoxarifado)) <> 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        If objProduto Is Nothing Then

            Set objProduto = New ClassProduto

            objProduto.sCodigo = sProdutoFormatado

            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 41351

            If lErro = 28030 Then gError 41352

        End If

        objAlmoxarifado.sNomeReduzido = sAlmoxarifado

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25060 Then gError 41353

        If lErro = 25060 Then gError 41354

        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        objEstoqueProduto.sProduto = sProdutoFormatado

        'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 41355

        'Se não encontrou EstoqueProduto no Banco de Dados
        If lErro = 21306 Then

            QuantDisponivel.Caption = Formata_Estoque(0)

        Else
            sUnidadeMed = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col)

            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objProduto.sSiglaUMEstoque, sUnidadeMed, dFator)
            If lErro <> SUCESSO Then gError 41356
            
            Call Combo_Obtem_ItemData(Benef, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Benef_Col), lBenef)
            
            If lBenef = 1 Then
                QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantBenef3 * dFator)
            ElseIf lBenef = 2 Then
                QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantOutras3 * dFator)
            Else
                QuantDisponivel.Caption = Formata_Estoque(objEstoqueProduto.dQuantDisponivel * dFator)
            End If

        End If

    Else

        'Limpa a Quantidade Disponível da Tela
        QuantDisponivel.Caption = ""

    End If

    QuantDisponivel_Calcula = SUCESSO

    Exit Function

Erro_QuantDisponivel_Calcula:

    QuantDisponivel_Calcula = gErr

    Select Case gErr

        Case 41350, 41351, 41353, 41355, 41356

        Case 41352
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 41354
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163862)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ORDEM_PRODUCAO
    Set Form_Load_Ocx = Me
    Caption = "Ordem de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OrdemProducao"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub UpDownItemOP_DownClick()

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_UpDownItemOP_DownClick

    ItemOP.SetFocus

    If Len(ItemOP.ClipText) > 0 Then

        iItem = StrParaInt(ItemOP.Text)
        
        iItem = iItem + 1
        
        If iItem > objGrid.iLinhasExistentes Then
        
            iItem = objGrid.iLinhasExistentes
            
        End If
        
    Else
    
        iItem = 1
        
    End If

    ItemOP.PromptInclude = False
    ItemOP.Text = CStr(iItem)
    ItemOP.PromptInclude = True
    
    Call ItemOP_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownItemOP_DownClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163863)

    End Select

    Exit Sub


End Sub


Public Sub UpDownItemOP_UpClick()

Dim lErro As Long
Dim iItem As Integer

On Error GoTo Erro_UpDownItemOP_UpClick

    ItemOP.SetFocus

    If Len(Trim(ItemOP.ClipText)) > 0 Then

        iItem = StrParaInt(ItemOP.Text)
        
        iItem = iItem - 1
        
        If iItem < 1 Then
        
            iItem = 1
        
        End If
    
    Else
               
        iItem = 1
        
    End If

    ItemOP.PromptInclude = False
    ItemOP.Text = CStr(iItem)
    ItemOP.PromptInclude = True
    
    Call ItemOP_Validate(bSGECancelDummy)

    Exit Sub

Erro_UpDownItemOP_UpClick:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163864)

    End Select

    Exit Sub

End Sub


Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call CodigoOPLabel_Click
        ElseIf Me.ActiveControl Is CclPadrao Then
            Call CclPadraoLabel_Click
        ElseIf Me.ActiveControl Is Maquina Then
            Call BotaoMaquinas_Click
        ElseIf Me.ActiveControl Is AlmoxPadrao Then
            Call AlmoxPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call BotaoCcls_Click
        ElseIf Me.ActiveControl Is PedidoDeVendaId Then
            Call BotaoPedidoDeVenda_Click
        ElseIf Me.ActiveControl Is CodigoCompetencia Then
            Call CompetenciaLabel_Click
        ElseIf Me.ActiveControl Is CodigoCTPadrao Then
            Call CTLabel_Click
        ElseIf Me.ActiveControl Is NomeMaquinaReal Then
            Call BotaoMaquinasReal_Click
        ElseIf Me.ActiveControl Is TipoMaoDeObraReal Then
            Call BotaoMO_Click
        ElseIf Me.ActiveControl Is ClienteTerc Then
            Call ClienteLabel_Click
        ElseIf Me.ActiveControl Is FornecedorTerc Then
            Call FornecedorLabel_Click
        End If

        Call CF2(Me, "OP_UserControl_KeyDown", KeyCode, Shift)
        
        ElseIf KeyCode = KEYCODE_CODBARRAS Then
            Call Trata_CodigoBarras1
        
    End If

End Sub

Public Sub TabStrip1_Click()

Dim lErro As Long
Dim dQuantidade As Double

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index
        
        'Incluido por Jorge Specian - 25/05/2005
        '---------------------------------------
        'Se Frame selecionado foi o de Operacoes
        If TabStrip1.SelectedItem.Index = TAB_Operacoes Then
            
            If IsNumeric(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) Then
                dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
            Else
                dQuantidade = 0
            End If
        
            lErro = Mostra_Arvore(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), dQuantidade)
            If lErro <> SUCESSO Then gError 137133
            
        End If
        '---------------------------------------
        
    
        lErro = CF2(Me, "OP_TabStrip1_Click")
        If lErro <> SUCESSO Then gError 99999
    
    End If
    
    Exit Sub
    
Erro_TabStrip1_Click:

    Select Case gErr
    
        Case 137133
            'erro tratada na rotina chamada
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163865)
    
    End Select
    
    Exit Sub

End Sub

Public Sub AlmoxPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadrao_Validate

    'Se almoxarifado estivar vazio sai da rotina
    If Len(Trim(AlmoxPadrao.Text)) = 0 Then Exit Sub

    'Verifica existência do Almoxarifado
    lErro = TP_Almoxarifado_Filial_Le(AlmoxPadrao, objAlmoxarifado, 0)
    If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 22980

    If lErro = 25136 Then gError 22981

    If lErro = 25143 Then gError 22982

    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 22980

        Case 22981, 22982
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163866)

    End Select

    Exit Sub

End Sub

Public Sub Data_Validate(Cancel As Boolean)
'Critica se a Data da OP está preenchida corretamente
Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(Data.Text)
    If lErro <> SUCESSO Then gError 22933

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 22933

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163867)

    End Select

    Exit Sub

End Sub

Public Sub DataFimPadrao_Validate(Cancel As Boolean)
'Critica se a Data Fim Padrao da OP está preenchida corretamente

Dim lErro As Long

On Error GoTo Erro_DataFimPadrao_Validate

    If Len(Trim(DataFimPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(DataFimPadrao.Text)
    If lErro <> SUCESSO Then gError 55247

    Exit Sub

Erro_DataFimPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 55247

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163868)

    End Select

    Exit Sub

End Sub

Public Sub DataInicioPadrao_Validate(Cancel As Boolean)
'Critica se a Data de Inicio Padrao da OP está preenchida corretamente

Dim lErro As Long

On Error GoTo Erro_DataInicioPadrao_Validate

    If Len(Trim(DataInicioPadrao.ClipText)) = 0 Then Exit Sub

    lErro = Data_Critica(DataInicioPadrao.Text)
    If lErro <> SUCESSO Then gError 55246

    Exit Sub

Erro_DataInicioPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 55246

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163869)

    End Select

    Exit Sub

End Sub

Public Sub CclPadrao_Validate(Cancel As Boolean)
'verifica existência da Ccl informada

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_CclPadrao_Validate

    'se Ccl não estiver preenchida sai da rotina
    If Len(Trim(CclPadrao.Text)) = 0 Then Exit Sub

    lErro = CF("Ccl_Critica", CclPadrao.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError 31558

    If lErro = 5703 Then gError 31559

    Exit Sub

Erro_CclPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 31558

        Case 31559
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, CclPadrao.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163870)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    'Se houve alteração nos dados da tela
    If (iCodigoAlterado = REGISTRO_ALTERADO) Then

        If Len(Trim(Codigo.Text)) > 0 Then

            'limpa a coleção global
'            Call Limpa_gcolItemOP(gcolItemOP)

            objOrdemDeProducao.sCodigo = Codigo.Text
            objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa

            'tenta ler a OP desejada
            lErro = CF("OrdemProducao_Le", objOrdemDeProducao)
            If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 41306

            If lErro = SUCESSO And objOrdemDeProducao.iTipo = OP_TIPO_OC Then gError 117659

            'ordem de producao baixada
            If lErro = 55316 Then gError 55319

            'se existir
            If lErro = SUCESSO Then

                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")

                If vbMsg = vbNo Then gError 41307

                'traz a OP para a tela
                lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
                If lErro <> SUCESSO And lErro <> 21966 Then gError 21973

                Call ComandoSeta_Fechar(Me.Name)

            End If

        End If

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 21973
    
        Case 41306
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_ORDENSDEPRODUCAO", gErr)
    
        Case 41307
    
        Case 55319
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, Codigo.Text)

        Case 117659
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDECORTE", gErr, Codigo.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163871)

    End Select

    Exit Sub

End Sub

Public Sub GridMovimentos_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid)

End Sub



Private Sub Carrega_ComboVersoes(ByVal sProdutoRaiz As String)
    
Dim lErro As Long
Dim objKit As New ClassKit
Dim colKits As New Collection
Dim iPadrao As Integer
Dim iIndice As Integer
    
On Error GoTo Erro_Carrega_ComboVersoes
    
    Versao.Enabled = True
    
    'Limpa a Combo
    Versao.Clear
    
    'Armazena o Produto Raiz do kit
    objKit.sProdutoRaiz = sProdutoRaiz
    
    'Le as Versoes Ativas e a Padrao
    lErro = CF("Kit_Le_Produziveis", objKit, colKits)
    If lErro <> SUCESSO And lErro <> 106333 Then gError 106321
    
    iPadrao = -1
    
    'Carrega a Combo com os Dados da Colecao
    For Each objKit In colKits
    
        Versao.AddItem (objKit.sVersao)
        
        'Se for a padrao -> Armazena
        If objKit.iSituacao = KIT_SITUACAO_PADRAO Then iPadrao = iIndice
        
        iIndice = iIndice + 1
        
    Next
    
    If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)) > 0 Then
    
        Versao.Text = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)
    
    ElseIf iPadrao <> -1 Then
    
        'Seleciona a Padrao na Combo
        Versao.ListIndex = iPadrao
        
        GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col) = Versao.Text

    End If

    Exit Sub
    
Erro_Carrega_ComboVersoes:

    Select Case gErr
    
        Case 106321
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163872)
    
    End Select
    
End Sub

Private Function Saida_Celula_Versao(objGridInt As AdmGrid) As Long
'faz a critica da celula de Versao do grid que está deixando de ser a corrente

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProdutoFilial As New ClassProdutoFilial

On Error GoTo Erro_Saida_Celula_Versao

    Set objGridInt.objControle = Versao

    'Incluido por Jorge Specian - 15/06/2005
    '---------------------------------------
    If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) <> 0 Then
        
        lErro = Trata_Arvore(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), Versao.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)))
        If lErro <> SUCESSO Then gError 137124
        
    End If
    '---------------------------------------
    
    lErro = CF2(Me, "OP_Saida_Celula_Versao", objGridInt)
    If lErro <> SUCESSO Then gError 99999

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 106338

    Saida_Celula_Versao = SUCESSO

    Exit Function

Erro_Saida_Celula_Versao:

    Saida_Celula_Versao = gErr

    Select Case gErr

        Case 106338, 137124
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163873)

    End Select

End Function

Private Function Saida_Celula_Maquina(objGridInt As AdmGrid) As Long
'faz a critica da celula de Equipamento do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objMaquina As New ClassMaquinas

On Error GoTo Erro_Saida_Celula_Maquina

    Set objGridInt.objControle = Maquina

    'Se a Máquina foi especificada => Faz a Validacao da Máquina
    If Len(Trim(Maquina.Text)) > 0 Then
        
        'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", Maquina, objMaquina)
        If lErro <> SUCESSO Then gError 106341
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 106338

    Saida_Celula_Maquina = SUCESSO

    Exit Function

Erro_Saida_Celula_Maquina:

    Saida_Celula_Maquina = gErr

    Select Case gErr

        Case 106338, 106341
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163874)

    End Select

End Function

Public Sub Maquina_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Maquina_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Maquina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Maquina
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Versao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Versao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Versao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Versao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Versao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function OP_Executa_Relatorio(ByVal objCT As Object, ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
'Executa o(s) relatorio(s) de acordo com a selecao no frame de relatorios

Dim lErro As Long, lNumIntRel As Long
Dim objRelatorio1 As New AdmRelatorio, objRelatorio2 As New AdmRelatorio
Dim lNumIntRel2 As Long

On Error GoTo Erro_OP_Executa_Relatorio

    'Executa o Relatorio de OP caso o usuário tenha marcado somente o relatório de OP ou ambos
    If OpcaoRelatorio.Item(0).Value = True Or OpcaoRelatorio.Item(2).Value = True Then
    
        lErro = CF("ItensOPRel_Prepara", objOrdemDeProducao, lNumIntRel)
        If lErro <> SUCESSO Then gError 111828
    
        '#####################################
        'Alterado por Wagner
        'Imprime o Relatorio de OP
        If OpcaoSimples.Value = True Then
            lErro = objRelatorio1.ExecutarDireto("Ordens de Produção", "OrdemProducao = @TORDPROD", 0, "OPINPAL", "TORDPROD", objOrdemDeProducao.sCodigo, "NNUMINTREL", CStr(lNumIntRel))
        Else
            lErro = objRelatorio1.ExecutarDireto("Ordens de Produção Detalhado", "OrdemProducao = @TORDPROD", 0, "OPINPALD", "TORDPROD", objOrdemDeProducao.sCodigo, "NNUMINTREL", CStr(lNumIntRel))
        End If
        If lErro <> SUCESSO Then gError 106406
        '#####################################
    
    End If
    
    'Executa o Relatório de Rótulos caso o usuário tenha marcado somente o relatório de Rótulos ou ambos
    If OpcaoRelatorio.Item(1).Value = True Or OpcaoRelatorio.Item(2).Value = True Then
    
        lErro = CF("RelRotulo_Customiza", objOrdemDeProducao, lNumIntRel2)
        If lErro <> SUCESSO Then gError 189053
    
        'Imprime o Relatório de Rótulos
        lErro = objRelatorio2.ExecutarDireto("Rótulos para Ordens de Produção", "OrdemProducao = @TORDPROD", 0, "SOLROTUL", "TORDPROD", objOrdemDeProducao.sCodigo, "NNUMINTREL", CStr(lNumIntRel2))
        If lErro <> SUCESSO Then gError 106408
        
    End If
    
    OP_Executa_Relatorio = SUCESSO
    
    Exit Function
    
Erro_OP_Executa_Relatorio:

    OP_Executa_Relatorio = gErr
    
    Select Case gErr
    
        Case 106406, 106408, 111828, 189053
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163875)
    
    End Select

End Function

Private Function Grid_Possui_Grade() As Boolean

Dim iIndice As Integer

    For iIndice = 1 To objGrid.iLinhasExistentes
        If left(GridMovimentos.TextMatrix(iIndice, 0), 2) = "# " Then
            Grid_Possui_Grade = True
            Exit Function
        End If
    Next
    
    Grid_Possui_Grade = False
        
    Exit Function
    
End Function

Sub Atualiza_Grid_Movimentos(objItemOP As ClassItemOP)

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim dQuantidade As Double
Dim objItemRomaneioGrade As ClassItemRomaneioGrade
    
    For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
            
        dQuantidade = dQuantidade + objItemRomaneioGrade.dQuantidade
        
    Next

    GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)

    objItemOP.dQuantidade = dQuantidade

    Exit Sub

End Sub

Public Sub BotaoImprimirPrevia_Click()

Dim lErro As Long
Dim objCalcNecesProd As New ClassCalcNecesProd
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProduto As String, sCcl As String
Dim sProdutoFormatado As String, bAchou As Boolean
Dim objNecesProdInfo As ClassNecesProdInfo
Dim objItemOP As ClassItemOP
Dim objItemRomaneioGrade As ClassItemRomaneioGrade

On Error GoTo Erro_BotaoImprimirPrevia_Click

    bAchou = False
    
    'Para cada item do grid, guarda em um objeto os dados do grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        If left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then

            Set objNecesProdInfo = New ClassNecesProdInfo
    
            sProduto = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)

            'Critica o formato do Produto
            lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 124215
    
            objNecesProdInfo.sProduto = sProdutoFormatado
    
            objNecesProdInfo.sUMNecesInfo = GridMovimentos.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
            objNecesProdInfo.sVersao = GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col)
    
            If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
                objNecesProdInfo.dQuantNecesInfo = CDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))
                If objNecesProdInfo.dQuantNecesInfo > 0 Then bAchou = True
            Else
                objNecesProdInfo.dQuantNecesInfo = 0
            End If
        
            Set objNecesProdInfo.objItemOP = gobjOP.colItens(iIndice)
        
            objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
    
        Else
    
            Set objItemOP = gobjOP.colItens(iIndice)
            
            For Each objItemRomaneioGrade In objItemOP.colItensRomaneioGrade
            
                Set objNecesProdInfo = New ClassNecesProdInfo
        
                objNecesProdInfo.sProduto = objItemRomaneioGrade.sProduto
        
                objNecesProdInfo.sUMNecesInfo = objItemOP.sSiglaUM
                objNecesProdInfo.sVersao = GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col)
        
                If objItemRomaneioGrade.dQuantidade > 0 Then
                    objNecesProdInfo.dQuantNecesInfo = objItemRomaneioGrade.dQuantidade
                    If objNecesProdInfo.dQuantNecesInfo > 0 Then bAchou = True
                Else
                    objNecesProdInfo.dQuantNecesInfo = 0
                End If
            
                objCalcNecesProd.colNecesInfProd.Add objNecesProdInfo
            
            Next
    
        End If
        
    Next
    
    If bAchou = False Then gError 124216
    
    objCalcNecesProd.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("Producao_Calcula_Necessidades", objCalcNecesProd)
    If lErro <> SUCESSO Then gError 124217
    
    lErro = CF("Rel_Producao_Calcula_Necessidades", objCalcNecesProd)
    If lErro <> SUCESSO Then gError 124217
    
    Exit Sub
    
Erro_BotaoImprimirPrevia_Click:

    Select Case gErr
        
        Case 124215, 124217
        
        Case 124216
            Call Rotina_Erro(vbOKOnly, "ERRO_NAO_DEFINIU_QTDE_PROD", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163876)
     
    End Select
     
    Exit Sub

End Sub

'###################################################
'Inserido por Wagner
Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigoOP As String

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("OrdemProducao_Automatico", sCodigoOP, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 131860

    Codigo.Text = sCodigoOP

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 131860
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163877)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoProxNum2_Click()

Dim lErro As Long
Dim sCodigoOP As String

On Error GoTo Erro_BotaoProxNum2_Click

    'Mostra número do proximo lote disponível
    lErro = CF("OrdemProducao_Automatico", sCodigoOP, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 131875

    CodOPGerada.Text = sCodigoOP

    Exit Sub

Erro_BotaoProxNum2_Click:

    Select Case gErr

        Case 131875
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163878)
    
    End Select

    Exit Sub

End Sub
'###############################################################################

'Inserido por Jorge Specian - Alterações ref. ao PCP ...

Function Trata_Arvore(ByVal sProdutoRaiz As String, ByVal sVersao As String, ByVal sUMedida As String, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim bTrazDefault As Boolean
Dim objItemOP As New ClassItemOP
Dim objRoteirosDeFabricacao As ClassRoteirosDeFabricacao
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim objItOPcol As New ClassItemOP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iNivel As Integer, iNivelRoteiro As Integer, iSeqRoteiro As Integer, iSeq As Integer

On Error GoTo Erro_Trata_Arvore

    If gbOPValidaParaMRP = True And Len(Trim(sProdutoRaiz)) <> 0 Then
    
        lErro = CF("Produto_Formata", sProdutoRaiz, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137126
        
        bTrazDefault = True
        
        'verifica se o item da OP já está na coleção
        For Each objItemOP In colOPItens
        
            'se está ...
            If objItemOP.sProduto = sProdutoFormatado And _
                objItemOP.sVersao = sVersao Then
                
                'Se alterou a Quantidade
                If iQtdeAlterada = REGISTRO_ALTERADO Then
                    
                    'Muda a Quantidades da coleção de Operações do objItemOP
                    lErro = AlteraQuantidade(objItemOP, dQuantidade)
                    If lErro <> SUCESSO Then gError 137072
                    
                    'Altera o objItemOP guardado na coleção com a nova Quantidade
                    colOPItens.Remove (sProdutoFormatado)
                    colOPItens.Add objItemOP, sProdutoFormatado
                    
                    iQtdeAlterada = 0
                    
                End If
                
                bTrazDefault = False
                Exit For
                
            End If
            
        Next
        
        'se não está -> traz a árvore default
        If bTrazDefault Then
        
            If Len(Trim(sVersao)) > 0 Then
        
                objItemOP.sProduto = sProdutoFormatado
                objItemOP.sVersao = sVersao
                objItemOP.sSiglaUM = sUMedida
                objItemOP.dQuantidade = dQuantidade
            
                Set objRoteirosDeFabricacao = New ClassRoteirosDeFabricacao
            
                objRoteirosDeFabricacao.sProdutoRaiz = sProdutoFormatado
                objRoteirosDeFabricacao.sVersao = sVersao
                
                Set objOrdemProducaoInsumos = New ClassOrdemProducaoInsumos
                
                objOrdemProducaoInsumos.sUMProduto = sUMedida
                objOrdemProducaoInsumos.dQuantidade = dQuantidade
                
                'Gera a árvore de operações
                lErro = CF("OrdemDeProducao_Gera_Operacoes1", objRoteirosDeFabricacao, objOrdemProducaoInsumos, objItemOP, iNivel, iNivelRoteiro, iSeqRoteiro, iSeq)
                If lErro <> SUCESSO Then gError 137083
                
                'ajusta as sequencias dos itens gerados
                lErro = CF("OrdemDeProducao_Gera_Operacoes2", objItemOP)
                If lErro <> SUCESSO Then gError 139268
                
                'ajusta a coleção
                For Each objItOPcol In colOPItens
                    If objItOPcol.sProduto = sProdutoFormatado Then
                        colOPItens.Remove (sProdutoFormatado)
                        Exit For
                    End If
                Next
                colOPItens.Add objItemOP, sProdutoFormatado
                                
            End If
            
            iQtdeAlterada = 0
            
        End If
            
    End If
    
    Call CF2(Me, "OP_GridMov_AtualizaLinha")
        
    Trata_Arvore = SUCESSO
    
    Exit Function
    
Erro_Trata_Arvore:

    Trata_Arvore = gErr
    
    Select Case gErr
        
        Case 137072, 137083, 137126, 139268
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163879)
            
    End Select
    
    Exit Function

End Function

Function Mostra_Arvore(ByVal sProdutoRaiz As String, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim objItemOP As New ClassItemOP
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Mostra_Arvore

    'limpa Operacoes
    Call Limpa_Operacoes
    
    LabelCodigoCompetencia.Caption = ""
    LabelCodigoCTPadrao.Caption = ""
    
    'limpa a Arvore
    Call Limpa_Arvore_Roteiro
    
    'limpa MRP
    MRP.Value = vbUnchecked
    Call Limpa_MRP
    
    If GridMovimentos.Row = 0 Or GridMovimentos.Row > objGrid.iLinhasExistentes Then
    
        GridMovimentos.Row = 1
        sProdutoRaiz = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
        dQuantidade = StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))
    
    End If
            
    If gbOPValidaParaMRP = True And Len(Trim(sProdutoRaiz)) <> 0 Then
            
        If dQuantidade = 0 Then gError 137025
        
        If GridMovimentos.Row <> StrParaInt(ItemOP.Text) Then
        
            ItemOP.PromptInclude = False
            ItemOP.Text = GridMovimentos.Row
            ItemOP.PromptInclude = True
            
            giItemOP = GridMovimentos.Row
        
        End If
        
        lErro = CF("Produto_Formata", sProdutoRaiz, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137134
        
        'verifica se o item da OP está na coleção
        For Each objItemOP In colOPItens
        
            'se está ...
            If objItemOP.sProduto = sProdutoFormatado Then
                
                'Monta a árvore de operações
                lErro = Carrega_Arvore(objItemOP)
                If lErro <> SUCESSO Then gError 137082
                
                Exit For
                
            End If
            
        Next
                    
    End If
    
    Mostra_Arvore = SUCESSO
    
    Exit Function
    
Erro_Mostra_Arvore:

    Mostra_Arvore = gErr
    
    Select Case gErr
    
        Case 137025
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, GridMovimentos.Row)
               
        Case 137082, 137134
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163880)
            
    End Select
    
    Exit Function

End Function

Public Sub BotaoAlterar_Click()

Dim lErro As Long
Dim objNode As Node
Dim sChave As String
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objProduto As ClassProduto
Dim objItemOP As New ClassItemOP
Dim sTexto As String
Dim objOPOperacoes As New ClassOrdemProducaoOperacoes
Dim sProdutoRaiz As String

On Error GoTo Erro_BotaoAlterar_Click

    Set objNode = Roteiro.SelectedItem

    If objNode Is Nothing Then gError 137026
    If objNode.Selected = False Then gError 137027
    
    If Len(Trim(CodigoCompetencia.ClipText)) = 0 Then gError 137028
        
    Set objOrdemProducaoOperacoes = New ClassOrdemProducaoOperacoes
    Set objCompetencias = New ClassCompetencias
    Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
    Set objOrdemProducaoOperacoes = colComponentes.Item(objNode.Tag)
    
    'preenche objOrdemProducaoOperacoes à partir dos dados da tela
    lErro = Move_Operacoes_Memoria(objOrdemProducaoOperacoes, objCompetencias, objCentrodeTrabalho)
    If lErro <> SUCESSO Then gError 137029

    sChave = objNode.Tag
        
    'prepara texto que identificará a nova Operação que está sendo incluida
    sTexto = objCompetencias.sNomeReduzido
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objOrdemProducaoOperacoes.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137692
    
    sTexto = sTexto & " (" & objProduto.sNomeReduzido

    If Len(Trim(CodigoCTPadrao.ClipText)) <> 0 Then
       sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
    End If
        
    sTexto = sTexto & ")"

    objNode.Text = sTexto

    colComponentes.Remove (sChave)
    colComponentes.Add objOrdemProducaoOperacoes, sChave
    
    'Inserido por Wagner 05/05/2006
    'A Função está errada e não á necessidade dela aqui
    'Call Recalcula_Nivel_Sequencial
    
    For Each objItemOP In colOPItens
    
        If objItemOP.sProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) Then
                    
            sProdutoRaiz = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
            
            For Each objOPOperacoes In objItemOP.colOrdemProducaoOperacoes
            
                If objOPOperacoes.iSeqRoteiro = objOrdemProducaoOperacoes.iSeqRoteiro Then
                
                    objOPOperacoes.lNumIntDocCompet = objOrdemProducaoOperacoes.lNumIntDocCompet
                    objOPOperacoes.lNumIntDocCT = objOrdemProducaoOperacoes.lNumIntDocCT
                    objOPOperacoes.sObservacao = objOrdemProducaoOperacoes.sObservacao
                    objOPOperacoes.iNivel = objOrdemProducaoOperacoes.iNivel
                    objOPOperacoes.iSeq = objOrdemProducaoOperacoes.iSeq
                    objOPOperacoes.iSeqArvore = objOrdemProducaoOperacoes.iSeqArvore
                    objOPOperacoes.iSeqPai = objOrdemProducaoOperacoes.iSeqPai
                    objOPOperacoes.iNivelRoteiro = objOrdemProducaoOperacoes.iNivelRoteiro
                    objOPOperacoes.iSeqRoteiro = objOrdemProducaoOperacoes.iSeqRoteiro
                    objOPOperacoes.iSeqRoteiroPai = objOrdemProducaoOperacoes.iSeqRoteiroPai
                    objOPOperacoes.sProduto = objOrdemProducaoOperacoes.sProduto
                    objOPOperacoes.sVersao = objOrdemProducaoOperacoes.sVersao
                    objOPOperacoes.iConsideraCarga = objOrdemProducaoOperacoes.iConsideraCarga
                    
                    Set objOPOperacoes.colOPInsumos = objOrdemProducaoOperacoes.colOPInsumos
                
                    Exit For
            
                End If
                
            Next
            
            colOPItens.Remove (sProdutoRaiz)
            colOPItens.Add objItemOP, sProdutoRaiz
            
            Exit For
        
        End If
        
    Next

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_BotaoAlterar_Click:

    Select Case gErr

        Case 137026, 137027
            Call Rotina_Erro(vbOKOnly, "AVISO_SELECIONAR_ESTRUTURA_ROTEIRO", gErr)

        Case 137029, 137692
            'erro tratado na rotina chamada
        
        Case 137028
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_COMPETENCIA_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163881)

    End Select

    Exit Sub

End Sub

Public Sub CodigoCompetencia_Change()

    iAlterado = REGISTRO_ALTERADO
    iCompetenciaAlterada = REGISTRO_ALTERADO

End Sub

Public Sub CodigoCompetencia_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoCompetencia, iAlterado)
    
End Sub

Public Sub CodigoCompetencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCompetencias As ClassCompetencias
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho

On Error GoTo Erro_CodigoCompetencia_Validate

    DescricaoCompetencia.Caption = ""
    
    If GridMovimentos.Row = 0 Then Exit Sub

    'Verifica se CodigoCompetencia não está preenchida
    If Len(Trim(CodigoCompetencia.Text)) = 0 Then

        Exit Sub
    
    End If

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 137042

    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 137085

    Set objCompetencias = New ClassCompetencias
    
    'Verifica sua existencia
    lErro = CF("TP_Competencia_Le", CodigoCompetencia, objCompetencias)
    If lErro <> SUCESSO Then gError 137086
    
    DescricaoCompetencia.Caption = objCompetencias.sDescricao
    
    'Verifica se CodigoCompetencia foi alterado
    If iCompetenciaAlterada = REGISTRO_ALTERADO Then
            
        CodigoCTPadrao.Text = ""
        DescricaoCTPadrao.Caption = ""
        
        'Verifica se existe CTPadrao cadastrado na Competencia e traz seus dados
        lErro = CF("Competencias_Le_CTPadrao", objCompetencias, objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134909 Then gError 137087
        
        If lErro = SUCESSO Then
        
           CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
           
           Call CodigoCTPadrao_Validate(bSGECancelDummy)
        
        End If
        
        iCompetenciaAlterada = 0
    
    End If
       
    Exit Sub

Erro_CodigoCompetencia_Validate:

    Cancel = True

    Select Case gErr

        Case 137042, 137086, 137087
            'erro tratado na rotina chamada

        Case 137085
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_PRODUTORAIZKIT_NAO_PREENCHIDO3", gErr)
            CodigoCompetencia.Text = ""
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163882)

    End Select

    Exit Sub

End Sub

Public Sub CodigoCTPadrao_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub


Public Sub CodigoCTPadrao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodigoCTPadrao, iAlterado)
    
End Sub

Public Sub CodigoCTPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objCTCompetencias As New ClassCTCompetencias
Dim objCompetencias As ClassCompetencias
Dim bCompetenciaCadastrada As Boolean

On Error GoTo Erro_CodigoCTPadrao_Validate

    DescricaoCTPadrao.Caption = ""
    
    If GridMovimentos.Row = 0 Then Exit Sub

    'Verifica se CodigoCTPadrao não está preenchido
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then

        'Verifica se CodigoCompetencia está preenchido
        If Len(Trim(CodigoCompetencia.Text)) = 0 Then gError 137088
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = EMPRESA_TODA
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CodigoCTPadrao, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 137089
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.sNomeReduzido = CodigoCompetencia.Text
        
        'Lê a Competencia pelo NomeReduzido para verificar seu NumIntDoc
        lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134937 Then gError 137090
    
        If lErro <> SUCESSO Then gError 137092
        
        lErro = CF("CentrodeTrabalho_Le_CTCompetencias", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134453 Then gError 137091
    
        bCompetenciaCadastrada = False
        
        For Each objCTCompetencias In objCentrodeTrabalho.colCompetencias
        
            If objCTCompetencias.lNumIntDocCompet = objCompetencias.lNumIntDoc Then
            
                bCompetenciaCadastrada = True
                Exit For
                
            End If
        
        Next
            
        If bCompetenciaCadastrada = False Then gError 137093
            
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
        
    End If
       
    Exit Sub

Erro_CodigoCTPadrao_Validate:

    Cancel = True

    Select Case gErr

        Case 137089, 134090, 137091
            'erro tratado na rotina chamada
        
        Case 137088
            Call Rotina_Erro(vbOKOnly, "ERRO_OPERACAO_NAO_PREENCHIDA", gErr)
            CodigoCTPadrao.Text = ""
                        
        Case 137092, 137093
            Call Rotina_Erro(vbOKOnly, "ERRO_COMPETENCIA_NAO_CADASTRADA_CT", gErr, objCentrodeTrabalho.lCodigo)
            CodigoCTPadrao.SetFocus
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163883)

    End Select

    Exit Sub

End Sub

Public Sub CompetenciaLabel_Click()

Dim lErro As Long
Dim objCompetencias As New ClassCompetencias
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CompetenciaLabel_Click

    If GridMovimentos.Row = 0 Then Exit Sub
    
    If MRP.Value = vbChecked Then Exit Sub

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 137094

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        'Verifica se a Competencia foi preenchida
        If Len(Trim(CodigoCompetencia.Text)) <> 0 Then
                    
            objCompetencias.sNomeReduzido = CodigoCompetencia.Text
            
            'Verifica a Competencia no BD a partir do NomeReduzido
            lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
            If lErro <> SUCESSO And lErro <> 134937 Then gError 139063
    
        End If
    
        Call Chama_Tela("CompetenciasLista", colSelecao, objCompetencias, objEventoCompetencias)

    End If
    
    Exit Sub

Erro_CompetenciaLabel_Click:

    Select Case gErr
    
        Case 137094, 139063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163884)

    End Select

    Exit Sub

End Sub

Public Sub CTLabel_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_CTLabel

    If GridMovimentos.Row = 0 Then Exit Sub
    
    If MRP.Value = vbChecked Then Exit Sub

    lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 137095

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        'Verifica se o CodigoCTPadrao foi preenchido
        If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
            
            objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
            
            'Verifica o CodigoCTPadrao, lendo no BD a partir do NomeReduzido
            lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134941 Then gError 139066
            
        End If
    
        Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCentroDeTrabalho)

    End If

    Exit Sub

Erro_CTLabel:

    Select Case gErr
    
        Case 137095, 139066

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163885)

    End Select

    Exit Sub

End Sub

Public Sub MRP_Click()

    Call Habilita_MRP

End Sub

Private Sub objEventoCentroDeTrabalho_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCentrodeTrabalho = obj1

    CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CodigoCTPadrao_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163886)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCompetencias_evSelecao(obj1 As Object)

Dim objCompetencias As New ClassCompetencias
Dim lErro As Long

On Error GoTo Erro_objEventoCompetencia_evSelecao

    Set objCompetencias = obj1

    CodigoCompetencia.Text = objCompetencias.sNomeReduzido
    
    Call CodigoCompetencia_Validate(bSGECancelDummy)

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCompetencia_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163887)

    End Select

    Exit Sub

End Sub

Public Sub Roteiro_NodeClick(ByVal Node As MSComctlLib.Node)

Dim lErro As Long
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim objNode As Node
Dim sProdutoRaiz As String
Dim sVersao As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objPMPItens As New ClassPMPItens
Dim objPO As New ClassPlanoOperacional
Dim bTemMRP As Boolean

On Error GoTo Erro_Roteiro_NodeClick

    Set objNode = Roteiro.SelectedItem

    Set objOrdemProducaoOperacoes = colComponentes.Item(objNode.Tag)
    
    bTemMRP = False
    
    'Limpa a Tab MRP
    Call Limpa_MRP
        
    sProdutoRaiz = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)
    sVersao = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col)
    
    'Formata o produto do grid
    lErro = CF("Produto_Formata", sProdutoRaiz, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 137145
    
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
        'Para cada item da coleção ...
        For Each objPMPItens In colPMPItens
        
            If objOrdemProducaoOperacoes.lNumIntDocOperOrigem = 0 Then
            
'                'Localiza o produto selecionado
'                If objPMPItens.sProduto = sProdutoFormatado And objPMPItens.sVersao = sVersao Then
            
                    'Para cada Plano Operacional ...
                    For Each objPO In objPMPItens.colPO
                        
                        'Localiza o nivel/sequencia selecionada
                        If ((UCase(objPO.sCodOPOrigem) = UCase(Codigo.Text) And objPO.sProduto = sProdutoFormatado) Or (objPMPItens.sProduto = sProdutoFormatado And objPMPItens.sVersao = sVersao And objPMPItens.sCodOPOrigem = Codigo.Text)) And objPO.iNivel = objOrdemProducaoOperacoes.iNivel And objOrdemProducaoOperacoes.iSeqArvore = objPO.iSeq Then
                            bTemMRP = True
                            Exit For
                        End If
                    Next
                    Exit For
                'End If
            Else
        
                'Para cada Plano Operacional ...
                For Each objPO In objPMPItens.colPO
                    
                    'Localiza a Operação Origem
                    If objPO.lNumIntDocOper = objOrdemProducaoOperacoes.lNumIntDocOperOrigem Then
                        bTemMRP = True
                        Exit For
                    End If
                Next
            End If
            
        Next
    
    End If
    
    If bTemMRP Then
        
        'Preenche a Tab MRP
        lErro = Preenche_MRP(objOrdemProducaoOperacoes, objPO)
        If lErro <> SUCESSO Then gError 137146
    
    End If
    
    'Preenche as Operações
    lErro = Preenche_Operacoes(objOrdemProducaoOperacoes, objPO)
    If lErro <> SUCESSO Then gError 137096

    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Roteiro_NodeClick:

    Select Case gErr

        Case 137096, 137145, 137146
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163888)

    End Select

    Exit Sub

End Sub

Public Sub TabStrip2_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip2.SelectedItem.Index <> iFrameAtualOper Then

        If TabStrip_PodeTrocarTab(iFrameAtualOper, TabStrip2, Me) <> SUCESSO Then Exit Sub

        'Torna Frame correspondente ao Tab selecionado visivel
        Frame2(TabStrip2.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame2(iFrameAtualOper).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtualOper = TabStrip2.SelectedItem.Index
        
    End If

End Sub

Private Function Inicializa_GridOperacaoInsumos(objGrid As AdmGrid) As Long
'Inserido por Jorge Specian - 29/04/2005

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Produto")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("C/P")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("UM")
    objGrid.colColuna.Add ("Versão")
    objGrid.colColuna.Add ("Perda")
    objGrid.colColuna.Add ("Composição")
    objGrid.colColuna.Add ("C.Stand.")

    'Controles que participam do Grid
    objGrid.colCampo.Add (ProdutoInsumos.Name)
    objGrid.colCampo.Add (DescricaoProduto.Name)
    objGrid.colCampo.Add (OrigemProduto.Name)
    objGrid.colCampo.Add (QuantidadeProduto.Name)
    objGrid.colCampo.Add (UMProduto.Name)
    objGrid.colCampo.Add (VersaoKitComp.Name)
    objGrid.colCampo.Add (PerdaInsumos.Name)
    objGrid.colCampo.Add (ComposicaoInsumos.Name)
    objGrid.colCampo.Add (CustoStandardInsumos.Name)

    'Colunas do Grid
    iGrid_ProdutoInsumos_Col = 1
    iGrid_DescricaoProduto_Col = 2
    iGrid_OrigemProduto_Col = 3
    iGrid_QuantidadeProduto_Col = 4
    iGrid_UMProduto_Col = 5
    iGrid_VersaoKitComp_Col = 6
    iGrid_PerdaInsumos_Col = 7
    iGrid_ComposicaoInsumos_Col = 8
    iGrid_CustoStandardInsumos_Col = 9

    objGrid.objGrid = GridOperacaoInsumos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    If m_objUserControl.Height > 9000 Then
        objGrid.iLinhasVisiveis = 8
    Else
        objGrid.iLinhasVisiveis = 5
    End If

    'Largura da primeira coluna
    GridOperacaoInsumos.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridOperacaoInsumos = SUCESSO

End Function

Public Sub GridOperacaoInsumos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridOperacaoInsumos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOperacaoInsumos, iAlterado)
    End If

End Sub

Public Sub GridOperacaoInsumos_GotFocus()
    
    Call Grid_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub GridOperacaoInsumos_EnterCell()

    Call Grid_Entrada_Celula(objGridOperacaoInsumos, iAlterado)

End Sub

Public Sub GridOperacaoInsumos_LeaveCell()
    
    Call Saida_Celula(objGridOperacaoInsumos)

End Sub

Public Sub GridOperacaoInsumos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridOperacaoInsumos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridOperacaoInsumos, iAlterado)
    End If

End Sub

Public Sub GridOperacaoInsumos_Scroll()

    Call Grid_Scroll(objGridOperacaoInsumos)

End Sub

Public Sub DescricaoProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescricaoProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub DescricaoProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub DescricaoProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = DescricaoProduto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantidadeProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantidadeProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub QuantidadeProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub QuantidadeProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = QuantidadeProduto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UMProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UMProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub UMProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub UMProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = UMProduto
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub VersaoKitComp_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VersaoKitComp_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub VersaoKitComp_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub VersaoKitComp_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = VersaoKitComp
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub PerdaInsumos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PerdaInsumos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub PerdaInsumos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub PerdaInsumos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = PerdaInsumos
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ComposicaoInsumos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComposicaoInsumos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub ComposicaoInsumos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub ComposicaoInsumos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = ComposicaoInsumos
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CustoStandardInsumos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CustoStandardInsumos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub CustoStandardInsumos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub CustoStandardInsumos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = CustoStandardInsumos
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ProdutoInsumos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoInsumos_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridOperacaoInsumos)

End Sub

Public Sub ProdutoInsumos_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridOperacaoInsumos)

End Sub

Public Sub ProdutoInsumos_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridOperacaoInsumos.objControle = ProdutoInsumos
    lErro = Grid_Campo_Libera_Foco(objGridOperacaoInsumos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Carrega_Arvore(ByVal objItemOP As ClassItemOP) As Long
'preenche a treeview Roteiro com a composicao de objRoteirosDeFabricacao
   
Dim objNode As Node
Dim lErro As Long, sChave As String, sChaveTvw As String
Dim iIndice As Integer
Dim sTexto As String
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim objCompetencias As ClassCompetencias
Dim objProduto As ClassProduto
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim iIndiceInicial As Integer
Dim iCount As Integer

On Error GoTo Erro_Carrega_Arvore

    iCount = 0

    For Each objOrdemProducaoOperacoes In objItemOP.colOrdemProducaoOperacoes
    
        iCount = iCount + 1
        
        Set objCompetencias = New ClassCompetencias
        
        objCompetencias.lNumIntDoc = objOrdemProducaoOperacoes.lNumIntDocCompet
        
        lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
        If lErro <> SUCESSO And lErro <> 134336 Then gError 137101

        'prepara texto que identificará a nova Operação que está sendo incluida
        sTexto = objCompetencias.sNomeReduzido
        
        Set objProduto = New ClassProduto
        
        objProduto.sCodigo = objOrdemProducaoOperacoes.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 137429
        
        sTexto = sTexto & " (" & objProduto.sNomeReduzido

        If objOrdemProducaoOperacoes.lNumIntDocCT > 0 Then
        
            Set objCentrodeTrabalho = New ClassCentrodeTrabalho
            
            objCentrodeTrabalho.lNumIntDoc = objOrdemProducaoOperacoes.lNumIntDocCT
            
            lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
            If lErro <> SUCESSO And lErro <> 134590 Then gError 137102
            
            If lErro = SUCESSO Then
        
                sTexto = sTexto & " - " & objCentrodeTrabalho.sNomeReduzido
                
            End If
           
        End If
        
        sTexto = sTexto & ")"
        
        'prepara uma chave para relacionar colComponentes ao node que está sendo incluido
        Call Calcula_Proxima_Chave(sChaveTvw)
        
        sChave = sChaveTvw
        sChaveTvw = sChaveTvw & objCompetencias.lCodigo

        If objOrdemProducaoOperacoes.iNivel = 0 Or iCount = 1 Then

            Set objNode = Roteiro.Nodes.Add(, tvwFirst, sChaveTvw, sTexto)

            If iIndiceInicial = 0 Then
                iIndiceInicial = objOrdemProducaoOperacoes.iSeq - 1
            End If

        Else

            Set objNode = Roteiro.Nodes.Add(objOrdemProducaoOperacoes.iSeqPai - iIndiceInicial, tvwChild, sChaveTvw, sTexto)

        End If
                
        Roteiro.Nodes.Item(objNode.Index).Expanded = True
        
        colComponentes.Add objOrdemProducaoOperacoes, sChave
        
        objNode.Tag = sChave
        
    Next
    
    Call Habilita_Operacoes(Roteiro.Nodes.Count)
    
    Carrega_Arvore = SUCESSO

    Exit Function

Erro_Carrega_Arvore:

    Carrega_Arvore = gErr

    Select Case gErr

        Case 137101, 137102, 137429
            'erros tratados nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163889)

    End Select

    Exit Function

End Function

Function Preenche_Operacoes(objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes, objPO As ClassPlanoOperacional) As Long
'preenche as tabs de Detalhes, Insumos e Produção à partir dos dados de objOrdemProducaoOperacoes

Dim lErro As Long
Dim iAlteradoAnterior As Integer
Dim objProduto As ClassProduto
Dim sProdutoMascarado As String
Dim objCompetencias As ClassCompetencias
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_Preenche_Operacoes
    
    'Limpa as Tabs de Detalhes e Insumos
    lErro = Limpa_Operacoes()
    If lErro <> SUCESSO Then gError 137103
    
    iAlteradoAnterior = iAlterado

    Nivel.Caption = objOrdemProducaoOperacoes.iNivel
    Sequencial.Caption = objOrdemProducaoOperacoes.iSeqArvore
    
    Set objProduto = New ClassProduto
    
    objProduto.sCodigo = objOrdemProducaoOperacoes.sProduto
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 137533
    
    lErro = Mascara_RetornaProdutoTela(objProduto.sCodigo, sProdutoMascarado)
    If lErro <> SUCESSO Then gError 137534
    
    ProdutoLabel.Caption = sProdutoMascarado & SEPARADOR & objProduto.sDescricao
    VersaoLabel.Caption = objOrdemProducaoOperacoes.sVersao
    
    If MRP.Value = vbChecked Then
        QtdeLabel.Caption = Formata_Estoque(objPO.dQuantidade)
        UMLabel.Caption = objPO.sUM
    Else
        QtdeLabel.Caption = Formata_Estoque(objOrdemProducaoOperacoes.dQuantidade)
        UMLabel.Caption = objOrdemProducaoOperacoes.sUMedida
    End If
    
    Set objCompetencias = New ClassCompetencias
    
    objCompetencias.lNumIntDoc = objOrdemProducaoOperacoes.lNumIntDocCompet
    
    lErro = CF("Competencias_Le_NumIntDoc", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134336 Then gError 137104
    
    CodigoCompetencia.PromptInclude = False
    CodigoCompetencia.Text = objCompetencias.sNomeReduzido
    CodigoCompetencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = objCompetencias.sDescricao
    
    If objOrdemProducaoOperacoes.iNumMaxMaqPorOper <> 0 Then
        NumMaxMaqPorOper.PromptInclude = False
        NumMaxMaqPorOper.Text = objOrdemProducaoOperacoes.iNumMaxMaqPorOper
        NumMaxMaqPorOper.PromptInclude = True
    End If
    
    If objOrdemProducaoOperacoes.iNumRepeticoes <> 0 Then
        Repeticao.PromptInclude = False
        Repeticao.Text = objOrdemProducaoOperacoes.iNumRepeticoes
        Repeticao.PromptInclude = True
    End If
    
    If objOrdemProducaoOperacoes.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = objOrdemProducaoOperacoes.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 137105
        
        CodigoCTPadrao.PromptInclude = False
        CodigoCTPadrao.Text = objCentrodeTrabalho.sNomeReduzido
        CodigoCTPadrao.PromptInclude = True
        
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If
    
    Observacao.Text = objOrdemProducaoOperacoes.sObservacao
    
    If objOrdemProducaoOperacoes.iConsideraCarga = MARCADO Then
        OptionMRPII.Value = True
    Else
        OptionMRP.Value = True
    End If
    
    If objOrdemProducaoOperacoes.iOrigem = ORIGEM_KIT Then
    
        OptionMRP.Enabled = False
        OptionMRPII.Enabled = False
        FrameConsideraAlgoritmo.Enabled = False
    
    Else
    
        OptionMRP.Enabled = True
        OptionMRPII.Enabled = True
        FrameConsideraAlgoritmo.Enabled = True
    
    End If
    
    lErro = Preenche_GridOperacaoInsumos(objOrdemProducaoOperacoes)
    If lErro <> SUCESSO Then gError 137106
    
    '##############################################################
    'Inserido por Wagner 08/11/2005
    lErro = Preenche_GridMaquinasReal(objOrdemProducaoOperacoes)
    If lErro <> SUCESSO Then gError 140821
    '##############################################################
    
    Call AlteraHabilitacao_Operacoes
    
    iCompetenciaAlterada = 0
    iAlterado = iAlteradoAnterior

    Preenche_Operacoes = SUCESSO
    
    Exit Function

Erro_Preenche_Operacoes:

    Preenche_Operacoes = gErr

    Select Case gErr
    
        Case 137103 To 137106, 137533, 137534, 140821
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163890)

    End Select

    Exit Function

End Function

Function Limpa_Operacoes() As Long

Dim lErro As Long
Dim iAlteradoAnterior As Integer

On Error GoTo Erro_Limpa_Operacoes
        
    iAlteradoAnterior = iAlterado
    
    Nivel.Caption = ""
    Sequencial.Caption = ""
    
    ProdutoLabel.Caption = ""
    VersaoLabel.Caption = ""
    QtdeLabel.Caption = ""
    UMLabel.Caption = ""
    
    CodigoCompetencia.PromptInclude = False
    CodigoCompetencia.Text = ""
    CodigoCompetencia.PromptInclude = True
    
    DescricaoCompetencia.Caption = ""
    
    CodigoCTPadrao.PromptInclude = False
    CodigoCTPadrao.Text = ""
    CodigoCTPadrao.PromptInclude = True
    
    DescricaoCTPadrao.Caption = ""
    
    Observacao.Text = ""
    
    Call Grid_Limpa(objGridOperacaoInsumos)
    
    '######################################
    'Inserido por Wagner 08/11/2005
    Call Grid_Limpa(objGridMaquinasReal)
    Call Grid_Limpa(objGridMOReal)
    '######################################
        
    iCompetenciaAlterada = 0
    iAlterado = iAlteradoAnterior

    Limpa_Operacoes = SUCESSO

    Exit Function

Erro_Limpa_Operacoes:

    Limpa_Operacoes = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163891)

    End Select

    Exit Function

End Function

Private Function Preenche_GridOperacaoInsumos(objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objProdutos As ClassProduto
Dim sProdutoMascarado As String

On Error GoTo Erro_Preenche_GridOperacaoInsumos
    
    Call Grid_Limpa(objGridOperacaoInsumos)
    
    'Exibe os dados da coleção na tela
    For iIndice = 1 To objOrdemProducaoOperacoes.colOPInsumos.Count
        
        Set objProdutos = New ClassProduto
        
        objProdutos.sCodigo = objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).sProduto
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 137107
        
        lErro = Mascara_RetornaProdutoTela(objProdutos.sCodigo, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 137108
                
        'Insere no GridOrdemProducaoInsumos
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_ProdutoInsumos_Col) = sProdutoMascarado
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_DescricaoProduto_Col) = objProdutos.sDescricao
        
        If objProdutos.iCompras = PRODUTO_COMPRAVEL Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_OrigemProduto_Col) = INSUMO_COMPRADO
        Else
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_OrigemProduto_Col) = INSUMO_PRODUZIDO
        End If
        
        If objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).dQuantidade > 0 Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_QuantidadeProduto_Col) = Formata_Estoque(objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).dQuantidade)
        End If
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_UMProduto_Col) = objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).sUMProduto
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_VersaoKitComp_Col) = objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).sVersaoKitComp
        If objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).dPercentualPerda > 0 Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_PerdaInsumos_Col) = Format(objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).dPercentualPerda, "Percent")
        End If
        Call Combo_Seleciona_ItemData(ComposicaoInsumos, objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).iComposicao)
        GridOperacaoInsumos.TextMatrix(iIndice, iGrid_ComposicaoInsumos_Col) = ComposicaoInsumos.Text
        If objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).dCustoStandard > 0 Then
            GridOperacaoInsumos.TextMatrix(iIndice, iGrid_CustoStandardInsumos_Col) = Format(objOrdemProducaoOperacoes.colOPInsumos.Item(iIndice).dCustoStandard, "Standard")
        End If

    Next

    objGridOperacaoInsumos.iLinhasExistentes = objOrdemProducaoOperacoes.colOPInsumos.Count
    
    Preenche_GridOperacaoInsumos = SUCESSO
    
    Exit Function

Erro_Preenche_GridOperacaoInsumos:

    Preenche_GridOperacaoInsumos = gErr

    Select Case gErr

        Case 137107, 137108

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163892)

    End Select

    Exit Function

End Function

Function Limpa_Arvore_Roteiro() As Long
'Limpa a Arvore do Roteiro

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Arvore_Roteiro

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 137109

    Roteiro.Nodes.Clear
    Set colComponentes = New Collection
    
    iProxChave = 1

    Limpa_Arvore_Roteiro = SUCESSO

    Exit Function

Erro_Limpa_Arvore_Roteiro:

    Limpa_Arvore_Roteiro = gErr
    
    Select Case gErr

        Case 137109

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163893)

    End Select

    Exit Function

End Function

Sub Calcula_Proxima_Chave(sChave As String)

Dim iNumero As Integer

    iNumero = iProxChave
    iProxChave = iProxChave + 1
    sChave = "X" & right$(CStr(100000 + iNumero), 5)

End Sub

Private Function Move_Operacoes_Memoria(ByVal objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes, ByVal objCompetencias As ClassCompetencias, ByVal objCentrodeTrabalho As ClassCentrodeTrabalho) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Operacoes_Memoria
        
    objCompetencias.sNomeReduzido = CodigoCompetencia.Text
    
    'Verifica a Competencia no BD a partir do Código
    lErro = CF("Competencias_Le_NomeReduzido", objCompetencias)
    If lErro <> SUCESSO And lErro <> 134937 Then gError 137110

    objOrdemProducaoOperacoes.lNumIntDocCompet = objCompetencias.lNumIntDoc
    
    If Len(Trim(CodigoCTPadrao.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137111
        
        objOrdemProducaoOperacoes.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Trim(Observacao.Text)) <> 0 Then objOrdemProducaoOperacoes.sObservacao = Observacao.Text
    
    objOrdemProducaoOperacoes.iConsideraCarga = IIf(OptionMRPII.Value = True, MARCADO, DESMARCADO)
    
    lErro = Move_OperacaoInsumos_Memoria(objOrdemProducaoOperacoes)
    If lErro <> SUCESSO Then gError 137112
    
    '##############################################
    'Inserido por Wagner
    lErro = Move_MaquinasReal_Memoria(objOrdemProducaoOperacoes)
    If lErro <> SUCESSO Then gError 140819
    '##############################################
    
    objOrdemProducaoOperacoes.iNumMaxMaqPorOper = StrParaInt(NumMaxMaqPorOper.Text)
    objOrdemProducaoOperacoes.iNumRepeticoes = StrParaInt(Repeticao.Text)
                    
    Move_Operacoes_Memoria = SUCESSO

    Exit Function

Erro_Move_Operacoes_Memoria:

    Move_Operacoes_Memoria = gErr

    Select Case gErr

        Case 137110 To 137112, 140819
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163894)

    End Select

    Exit Function

End Function

Private Function Move_OperacaoInsumos_Memoria(objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim objOrdemProducaoInsumos As ClassOrdemProducaoInsumos
Dim objProdutos As ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iIndice As Integer

On Error GoTo Erro_Move_OperacaoInsumos_Memoria

    'Limpa coleção para reiniciar
    Set objOrdemProducaoOperacoes.colOPInsumos = New Collection

    'Ir preenchendo a colecao no objOperacoes com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridOperacaoInsumos.iLinhasExistentes

        'Se o Item não estiver preenchido caio fora
        If Len(Trim(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_Produto_Col))) = 0 Then Exit For
                
        Set objProdutos = New ClassProduto
        
        lErro = CF("Produto_Formata", GridOperacaoInsumos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 137113
        
        objProdutos.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProdutos)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 137114
        
        Set objOrdemProducaoInsumos = New ClassOrdemProducaoInsumos
        
        objOrdemProducaoInsumos.lNumIntDocOper = objOrdemProducaoOperacoes.lNumIntDoc
        objOrdemProducaoInsumos.sProduto = objProdutos.sCodigo
        objOrdemProducaoInsumos.dQuantidade = StrParaDbl(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_QuantidadeProduto_Col))
        objOrdemProducaoInsumos.sUMProduto = GridOperacaoInsumos.TextMatrix(iIndice, iGrid_UMProduto_Col)
        objOrdemProducaoInsumos.iComposicao = Composicao_Extrai(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_ComposicaoInsumos_Col))
        objOrdemProducaoInsumos.dPercentualPerda = StrParaDbl(Val(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_PerdaInsumos_Col)) / 100)
        objOrdemProducaoInsumos.dCustoStandard = StrParaDbl(GridOperacaoInsumos.TextMatrix(iIndice, iGrid_CustoStandardInsumos_Col))
        objOrdemProducaoInsumos.sVersaoKitComp = GridOperacaoInsumos.TextMatrix(iIndice, iGrid_VersaoKitComp_Col)
    
        objOrdemProducaoOperacoes.colOPInsumos.Add objOrdemProducaoInsumos
    
    Next

    Move_OperacaoInsumos_Memoria = SUCESSO

    Exit Function

Erro_Move_OperacaoInsumos_Memoria:

    Move_OperacaoInsumos_Memoria = gErr

    Select Case gErr
            
        Case 137113, 137114
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163895)

    End Select

    Exit Function

End Function

Sub Recalcula_Nivel_Sequencial()
'(re)calcula niveis e sequencias de toda a estrutura
'deve ser chamada apos a remocao de algum node

Dim iIndice As Integer

    If Roteiro.Nodes.Count = 0 Then Exit Sub

    For iIndice = LBound(aNivelSequencial) To UBound(aNivelSequencial)
        aNivelSequencial(iIndice) = 0
    Next

    iUltimoNivel = 0

    'chamar rotina que recalcula recursivamente os campos nivel e sequencial (Nivel e SeqArvore)
    Call Calcula_Nivel_Sequencial(Roteiro.Nodes.Item(1), 0, 0)

End Sub

Sub Calcula_Nivel_Sequencial(objNode As Node, iNivel As Integer, iPosicaoAtual As Integer)
'parte recursiva do recalculo de nivel e sequencial, atuando a partir do node passado
'iNivel informa o nivel deste node

Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim sChave1 As String

    sChave1 = objNode.Tag

    Set objOrdemProducaoOperacoes = colComponentes.Item(sChave1)

    aNivelSequencial(iNivel) = aNivelSequencial(iNivel) + 1

    iPosicaoAtual = iPosicaoAtual + 1
    aSeqPai(iNivel) = iPosicaoAtual

    objOrdemProducaoOperacoes.iSeqArvore = aNivelSequencial(iNivel)

    If iNivel > 0 Then
        objOrdemProducaoOperacoes.iSeqRoteiroPai = aSeqPai(iNivel - 1)
    Else
        objOrdemProducaoOperacoes.iSeqRoteiroPai = 0
    End If
    
    objOrdemProducaoOperacoes.iSeqRoteiro = iPosicaoAtual

    objOrdemProducaoOperacoes.iNivelRoteiro = iNivel
    
    colComponentes.Remove sChave1
    colComponentes.Add objOrdemProducaoOperacoes, sChave1

    If objNode.Children > 0 Then
        Call Calcula_Nivel_Sequencial(objNode.Child, iNivel + 1, iPosicaoAtual)
    End If

    If objNode.Index <> objNode.LastSibling.Index Then Call Calcula_Nivel_Sequencial(objNode.Next, iNivel, iPosicaoAtual)

    If iNivel > iUltimoNivel Then iUltimoNivel = iNivel
   
End Sub

Private Function Composicao_Extrai(sComposicaoGrid As String) As Integer

Dim iIndice As Integer

    For iIndice = 0 To ComposicaoInsumos.ListCount - 1
        If ComposicaoInsumos.List(iIndice) = sComposicaoGrid Then
            Composicao_Extrai = iIndice
            Exit For
        End If
    Next

End Function


Private Function Habilita_MRP() As Long
        
Dim lErro As Long

On Error GoTo Erro_Habilita_MRP

    If MRP.Value = vbChecked Then
    
        LabelMRP.Enabled = True
        
        LabelDataInicio.Enabled = True
        
        LabelDataFinal.Enabled = True
        
        LabelOP.Enabled = True
        OPCodigoMRP.Enabled = True
        
        BotaoAbrirOP.Enabled = False 'é False mesmo... quem vai habilitá-lo é onde preenche com Código da SubOP <> Código da OP
        BotaoMaquinaPMP.Enabled = True
        BotaoCompetencia.Enabled = True
        BotaoCT.Enabled = True
        BotaoGrafico.Enabled = True
        
        LabelMaquinas.Enabled = True
                    
    Else
    
        LabelMRP.Enabled = False
    
        LabelDataInicio.Enabled = False
        DataInicio.Caption = ""
        
        LabelDataFinal.Enabled = False
        DataFinal.Caption = ""
        
        LabelOP.Enabled = False
        OPCodigoMRP.Text = ""
        OPCodigoMRP.Enabled = False
        
        BotaoAbrirOP.Enabled = False
        BotaoMaquinaPMP.Enabled = False
        BotaoCompetencia.Enabled = False
        BotaoCT.Enabled = False
        BotaoGrafico.Enabled = False
        
        LabelMaquinas.Enabled = False
        
        Call Grid_Limpa(objGridMaquinas)
            
    End If
    
    Habilita_MRP = SUCESSO
    
    Exit Function
    
Erro_Habilita_MRP:

    Habilita_MRP = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163896)

    End Select

    Exit Function

End Function

Public Sub GridMaquinas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaquinas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
    End If

End Sub

Public Sub GridMaquinas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaquinas)

End Sub

Public Sub GridMaquinas_EnterCell()

    Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)

End Sub

Public Sub GridMaquinas_LeaveCell()
    
    Call Saida_Celula(objGridMaquinas)

End Sub

Public Sub GridMaquinas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridMaquinas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinas, iAlterado)
    End If

End Sub

Public Sub GridMaquinas_Scroll()

    Call Grid_Scroll(objGridMaquinas)

End Sub

Private Function Inicializa_GridMaquinas(objGrid As AdmGrid) As Long
'Inserido por Jorge Specian - 10/05/2005

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Data")
    objGrid.colColuna.Add ("Máquina")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Horas")
    objGrid.colColuna.Add ("Tx.Prod.")

    'Controles que participam do Grid
    objGrid.colCampo.Add (DataPO.Name)
    objGrid.colCampo.Add (NomeMaquina.Name)
    objGrid.colCampo.Add (QuantidadeMaquina.Name)
    objGrid.colCampo.Add (HorasPO.Name)
    objGrid.colCampo.Add (TaxaProducao.Name)

    'Colunas do Grid
    iGrid_DataPO_Col = 1
    iGrid_NomeMaquina_Col = 2
    iGrid_QuantidadeMaquina_Col = 3
    iGrid_HorasPO_Col = 4
    iGrid_TaxaProducao_Col = 5

    objGrid.objGrid = GridMaquinas

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    If m_objUserControl.Height > 9000 Then
        objGrid.iLinhasVisiveis = 5
    Else
        objGrid.iLinhasVisiveis = 4
    End If

    'Largura da primeira coluna
    GridMaquinas.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGrid.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGrid.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaquinas = SUCESSO

End Function

Public Sub NomeMaquina_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Public Sub NomeMaquina_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Public Sub NomeMaquina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = NomeMaquina
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Public Sub QuantidadeMaquina_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Public Sub QuantidadeMaquina_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Public Sub QuantidadeMaquina_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = QuantidadeMaquina
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TaxaProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Public Sub TaxaProducao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Public Sub TaxaProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = TaxaProducao
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function AlteraQuantidade(ByVal objItemOP As ClassItemOP, ByVal dQuantidade As Double) As Long

Dim lErro As Long
Dim dFatorQtdeOP As Double
Dim objOrdemProducaoOperacoes As New ClassOrdemProducaoOperacoes
Dim objOrdemProducaoInsumos As New ClassOrdemProducaoInsumos
Dim objOperacoesTempo As ClassOperacoesTempo
Dim dTempoCorrigido As Double
Dim dQuantidadeCorrigida As Double

On Error GoTo Erro_AlteraQuantidade

    dFatorQtdeOP = dQuantidade / objItemOP.dQuantidade
                            
    For Each objOrdemProducaoOperacoes In objItemOP.colOrdemProducaoOperacoes
                
        If objOrdemProducaoOperacoes.iIgnoraTaxaProducao = MARCADO Then
                                    
            Set objOperacoesTempo = objOrdemProducaoOperacoes.objOperacoesTempo
            
            If objOperacoesTempo.iTipo <> ITEM_TIPO_TAXAPRODUCAO_FIXO Then
            
                dTempoCorrigido = objOperacoesTempo.dTempoOperacao * dFatorQtdeOP
                objOperacoesTempo.dTempoOperacao = dTempoCorrigido
            
            End If
            
            Set objOrdemProducaoOperacoes.objOperacoesTempo = objOperacoesTempo
        
        End If
        
        For Each objOrdemProducaoInsumos In objOrdemProducaoOperacoes.colOPInsumos
        
            If objOrdemProducaoInsumos.iComposicao = PRODUTOKIT_COMPOSICAO_VARIAVEL Then
                
                dQuantidadeCorrigida = objOrdemProducaoInsumos.dQuantidade * dFatorQtdeOP
                objOrdemProducaoInsumos.dQuantidade = dQuantidadeCorrigida
            
            End If
            
        Next
                
    Next
    
    objItemOP.dQuantidade = dQuantidade
    
    AlteraQuantidade = SUCESSO
    
    Exit Function

Erro_AlteraQuantidade:

    AlteraQuantidade = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163897)
    
    End Select
    
    Exit Function

End Function

'###############################################################################
'Inserido por Wagner
Function Preenche_MRP(objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes, objPO As ClassPlanoOperacional) As Long
'preenche as tabs de Detalhes, Insumos e Produção à partir dos dados de objOrdemProducaoOperacoes

Dim lErro As Long
Dim objPOMaquinas As ClassPOMaquinas
Dim iIndice As Integer
Dim objTaxa As ClassTaxaDeProducao
Dim objMaquina As ClassMaquinas
Dim dParteFracionaria As Double
Dim dMinutos As Double

On Error GoTo Erro_Preenche_MRP

    MRP.Value = vbChecked
    
    Call Grid_Limpa(objGridMaquinas)
    
    Call Habilita_MRP
    
    Call AlteraHabilitacao_Operacoes
    
    DataInicio.Caption = Format(objPO.dtDataInicio, "dd/mm/yyyy")
    
    DataFinal.Caption = Format(objPO.dtDataFim, "dd/mm/yyyy")

    OPCodigoMRP.Text = objPO.sCodOPOrigem
        
    If OPCodigoMRP.Text <> Codigo.Text Then
        BotaoAbrirOP.Enabled = True
    End If
        
    iIndice = 0
    
    For Each objPOMaquinas In objPO.colAlocacaoMaquinas
    
        iIndice = iIndice + 1
        
        If objOrdemProducaoOperacoes.iIgnoraTaxaProducao = MARCADO Then
        
            Set objTaxa = New ClassTaxaDeProducao
            
            objTaxa.dLoteMax = objOrdemProducaoOperacoes.objOperacoesTempo.dLoteMax
            objTaxa.dLoteMin = objOrdemProducaoOperacoes.objOperacoesTempo.dLoteMin
            objTaxa.dLotePadrao = objOrdemProducaoOperacoes.objOperacoesTempo.dLotePadrao
            objTaxa.dQuantidade = objPO.dQuantidade
            objTaxa.dTempoDescarga = objOrdemProducaoOperacoes.objOperacoesTempo.dTempoDescarga
            objTaxa.dTempoMovimentacao = objOrdemProducaoOperacoes.objOperacoesTempo.dTempoMovimentacao
            objTaxa.dTempoOperacao = objOrdemProducaoOperacoes.objOperacoesTempo.dTempoOperacao
            objTaxa.dTempoPreparacao = objOrdemProducaoOperacoes.objOperacoesTempo.dTempoPreparacao
            objTaxa.iTipo = objOrdemProducaoOperacoes.objOperacoesTempo.iTipo
            objTaxa.sUMProduto = objPO.sUM
            objTaxa.sUMTempo = objOrdemProducaoOperacoes.objOperacoesTempo.sUMTempo
            objTaxa.lNumIntDocCompet = objOrdemProducaoOperacoes.lNumIntDocCompet
            objTaxa.sProduto = objPO.sProduto
        
        Else
        
            Set objTaxa = objPOMaquinas.objTaxaProducao
                
            objTaxa.lNumIntDoc = objPOMaquinas.lNumIntDocTxProd
            
            lErro = CF("TaxaDeProducao_Le_NumIntDoc", objTaxa)
            If lErro <> SUCESSO And lErro <> 137151 Then gError 137152
        
        End If
        
        Set objMaquina = New ClassMaquinas
        
        objMaquina.lNumIntDoc = objPOMaquinas.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 136409
    
        GridMaquinas.TextMatrix(iIndice, iGrid_NomeMaquina_Col) = objMaquina.sNomeReduzido
        GridMaquinas.TextMatrix(iIndice, iGrid_QuantidadeMaquina_Col) = CStr(objPOMaquinas.iQuantidade)
        If objPOMaquinas.dtData <> DATA_NULA Then
            GridMaquinas.TextMatrix(iIndice, iGrid_DataPO_Col) = objPOMaquinas.dtData
        End If
        If objPOMaquinas.dHorasMaquina <> 0 Then
            GridMaquinas.TextMatrix(iIndice, iGrid_HorasPO_Col) = Formata_Estoque(objPOMaquinas.dHorasMaquina)
        End If
        If objTaxa.iTipo = ITEM_TIPO_TAXAPRODUCAO_FIXO Then
            GridMaquinas.TextMatrix(iIndice, iGrid_TaxaProducao_Col) = Formata_Estoque(objTaxa.dTempoOperacao) & " /" & objTaxa.sUMTempo
        Else
            GridMaquinas.TextMatrix(iIndice, iGrid_TaxaProducao_Col) = Formata_Estoque(objTaxa.dQuantidade / objTaxa.dTempoOperacao) & " " & objTaxa.sUMProduto & "/" & objTaxa.sUMTempo
        End If
    
    Next
    
    objGridMaquinas.iLinhasExistentes = iIndice
    
    Preenche_MRP = SUCESSO

    Exit Function

Erro_Preenche_MRP:

    Preenche_MRP = gErr

    Select Case gErr
    
        Case 136409, 137152
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163898)

    End Select

    Exit Function

End Function

Function Limpa_MRP() As Long
'Limpa os dados relativos a simulação do MRP

Dim lErro As Long

On Error GoTo Erro_Limpa_MRP

    MRP.Value = vbUnchecked
    MRP.Enabled = False

    DataInicio.Caption = ""
    
    DataFinal.Caption = ""

    OPCodigoMRP.Text = ""
        
    Call Grid_Limpa(objGridMaquinas)
    
    Limpa_MRP = SUCESSO

    Exit Function

Erro_Limpa_MRP:

    Limpa_MRP = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163899)

    End Select

    Exit Function

End Function

Function Habilita_Operacoes(ByVal lArvore As Long) As Long

Dim lErro As Long

On Error GoTo Erro_Habilita_Operacoes

    'se houver árvore ...
    If lArvore > 0 Then
        
        'selecionar a raiz
        Set Roteiro.SelectedItem = Roteiro.Nodes.Item(1)
        Roteiro.SelectedItem.Selected = True
        
        'e carregar as operações pertinentes
        Call Roteiro_NodeClick(Roteiro.Nodes.Item(1))
        
        LabelDetProduto.Enabled = True
        LabelDetVersao.Enabled = True
        LabelDetQtde.Enabled = True
        LabelDetUM.Enabled = True
        
        BotaoAlterar.Enabled = True
        BotaoImprimirOper.Enabled = True
        FrameItemOP.Enabled = True
        ItemOP.Enabled = False
        UpDownItemOP.Enabled = False
        'só habilita se tiver mais de um item na OP
        If objGrid.iLinhasExistentes > 1 Then
            ItemOP.Enabled = True
            UpDownItemOP.Enabled = True
        End If
        
        CompetenciaLabel.Enabled = True
        CodigoCompetencia.Enabled = True
        
        CTLabel.Enabled = True
        CodigoCTPadrao.Enabled = True
        
        LabelObservacao.Enabled = True
        Observacao.Enabled = True
        
        LabelNivel.Enabled = True
        
        LabelSeq.Enabled = True
        
        LabelRepeticao.Enabled = True
        Repeticao.Enabled = True
        
        LabelNumMaxMaqPorOper.Enabled = True
        NumMaxMaqPorOper.Enabled = True
        
    Else
    
        LabelDetProduto.Enabled = False
        LabelDetVersao.Enabled = False
        LabelDetQtde.Enabled = False
        LabelDetUM.Enabled = False
    
        BotaoAlterar.Enabled = False
        BotaoImprimirOper.Enabled = False
        FrameItemOP.Enabled = False
        ItemOP.Enabled = False
        UpDownItemOP.Enabled = False
        
        CompetenciaLabel.Enabled = False
        CodigoCompetencia.Text = ""
        CodigoCompetencia.Enabled = False
        
        DescricaoCompetencia.Caption = ""
        
        CTLabel.Enabled = False
        CodigoCTPadrao.Text = ""
        CodigoCTPadrao.Enabled = False
        
        DescricaoCTPadrao.Caption = ""
        
        LabelObservacao.Enabled = False
        Observacao.Text = ""
        Observacao.Enabled = False
        
        LabelNivel.Enabled = False
        Nivel.Caption = ""
        
        LabelSeq.Enabled = False
        Sequencial.Caption = ""
        
        OptionMRP.Enabled = False
        OptionMRPII.Value = True
        OptionMRPII.Enabled = False
        FrameConsideraAlgoritmo.Enabled = False
        
        LabelRepeticao.Enabled = False
        Repeticao.Enabled = False
        
        LabelNumMaxMaqPorOper.Enabled = False
        NumMaxMaqPorOper.Enabled = False
        
    End If
    
    Call AlteraHabilitacao_Operacoes
    
    Habilita_Operacoes = SUCESSO
    
    Exit Function
    
Erro_Habilita_Operacoes:
    
    Habilita_Operacoes = gErr
    
    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163900)

    End Select
    
    Exit Function

End Function

Function AlteraHabilitacao_Operacoes() As Long

Dim lErro As Long

On Error GoTo Erro_AlteraHabilitacao_Operacoes

    If MRP.Value = vbChecked Then
    
        CompetenciaLabel.ForeColor = &H80000012
        CompetenciaLabel.MousePointer = 0        'Default
        CodigoCompetencia.Enabled = False
        CodigoCompetencia.Visible = False
        
        With LabelCodigoCompetencia
        
            .left = 1380
            .top = 945
            .Width = 2445
            .Height = 315
            .Caption = CodigoCompetencia.Text
            .Visible = True
        
        End With
        
        CTLabel.ForeColor = &H80000012
        CTLabel.MousePointer = 0                 'Default
        CodigoCTPadrao.Enabled = False
        CodigoCTPadrao.Visible = False
        
        With LabelCodigoCTPadrao
        
            .left = 1380
            .top = 1320
            .Width = 2445
            .Height = 315
            .Caption = CodigoCTPadrao.Text
            .Visible = True
            
        End With
                
        OptionMRP.Enabled = False
        OptionMRPII.Enabled = False
        FrameConsideraAlgoritmo.Enabled = False
        NumMaxMaqPorOper.Enabled = False

    Else
    
        CompetenciaLabel.ForeColor = &H80&
        CompetenciaLabel.MousePointer = 14       'Arrow and Question
        LabelCodigoCompetencia.Caption = ""
        LabelCodigoCompetencia.Visible = False
        
        CTLabel.ForeColor = &H80&
        CTLabel.MousePointer = 14                'Arrow and Question
        LabelCodigoCTPadrao.Caption = ""
        LabelCodigoCTPadrao.Visible = False
    
        If Roteiro.Nodes.Count > 0 Then
            
            CodigoCompetencia.Enabled = True
            CodigoCompetencia.Visible = True
        
            CodigoCTPadrao.Enabled = True
            CodigoCTPadrao.Visible = True
    
        End If
        
        NumMaxMaqPorOper.Enabled = True
    
    End If
    
    AlteraHabilitacao_Operacoes = SUCESSO
    
    Exit Function
    
Erro_AlteraHabilitacao_Operacoes:

    AlteraHabilitacao_Operacoes = gErr
    
    Select Case gErr
    
    End Select
    
    Exit Function

End Function

Public Sub HorasPO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Public Sub HorasPO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Public Sub HorasPO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = HorasPO
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataPO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinas)

End Sub

Public Sub DataPO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinas)

End Sub

Public Sub DataPO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinas.objControle = DataPO
    lErro = Grid_Campo_Libera_Foco(objGridMaquinas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Roteiro_Collapse(ByVal Node As MSComctlLib.Node)
    Roteiro_NodeClick Node
End Sub

Public Sub BotaoImprimirOper_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimirOper_Click

    If Len(Trim(Codigo.Text)) <> 0 Then

        lErro = objRelatorio.ExecutarDireto("Ordem de Produção - Operações", "", 0, "", "TCODOP", Codigo.Text)
        If lErro <> SUCESSO Then gError 138252
    
    End If
    
    Exit Sub
    
Erro_BotaoImprimirOper_Click:
    
    Select Case gErr
    
        Case 138252
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163901)
    
    End Select
    
    Exit Sub
    
End Sub

Public Function Atualiza_Cronograma(objTelaGrafico As ClassTelaGrafico) As Long
'Acerta a tela de cronograma gráfico após o retorno da tela de etapas
'Para isso remonta o objTelaGrafico com os novos dados

Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objTelaGraficoItem As New ClassTelaGraficoItens
Dim objPMPItem As ClassPMPItens
Dim objPO As ClassPlanoOperacional
Dim objPOAux As ClassPlanoOperacional
Dim iIndice As Integer
Dim iCont As Integer
Dim sNL As String
Dim objCT As ClassCentrodeTrabalho
Dim lErro As Long
Dim bPrimeira As Boolean
Dim sProdutoAnterior As String

On Error GoTo Erro_Atualiza_Cronograma

    Set objTelaGrafico.colBotoes = New Collection
    Set objTelaGrafico.colItens = New Collection
    Set objTelaGrafico.colParametros = New Collection

    Set objTelaGrafico.objTela = Me
    
    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 139062
    
    objTelaGrafico.sNomeTela = "Cronograma das Operações do Item da OP"
    objTelaGrafico.iTamanhoDia = 540
    objTelaGrafico.iModal = DESMARCADO
    objTelaGrafico.iAtualizaRetornoClick = DESMARCADO

    iIndice = 0
    
    sNL = Chr(13) & Chr(10)

    'Para cada item do Plano Mestre
    For Each objPMPItem In colPMPItens
        
        If objPMPItem.sProduto = sProdutoFormatado And _
            sProdutoAnterior <> sProdutoFormatado Or objPMPItem.sCodOPOrigem <> Codigo.Text Then
        
            sProdutoAnterior = sProdutoFormatado
        
            iIndice = iIndice + 1
        
            iCont = 0
        
            'Para cada etapa
            For Each objPO In objPMPItem.colPO
            
                iCont = iCont + 1
            
                Set objTelaGraficoItem = New ClassTelaGraficoItens
                Set objCT = New ClassCentrodeTrabalho
                            
                bPrimeira = True
                
                For Each objPOAux In objPMPItem.colPO
                    'Se existe outra com data de início menor ou igual desde que
                    'o nó esteja depois na estrutura de árvore então o PO corrente
                    'não é o primeiro
                    If (objPOAux.dtDataInicio < objPO.dtDataInicio) Or (objPOAux.dtDataInicio = objPO.dtDataInicio And ((objPOAux.iNivel > objPO.iNivel) Or (objPOAux.iSeq > objPO.iSeq))) Then
                        bPrimeira = False
                    End If
                Next
                
                'Se for a primeira é a final em termos de data
                If iCont = 1 Then
                    objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_FIM
                End If
                           
                If bPrimeira Then
                    
                    If iCont = 1 Then
                        objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO_E_FIM
                    Else
                        objTelaGraficoItem.iIcone = TELA_GRAFICO_ICONE_INICIO
                    End If
                    
                End If
                
                objCT.lNumIntDoc = objPO.lNumIntDocCT
                
                'Le o centro de trabalho
                lErro = CF("CentrodeTrabalho_Le_NumIntDoc", objCT)
                If lErro <> SUCESSO And lErro <> 134590 Then gError 139063
                        
                objTelaGraficoItem.dtDataFim = objPO.dtDataFim
                objTelaGraficoItem.dtDataInicio = objPO.dtDataInicio
                objTelaGraficoItem.sTextoExibicao = "OP Origem: " & objPMPItem.sCodOPOrigem & sNL & "OP: " & objPO.sCodOPOrigem & sNL & "Prioridade: " & objPMPItem.iPrioridade & sNL & "Data Início: " & Format(objPO.dtDataInicio, "dd/mm/yyyy") & sNL & "Data Fim: " & Format(objPO.dtDataFim, "dd/mm/yyyy") & sNL & "CT: " & objCT.sNomeReduzido
                objTelaGraficoItem.sNome = objPO.sCodOPOrigem
                objTelaGraficoItem.iIndiceCor = iIndice
                
                objTelaGrafico.colItens.Add objTelaGraficoItem
                
            Next
        
        End If
        
    Next
    
    Atualiza_Cronograma = SUCESSO

    Exit Function

Erro_Atualiza_Cronograma:

    Atualiza_Cronograma = gErr

    Select Case gErr
    
        Case 139062, 139063

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163902)

    End Select
    
    Exit Function
    
End Function

'################################################################
'Inserido por Wagner
Private Function Inicializa_GridMaquinasReal(objGrid As AdmGrid) As Long
'Inserido por Jorge Specian - 10/05/2005

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Data")
    objGrid.colColuna.Add ("Máquina")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Horas")

    'Controles que participam do Grid
    objGrid.colCampo.Add (DataReal.Name)
    objGrid.colCampo.Add (NomeMaquinaReal.Name)
    objGrid.colCampo.Add (QuantidadeMaquinaReal.Name)
    objGrid.colCampo.Add (HorasReal.Name)

    'Colunas do Grid
    iGrid_DataReal_Col = 1
    iGrid_NomeMaquinaReal_Col = 2
    iGrid_QuantidadeMaquinaReal_Col = 3
    iGrid_HorasReal_Col = 4

    objGrid.objGrid = GridMaquinasReal

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    If m_objUserControl.Height > 9000 Then
        objGrid.iLinhasVisiveis = 8
    Else
        objGrid.iLinhasVisiveis = 5
    End If

    'Largura da primeira coluna
    GridMaquinasReal.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridMaquinasReal = SUCESSO

End Function

Public Sub NomeMaquinaReal_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NomeMaquinaReal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinasReal)

End Sub

Public Sub NomeMaquinaReal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinasReal)

End Sub

Public Sub NomeMaquinaReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinasReal.objControle = NomeMaquinaReal
    lErro = Grid_Campo_Libera_Foco(objGridMaquinasReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantidadeMaquinaReal_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantidadeMaquinaReal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinasReal)

End Sub

Public Sub QuantidadeMaquinaReal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinasReal)

End Sub

Public Sub QuantidadeMaquinaReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinasReal.objControle = QuantidadeMaquinaReal
    lErro = Grid_Campo_Libera_Foco(objGridMaquinasReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub HorasReal_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub HorasReal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinasReal)

End Sub

Public Sub HorasReal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinasReal)

End Sub

Public Sub HorasReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinasReal.objControle = HorasReal
    lErro = Grid_Campo_Libera_Foco(objGridMaquinasReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataReal_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataReal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMaquinasReal)

End Sub

Public Sub DataReal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMaquinasReal)

End Sub

Public Sub DataReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMaquinasReal.objControle = DataReal
    lErro = Grid_Campo_Libera_Foco(objGridMaquinasReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Inicializa_GridMOReal(objGrid As AdmGrid) As Long
'Inserido por Jorge Specian - 10/05/2005

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Data")
    objGrid.colColuna.Add ("Máquina")
    objGrid.colColuna.Add ("Operador")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Quantidade")
    objGrid.colColuna.Add ("Horas")

    'Controles que participam do Grid
    objGrid.colCampo.Add (DataRealMO.Name)
    objGrid.colCampo.Add (NomeMaquinaRealMO.Name)
    objGrid.colCampo.Add (TipoMaoDeObraReal.Name)
    objGrid.colCampo.Add (DescricaoTipoMO.Name)
    objGrid.colCampo.Add (QuantidadeMaquinaRealMO.Name)
    objGrid.colCampo.Add (HorasRealMO.Name)

    'Colunas do Grid
    iGrid_DataRealMO_Col = 1
    iGrid_NomeMaquinaRealMO_Col = 2
    iGrid_TipoMaoDeObraReal_Col = 3
    iGrid_DescricaoTipoMO_Col = 4
    iGrid_QuantidadeMaquinaRealMO_Col = 5
    iGrid_HorasRealMO_Col = 6

    objGrid.objGrid = GridMOReal

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    If m_objUserControl.Height > 9000 Then
        objGrid.iLinhasVisiveis = 8
    Else
        objGrid.iLinhasVisiveis = 5
    End If

    'Largura da primeira coluna
    GridMOReal.ColWidth(0) = 250

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL
    
    Call Grid_Inicializa(objGrid)

    Inicializa_GridMOReal = SUCESSO

End Function

Public Sub NomeMaquinaRealMO_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NomeMaquinaRealMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOReal)

End Sub

Public Sub NomeMaquinaRealMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOReal)

End Sub

Public Sub NomeMaquinaRealMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOReal.objControle = NomeMaquinaRealMO
    lErro = Grid_Campo_Libera_Foco(objGridMOReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantidadeMaquinaRealMO_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantidadeMaquinaRealMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOReal)

End Sub

Public Sub QuantidadeMaquinaRealMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOReal)

End Sub

Public Sub QuantidadeMaquinaRealMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOReal.objControle = QuantidadeMaquinaRealMO
    lErro = Grid_Campo_Libera_Foco(objGridMOReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub HorasRealMO_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub HorasRealMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOReal)

End Sub

Public Sub HorasRealMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOReal)

End Sub

Public Sub HorasRealMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOReal.objControle = HorasRealMO
    lErro = Grid_Campo_Libera_Foco(objGridMOReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DataRealMO_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataRealMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOReal)

End Sub

Public Sub DataRealMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOReal)

End Sub

Public Sub DataRealMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOReal.objControle = DataRealMO
    lErro = Grid_Campo_Libera_Foco(objGridMOReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub TipoMaodeObraReal_change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoMaodeObraReal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOReal)

End Sub

Private Sub TipoMaodeObraReal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOReal)

End Sub

Private Sub TipoMaodeObraReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOReal.objControle = TipoMaoDeObraReal
    lErro = Grid_Campo_Libera_Foco(objGridMOReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub DescricaoTipoMO_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescricaoTipoMO_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridMOReal)

End Sub

Public Sub DescricaoTipoMO_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMOReal)

End Sub

Public Sub DescricaoTipoMO_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridMOReal.objControle = DescricaoTipoMO
    lErro = Grid_Campo_Libera_Foco(objGridMOReal)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoMO_Click()

Dim lErro As Long
Dim objTiposDeMaodeObras As New ClassTiposDeMaodeObra
Dim colSelecao As New Collection
Dim sFiltro1 As String
Dim sFiltro2 As String

On Error GoTo Erro_BotaoMO_Click

    If Me.ActiveControl Is TipoMaoDeObraReal Then
    
        objTiposDeMaodeObras.iCodigo = StrParaInt(TipoMaoDeObraReal.Text)
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridMOReal.Row = 0 Then gError 140778

        objTiposDeMaodeObras.iCodigo = StrParaInt(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_TipoMaoDeObraReal_Col))
        
    End If
    
    'Tem que estar no CT
    sFiltro1 = "Codigo IN (SELECT CTO.TipoMaoDeObra FROM CentrodeTrabalho AS CT,  CTOperadores AS CTO WHERE CTO.NumIntDocCT = CT.NumIntDoc AND CT.NomeReduzido = ?)"
    
    'Tem que ser da máquina
    sFiltro2 = "Codigo IN (SELECT MO.TipoMaoDeObra FROM CentrodeTrabalho AS CT, CTMaquinas AS CTM, Maquinas AS M, MaquinaOperadores AS MO WHERE CTM.NumIntDocCT = CT.NumIntDoc AND M.NumIntDoc = CTM.NumIntDocMaq AND MO.NumIntDocMaq = M.NumIntDoc AND CT.NomeReduzido = ? AND M.NomeReduzido = ?)"
    
    colSelecao.Add CodigoCTPadrao.Text
    colSelecao.Add CodigoCTPadrao.Text
    colSelecao.Add GridMOReal.TextMatrix(GridMOReal.Row, iGrid_NomeMaquinaRealMO_Col)

    Call Chama_Tela("TiposDeMaodeObraLista", colSelecao, objTiposDeMaodeObras, objEventoMO, sFiltro1 & " AND " & sFiltro2)

    Exit Sub

Erro_BotaoMO_Click:

    Select Case gErr
        
        Case 140778
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163903)

    End Select

    Exit Sub

End Sub

Private Sub objEventoMO_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTiposDeMaodeObra As ClassTiposDeMaodeObra
Dim iLinha As Integer
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMO As ClassCTOperadores
Dim bAchou As Boolean
Dim iIndice As Integer
Dim dtData As Date
Dim sMaquina As String

On Error GoTo Erro_objEventoMO_evSelecao

    Set objTiposDeMaodeObra = obj1

    iLinha = GridMOReal.Row
    dtData = StrParaDate(GridMOReal.TextMatrix(iLinha, iGrid_DataRealMO_Col))
    sMaquina = GridMOReal.TextMatrix(iLinha, iGrid_NomeMaquinaRealMO_Col)

    objCT.sNomeReduzido = CodigoCTPadrao.Text
    objCT.iFilialEmpresa = giFilialEmpresa
        
     'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
    If lErro <> SUCESSO And lErro <> 137212 Then gError 141044

    bAchou = False
    For Each objCTMO In objCT.colOperadores
        If objCTMO.iCodTipoMO = objTiposDeMaodeObra.iCodigo Then
            bAchou = True
            Exit For
        End If
    Next

    If Not bAchou Then gError 141045
            
    For iIndice = 1 To objGridMOReal.iLinhasExistentes
        If iIndice <> iLinha Then
            'Para mesma data
            If dtData = StrParaDate(GridMOReal.TextMatrix(iIndice, iGrid_DataRealMO_Col)) Then
                'Mesma Máquina
                If sMaquina = GridMOReal.TextMatrix(iIndice, iGrid_NomeMaquinaRealMO_Col) Then
                    'Mesmo Operador
                    If objTiposDeMaodeObra.iCodigo = StrParaInt(GridMOReal.TextMatrix(iIndice, iGrid_TipoMaoDeObraReal_Col)) Then
                        gError 141051
                    End If
                End If
            End If
        End If
    Next
    
    TipoMaoDeObraReal.Text = objTiposDeMaodeObra.iCodigo
    
    If Not (Me.ActiveControl Is TipoMaoDeObraReal) Then

        GridMOReal.TextMatrix(GridMOReal.Row, iGrid_TipoMaoDeObraReal_Col) = CStr(objTiposDeMaodeObra.iCodigo)
        GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DescricaoTipoMO_Col) = objTiposDeMaodeObra.sDescricao

    End If
    
    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoMO_evSelecao:

    Select Case gErr
    
        Case 141043

        Case 141044
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_PERTENTECE_CT", gErr, objTiposDeMaodeObra.iCodigo, objCT.lCodigo, objCT.iFilialEmpresa)

        Case 141051
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_REPETIDO", gErr, objTiposDeMaodeObra.iCodigo, dtData, sMaquina, iIndice)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163904)

    End Select

    Exit Sub

End Sub

Public Sub BotaoMaquinasReal_Click()

Dim lErro As Long
Dim objMaquinas As New ClassMaquinas
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoMaquinasReal_Click

    If GridMaquinasReal.Row <= objGridMaquinasReal.iLinhasExistentes Then

        If Me.ActiveControl Is NomeMaquinaReal Then
                
            objMaquinas.sNomeReduzido = NomeMaquinaReal.Text
            
        Else
        
            'Verifica se tem alguma linha selecionada no Grid
            If GridMaquinasReal.Row = 0 Then gError 140779
    
            objMaquinas.sNomeReduzido = GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_NomeMaquinaReal_Col)
            
        End If
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 140780
        
        colSelecao.Add CodigoCTPadrao.Text
        
        Call Chama_Tela("MaquinasLista", colSelecao, objMaquinas, objEventoMaquinas, "NumIntdoc IN (SELECT CTM.NumIntDocMaq FROM CTMaquinas AS CTM, CentrodeTrabalho AS CT WHERE CT.NumIntDoc = CTM.NumIntDocCT AND CT.NomeReduzido = ?)")

    End If
    
    Exit Sub

Erro_BotaoMaquinasReal_Click:

    Select Case gErr

        Case 140779
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 140780

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163905)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoMaquinas_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objMaquinas As ClassMaquinas
Dim iLinha As Integer
Dim objCTMaquina As ClassCTMaquinas
Dim objCT As New ClassCentrodeTrabalho
Dim bAchou As Boolean
Dim iIndice As Integer
Dim dtData As String

On Error GoTo Erro_objEventoMaquinas_evSelecao

    Set objMaquinas = obj1

    'Lê o Maquinas
    lErro = CF("TP_Maquina_Le", NomeMaquinaReal, objMaquinas)
    If lErro <> SUCESSO Then gError 140781
                
    objCT.sNomeReduzido = CodigoCTPadrao.Text
    objCT.iFilialEmpresa = giFilialEmpresa
    
    'Lê o CentrodeTrabalho que está sendo Passado
    lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
    If lErro <> SUCESSO And lErro <> 137212 Then gError 141054
    
    bAchou = False
    
    For Each objCTMaquina In objCT.colMaquinas
        If objMaquinas.lNumIntDoc = objCTMaquina.lNumIntDocMaq Then
            bAchou = True
        End If
    Next
    
    If Not bAchou Then gError 141055
    
    iLinha = GridMaquinasReal.Row
    dtData = StrParaDate(GridMaquinasReal.TextMatrix(iLinha, iGrid_DataReal_Col))
    
    For iIndice = 1 To objGridMaquinasReal.iLinhasExistentes
        If iIndice <> iLinha Then
            'Se para mesma data
            If dtData = StrParaDate(GridMaquinasReal.TextMatrix(iIndice, iGrid_DataReal_Col)) Then
                'Tem a mesma máquina, erro
                If objMaquinas.sNomeReduzido = GridMaquinasReal.TextMatrix(iIndice, iGrid_NomeMaquinaReal_Col) Then
                    gError 141056
                End If
            End If
        End If
    Next
        
    'Mostra os dados do Maquinas na tela
    GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_NomeMaquinaReal_Col) = objMaquinas.sNomeReduzido
    
    NomeMaquinaReal.Text = objMaquinas.sNomeReduzido
       
    iAlterado = REGISTRO_ALTERADO
    
    Me.Show

    Exit Sub

Erro_objEventoMaquinas_evSelecao:

    Select Case gErr

        Case 140781
            'erro tratado na rotina chamada
        
        Case 141054
        
        Case 141055
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINAS_NAO_CADASTRADA_CT", gErr, objCT.lCodigo, objCT.iFilialEmpresa)
            
        Case 141056
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA2", gErr, objMaquinas.sNomeReduzido, iIndice, dtData)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163906)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_DataReal(objGridInt As AdmGrid) As Long
'faz a critica da celula de data do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataReal

    Set objGridInt.objControle = DataReal

    'verifica se a data está preenchida
    If Len(Trim(DataReal.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataReal.Text)
        If lErro <> SUCESSO Then gError 140792

        'verifica se precisa preencher o grid com uma nova linha
        If GridMaquinasReal.Row - GridMaquinasReal.FixedRows = objGridMaquinasReal.iLinhasExistentes Then
            objGridMaquinasReal.iLinhasExistentes = objGridMaquinasReal.iLinhasExistentes + 1
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140793

    Saida_Celula_DataReal = SUCESSO

    Exit Function

Erro_Saida_Celula_DataReal:

    Saida_Celula_DataReal = gErr

    Select Case gErr

        Case 140792, 140793
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163907)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NomeMaquinaReal(objGridInt As AdmGrid) As Long
'Faz a crítica da célula CodigoItem do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim objMaquinas As New ClassMaquinas
Dim objCTMaquina As ClassCTMaquinas
Dim objCT As New ClassCentrodeTrabalho
Dim bAchou As Boolean
Dim iIndice As Integer
Dim dtData As String

On Error GoTo Erro_Saida_Celula_NomeMaquinaReal

    Set objGridInt.objControle = NomeMaquinaReal

    'Se o campo foi preenchido
    If Len(Trim(NomeMaquinaReal.Text)) > 0 Then
    
         'Verifica sua existencia
        lErro = CF("TP_Maquina_Le", NomeMaquinaReal, objMaquinas)
        If lErro <> SUCESSO Then gError 140796
            
        objCT.sNomeReduzido = CodigoCTPadrao.Text
        objCT.iFilialEmpresa = giFilialEmpresa
        
         'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137212 Then gError 141052
        
        bAchou = False
        
        For Each objCTMaquina In objCT.colMaquinas
            If objMaquinas.lNumIntDoc = objCTMaquina.lNumIntDocMaq Then
                bAchou = True
            End If
        Next
        
        If Not bAchou Then gError 141053
        
        iLinha = GridMaquinasReal.Row
        dtData = StrParaDate(GridMaquinasReal.TextMatrix(iLinha, iGrid_DataReal_Col))
        
        For iIndice = 1 To objGridMaquinasReal.iLinhasExistentes
            If iIndice <> iLinha Then
                'Se para mesma data
                If dtData = StrParaDate(GridMaquinasReal.TextMatrix(iIndice, iGrid_DataReal_Col)) Then
                    'Tem a mesma máquina, erro
                    If objMaquinas.sNomeReduzido = GridMaquinasReal.TextMatrix(iIndice, iGrid_NomeMaquinaReal_Col) Then
                        gError 141057
                    End If
                End If
            End If
        Next
    
        GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_NomeMaquinaReal_Col) = objMaquinas.sNomeReduzido
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140797

    Saida_Celula_NomeMaquinaReal = SUCESSO

    Exit Function

Erro_Saida_Celula_NomeMaquinaReal:

    Saida_Celula_NomeMaquinaReal = gErr

    Select Case gErr
        
        Case 140796, 140797, 141052
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 141053
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINAS_NAO_CADASTRADA_CT", gErr, objCT.lCodigo, objCT.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 141057
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_REPETIDA2", gErr, objMaquinas.sNomeReduzido, iIndice, dtData)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163908)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HorasReal(objGridInt As AdmGrid) As Long
'faz a critica da celula de Horas do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim dHoras As Double
Dim iQtd As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As New ClassMaquinas
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_HorasReal

    Set objGridInt.objControle = HorasReal

    'verifica se a Horas está preenchida
    If Len(Trim(HorasReal.Text)) > 0 Then

        'verifica se a Horas é válida
        lErro = Valor_Positivo_Critica(HorasReal.Text)
        If lErro <> SUCESSO Then gError 140798
        
        HorasReal.Text = Formata_Estoque(StrParaDbl(HorasReal.Text))

        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 140987
        
        objCTMaquinas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
        
        objMaquinas.sNomeReduzido = GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_NomeMaquinaReal_Col)
    
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 140988
        
        objCTMaquinas.lNumIntDocMaq = objMaquinas.lNumIntDoc
       
        lErro = CF("Verifica_Maquinas_Disponiveis_Dia", objCTMaquinas, StrParaDate(GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_DataReal_Col)), giFilialEmpresa, iQtd, dHoras)
        If lErro <> SUCESSO Then gError 140989
        
        If StrParaDbl(HorasReal.Text) - dHoras > QTDE_ESTOQUE_DELTA Then
            vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_HORAS_DISP_MENOR_HORAS_CAD", HorasReal.Text, Formata_Estoque(dHoras))
            If vbMsgBox = vbNo Then gError 140990
        End If
        
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140799

    Saida_Celula_HorasReal = SUCESSO

    Exit Function

Erro_Saida_Celula_HorasReal:

    Saida_Celula_HorasReal = gErr

    Select Case gErr

        Case 140798, 140799, 140987 To 140990
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163909)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeMaquinaReal(objGridInt As AdmGrid) As Long
'faz a critica da celula de QuantidadeMaquina do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim dHoras As Double
Dim iQtd As Integer
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim objMaquinas As New ClassMaquinas
Dim vbMsgBox As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_QuantidadeMaquinaReal

    Set objGridInt.objControle = QuantidadeMaquinaReal

    'verifica se a QuantidadeMaquina está preenchida
    If Len(Trim(QuantidadeMaquinaReal.Text)) > 0 Then

        'verifica se a QuantidadeMaquina é válida
        lErro = Valor_Positivo_Critica(QuantidadeMaquinaReal.Text)
        If lErro <> SUCESSO Then gError 140800
        
        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        objCentrodeTrabalho.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 140983
        
        objCTMaquinas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
        
        objMaquinas.sNomeReduzido = GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_NomeMaquinaReal_Col)
    
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 140984
        
        objCTMaquinas.lNumIntDocMaq = objMaquinas.lNumIntDoc
       
        lErro = CF("Verifica_Maquinas_Disponiveis_Dia", objCTMaquinas, StrParaDate(GridMaquinasReal.TextMatrix(GridMaquinasReal.Row, iGrid_DataReal_Col)), giFilialEmpresa, iQtd, dHoras)
        If lErro <> SUCESSO Then gError 140985
        
        If StrParaInt(QuantidadeMaquinaReal.Text) > iQtd Then
            vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_QTD_DISP_MENOR_QTD_CAD", QuantidadeMaquinaReal.Text, iQtd)
            If vbMsgBox = vbNo Then gError 140986
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140801

    Saida_Celula_QuantidadeMaquinaReal = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeMaquinaReal:

    Saida_Celula_QuantidadeMaquinaReal = gErr

    Select Case gErr

        Case 140800, 140801, 140983 To 140986
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163910)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataRealMO(objGridInt As AdmGrid) As Long
'faz a critica da celula de data do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim bAchou As Boolean

On Error GoTo Erro_Saida_Celula_DataRealMO

    Set objGridInt.objControle = DataRealMO

    'verifica se a data está preenchida
    If Len(Trim(DataRealMO.ClipText)) > 0 Then

        'verifica se a data é válida
        lErro = Data_Critica(DataRealMO.Text)
        If lErro <> SUCESSO Then gError 140794
        
        bAchou = False
        NomeMaquinaRealMO.Clear
        For iLinha = 1 To objGridMaquinasReal.iLinhasExistentes
            If StrParaDate(GridMaquinasReal.TextMatrix(iLinha, iGrid_DataReal_Col)) = StrParaDate(DataRealMO.Text) Then
                NomeMaquinaRealMO.AddItem GridMaquinasReal.TextMatrix(iLinha, iGrid_NomeMaquinaReal_Col)
                bAchou = True
            End If
        Next
        
        If Not bAchou Then gError 140967

        'verifica se precisa preencher o grid com uma nova linha
        If GridMOReal.Row - GridMOReal.FixedRows = objGridMOReal.iLinhasExistentes Then
            objGridMOReal.iLinhasExistentes = objGridMOReal.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140795

    Saida_Celula_DataRealMO = SUCESSO

    Exit Function

Erro_Saida_Celula_DataRealMO:

    Saida_Celula_DataRealMO = gErr

    Select Case gErr

        Case 140794, 140795
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 140967
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_MAQUINA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163911)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_HorasRealMO(objGridInt As AdmGrid) As Long
'faz a critica da celula de Horas do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMO As ClassCTOperadores
Dim iMO As Integer
Dim dHoras As Double
Dim vbMsgBox As VbMsgBoxResult
Dim bAchou As Boolean

On Error GoTo Erro_Saida_Celula_HorasRealMO

    Set objGridInt.objControle = HorasRealMO

    'verifica se a Horas está preenchida
    If Len(Trim(HorasRealMO.Text)) > 0 Then

        'verifica se a Horas é válida
        lErro = Valor_Positivo_Critica(HorasRealMO.Text)
        If lErro <> SUCESSO Then gError 140802

        HorasRealMO.Text = Formata_Estoque(StrParaDbl(HorasRealMO.Text))

        objCT.iFilialEmpresa = giFilialEmpresa
        objCT.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137212 Then gError 141038
       
        'Pega a quantidade de Horas disponível no CT
        lErro = CF("Verifica_CT_Disponiveis_Dia", objCT, StrParaDate(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DataRealMO_Col)), dHoras)
        If lErro <> SUCESSO Then gError 141039
        
        iMO = StrParaInt(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_TipoMaoDeObraReal_Col))
        
        bAchou = False
        
        'Multiplica pela quantidade de Mão de Obra
        For Each objCTMO In objCT.colOperadores
            If iMO = objCTMO.iCodTipoMO Then
                dHoras = dHoras * objCTMO.iQuantidade
                bAchou = True
                Exit For
            End If
        Next
        
        If bAchou Then
            'Se a qtd de horas for maior que a diponível informa ao usuário
            If StrParaDbl(HorasRealMO.Text) - dHoras > QTDE_ESTOQUE_DELTA Then
                vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_HORAS_DISP_MENOR_HORAS_CAD", HorasRealMO.Text, Formata_Estoque(dHoras))
                If vbMsgBox = vbNo Then gError 141040
            End If
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140803

    Saida_Celula_HorasRealMO = SUCESSO

    Exit Function

Erro_Saida_Celula_HorasRealMO:

    Saida_Celula_HorasRealMO = gErr

    Select Case gErr

        Case 140802, 140803, 141038 To 141040
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163912)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_QuantidadeMaquinaRealMO(objGridInt As AdmGrid) As Long
'faz a critica da celula de QuantidadeMaquina do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMO As ClassCTOperadores
Dim iMO As Integer
Dim vbMsgBox As VbMsgBoxResult
Dim bAchou As Boolean

On Error GoTo Erro_Saida_Celula_QuantidadeMaquinaRealMO

    Set objGridInt.objControle = QuantidadeMaquinaRealMO

    'verifica se a QuantidadeMaquina está preenchida
    If Len(Trim(QuantidadeMaquinaRealMO.Text)) > 0 Then

        'verifica se a QuantidadeMaquina é válida
        lErro = Valor_Positivo_Critica(QuantidadeMaquinaRealMO.Text)
        If lErro <> SUCESSO Then gError 140804

        objCT.iFilialEmpresa = giFilialEmpresa
        objCT.sNomeReduzido = CodigoCTPadrao.Text
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137212 Then gError 141041
       
        iMO = StrParaInt(GridMOReal.TextMatrix(GridMOReal.Row, iGrid_TipoMaoDeObraReal_Col))
        
        bAchou = False
        
        'Multiplica pela quantidade de Mão de Obra
        For Each objCTMO In objCT.colOperadores
            If iMO = objCTMO.iCodTipoMO Then
                bAchou = True
                Exit For
            End If
        Next
        
        If bAchou Then
            If StrParaInt(QuantidadeMaquinaRealMO.Text) - objCTMO.iQuantidade > QTDE_ESTOQUE_DELTA Then
                vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_QTD_DISP_MENOR_QTD_CAD", QuantidadeMaquinaRealMO.Text, objCTMO.iQuantidade)
                If vbMsgBox = vbNo Then gError 141042
            End If
        End If
            
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140805

    Saida_Celula_QuantidadeMaquinaRealMO = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantidadeMaquinaRealMO:

    Saida_Celula_QuantidadeMaquinaRealMO = gErr

    Select Case gErr

        Case 140804, 140805, 141041, 141042
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163913)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoMaodeObraReal(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sCodTipoMO As String
Dim iLinha As Integer
Dim objTiposDeMaodeObra As New ClassTiposDeMaodeObra
Dim objCT As New ClassCentrodeTrabalho
Dim objCTMO As ClassCTOperadores
Dim bAchou As Boolean
Dim iIndice As Integer
Dim dtData As Date
Dim sMaquina As String

On Error GoTo Erro_Saida_Celula_TipoMaodeObraReal

    Set objGridInt.objControle = TipoMaoDeObraReal
    
    iLinha = GridMOReal.Row
    dtData = StrParaDate(GridMOReal.TextMatrix(iLinha, iGrid_DataRealMO_Col))
    sMaquina = GridMOReal.TextMatrix(iLinha, iGrid_NomeMaquinaRealMO_Col)
                    
    'Se o campo foi preenchido
    If Len(TipoMaoDeObraReal.Text) > 0 Then
        
        objTiposDeMaodeObra.iCodigo = StrParaInt(TipoMaoDeObraReal.Text)
        
        objCT.sNomeReduzido = CodigoCTPadrao.Text
        objCT.iFilialEmpresa = giFilialEmpresa
        
         'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_Completo", objCT)
        If lErro <> SUCESSO And lErro <> 137212 Then gError 141042
        
        'Lê o TiposDeMaodeObra que está sendo Passado
        lErro = CF("TiposDeMaodeObra_Le", objTiposDeMaodeObra)
        If lErro <> SUCESSO And lErro <> 137598 Then gError 140806
    
        If lErro = SUCESSO Then
            
            bAchou = False
            For Each objCTMO In objCT.colOperadores
                If objCTMO.iCodTipoMO = objTiposDeMaodeObra.iCodigo Then
                    bAchou = True
                    Exit For
                End If
            Next
        
            If Not bAchou Then gError 141043
            
            For iIndice = 1 To objGridMOReal.iLinhasExistentes
                If iIndice <> iLinha Then
                    'Para mesma data
                    If dtData = StrParaDate(GridMOReal.TextMatrix(iIndice, iGrid_DataRealMO_Col)) Then
                        'Mesma Máquina
                        If sMaquina = GridMOReal.TextMatrix(iIndice, iGrid_NomeMaquinaRealMO_Col) Then
                            'Mesmo Operador
                            If objTiposDeMaodeObra.iCodigo = StrParaInt(GridMOReal.TextMatrix(iIndice, iGrid_TipoMaoDeObraReal_Col)) Then
                                gError 141050
                            End If
                        End If
                    End If
                End If
            Next

            GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DescricaoTipoMO_Col) = objTiposDeMaodeObra.sDescricao
        
        Else
        
            TipoMaoDeObraReal.Text = ""
            GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DescricaoTipoMO_Col) = ""
            gError 140807
        
        End If
        
    Else
        GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DescricaoTipoMO_Col) = ""

    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140808

    Saida_Celula_TipoMaodeObraReal = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoMaodeObraReal:

    Saida_Celula_TipoMaodeObraReal = gErr

    Select Case gErr
                
        Case 140806, 140808, 141042
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 140807
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_CADASTRADO", gErr, objTiposDeMaodeObra.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 141043
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_PERTENTECE_CT", gErr, objTiposDeMaodeObra.iCodigo, objCT.lCodigo, objCT.iFilialEmpresa)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 141050
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOSDEMAODEOBRA_NAO_REPETIDO", gErr, objTiposDeMaodeObra.iCodigo, dtData, sMaquina, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163914)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_NomeMaquinaRealMO(objGridInt As AdmGrid) As Long
'Faz a crítica da célula NomeMaquinaRealMO do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iLinha As Integer
Dim objMaquinas As New ClassMaquinas
Dim sMaquina As String
Dim sMaquinaAnterior As String

On Error GoTo Erro_Saida_Celula_NomeMaquinaRealMO

    Set objGridInt.objControle = NomeMaquinaRealMO
    
    sMaquinaAnterior = GridMOReal.TextMatrix(GridMOReal.Row, iGrid_NomeMaquinaRealMO_Col)
   
    If NomeMaquinaRealMO.Text <> sMaquinaAnterior Then
   
        GridMOReal.TextMatrix(GridMOReal.Row, iGrid_TipoMaoDeObraReal_Col) = ""
        GridMOReal.TextMatrix(GridMOReal.Row, iGrid_DescricaoTipoMO_Col) = ""
   
    End If
   
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 140808

    Saida_Celula_NomeMaquinaRealMO = SUCESSO

    Exit Function

Erro_Saida_Celula_NomeMaquinaRealMO:

    Saida_Celula_NomeMaquinaRealMO = gErr

    Select Case gErr
        
        Case 140808
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163915)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub BotaoTrazerMaquinasPrevistas_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iLinha2 As Integer
Dim colMO As New Collection
Dim objMaquina As ClassMaquinas
Dim objMaquinaMO As ClassMaquinaOperadores
Dim sNomeMaquina As String
Dim iQtdMaq As Integer
Dim dHorasMaq As Double
Dim dtData As Date
Dim objMO As ClassTiposDeMaodeObra

On Error GoTo Erro_BotaoTrazerMaquinasPrevistas_Click

    Call Grid_Limpa(objGridMaquinasReal)
    Call Grid_Limpa(objGridMOReal)
    
    iLinha2 = 0
    iLinha = 0

    For iLinha = 1 To objGridMaquinas.iLinhasExistentes
    
        sNomeMaquina = GridMaquinas.TextMatrix(iLinha, iGrid_NomeMaquina_Col)
        iQtdMaq = StrParaInt(GridMaquinas.TextMatrix(iLinha, iGrid_QuantidadeMaquina_Col))
        dHorasMaq = StrParaDbl(GridMaquinas.TextMatrix(iLinha, iGrid_HorasPO_Col))
        dtData = StrParaDate(GridMaquinas.TextMatrix(iLinha, iGrid_DataPO_Col))
        
        GridMaquinasReal.TextMatrix(iLinha, iGrid_NomeMaquinaReal_Col) = sNomeMaquina
        GridMaquinasReal.TextMatrix(iLinha, iGrid_QuantidadeMaquinaReal_Col) = iQtdMaq
        GridMaquinasReal.TextMatrix(iLinha, iGrid_HorasReal_Col) = Formata_Estoque(dHorasMaq)
        GridMaquinasReal.TextMatrix(iLinha, iGrid_DataReal_Col) = Format(dtData, "dd/mm/yyyy")
    
        Set objMaquina = New ClassMaquinas
        
        objMaquina.sNomeReduzido = sNomeMaquina
    
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 140810
        
        'Lê o MaquinasInsumos que está sendo Passado
        lErro = CF("Maquinas_Le_Itens", objMaquina)
        If lErro <> SUCESSO Then gError 140822
       
        
        For Each objMaquinaMO In objMaquina.colTipoOperadores
        
            iLinha2 = iLinha2 + 1
            
            Set objMO = New ClassTiposDeMaodeObra
            
            objMO.iCodigo = objMaquinaMO.iTipoMaoDeObra
            
            lErro = CF("TiposDeMaodeObra_Le", objMO)
            If lErro <> SUCESSO And lErro <> 137598 Then gError 140811
                       
            GridMOReal.TextMatrix(iLinha2, iGrid_NomeMaquinaRealMO_Col) = sNomeMaquina
            GridMOReal.TextMatrix(iLinha2, iGrid_QuantidadeMaquinaRealMO_Col) = objMaquinaMO.iQuantidade * iQtdMaq
            GridMOReal.TextMatrix(iLinha2, iGrid_HorasRealMO_Col) = Formata_Estoque(objMaquinaMO.dPercentualUso * dHorasMaq)
            GridMOReal.TextMatrix(iLinha2, iGrid_DataRealMO_Col) = Format(dtData, "dd/mm/yyyy")
            GridMOReal.TextMatrix(iLinha2, iGrid_TipoMaoDeObraReal_Col) = objMO.iCodigo
            GridMOReal.TextMatrix(iLinha2, iGrid_DescricaoTipoMO_Col) = objMO.sDescricao
        
        Next
            
    Next

    objGridMaquinasReal.iLinhasExistentes = objGridMaquinas.iLinhasExistentes
    objGridMOReal.iLinhasExistentes = iLinha2

    Exit Sub

Erro_BotaoTrazerMaquinasPrevistas_Click:

    Select Case gErr
    
        Case 140810, 140811, 140822
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163916)
    
    End Select

    Exit Sub

End Sub

Public Sub GridMaquinasReal_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMaquinasReal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinasReal, iAlterado)
    End If

End Sub

Public Sub GridMaquinasReal_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMaquinasReal)

End Sub

Public Sub GridMaquinasReal_EnterCell()

    Call Grid_Entrada_Celula(objGridMaquinasReal, iAlterado)

End Sub

Public Sub GridMaquinasReal_LeaveCell()
    
    Call Saida_Celula(objGridMaquinasReal)

End Sub

Public Sub GridMaquinasReal_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridMaquinasReal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMaquinasReal, iAlterado)
    End If

End Sub

Public Sub GridMaquinasReal_RowColChange()

    Call Grid_RowColChange(objGridMaquinasReal)

End Sub

Public Sub GridMaquinasReal_Scroll()

    Call Grid_Scroll(objGridMaquinasReal)

End Sub

Public Sub GridMaquinasReal_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long

On Error GoTo Erro_GridMaquinasReal_KeyDown
        
    'Verifica se é uma exclusão
    If KeyCode = vbKeyDelete Then
    
        lErro = GridMaquinasReal_Valida_Exclusao_Linha(GridMaquinasReal.Row)
        If lErro <> SUCESSO Then gError 140851
            
    End If
        
    Call Grid_Trata_Tecla1(KeyCode, objGridMaquinasReal)
        
    Exit Sub
    
Erro_GridMaquinasReal_KeyDown:
    
    Select Case gErr
    
        Case 140851

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163917)

    End Select

    Exit Sub
    
End Sub

Public Sub GridMaquinasReal_LostFocus()

    Call Grid_Libera_Foco(objGridMaquinasReal)

End Sub

Public Sub GridMOReal_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridMOReal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMOReal, iAlterado)
    End If

End Sub

Public Sub GridMOReal_GotFocus()
    
    Call Grid_Recebe_Foco(objGridMOReal)

End Sub

Public Sub GridMOReal_EnterCell()

    Call Grid_Entrada_Celula(objGridMOReal, iAlterado)

End Sub

Public Sub GridMOReal_LeaveCell()
    
    Call Saida_Celula(objGridMOReal)

End Sub

Public Sub GridMOReal_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridMOReal, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMOReal, iAlterado)
    End If

End Sub

Public Sub GridMOReal_RowColChange()

    Call Grid_RowColChange(objGridMOReal)

End Sub

Public Sub GridMOReal_Scroll()

    Call Grid_Scroll(objGridMOReal)

End Sub

Public Sub GridMOReal_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridMOReal)
        
End Sub

Public Sub GridMOReal_LostFocus()

    Call Grid_Libera_Foco(objGridMOReal)

End Sub

Private Function Move_MaquinasReal_Memoria(objOPOperacoes As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objItemOPMaquinas As ClassItemOPOperacoesMaquinas
Dim objItemOPMO As ClassItemOPOperacoesMO
Dim objMaquina As ClassMaquinas

On Error GoTo Erro_Move_MaquinasReal_Memoria

    'Limpa coleção para reiniciar
    Set objOPOperacoes.colUsoMaquinas = New Collection

    'Ir preenchendo a colecao no objOperacoes com todas as linhas "existentes" do grid
    For iIndice = 1 To objGridMaquinasReal.iLinhasExistentes
    
        Set objItemOPMaquinas = New ClassItemOPOperacoesMaquinas
        Set objMaquina = New ClassMaquinas
        
        objMaquina.sNomeReduzido = GridMaquinasReal.TextMatrix(iIndice, iGrid_NomeMaquinaReal_Col)
    
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquina)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 140810
        
        objItemOPMaquinas.dHoras = StrParaDbl(GridMaquinasReal.TextMatrix(iIndice, iGrid_HorasReal_Col))
        objItemOPMaquinas.dtData = StrParaDate(GridMaquinasReal.TextMatrix(iIndice, iGrid_DataReal_Col))
        objItemOPMaquinas.iQuantidade = StrParaInt(GridMaquinasReal.TextMatrix(iIndice, iGrid_QuantidadeMaquinaReal_Col))
        objItemOPMaquinas.lNumIntDocMaq = objMaquina.lNumIntDoc
        
        If objMaquina.sNomeReduzido = "" Then gError 140812
        If objItemOPMaquinas.dHoras = 0 Then gError 140813
        If objItemOPMaquinas.iQuantidade = 0 Then gError 140814
        
        For iIndice2 = 1 To objGridMOReal.iLinhasExistentes
        
            'Se a mão de Obra é dessa máquina nesse dia
            If objItemOPMaquinas.dtData = StrParaDate(GridMOReal.TextMatrix(iIndice2, iGrid_DataRealMO_Col)) And _
                objMaquina.sNomeReduzido = GridMOReal.TextMatrix(iIndice2, iGrid_NomeMaquinaRealMO_Col) Then
        
                Set objItemOPMO = New ClassItemOPOperacoesMO
            
                objItemOPMO.dHoras = StrParaDbl(GridMOReal.TextMatrix(iIndice2, iGrid_HorasRealMO_Col))
                objItemOPMO.iQuantidade = StrParaInt(GridMOReal.TextMatrix(iIndice2, iGrid_QuantidadeMaquinaRealMO_Col))
                objItemOPMO.iTipoMO = StrParaInt(GridMOReal.TextMatrix(iIndice2, iGrid_TipoMaoDeObraReal_Col))
            
                objItemOPMaquinas.colMO.Add objItemOPMO
                
            End If
            
            If Len(Trim(GridMOReal.TextMatrix(iIndice2, iGrid_NomeMaquinaRealMO_Col))) = 0 Then gError 140815
            If StrParaInt(GridMOReal.TextMatrix(iIndice2, iGrid_TipoMaoDeObraReal_Col)) = 0 Then gError 140816
            If StrParaInt(GridMOReal.TextMatrix(iIndice2, iGrid_QuantidadeMaquinaRealMO_Col)) = 0 Then gError 140817
            If StrParaDbl(GridMOReal.TextMatrix(iIndice2, iGrid_HorasRealMO_Col)) = 0 Then gError 140818
        
        Next
        
        objOPOperacoes.colUsoMaquinas.Add objItemOPMaquinas
      
    Next

    Move_MaquinasReal_Memoria = SUCESSO

    Exit Function

Erro_Move_MaquinasReal_Memoria:

    Move_MaquinasReal_Memoria = gErr

    Select Case gErr
    
        Case 140812
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 140813
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAS_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 140814
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case 140815
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_PREENCHIDA_GRID2", gErr, iIndice2)

        Case 140816
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMAODEOBRA_NAO_PREENCHIDA_GRID", gErr, iIndice2)

        Case 140817
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA_GRID2", gErr, iIndice2)

        Case 140818
            Call Rotina_Erro(vbOKOnly, "ERRO_HORAS_NAO_PREENCHIDA_GRID2", gErr, iIndice2)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163918)

    End Select

    Exit Function

End Function

Private Function Preenche_GridMaquinasReal(ByVal objOPOperacoesAux As ClassOrdemProducaoOperacoes) As Long

Dim lErro As Long
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim objItemOPMaq As ClassItemOPOperacoesMaquinas
Dim objItemOPMO As ClassItemOPOperacoesMO
Dim objMaquina As ClassMaquinas
Dim objMO As ClassTiposDeMaodeObra
Dim objOPOperacoes As New ClassOrdemProducaoOperacoes

On Error GoTo Erro_Preenche_GridMaquinasReal
    
    Call Grid_Limpa(objGridMaquinasReal)
    Call Grid_Limpa(objGridMOReal)
    
    If objOPOperacoesAux.lNumIntDocOperOrigem = 0 Then
    
        Set objOPOperacoes = objOPOperacoesAux
    
    Else
    
        objOPOperacoes.lNumIntDoc = objOPOperacoesAux.lNumIntDocOperOrigem
    
        lErro = CF("OrdemDeProducao_Le_Oper_NumIntDoc", objOPOperacoes)
        If lErro <> SUCESSO And lErro <> 137037 Then gError 140853
    
    End If
    
    'Exibe os dados da coleção na tela
    For Each objItemOPMaq In objOPOperacoes.colUsoMaquinas
    
        Set objMaquina = New ClassMaquinas
        
        objMaquina.lNumIntDoc = objItemOPMaq.lNumIntDocMaq
    
        'Le a Máquina atraves do NumIntDoc
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquina)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 140819
        
        'Se nao encontrou => Erro
        If lErro = 106353 Then gError 140820
            
        iIndice1 = iIndice1 + 1
    
        GridMaquinasReal.TextMatrix(iIndice1, iGrid_HorasReal_Col) = Formata_Estoque(objItemOPMaq.dHoras)
        GridMaquinasReal.TextMatrix(iIndice1, iGrid_NomeMaquinaReal_Col) = objMaquina.sNomeReduzido
        GridMaquinasReal.TextMatrix(iIndice1, iGrid_QuantidadeMaquinaReal_Col) = objItemOPMaq.iQuantidade
        GridMaquinasReal.TextMatrix(iIndice1, iGrid_DataReal_Col) = Format(objItemOPMaq.dtData, "dd/mm/yyyy")
        
        For Each objItemOPMO In objItemOPMaq.colMO
        
            Set objMO = New ClassTiposDeMaodeObra
            
            objMO.iCodigo = objItemOPMO.iTipoMO
        
            'Lê o TiposDeMaodeObra que está sendo Passado
            lErro = CF("TiposDeMaodeObra_Le", objMO)
            If lErro <> SUCESSO And lErro <> 137598 Then gError 140821
        
            iIndice2 = iIndice2 + 1
        
            GridMOReal.TextMatrix(iIndice2, iGrid_HorasRealMO_Col) = Formata_Estoque(objItemOPMO.dHoras)
            GridMOReal.TextMatrix(iIndice2, iGrid_DataRealMO_Col) = Format(objItemOPMaq.dtData, "dd/mm/yyyy")
            GridMOReal.TextMatrix(iIndice2, iGrid_QuantidadeMaquinaRealMO_Col) = objItemOPMO.iQuantidade
            GridMOReal.TextMatrix(iIndice2, iGrid_NomeMaquinaRealMO_Col) = objMaquina.sNomeReduzido
            GridMOReal.TextMatrix(iIndice2, iGrid_TipoMaoDeObraReal_Col) = objMO.iCodigo
            GridMOReal.TextMatrix(iIndice2, iGrid_DescricaoTipoMO_Col) = objMO.sDescricao
        
        Next

    Next

    objGridMaquinasReal.iLinhasExistentes = iIndice1
    objGridMOReal.iLinhasExistentes = iIndice2
    
    Preenche_GridMaquinasReal = SUCESSO
    
    Exit Function

Erro_Preenche_GridMaquinasReal:

    Preenche_GridMaquinasReal = gErr

    Select Case gErr

        Case 140819, 140821, 140853
        
        Case 140820
            Call Rotina_Erro(vbOKOnly, "ERRO_MAQUINA_NAO_CADASTRADA", gErr, objMaquina.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163919)

    End Select

    Exit Function

End Function

Private Function GridMaquinasReal_Valida_Exclusao_Linha(ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sMaquina As String
Dim dtData As Date
Dim iIndice As Integer

On Error GoTo Erro_GridMaquinasReal_Valida_Exclusao_Linha

    sMaquina = GridMaquinasReal.TextMatrix(iLinha, iGrid_NomeMaquinaReal_Col)
    dtData = StrParaDate(GridMaquinasReal.TextMatrix(iLinha, iGrid_DataReal_Col))

    For iIndice = 1 To objGridMOReal.iLinhasExistentes
    
        If sMaquina = GridMOReal.TextMatrix(iIndice, iGrid_NomeMaquinaRealMO_Col) And dtData = StrParaDate(GridMOReal.TextMatrix(iIndice, iGrid_DataRealMO_Col)) Then gError 140852
    
    Next
    
    GridMaquinasReal_Valida_Exclusao_Linha = SUCESSO
    
    Exit Function

Erro_GridMaquinasReal_Valida_Exclusao_Linha:

    GridMaquinasReal_Valida_Exclusao_Linha = gErr

    Select Case gErr
    
        Case 140852
            Call Rotina_Erro(vbOKOnly, "ERRO_MAODEOBRA_MAQUINA_EXISTENTE", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163920)

    End Select

    Exit Function

End Function
'#############################################################

'##################################################
'Inserido por Wagner
Public Sub BotaoImprimirRotulos_Click()

Dim objRelatorio As New AdmRelatorio
Dim lErro As Long
Dim iIndice As Integer
Dim objOP As New ClassOrdemDeProducao

On Error GoTo Erro_BotaoImprimirRotulos_Click

    If objGrid.iLinhasExistentes = 0 Then gError 141501

    If Len(Trim(Codigo.Text)) = 0 Then gError 141502
    
    'Loop de Validação dos dados do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a quantidade foi digitada
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 141503

        'Verifica se o almoxarifado foi informado
        If Len(Trim(GridMovimentos.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 And left(GridMovimentos.TextMatrix(iIndice, 0), 1) <> "#" Then gError 141504
        
    Next
                    
    'Move os dados para o objMovimentoEstoque
    lErro = Move_Tela_Memoria(objOP)
    If lErro <> SUCESSO Then gError 30989
    
    Call objRelatorio.Rel_Menu_Executar("Rótulos de Produção", objOP)

    Exit Sub

Erro_BotaoImprimirRotulos_Click:
    
    Select Case gErr
    
        Case 141501
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEMOP_INFORMADO", gErr)

        Case 141502
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGOOP_NAO_PREENCHIDO", gErr)

        Case 141503
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 141504
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163921)

    End Select

    Exit Sub
    
End Sub
'##################################################

Public Sub NumMaxMaqPorOper_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumMaxMaqPorOper_Validate

    'Veifica se NumMaxMaqPorOper está preenchida
    If Len(Trim(NumMaxMaqPorOper.Text)) <> 0 Then

       'Critica a NumMaxMaqPorOper
       lErro = Valor_Positivo_Critica(NumMaxMaqPorOper.Text)
       If lErro <> SUCESSO Then gError 141610
       
    End If

    Exit Sub

Erro_NumMaxMaqPorOper_Validate:

    Cancel = True

    Select Case gErr

        Case 141610

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163922)

    End Select

    Exit Sub

End Sub

Public Sub NumMaxMaqPorOper_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumMaxMaqPorOper, iAlterado)
    
End Sub

Public Sub NumMaxMaqPorOper_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoRoteiros_Click()

Dim lErro As Long
Dim objRoteirosDeFabricacao As New ClassRoteirosDeFabricacao
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objNode As Object

On Error GoTo Erro_BotaoRoteiros_Click

'    lErro = CF("Produto_Formata", SCodigo_Extrai(ProdutoLabel.Caption), sProdutoFormatado, iProdutoPreenchido)
'    If lErro <> SUCESSO Then gError 177340
'
'    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    Set objNode = Roteiro.SelectedItem

    Set objOrdemProducaoOperacoes = colComponentes.Item(objNode.Tag)

    objRoteirosDeFabricacao.sProdutoRaiz = objOrdemProducaoOperacoes.sProduto
    
    If Len(Trim(VersaoLabel.Caption)) <> 0 Then
        objRoteirosDeFabricacao.sVersao = VersaoLabel.Caption
    End If

    Call Chama_Tela("RoteirosDeFabricacao", objRoteirosDeFabricacao)

    Exit Sub

Erro_BotaoRoteiros_Click:

    Select Case gErr

        Case 177340

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177341)

    End Select

    Exit Sub

End Sub

Public Sub BotaoKit_Click()

Dim lErro As Long
Dim objKit As New ClassKit
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objOrdemProducaoOperacoes As ClassOrdemProducaoOperacoes
Dim objNode As Object

On Error GoTo Erro_BotaoKit_Click

'    lErro = CF("Produto_Formata", SCodigo_Extrai(ProdutoLabel.Caption), sProdutoFormatado, iProdutoPreenchido)
'    If lErro <> SUCESSO Then gError 177342
'
'    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then sProdutoFormatado = ""

    Set objNode = Roteiro.SelectedItem

    Set objOrdemProducaoOperacoes = colComponentes.Item(objNode.Tag)

    objKit.sProdutoRaiz = objOrdemProducaoOperacoes.sProduto
    
    If Len(Trim(VersaoLabel.Caption)) <> 0 Then
        objKit.sVersao = VersaoLabel.Caption
    End If

    Call Chama_Tela("Kit", objKit)

    Exit Sub

Erro_BotaoKit_Click:

    Select Case gErr

        Case 177342

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 177343)

    End Select

    Exit Sub

End Sub

Public Sub Repeticao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Repeticao_Validate

    'Veifica se Repeticao está preenchida
'    If Len(Trim(Repeticao.Text)) <> 0 Then

       'Critica a Repeticao
       lErro = Valor_Positivo_Critica(Repeticao.Text)
       If lErro <> SUCESSO Then gError 141611
            
'    End If

    Exit Sub

Erro_Repeticao_Validate:

    Cancel = True

    Select Case gErr

        Case 141611

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174307)

    End Select

    Exit Sub

End Sub

Public Sub Repeticao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Repeticao, iAlterado)
    
End Sub

Public Sub Repeticao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BotaoOPRC_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim objItemOP As ClassItemOP

On Error GoTo Erro_BotaoOPRC_Click:

    gobjOP.sCodigo = Codigo.Text
    
    For iIndice = 1 To objGrid.iLinhasExistentes
    
        Set objItemOP = gobjOP.colItens.Item(iIndice)
        
        objItemOP.sVersao = GridMovimentos.TextMatrix(iIndice, iGrid_Versao_Col)
        
    Next
                
    Call Chama_Tela_Modal("OrdemProducaoRC", gobjOP)

    Exit Sub

Erro_BotaoOPRC_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 174307)

    End Select

    Exit Sub
    
End Sub

'#####################################
'Inserido por Wagner 03/08/2006
Sub BotaoProjetos_Click()
    Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Sub

Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Sub Projeto_GotFocus()
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Sub

Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Projeto_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    For iIndice = 1 To objGrid.iLinhasExistentes

        Set objItensPRJCR = New ClassItensPRJCR
        
        objItensPRJCR.sItem = CStr(iIndice)
        objItensPRJCR.sDescricao = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col)
        objItensPRJCR.sObservacao = GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col) & SEPARADOR & GridMovimentos.TextMatrix(iIndice, iGrid_DescricaoItem_Col)

        objItensPRJCR.dQuantidadeOriginal = StrParaDbl(GridMovimentos.TextMatrix(iIndice, iGrid_Quantidade_Col))

        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_OP
        
        colItensPRJ.Add objItensPRJCR

    Next
    
    If colItensPRJ.Count = 0 Then
        Set objItensPRJCR = New ClassItensPRJCR
        objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_OP
        colItensPRJ.Add objItensPRJCR
    End If
    
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

Sub Etapa_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Etapa_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Sub Etapa_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub
'#####################################
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

Public Function OP_Form_UnLoad(ByVal objCT As Object, Cancel As Integer) As Long
    OP_Form_UnLoad = SUCESSO
End Function

Public Function OP_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long
    OP_Rotina_Grid_Enable = SUCESSO
End Function

Public Function OP_GridMovimentos_KeyDown(ByVal objCT As Object, ByVal iLinhaAnterior As Integer) As Long
    OP_GridMovimentos_KeyDown = SUCESSO
End Function

Public Function OP_Form_Load(ByVal objCT As Object) As Long
    OP_Form_Load = SUCESSO
End Function

Function OP_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    OP_Saida_Celula = SUCESSO
End Function

Public Function OP_Item_IncluirGrid(ByVal objCT As Object) As Long
    OP_Item_IncluirGrid = SUCESSO
End Function

Public Function OP_Item_PreencheGridMov(ByVal objCT As Object, ByVal objItemOP As ClassItemOP, ByVal iIndice As Integer) As Long
    OP_Item_PreencheGridMov = SUCESSO
End Function

Public Function OP_PreencheGridMov(ByVal objCT As Object) As Long
    OP_PreencheGridMov = SUCESSO
End Function

Private Function Limpa_Linha_GridMovimentos(ByVal iLinha As Integer, ByVal bInseriuLinha As Boolean) As Long

On Error GoTo Erro_Limpa_Linha_GridMovimentos

    If bInseriuLinha Then
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes - 1
    End If

    GridMovimentos.TextMatrix(iLinha, iGrid_FilialCliente_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_FilialPedido_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Destinacao_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Situacao_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Ccl_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Almoxarifado_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_DescricaoItem_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_UnidadeMed_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Quantidade_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Benef_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_DataPrevInicio_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_DataPrevFim_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_PedidoDeVenda_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Prioridade_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Versao_Col) = ""
    GridMovimentos.TextMatrix(iLinha, iGrid_Maquina_Col) = ""

    Call CF2(Me, "OP_Limpa_Linha_GridMov", iLinha, bInseriuLinha)

    Limpa_Linha_GridMovimentos = SUCESSO
    
    Exit Function
    
Erro_Limpa_Linha_GridMovimentos:

    Limpa_Linha_GridMovimentos = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function
    
End Function

Public Function OP_Limpa_Linha_GridMov(ByVal objCT As Object, ByVal iLinha As Integer, ByVal bInseriuLinha As Boolean) As Long
    OP_Limpa_Linha_GridMov = SUCESSO
End Function

Public Function OP_GridMov_AtualizaLinha(ByVal objCT As Object, Optional ByVal bForcar As Boolean = False) As Long
    OP_GridMov_AtualizaLinha = SUCESSO
End Function

Public Function OP_GravarRegistro(ByVal objCT As Object) As Long
    OP_GravarRegistro = SUCESSO
End Function

Public Function OP_Move_Grid_Memoria(ByVal objCT As Object, ByVal iIndice As Integer, ByVal objItemOP As ClassItemOP) As Long
    OP_Move_Grid_Memoria = SUCESSO
End Function

Public Function OP_Limpa_Tela_OP(ByVal objCT As Object) As Long
    OP_Limpa_Tela_OP = SUCESSO
End Function

Public Function OP_Move_Tela_Memoria(ByVal objCT As Object, ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
    OP_Move_Tela_Memoria = SUCESSO
End Function

Public Function OP_Traz_Tela_OP(ByVal objCT As Object, ByVal objOrdemDeProducao As ClassOrdemDeProducao) As Long
    OP_Traz_Tela_OP = SUCESSO
End Function

Public Function OP_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long
    OP_UserControl_KeyDown = SUCESSO
End Function

Public Function OP_TabStrip1_Click(ByVal objCT As Object) As Long
    OP_TabStrip1_Click = SUCESSO
End Function

Public Function OP_Saida_Celula_Versao(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    OP_Saida_Celula_Versao = SUCESSO
End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_OP
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158093)
     
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158094)
     
    End Select
     
    Exit Sub

End Sub

Public Sub BotaoConsultaPV_Click()

Dim lErro As Long
Dim objPV As New ClassPedidoDeVenda
Dim colSelecao As New Collection
Dim sCodProduto As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoConsultaPV_Click

    If GridMovimentos.Row = 0 Then gError 52016

    sCodProduto = GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 52017

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        'verifica se a destinação é Pedido de Venda - - Se não for - - > Erro
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Destinacao_Col) <> STRING_PEDIDOVENDA Then gError 52018
        
        'chama a tela de lista de estoque do produto corrente
        If GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col) <> "VÁRIOS" Then
        
            objPV.lCodigo = StrParaLong(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_PedidoDeVenda_Col))
            objPV.iFilialEmpresa = Codigo_Extrai(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_FilialPedido_Col))
            
            Call Chama_Tela("PedidoVenda", objPV)
        
        Else
            colSelecao.Add Codigo.Text
            colSelecao.Add sProdutoFormatado
        
            Call Chama_Tela("ItemOPItemPVLista", colSelecao, objPV, Nothing, "CodigoOP = ? AND Produto = ?")
        End If
    Else
        Error 52019
    End If

    Exit Sub

Erro_BotaoConsultaPV_Click:

    Select Case gErr
        
        Case 52016
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 52017
                
        Case 52018
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINACAO_DEPENDENTE", gErr)
        
        Case 52019
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163809)

    End Select

    Exit Sub
End Sub

Public Sub ClienteLabel_Click()
'chama o browser de clientes

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'se o cliente foi preenchido
    If Len(Trim(ClienteTerc.Text)) > 0 Then
        'Prenche o nome reduzido do Cliente
        objCliente.sNomeReduzido = Trim(ClienteTerc.Text)
    End If
    
    'chama a tela de browser clienteslista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Public Sub FilialTerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FornecedorTerc_Change()
    iAlterado = REGISTRO_ALTERADO
    iFornecedorTercAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FornecedorTerc_Validate(Cancel As Boolean)
'verifica se o fornecedor selecionado é valido

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornecedorTerc_Validate
    
    If iFornecedorTercAlterado = REGISTRO_ALTERADO Then
    
        'limpa a filial e sai da rotina
        FilialTerc.Clear
    
        'Verifica preenchimento de Fornecedor, se não foi preenchido
        If Len(Trim(FornecedorTerc.Text)) <> 0 Then
    
            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(FornecedorTerc, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialTerc, colCodigoNome)
        
            'verifica se foi digitado nome ou cód. do fornecedor
            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
                
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                    
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", FilialTerc, iCodFilial)
                
            End If
            
        End If
        
    End If

    Exit Sub

Erro_FornecedorTerc_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209631)

    End Select

    Exit Sub

End Sub

Public Sub ClienteTerc_Change()
    iAlterado = REGISTRO_ALTERADO
    iClienteTercAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ClienteTerc_Validate(Cancel As Boolean)
'verifica se o cliente é valido

Dim lErro As Long
Dim iCodFilial As Integer
Dim objCliente As New ClassCliente
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_ClienteTerc_Validate

    If iClienteTercAlterado = REGISTRO_ALTERADO Then

        'limpa a filial e sai da rotina
        FilialTerc.Clear

        'se o cliente não foi preenchido, sai da rotina
        If Len(Trim(ClienteTerc.Text)) <> 0 Then
        
            'Busca o Cliente no BD
            lErro = TP_Cliente_Le(ClienteTerc, objCliente, iCodFilial)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            'busca no bd a relação de filiais referentes ao cliente
            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Preenche ComboBox de Filiais do cliente
            Call CF("Filial_Preenche", FilialTerc, colCodigoNome)
            
            'verifica se foi digitado nome ou cód. do cliente
            If colCodigoNome.Count = 1 Or iCodFilial <> 0 Then
                
                If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ
                    
                'Seleciona filial na Combo Filial
                Call CF("Filial_Seleciona", FilialTerc, iCodFilial)
                
            End If
        
        End If
        
    End If
    
    Exit Sub
        
Erro_ClienteTerc_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
       
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209632)
    
    End Select
    
    Exit Sub

End Sub

Public Sub FilialTerc_Validate(Cancel As Boolean)
'verifica se a filial do cliente\fornecedor é válida

Dim lErro As Long
Dim objFilialCliente As ClassFilialCliente
Dim objFilialFornecedor As ClassFilialFornecedor
Dim iCodigo As Integer

On Error GoTo Erro_FilialTerc_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(FilialTerc.Text)) = 0 Then Exit Sub

   'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If FilialTerc.ListIndex >= 0 Then Exit Sub

    'se o tipo de terc. for cliente
    If OptionTipoTerc(1).Value = True Then
    
        'verifica se o cliente foi preenchido
        If Len(Trim(ClienteTerc.Text)) = 0 Then gError 209633

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(FilialTerc, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    
        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then
    
            'instancia o obj
            Set objFilialCliente = New ClassFilialCliente
    
            'passa o nº preenchido como código
            objFilialCliente.iCodFilial = iCodigo
    
            'Tentativa de leitura da Filial com esse código no BD
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Trim(ClienteTerc.Text), objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError ERRO_SEM_MENSAGEM
    
            'Não encontrou Filial no  BD
            If lErro = 17660 Then gError 209634
    
            'Encontrou Filial no BD, coloca no Text da Combo
            FilialTerc.Text = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
    
        End If
            
        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then gError 209635
    
    'senão, é o fornecedor
    Else

        'verifica se o fornecedor foi preenchido
        If Len(Trim(FornecedorTerc.Text)) = 0 Then gError 209636

        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
        lErro = Combo_Seleciona(FilialTerc, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM
    
        'Nao existe o ítem com o CÓDIGO na List da ComboBox
        If lErro = 6730 Then
    
            'instancia o obj
            Set objFilialFornecedor = New ClassFilialFornecedor
    
            'passa o nº preenchido como código
            objFilialFornecedor.iCodFilial = iCodigo
    
            'Tentativa de leitura da Filial com esse código no BD
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Trim(FornecedorTerc.Text), objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError ERRO_SEM_MENSAGEM
    
            'Não encontrou Filial no  BD
            If lErro = 18272 Then gError 209637
    
            'Encontrou Filial no BD, coloca no Text da Combo
            FilialTerc.Text = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialFornecedor.sNome
    
        End If
            
        'Não existe o ítem com a STRING na List da ComboBox
        If lErro = 6731 Then gError 209638

    End If

    Exit Sub

Erro_FilialTerc_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 209633
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 209634, 209635
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, FilialTerc.Text)

        Case 209636
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 209637, 209638
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FornecedorTerc.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209639)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorLabel_Click()
'chama o browser referente ao fornecedor

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'se o fornecedor estiver preenchido
    If Len(Trim(FornecedorTerc.Text)) > 0 Then
        'Preenche nomeReduzido com o fornecedor da tela
        objFornecedor.sNomeReduzido = Trim(FornecedorTerc.Text)
    End If
    
    'chama a tela c/ a lista de fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Public Sub OptionTipoTerc_Click(Index As Integer)
    Call OptionTipoTerc_Trata
End Sub

Private Sub OptionTipoTerc_Trata()
'verifica qual tipo de terc. foi selecionado
   
    'se foi cliente
    If OptionTipoTerc(1).Value = True Then
    
        'limpa a filial
        If Not ClienteTerc.Visible Then
        
            FilialTerc.Clear
              
            'desabilita a label e o textbox do fornecedor
            FornecedorTerc.Visible = False
            FornecedorTerc.Text = ""
            FornecedorLabel.Visible = False
            
            'habilita o cliente
            ClienteTerc.Visible = True
            ClienteLabel.Visible = True
            
        End If
        
    'senão
    Else
        
        If Not FornecedorTerc.Visible Then
        
            FilialTerc.Clear
        
            'desabilita a label e o textbox do cliente
            ClienteTerc.Visible = False
            ClienteTerc.Text = ""
            ClienteLabel.Visible = False
            
            'habilita o fornecedor
            FornecedorTerc.Visible = True
            FornecedorLabel.Visible = True
    
        End If
        
    End If

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)
'traz os dados do item selecionado no browser

Dim objCliente As ClassCliente

    Set objCliente = obj1

    'Preenche o Cliente com o cod. do Cliente selecionado
    ClienteTerc.Text = objCliente.lCodigo
    
    'Dispara o Validate de Cliente p/ a validação do cliente e preencher a filial
    Call ClienteTerc_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)
'traz os dados do item selecionado no browser

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o cód. do fornecedor na Tela
    FornecedorTerc.Text = objFornecedor.lCodigo

    'dispara o validate do fornecedor p/ validar o fornecedor e preencher a filial
    Call FornecedorTerc_Validate(bSGECancelDummy)

    Me.Show

End Sub

Public Sub EscaninhoTerc_Click()
'    If EscaninhoTerc.ListIndex > 0 Then
'        FrameClienteFornec.Enabled = True
'    Else
'        FrameClienteFornec.Enabled = False
'        FilialTerc.Clear
'        ClienteTerc.Text = ""
'        FornecedorTerc.Text = ""
'    End If
End Sub

Public Function Trata_CodigoBarras1() As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoEnxuto As String
Dim sCodBarras As String
Dim sCodBarrasOriginal As String
Dim dCusto As Double
Dim objItemMovEst As New ClassItemMovEstoque
Dim objItemOP As New ClassItemOP
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim iPossuiGrade As Integer
Dim objKit As New ClassKit
Dim objProdutoFilial As New ClassProdutoFilial
Dim colItensRomaneioGrade As New Collection
Dim objItensRomaneio As ClassItemRomaneioGrade
Dim iIndice As Integer
Dim objRomaneioGrade As New ClassRomaneioGrade


On Error GoTo Erro_Trata_CodigoBarras1

    If objGrid.iLinhasExistentes + 1 = GridMovimentos.Row Then
    
        'Verifica se o Produto está preenchido
        If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col))) = 0 Then
            
            If ActiveControl Is Produto Then

                    Set objGrid.objControle = Produto

                    lErro = Grid_Abandona_Celula(objGrid)
                    If lErro <> SUCESSO Then gError 210853

            End If
            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 210854
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 210855
    
                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 210856
    
                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 210857
        
                Me.Show
        
                Produto.PromptInclude = False
                Produto.Text = sProdutoEnxuto
                Produto.PromptInclude = True
                
                GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Produto_Col) = Produto.Text
                
                gError 210863

                
            Else
            
                gError 210859
    
    
            End If
    
        End If
    
    End If

    Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_Trata_CodigoBarras1:

    Trata_CodigoBarras1 = gErr

    Select Case gErr

        Case 210853, 210854, 210855, 210858, 210859, 210861, 210863

        Case 210856
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 210857
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 210860)

    End Select

    Exit Function

End Function

'Private Function Saida_Celula_Produto1(objGridInt As AdmGrid) As Long
'
'Dim lErro As Long
'Dim sProduto As String
'Dim iProdutoPreenchido As Integer
'Dim sProdutoFormatado As String
'Dim colSelecao As New Collection
'Dim objProduto As New ClassProduto
'Dim vbMsg As VbMsgBoxResult
'Dim objMaquina As New ClassMaquinas
'Dim objKit As New ClassKit
'Dim objProdutoFilial As New ClassProdutoFilial
'Dim iPossuiGrade As Integer
'Dim iIndice As Integer
'Dim colItensRomaneioGrade As New Collection
'Dim objItensRomaneio As ClassItemRomaneioGrade
'Dim objItemOP As New ClassItemOP
'Dim objRomaneioGrade As New ClassRomaneioGrade
'
'On Error GoTo Erro_Saida_Celula_Produto1
'
'    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
'    If lErro <> SUCESSO Then gError 106433
'
'    'se o produto foi preenchido
'    If Len(Trim(Produto.ClipText)) <> 0 Then
'
'        lErro = CF("Produto_Critica2", Produto.Text, objProduto, iProdutoPreenchido)
'        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError 126489
'
'        'se produto estiver preenchido
'        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
'
'            'se é um produto gerencial e não é pai de grade ==> erro
'            If lErro = 25043 And Len(Trim(objProduto.sGrade)) = 0 Then gError 126490
'
'            'se o produto nao for gerencial e ainda assim deu erro ==> nao está cadastrado
'            If lErro <> SUCESSO And lErro <> 25043 Then gError 21988
'
'            'verifica se este produto já foi usado na OP
'            lErro = VerificaUso_Produto(objProduto)
'            If lErro <> SUCESSO And lErro <> 41316 Then gError 55326
'
'            If lErro = 41316 Then gError 41317
'
'            iPossuiGrade = DESMARCADO
'
'            If Len(Trim(objProduto.sGrade)) > 0 Then iPossuiGrade = MARCADO
'
'            If iPossuiGrade = DESMARCADO Then
'
'                'se o produto nao controla estoque ==> erro
'                If objProduto.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then gError 126491
'
'                If objProduto.iPCP = PRODUTO_PCP_NAOPODE Or objProduto.iCompras <> PRODUTO_PRODUZIVEL Then gError 22946
'
'                 'Preenche a linha do grid
'                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
'                If lErro <> SUCESSO Then gError 22945
'
'                'Calcula a Quantidade Disponível
'                lErro = QuantDisponivel_Calcula(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Almoxarifado_Col), objProduto)
'                If lErro <> SUCESSO Then gError 41318
'
'                'Verifica se é um kit
'                objKit.sProdutoRaiz = sProdutoFormatado
'                lErro = CF("Kit_Le_Padrao", objKit)
'                If lErro <> SUCESSO And lErro <> 106304 Then gError 106430
'
'                'Se encontrou => É UM KIT => Carrega a Combo com as Versoes
'                If lErro <> 106304 Then Call Carrega_ComboVersoes(objProduto.sCodigo)
'
'                If Len(Trim(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col))) = 0 Then
'
'                    objProdutoFilial.sProduto = sProdutoFormatado
'                    objProdutoFilial.iFilialEmpresa = giFilialEmpresa
'
'                    'Busca o Lote Economico do Produto/FilialEmpresa
'                    lErro = CF("ProdutoFilial_Le", objProdutoFilial)
'                    If lErro <> SUCESSO And lErro <> 28261 Then gError 106402
'
'                    'preenche com o lote econômico (caso exista)
'                    If objProdutoFilial.dLoteEconomico > 0 Then GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col) = objProdutoFilial.dLoteEconomico
'
'                End If
'
'            'se é um produto pai de grade
'            Else
'
'                'Verifica se há filhos válidos da grade pai
'                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneioGrade)
'                If lErro <> SUCESSO Then gError 126495
'
'                'Se nao existir, erro
'                If colItensRomaneioGrade.Count = 0 Then gError 126496
'
'                'Para cada filho de grade do produto
'                For Each objItensRomaneio In colItensRomaneioGrade
'
'                    'Verifica se ele já aparece no grid
'                    For iIndice = 1 To objGrid.iLinhasExistentes
'
'                        lErro = CF("Produto_Formata", GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
'                        If lErro <> SUCESSO Then gError 126498
'
'                        'Se aparecer ==> Erro
'                        If sProdutoFormatado = objItensRomaneio.sProduto Then gError 126497
'
'                    Next
'
'                Next
'
'                objItemOP.sProduto = objProduto.sCodigo
'                objItemOP.sSiglaUMEstoque = objProduto.sSiglaUMEstoque
'                objItemOP.sCodigo = Codigo.Text
'                objItemOP.iItem = GridMovimentos.Row
'                objItemOP.sDescricao = objProduto.sDescricao
'                objItemOP.sAlmoxarifadoNomeRed = AlmoxPadrao.Text
'
'                Set objRomaneioGrade = New ClassRomaneioGrade
'                Set objRomaneioGrade.objTela = Me
'
'                objRomaneioGrade.sNomeTela = Me.Name
'
'                Set objRomaneioGrade.objObjetoTela = objItemOP
'
'                Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
'                If giRetornoTela <> vbOK Then gError 126499
'
'                 'Preenche a linha do grid
'                lErro = ProdutoLinha_Preenche(objProduto, objItemOP)
'                If lErro <> SUCESSO Then gError 126552
'
'                Call Atualiza_Grid_Movimentos(objItemOP)
'
'            End If
'
'        End If
'
'        'Incluido por Jorge Specian - 15/06/2005
'        '---------------------------------------
'        If Len(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)) <> 0 Then
'
'            lErro = Trata_Arvore(Produto.Text, GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Versao_Col), GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_UnidadeMed_Col), StrParaDbl(GridMovimentos.TextMatrix(GridMovimentos.Row, iGrid_Quantidade_Col)))
'            If lErro <> SUCESSO Then gError 137123
'
'        End If
'        '---------------------------------------
'
'    End If
'
'    Saida_Celula_Produto1 = SUCESSO
'
'    Exit Function
'
'Erro_Saida_Celula_Produto1:
'
'    Saida_Celula_Produto1 = gErr
'
'    Select Case gErr
'
'        Case 21987, 22945, 41318, 21985, 55326, 106430, 106433, 106402, 126495, 126498, 126499, 126552, 137123
'
'        Case 21988
'
'        Case 22946
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PRODUZIVEL1", gErr, Produto.Text)
'
'        Case 41317
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_DUPLICADO", gErr, Produto.Text, Codigo.Text)
'
'        Case 126490
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
'
'        Case 126491
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ESTOQUE", gErr, objProduto.sCodigo)
'
'        Case 126496
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, Produto.Text)
'
'        Case 126497
'            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_FILHO_GRADE_GRID", gErr, objProduto.sCodigo, GridMovimentos.TextMatrix(iIndice, iGrid_Produto_Col))
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163830)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Sub IgnoraEst_Click()
     iAlterado = REGISTRO_ALTERADO
End Sub
