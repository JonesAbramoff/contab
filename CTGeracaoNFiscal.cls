VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTGeracaoNFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'OBSERVACAO:
'O frame de contabilidade foi incluido nesta tela apenas p/evitar duplicidade de codigo p/tratamento de campos de lote e data contabil. Por isso ele fica invisivel.
'Os grids de alocacao e itens ficam invisiveis e servem apenas p/auxiliar a contabilizacao

'Property Variables:
Dim m_Caption As String

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjNFiscal As ClassNFiscal
Private gobjPedidoVenda As ClassPedidoDeVenda
Private giExercicio As Integer, giPeriodo As Integer
Private gcolAlmoxFilial As New Collection

'Associados a contabilidade
Public objContabil As New ClassContabil
Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Public objGrid1 As AdmGrid
Private WithEvents objEventoCodigoViagem As AdmEvento
Attribute objEventoCodigoViagem.VB_VarHelpID = -1

Event Unload()

Public iAlterado As Integer
Dim iTabPrincipalAlterado As Integer
Dim iFrameAtual As Integer
Dim iClienteAlterado As Integer
Dim gobjGeracaoNFiscal As New ClassGeracaoNFiscal

Dim giVendedorInicial As Integer
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim objGrid As AdmGrid
Dim iGrid_GeraNFiscal_Col As Integer
Dim iGrid_Pedido_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_NomeRed_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Estado_Col As Integer
Dim iGrid_Cidade_Col As Integer
Dim iGrid_Bairro_Col As Integer
Dim iGrid_TransPortadora_Col As Integer
Dim iGrid_Emissao_Col As Integer
Dim iGrid_Entrega_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Motivo_Col As Integer

'Eventos de Browse
Private WithEvents objEventoPedidoDe As AdmEvento
Attribute objEventoPedidoDe.VB_VarHelpID = -1
Private WithEvents objEventoPedidoAte As AdmEvento
Attribute objEventoPedidoAte.VB_VarHelpID = -1
Private WithEvents objEventoClienteDe As AdmEvento
Attribute objEventoClienteDe.VB_VarHelpID = -1
Private WithEvents objEventoClienteAte As AdmEvento
Attribute objEventoClienteAte.VB_VarHelpID = -1

Dim asOrdenacao(3) As String
Dim asOrdenacaoString(3) As String

'Constantes públicas dos tabs
Private Const TAB_Selecao = 1
Private Const TAB_Pedidos = 2

'mnemonicos
Private Const CODIGO1 As String = "Codigo"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const CLIENTE1 As String = "Cliente"
Private Const FILIAL1 As String = "Filial"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const Serie1 As String = "Serie"
Private Const NFISCAL1 As String = "Nota_Fiscal"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_SAIDA As String = "Data_Saida"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Medida"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const PRECO_UNITARIO As String = "Preco_Unitario"
Private Const PRECO_TOTAL As String = "Preco_Total"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
'#######Valores####################
Private Const ICMS As String = "ICMS_Valor"

Private Const ICMS_DIFAL_REMETENTE As String = "ICMS_Difal_Remetente"
Private Const ICMS_DIFAL_DESTINO As String = "ICMS_Difal_Destino"
Private Const ICMS_FCP As String = "ICMS_FCP_Valor"

Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "Valor_IPI"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const VALOR_TOTAL As String = "Valor_Total"
'###########Almoxarifado############
Private Const PRODUTO_ALMOX As String = "Produto_Almox"
Private Const ALMOX1 As String = "Almoxarifado"
Private Const QUANT_ALOCADA As String = "Quant_Alocada"
Private Const UNIDADE_MED_EST As String = "Unidade_Med_Est"
'###########Tributação##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque" '??? ERRADO: nao está no bd
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
Private Const CONTACONTABILPROD As String = "ContaContabilProd"
'fim da contabilidade

'ver conceito de filial que fatura vs filial do pedido e de manter todo o historico p/poder "reimprimir" uma NF

'ver p/setar sIPICodProduto e numintdoc (do item da nf) a nivel de item trib nf
'incluir coluna no grid p/transportadora

'a principio nao poderá entrar como EMPRESA_TODA

'incluir outros controles p/:
    'transportadora, status (pulada por falta de estoque, pulada por bloqueio,...)

'trocar sNaturezaOpEntrada p/sNaturezaOpInterna em ClassTributacaoNF, no bd, type, etc.
'natop, deveria estar em tributacaoPV e nao em PV (ou nos dois)

'Será que com a checkbox de "todos os pedidos" marcada devo incluir mesmo os pedidos bloqueados ?
'':acho que nao
''
''
'' se o flag de "só fatura tudo" estiver setado
''    Se tem "Bloqueio Parcial" ou "Nao Reserva" pular o pedido deixando-o marcado como "faltou estoque p/poder faturar" (vamos guardar a data da ultima tentativa frustrada)
''
''0. Se tem bloqueio total ou outro bloqueio que nao seja de estoque (Credito, Endereco,...) pular o pedido.
''
''1. Para os Pedidos que tem SoFaturaTudo,
''    Se tem "Bloqueio Parcial" ou "Nao Reserva" pular o pedido deixando-o marcado como "faltou estoque p/poder faturar" (vamos guardar a data da ultima tentativa frustrada)
''    Senao, tenta gerar a NF atendendo a todos os itens completamente, se nao conseguir pular o pedido deixando-o marcado como "faltou estoque p/poder faturar" (vamos guardar a data da ultima tentativa frustrada)
''
''2. Para os Pedidos que não têm SoFaturaTudo,
''2.1. Para os ítens do tipo "estoque (sem reserva)",  tenta  "tirar" QAR no almoxarifado default. Se nao conseguir abrir dialogo.
''2.2. Para os itens "reserva+estoque", tenta "tirar" a qtde reservada de acordo com a reserva. Se a qtde reservada for zero, pular o item.
''Se conseguir, libera as reservas, senao erro (nao fatura o pedido, (vamos guardar a data da ultima tentativa frustrada)).
''2.3. Se houver algum item OK p/faturamento pode fatura-lo. No futuro podemos incluir algum criterio p/faturar apenas qdo o volume a ser faturado for "expressivo" em relacao ao total que poderia ser faturado se nao houvesse falta de estoque.
''
''Observacao:
''     a geracao da NF já deveria poder ler no cadastro de pedidos, diretamente, se o pedido é "faturável" ou nao.
        'numa versao melhorada poderia manter controle p/só faturar algo "significativo" em relacao ao pedido
''
'--- inicio dos properties get dos controles da tela

Public Property Get BotaoExportar() As Object
     Set BotaoExportar = objUserControl.Controls("BotaoExportar")
End Property

Public Property Get BotaoDesmarcar() As Object
     Set BotaoDesmarcar = objUserControl.Controls("BotaoDesmarcar")
End Property

Public Property Get BotaoMarcar() As Object
     Set BotaoMarcar = objUserControl.Controls("BotaoMarcar")
End Property

Public Property Get ListRegioes() As Object
     Set ListRegioes = objUserControl.Controls("ListRegioes")
End Property

Public Property Get VendedorFinal() As Object
     Set VendedorFinal = objUserControl.Controls("VendedorFinal")
End Property

Public Property Get VendedorInicial() As Object
     Set VendedorInicial = objUserControl.Controls("VendedorInicial")
End Property

Public Property Get LabelVendedorDe() As Object
     Set LabelVendedorDe = objUserControl.Controls("LabelVendedorDe")
End Property

Public Property Get LabelVendedorAte() As Object
     Set LabelVendedorAte = objUserControl.Controls("LabelVendedorAte")
End Property

Public Property Get optImprimirRomaneio() As Object
     Set optImprimirRomaneio = objUserControl.Controls("optImprimirRomaneio")
End Property

Public Property Get LabelViagem() As Object
     Set LabelViagem = objUserControl.Controls("LabelViagem")
End Property

Public Property Get CodigoViagem() As Object
     Set CodigoViagem = objUserControl.Controls("CodigoViagem")
End Property

Public Property Get BotaoImprimirPI() As Object
     Set BotaoImprimirPI = objUserControl.Controls("BotaoImprimirPI")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get DataEmissaoDe() As Object
     Set DataEmissaoDe = objUserControl.Controls("DataEmissaoDe")
End Property

Public Property Get UpDownEmissaoDe() As Object
     Set UpDownEmissaoDe = objUserControl.Controls("UpDownEmissaoDe")
End Property

Public Property Get DataEmissaoAte() As Object
     Set DataEmissaoAte = objUserControl.Controls("DataEmissaoAte")
End Property

Public Property Get UpDownEmissaoAte() As Object
     Set UpDownEmissaoAte = objUserControl.Controls("UpDownEmissaoAte")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get PedidoInicial() As Object
     Set PedidoInicial = objUserControl.Controls("PedidoInicial")
End Property

Public Property Get PedidoFinal() As Object
     Set PedidoFinal = objUserControl.Controls("PedidoFinal")
End Property

Public Property Get LabelPedidoDe() As Object
     Set LabelPedidoDe = objUserControl.Controls("LabelPedidoDe")
End Property

Public Property Get LabelPedidoAte() As Object
     Set LabelPedidoAte = objUserControl.Controls("LabelPedidoAte")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get ClienteDe() As Object
     Set ClienteDe = objUserControl.Controls("ClienteDe")
End Property

Public Property Get ClienteAte() As Object
     Set ClienteAte = objUserControl.Controls("ClienteAte")
End Property

Public Property Get LabelClienteAte() As Object
     Set LabelClienteAte = objUserControl.Controls("LabelClienteAte")
End Property

Public Property Get LabelClienteDe() As Object
     Set LabelClienteDe = objUserControl.Controls("LabelClienteDe")
End Property

Public Property Get ExibeTodos() As Object
     Set ExibeTodos = objUserControl.Controls("ExibeTodos")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get DataEntregaDe() As Object
     Set DataEntregaDe = objUserControl.Controls("DataEntregaDe")
End Property

Public Property Get UpDownEntregaDe() As Object
     Set UpDownEntregaDe = objUserControl.Controls("UpDownEntregaDe")
End Property

Public Property Get DataEntregaAte() As Object
     Set DataEntregaAte = objUserControl.Controls("DataEntregaAte")
End Property

Public Property Get UpDownEntregaAte() As Object
     Set UpDownEntregaAte = objUserControl.Controls("UpDownEntregaAte")
End Property

Public Property Get CTBLancAutomatico() As Object
     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
End Property

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBFrame7() As Object
     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
End Property

Public Property Get CTBCclDescricao() As Object
     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
End Property

Public Property Get CTBContaDescricao() As Object
     Set CTBContaDescricao = objUserControl.Controls("CTBContaDescricao")
End Property

Public Property Get CTBLabel7() As Object
     Set CTBLabel7 = objUserControl.Controls("CTBLabel7")
End Property

Public Property Get CTBCclLabel() As Object
     Set CTBCclLabel = objUserControl.Controls("CTBCclLabel")
End Property

Public Property Get CTBModelo() As Object
     Set CTBModelo = objUserControl.Controls("CTBModelo")
End Property

Public Property Get CTBBotaoLimparGrid() As Object
     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
End Property

Public Property Get CTBBotaoModeloPadrao() As Object
     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
End Property

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
End Property

Public Property Get CTBConta() As Object
     Set CTBConta = objUserControl.Controls("CTBConta")
End Property

Public Property Get CTBDebito() As Object
     Set CTBDebito = objUserControl.Controls("CTBDebito")
End Property

Public Property Get CTBCredito() As Object
     Set CTBCredito = objUserControl.Controls("CTBCredito")
End Property

Public Property Get CTBCcl() As Object
     Set CTBCcl = objUserControl.Controls("CTBCcl")
End Property

Public Property Get CTBUpDown3() As Object
     Set CTBUpDown3 = objUserControl.Controls("CTBUpDown3")
End Property

Public Property Get CTBDataContabil3() As Object
     Set CTBDataContabil3 = objUserControl.Controls("CTBDataContabil3")
End Property

Public Property Get CTBLote3() As Object
     Set CTBLote3 = objUserControl.Controls("CTBLote3")
End Property

Public Property Get CTBDocumento() As Object
     Set CTBDocumento = objUserControl.Controls("CTBDocumento")
End Property

Public Property Get CTBTvwCcls() As Object
     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
End Property

Public Property Get CTBTvwContas() As Object
     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
End Property

Public Property Get CTBGridContabil() As Object
     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
End Property

Public Property Get CTBLabelLote3() As Object
     Set CTBLabelLote3 = objUserControl.Controls("CTBLabelLote3")
End Property

Public Property Get CTBLabelDoc() As Object
     Set CTBLabelDoc = objUserControl.Controls("CTBLabelDoc")
End Property

Public Property Get CTBLabel8() As Object
     Set CTBLabel8 = objUserControl.Controls("CTBLabel8")
End Property

Public Property Get CTBTotalCredito() As Object
     Set CTBTotalCredito = objUserControl.Controls("CTBTotalCredito")
End Property

Public Property Get CTBTotalDebito() As Object
     Set CTBTotalDebito = objUserControl.Controls("CTBTotalDebito")
End Property

Public Property Get CTBLabelTotais() As Object
     Set CTBLabelTotais = objUserControl.Controls("CTBLabelTotais")
End Property

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
End Property

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
End Property

Public Property Get CTBLabelContas() As Object
     Set CTBLabelContas = objUserControl.Controls("CTBLabelContas")
End Property

Public Property Get CTBLabelHistoricos() As Object
     Set CTBLabelHistoricos = objUserControl.Controls("CTBLabelHistoricos")
End Property

Public Property Get CTBLabel5() As Object
     Set CTBLabel5 = objUserControl.Controls("CTBLabel5")
End Property

Public Property Get CTBLabel13() As Object
     Set CTBLabel13 = objUserControl.Controls("CTBLabel13")
End Property

Public Property Get CTBExercicio() As Object
     Set CTBExercicio = objUserControl.Controls("CTBExercicio")
End Property

Public Property Get CTBPeriodo() As Object
     Set CTBPeriodo = objUserControl.Controls("CTBPeriodo")
End Property

Public Property Get CTBLabel14() As Object
     Set CTBLabel14 = objUserControl.Controls("CTBLabel14")
End Property

Public Property Get CTBOrigem() As Object
     Set CTBOrigem = objUserControl.Controls("CTBOrigem")
End Property

Public Property Get CTBLabel21() As Object
     Set CTBLabel21 = objUserControl.Controls("CTBLabel21")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get UpDownSaida() As Object
     Set UpDownSaida = objUserControl.Controls("UpDownSaida")
End Property

Public Property Get Motivo() As Object
     Set Motivo = objUserControl.Controls("Motivo")
End Property

Public Property Get Transportadora() As Object
     Set Transportadora = objUserControl.Controls("Transportadora")
End Property

Public Property Get Serie() As Object
     Set Serie = objUserControl.Controls("Serie")
End Property

Public Property Get BotaoNFiscalFaturaImprime() As Object
     Set BotaoNFiscalFaturaImprime = objUserControl.Controls("BotaoNFiscalFaturaImprime")
End Property

Public Property Get BotaoNFiscalFatura() As Object
     Set BotaoNFiscalFatura = objUserControl.Controls("BotaoNFiscalFatura")
End Property

Public Property Get Bairro() As Object
     Set Bairro = objUserControl.Controls("Bairro")
End Property

Public Property Get BotaoPedido() As Object
     Set BotaoPedido = objUserControl.Controls("BotaoPedido")
End Property

Public Property Get Ordenados() As Object
     Set Ordenados = objUserControl.Controls("Ordenados")
End Property

Public Property Get BotaoDesmarcarTodos() As Object
     Set BotaoDesmarcarTodos = objUserControl.Controls("BotaoDesmarcarTodos")
End Property

Public Property Get BotaoMarcarTodos() As Object
     Set BotaoMarcarTodos = objUserControl.Controls("BotaoMarcarTodos")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get Pedido() As Object
     Set Pedido = objUserControl.Controls("Pedido")
End Property

Public Property Get GeraNFiscal() As Object
     Set GeraNFiscal = objUserControl.Controls("GeraNFiscal")
End Property

Public Property Get BotaoNFiscal() As Object
     Set BotaoNFiscal = objUserControl.Controls("BotaoNFiscal")
End Property

Public Property Get BotaoNFiscalImprime() As Object
     Set BotaoNFiscalImprime = objUserControl.Controls("BotaoNFiscalImprime")
End Property

Public Property Get DataEntrega() As Object
     Set DataEntrega = objUserControl.Controls("DataEntrega")
End Property

Public Property Get NomeReduzido() As Object
     Set NomeReduzido = objUserControl.Controls("NomeReduzido")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Estado() As Object
     Set Estado = objUserControl.Controls("Estado")
End Property

Public Property Get Cidade() As Object
     Set Cidade = objUserControl.Controls("Cidade")
End Property

Public Property Get FilialEmpresa() As Object
     Set FilialEmpresa = objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get GridPedido() As Object
     Set GridPedido = objUserControl.Controls("GridPedido")
End Property

Public Property Get DataSaida() As Object
     Set DataSaida = objUserControl.Controls("DataSaida")
End Property

Public Property Get CTBUpDown() As Object
     Set CTBUpDown = objUserControl.Controls("CTBUpDown")
End Property

Public Property Get CTBDataContabil() As Object
     Set CTBDataContabil = objUserControl.Controls("CTBDataContabil")
End Property

Public Property Get CTBLote() As Object
     Set CTBLote = objUserControl.Controls("CTBLote")
End Property

Public Property Get GridAlocacao() As Object
     Set GridAlocacao = objUserControl.Controls("GridAlocacao")
End Property

Public Property Get CTBLabelLote() As Object
     Set CTBLabelLote = objUserControl.Controls("CTBLabelLote")
End Property

Public Property Get LabelDataContabil() As Object
     Set LabelDataContabil = objUserControl.Controls("LabelDataContabil")
End Property

Public Property Get label() As Object
     Set label = objUserControl.Controls("label")
End Property

Public Property Get TabStrip1() As Object
     Set TabStrip1 = objUserControl.Controls("TabStrip1")
End Property

'################################################
'Inserido por Wagner
Public Property Get Label1() As Object
     Set Label1 = objUserControl.Controls("Label1")
End Property
'################################################

'--- fim dos properties get dos controles da tela

Public Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os pedidos do Grid

Dim iLinha As Integer

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Desmarca na tela o pedido em questão
        GridPedido.TextMatrix(iLinha, iGrid_GeraNFiscal_Col) = S_DESMARCADO

        'Desmarca no Obj o pedido em questão
        gobjGeracaoNFiscal.colNFiscalInfo.Item(iLinha).iMarcada = S_DESMARCADO

    Next

    'Atualiza na tela a checkbox desmarcada
    Call Grid_Refresh_Checkbox(objGrid)

End Sub

Public Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub BotaoMarcarTodos_Click()
'Marca todos os pedidos do Grid

Dim iLinha As Integer
Dim objNFiscalInfo As ClassNFiscalInfo

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes

        'Marca na tela o pedido em questão
        GridPedido.TextMatrix(iLinha, iGrid_GeraNFiscal_Col) = S_MARCADO

        gobjGeracaoNFiscal.colNFiscalInfo.Item(iLinha).iMarcada = S_MARCADO

    Next

    'Atualiza na tela a checkbox marcada
    Call Grid_Refresh_Checkbox(objGrid)

End Sub

Private Function Inicializa_Grid_Pedido(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Gera NF")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Nome")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Estado Entrega")
    objGridInt.colColuna.Add ("Cidade Entrega")
    objGridInt.colColuna.Add ("Bairro Entrega")
    objGridInt.colColuna.Add ("Transportadora")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Entrega")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Erro na Geração")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (GeraNFiscal.Name)
    objGridInt.colCampo.Add (Pedido.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (NomeReduzido.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Estado.Name)
    objGridInt.colCampo.Add (Cidade.Name)
    objGridInt.colCampo.Add (Bairro.Name)
    objGridInt.colCampo.Add (Transportadora.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (DataEntrega.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (Motivo.Name)

    'Colunas do Grid
    iGrid_GeraNFiscal_Col = 1
    iGrid_Pedido_Col = 2
    iGrid_Cliente_Col = 3
    iGrid_NomeRed_Col = 4
    iGrid_Filial_Col = 5
    iGrid_Estado_Col = 6
    iGrid_Cidade_Col = 7
    iGrid_Bairro_Col = 8
    iGrid_TransPortadora_Col = 9
    iGrid_Emissao_Col = 10
    iGrid_Entrega_Col = 11
    iGrid_Valor_Col = 12
    iGrid_FilialEmpresa_Col = 13
    iGrid_Motivo_Col = 14

    'Grid do GridInterno
    objGridInt.objGrid = GridPedido

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1

    'Largura da primeira coluna
    GridPedido.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    '???? Estava permitindo excluir linhas do grid de pedidos
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'GridPedido.Width = 8400

    Inicializa_Grid_Pedido = SUCESSO

    Exit Function

End Function

Public Sub BotaoNFiscal_Click()

Dim lErro As Long
Dim objNFiscalInfo As ClassNFiscalInfo
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colNumIntNF As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_BotaoNFiscal_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    objFilialEmpresa.lCodEmpresa = glEmpresa
    'Lê a FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 51354
    
    'Atribui o tipo de NotaFiscal Interna de Saída de Venda a nota a
    'ser gerada
    If objFilialEmpresa.iTipoTribAtividade = 3 Then
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISSPV
    Else
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISVPV
    End If
    
    objTipoDocInfo.iCodigo = gobjGeracaoNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 51354

    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) = 0 Then Error 51354
    
    'Recolhe a série e a Data de Saída
    gobjGeracaoNFiscal.sSerie = Serie.Text
    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
    gobjGeracaoNFiscal.iImprime = 0

    lErro = GeracaoNF_Prepara_CTB
    If lErro <> SUCESSO Then Error 59382

    'Chama a rotina que gera as notas ficais a partir dos pedidos selecionados
    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
    If lErro <> SUCESSO Then Error 44197
    
    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
          
        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
            If ISSerieEletronica(Serie.Text) Then
                colNumIntNF.Add objNFiscalInfo.lNumIntDoc
            End If
        End If
    Next
        
    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
    Call Grid_Limpa(objGrid)

    If colNumIntNF.Count > 0 Then
        lErro = CF("NFE_Grava1", 0, objTipoDocInfo, colNumIntNF)
        If lErro <> SUCESSO Then Error 51431
    End If
    
    If optImprimirRomaneio.Value = vbChecked Then
        lErro = Viagem_Imprime(StrParaLong(CodigoViagem.Text))
        If lErro <> SUCESSO Then Error 51431
    End If
    
    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
    If lErro = 51429 Then Error 51431
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoNFiscal_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 44197, 51427, 59382
        
        Case 51354
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case 51431

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155885)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoNFiscalFatura_Click()

Dim lErro As Long
Dim objNFiscalInfo As ClassNFiscalInfo
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colNumIntNF As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_BotaoNFiscalFatura_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    objFilialEmpresa.lCodEmpresa = glEmpresa
    'Lê a FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 59383
            
    'Atribui o tipo de Nota Fiscal Interna de Saida de Fatura de Venda
    'para as notas a serem geradas
    If objFilialEmpresa.iTipoTribAtividade = 3 Then
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFSPV
    Else
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
    End If
    
    objTipoDocInfo.iCodigo = gobjGeracaoNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 59383
    
    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) = 0 Then Error 51437

    'Recolhe a série e a data de saída da tela
    gobjGeracaoNFiscal.sSerie = Serie.Text
    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
    gobjGeracaoNFiscal.iImprime = 0
    
    lErro = GeracaoNF_Prepara_CTB
    If lErro <> SUCESSO Then Error 59383

    'Chama a rotina que gera as NFs a partir de pedidos
    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
    If lErro <> SUCESSO Then Error 44198
    
    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
          
        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
            If ISSerieEletronica(Serie.Text) Then
                colNumIntNF.Add objNFiscalInfo.lNumIntDoc
            End If
        End If
    Next
    
    
    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
    Call Grid_Limpa(objGrid)
    
    If colNumIntNF.Count > 0 Then
        lErro = CF("NFE_Grava1", 0, objTipoDocInfo, colNumIntNF)
        If lErro <> SUCESSO Then Error 51432
    End If
    
    If optImprimirRomaneio.Value = vbChecked Then
        lErro = Viagem_Imprime(StrParaLong(CodigoViagem.Text))
        If lErro <> SUCESSO Then Error 51432
    End If
    
    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
    If lErro <> SUCESSO Then Error 51432
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoNFiscalFatura_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 44198, 51425, 59383

        Case 51432

        Case 51437
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155886)

    End Select

    Exit Sub

End Sub

Public Sub BotaoNFiscalFaturaImprime_Click()

Dim lErro As Long
Dim objNFiscalInfo As ClassNFiscalInfo
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colNumIntNF As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_BotaoNFiscalFaturaImprime_Click

    GL_objMDIForm.MousePointer = vbHourglass

    objFilialEmpresa.iCodFilial = giFilialEmpresa
    objFilialEmpresa.lCodEmpresa = glEmpresa
    'Lê a FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 59384
    
    'Atribui o tipo de NotaFiscal Interna de Saída Fatura de Venda a nota a
    'ser gerada
    If objFilialEmpresa.iTipoTribAtividade = 3 Then
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFSPV
    Else
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV
    End If
    
    objTipoDocInfo.iCodigo = gobjGeracaoNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 59384

    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) = 0 Then Error 51438
    
    'Recolhe a série e a Data de Saída
    gobjGeracaoNFiscal.sSerie = Serie.Text
    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
    gobjGeracaoNFiscal.iImprime = 1
    
    lErro = GeracaoNF_Prepara_CTB
    If lErro <> SUCESSO Then Error 59384

    'Chama a rotina que gera as notas ficais a partir dos pedidos selecionados
    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
    If lErro <> SUCESSO Then Error 44199
    
    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
          
        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
            If ISSerieEletronica(Serie.Text) Then
                colNumIntNF.Add objNFiscalInfo.lNumIntDoc
            End If
        End If
    
    Next
    
    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
    Call Grid_Limpa(objGrid)
    
    If colNumIntNF.Count > 0 Then
        lErro = CF("NFE_Grava1", 0, objTipoDocInfo, colNumIntNF)
        If lErro <> SUCESSO Then Error 51440
    End If
    
    If optImprimirRomaneio.Value = vbChecked Then
        lErro = Viagem_Imprime(StrParaLong(CodigoViagem.Text))
        If lErro <> SUCESSO Then Error 51440
    End If
    
    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
    If lErro <> SUCESSO Then Error 51440
    
    Call NotaFiscal_Imprime
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoNFiscalFaturaImprime_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 44199, 51439, 51440, 59384
        
        Case 51438
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155887)

    End Select

    Exit Sub

End Sub

Public Sub BotaoNFiscalImprime_Click()

Dim lErro As Long
Dim objNFiscalInfo As ClassNFiscalInfo
Dim iIndice As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim colNumIntNF As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_BotaoNFiscal_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    objFilialEmpresa.iCodFilial = giFilialEmpresa
    objFilialEmpresa.lCodEmpresa = glEmpresa
    'Lê a FilialEmpresa
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then Error 51441
    
    'Atribui o tipo de NotaFiscal Interna de Saída de Venda a nota a
    'ser gerada
    If objFilialEmpresa.iTipoTribAtividade = 3 Then
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISSPV
    Else
        gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISVPV
    End If
    
    objTipoDocInfo.iCodigo = gobjGeracaoNFiscal.iTipoNFiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 51441

    'Verifica se a série está preenchida
    If Len(Trim(Serie.Text)) = 0 Then Error 51441

    'Recolhe a série e a Data de Saída
    gobjGeracaoNFiscal.sSerie = Serie.Text
    gobjGeracaoNFiscal.dtDataSaida = StrParaDate(DataSaida.Text)
    gobjGeracaoNFiscal.iImprime = 1
    
    lErro = GeracaoNF_Prepara_CTB
    If lErro <> SUCESSO Then Error 59385

    'Chama a rotina que gera as notas ficais a partir dos pedidos selecionados
    lErro = CF("GeracaoNFiscal_GerarNFs", gobjGeracaoNFiscal)
    If lErro <> SUCESSO Then Error 44200
    
    For iIndice = gobjGeracaoNFiscal.colNFiscalInfo.Count To 1 Step -1
        Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo(iIndice)
          
        If objNFiscalInfo.iMarcada = MARCADO And objNFiscalInfo.iMotivoNaoGerada = 0 Then
            gobjGeracaoNFiscal.colNFiscalInfo.Remove iIndice
            If ISSerieEletronica(Serie.Text) Then
                colNumIntNF.Add objNFiscalInfo.lNumIntDoc
            End If
        End If
    Next
    
    'Recarrega o grid de Pedidos excluíndo aqueles que já geraram NFs
    Call Grid_Limpa(objGrid)

    If colNumIntNF.Count > 0 Then
        lErro = CF("NFE_Grava1", 0, objTipoDocInfo, colNumIntNF)
        If lErro <> SUCESSO Then Error 51442
    End If
    
    If optImprimirRomaneio.Value = vbChecked Then
        lErro = Viagem_Imprime(StrParaLong(CodigoViagem.Text))
        If lErro <> SUCESSO Then Error 51442
    End If
    
    lErro = Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
    If lErro <> SUCESSO Then Error 51442
        
    'Exibe os pedidos desselecionados
    Call BotaoDesmarcarTodos_Click
    
    Call NotaFiscal_Imprime
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoNFiscal_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 44200, 51442, 59385
        
        Case 51441
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155888)

    End Select

    Exit Sub

End Sub

Public Sub BotaoPedido_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iAchou As Integer
Dim objNFiscalInfo As New ClassNFiscalInfo
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_BotaoPedido_Click
    
    If objGrid.iLinhasExistentes = 0 Then Exit Sub
    
    'Se nenhuma linha válida estiver com o foco
    If GridPedido.Row < 1 Or GridPedido.Row > objGrid.iLinhasExistentes Then Error 51453
    
    'Passa a linha do Grid para o Obj
    Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo.Item(GridPedido.Row)

    'Passa os dados do NFiscal para o Obj
    objPedidoDeVenda.iFilialEmpresa = objNFiscalInfo.iFilialEmpresa
    objPedidoDeVenda.lCodigo = objNFiscalInfo.lPedido

    If objPedidoDeVenda.iFilialEmpresa <> giFilialEmpresa Then Error 51454

    'Chama a tela de Pedidos de Venda
    Call Chama_Tela("PedidoVenda", objPedidoDeVenda)

    Exit Sub

Erro_BotaoPedido_Click:

    Select Case Err

        Case 51453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_SELECIONADO", Err)
        
        Case 51454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_DIFERENTE_FILIALEMPRESA", Err, objPedidoDeVenda.lCodigo, objPedidoDeVenda.iFilialEmpresa, giFilialEmpresa)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155889)

    End Select

    Exit Sub

End Sub

Public Sub ClienteAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    
End Sub


Public Sub ClienteAte_GotFocus()
Dim iTabAux As Integer
Dim iClienteAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    iClienteAux = iClienteAlterado
    
    Call MaskEdBox_TrataGotFocus(ClienteAte, iAlterado)
    iTabPrincipalAlterado = iTabAux
    iClienteAlterado = iClienteAux

End Sub

Public Sub ClienteAte_Validate(Cancel As Boolean)
'Verifica se o Cliente De é maior que o Cliente Até
'Verifica a integridade do cliente com o BD

Dim lErro As Long
Dim objClienteAte As New ClassCliente
Dim iCodFilial As Integer
Dim iCria As Integer
Dim colCodigoNome As AdmColCodigoNome

On Error GoTo Erro_ClienteAte_Validate
    
    If iClienteAlterado = 1 Then

        If Len(Trim(ClienteAte.Text)) > 0 Then
            
            'Se o Cliente De estiver preenchido
            If Len(Trim(ClienteDe.Text)) > 0 Then
                'Verifica se o Cliente De é maior que o Cliente Até ----->>> Erro
                If LCodigo_Extrai(ClienteDe.Text) > LCodigo_Extrai(ClienteAte.Text) Then Error 58014
                
            End If
            
            objClienteAte.lCodigo = ClienteAte.Text
            
            'Le o Cliente para testar sua integridade com o BD
            lErro = CF("Cliente_Le", objClienteAte)
            If lErro <> SUCESSO And lErro <> 12293 Then Error 58015
            
            'Se não encontrou ----> erro
            If lErro = 12293 Then Error 58016
            
        End If

        iClienteAlterado = 0

    End If
    
    Exit Sub
    
Erro_ClienteAte_Validate:

    Cancel = True

    Select Case Err
    
    Case 58014
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", Err)
    
    Case 58015 'Tratados nas rotinas chamadas
        
    Case 58016
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objClienteAte.lCodigo)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155890)

    End Select
    
    Exit Sub
    
End Sub

Public Sub ClienteDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ClienteDe_GotFocus()
Dim iTabAux As Integer
Dim iClienteAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    iClienteAux = iClienteAlterado
    
    Call MaskEdBox_TrataGotFocus(ClienteDe, iAlterado)
    iTabPrincipalAlterado = iTabAux
    iClienteAlterado = iClienteAux

End Sub

Public Sub ClienteDe_Validate(Cancel As Boolean)
'Verifica se o Cliente De é maior que o Cliente Até
'Verifica a integridade do cliente com o BD

Dim lErro As Long
Dim objClienteDe As New ClassCliente
Dim iCodFilial As Integer
Dim iCria As Integer
Dim colCodigoNome As AdmColCodigoNome

On Error GoTo Erro_ClienteDe_Validate
    
    'Se o ClienteDE não foi alterado --> sai
    If iClienteAlterado = 0 Then Exit Sub
    'Se algum clientefoi informado
    If Len(Trim(ClienteDe.Text)) > 0 Then
        'Se o cliente até estiver preenchido
        If Len(Trim(ClienteAte.Text)) > 0 Then
            'Verifica se o cliente De é menor que o cliente até
            If LCodigo_Extrai(ClienteDe.Text) > LCodigo_Extrai(ClienteAte.Text) Then Error 58011
        End If
        
        objClienteDe.lCodigo = StrParaLong(ClienteDe.Text)
        'Lê o cliente informado
        lErro = CF("Cliente_Le", objClienteDe)
        If lErro <> SUCESSO And lErro <> 12293 Then Error 58012
        If lErro = 12293 Then Error 58013 'Não encontrou
        
    End If
    'Zera flag de alteração de cliente de.
    iClienteAlterado = 0
    
    Exit Sub
    
Erro_ClienteDe_Validate:
    
    Cancel = True
    
    Select Case Err
    
    Case 58011
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", Err)
        
    Case 58012 'Tratados nas rotinas chamadas
    
    Case 58013
        Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, objClienteDe.lCodigo)

    Case Else
        Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155891)

    End Select
    
    Exit Sub
    
End Sub

Public Sub CTBUpDown_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_CTBUpDown_DownClick

    lErro = Data_Up_Down_Click(CTBDataContabil, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 71544

    Exit Sub

Erro_CTBUpDown_DownClick:

    Select Case gErr

        Case 71544

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155892)

    End Select

    Exit Sub

End Sub

Public Sub CTBUpDown_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_CTBUpDown_UpClick

    lErro = Data_Up_Down_Click(CTBDataContabil, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 71545

    Exit Sub

Erro_CTBUpDown_UpClick:

    Select Case gErr

        Case 71545

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155893)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissaoAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissaoAte_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub DataEmissaoAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Se a DataEmissaoAte está preenchida
    If Len(DataEmissaoAte.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoAte é válida
    lErro = Data_Critica(DataEmissaoAte.Text)
    If lErro <> SUCESSO Then Error 28458
    
    'Verifica se a data de emissao de está preenchida
    If Len(DataEmissaoDe.ClipText) = 0 Then Exit Sub
    
    'Verifica se a data emissão de é maior que a Data de emissão até
    If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then Error 58020

    Exit Sub

Erro_DataEmissaoAte_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 28458
        
        Case 58020
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
    
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155894)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissaoDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissaoDe_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub DataEmissaoDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Se a DataEmissaoDe está preenchida
    If Len(DataEmissaoDe.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEmissaoDe é válida
    lErro = Data_Critica(DataEmissaoDe.Text)
    If lErro <> SUCESSO Then Error 28457

    If Len(Trim(DataEmissaoAte.ClipText)) = 0 Then Exit Sub

    If CDate(DataEmissaoDe.Text) > CDate(DataEmissaoAte.Text) Then Error 31385

    Exit Sub

Erro_DataEmissaoDe_Validate:
    
    Cancel = True
    
    Select Case Err
    
        Case 28457
    
        Case 31385
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
    
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155895)

    End Select

    Exit Sub

End Sub

Public Sub DataEntregaAte_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEntregaAte_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaAte, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub DataEntregaAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEntregaAte_Validate

    'Se a DataEntregaAte está preenchida
    If Len(DataEntregaAte.ClipText) = 0 Then Exit Sub

    'Verifica se a DataEntregaAte é válida
    lErro = Data_Critica(DataEntregaAte.Text)
    If lErro <> SUCESSO Then Error 28464

    If Len(Trim(DataEntregaDe.ClipText)) = 0 Then Exit Sub

    If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then Error 31386

    Exit Sub

Erro_DataEntregaAte_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 28464

        Case 31386
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155896)

    End Select

    Exit Sub

End Sub

Public Sub DataEntregaDe_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEntregaDe_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(DataEntregaDe, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub DataEntregaDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEntregaDe_Validate

    'Se a DataEntregaDe está preenchida
    If Len(Trim(DataEntregaDe.ClipText)) = 0 Then Exit Sub

    'Verifica se a DataEntregaDe é válida
    lErro = Data_Critica(DataEntregaDe.Text)
    If lErro <> SUCESSO Then Error 28463
    
    If Len(Trim(DataEntregaAte.ClipText)) = 0 Then Exit Sub

    If CDate(DataEntregaDe.Text) > CDate(DataEntregaAte.Text) Then Error 58021

    Exit Sub

Erro_DataEntregaDe_Validate:
    
    Cancel = True
    
    Select Case Err

        Case 28464
        
        Case 58021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)
    
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155897)

    End Select

    Exit Sub

End Sub

Public Sub DataSaida_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataSaida_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataSaida, iAlterado)

End Sub

Public Sub ExibeTodos_Click()
'???? Quando a checkbox de "exibi todos os pedidos"  está selecionada todas as condições ( DataDe, DataAte, ClienteDe, ClienteAte…) são desabilitadas. OK. Mas só que os labels que chamam os browses esqueceram de ser desabilitados.

    iTabPrincipalAlterado = REGISTRO_ALTERADO

    'Limpa os campos da tela
    PedidoInicial.Text = ""
    PedidoFinal.Text = ""
    ClienteDe.Text = ""
    ClienteAte.Text = ""
    DataEmissaoDe.PromptInclude = False
    DataEmissaoDe.Text = ""
    DataEmissaoDe.PromptInclude = True
    DataEmissaoAte.PromptInclude = False
    DataEmissaoAte.Text = ""
    DataEmissaoAte.PromptInclude = True
    DataEntregaDe.PromptInclude = False
    DataEntregaDe.Text = ""
    DataEntregaDe.PromptInclude = True
    DataEntregaAte.PromptInclude = False
    DataEntregaAte.Text = ""
    DataEntregaAte.PromptInclude = True
    VendedorInicial.Text = ""
    VendedorFinal.Text = ""
    CodigoViagem.PromptInclude = False
    CodigoViagem.Text = ""
    CodigoViagem.PromptInclude = True
    Call BotaoDesmarcar_Click

    'Se marcar ExibeTodos, exibe todos os pedidos
    If ExibeTodos.Value = 1 Then
    
        PedidoInicial.Enabled = False
        PedidoFinal.Enabled = False
        ClienteDe.Enabled = False
        ClienteAte.Enabled = False
        DataEmissaoDe.Enabled = False
        DataEmissaoAte.Enabled = False
        DataEntregaDe.Enabled = False
        DataEntregaAte.Enabled = False
        UpDownEmissaoDe.Enabled = False
        UpDownEmissaoAte.Enabled = False
        UpDownEntregaDe.Enabled = False
        UpDownEntregaAte.Enabled = False
        '??? Esqueceu de desabilitar as Labels com browse
        LabelPedidoAte.Enabled = False
        LabelPedidoDe.Enabled = False
        LabelClienteAte.Enabled = False
        LabelClienteDe.Enabled = False
        Label1(0).Enabled = False
        Label1(1).Enabled = False
        Label1(2).Enabled = False
        Label1(3).Enabled = False
        VendedorInicial.Enabled = False
        VendedorFinal.Enabled = False
        LabelVendedorAte.Enabled = False
        LabelVendedorDe.Enabled = False
        ListRegioes.Enabled = False
        CodigoViagem.Enabled = False
        LabelViagem.Enabled = False
        
    Else
    
        PedidoInicial.Enabled = True
        PedidoFinal.Enabled = True
        ClienteDe.Enabled = True
        ClienteAte.Enabled = True
        DataEmissaoDe.Enabled = True
        DataEmissaoAte.Enabled = True
        DataEntregaDe.Enabled = True
        DataEntregaAte.Enabled = True
        UpDownEmissaoDe.Enabled = True
        UpDownEmissaoAte.Enabled = True
        UpDownEntregaDe.Enabled = True
        UpDownEntregaAte.Enabled = True
        '??? Esqueceu de habilitar as Labels com browse
        LabelPedidoAte.Enabled = True
        LabelPedidoDe.Enabled = True
        LabelClienteAte.Enabled = True
        LabelClienteDe.Enabled = True
        Label1(0).Enabled = True
        Label1(1).Enabled = True
        Label1(2).Enabled = True
        Label1(3).Enabled = True
        VendedorInicial.Enabled = True
        VendedorFinal.Enabled = True
        LabelVendedorAte.Enabled = True
        LabelVendedorDe.Enabled = True
        ListRegioes.Enabled = True
        CodigoViagem.Enabled = True
        LabelViagem.Enabled = True
    
    End If
    
    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    Set objEventoPedidoDe = New AdmEvento
    Set objEventoPedidoAte = New AdmEvento
    Set objEventoClienteDe = New AdmEvento
    Set objEventoClienteAte = New AdmEvento
    Set objEventoCodigoViagem = New AdmEvento
    Set objEventoVendedor = New AdmEvento

    asOrdenacao(0) = " FilialEmpresaPV, CodigoPV"
    asOrdenacao(1) = " NomeCliente, FilialEmpresaPV, CodigoPV"
    asOrdenacao(2) = " EmissaoPedido , FilialEmpresaPV, CodigoPV"
    asOrdenacao(3) = " SiglaEstadoEntrega, CidadeEntrega,  BairroEntrega "

    asOrdenacaoString(0) = "Filial da Empresa + Pedido"
    asOrdenacaoString(1) = "Cliente + Filial da Empresa + Pedido"
    asOrdenacaoString(2) = "Data de Emissão do Pedido + Filial da Empresa + Pedido"
    asOrdenacaoString(3) = "Estado + Cidade + Bairro"

    iFrameAtual = 1
    giVendedorInicial = 1

    DataSaida.PromptInclude = False
    DataSaida.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataSaida.PromptInclude = True

    Set objGrid = New AdmGrid

    'Executa a Inicialização do grid Pedido
    lErro = Inicializa_Grid_Pedido(objGrid)
    If lErro <> SUCESSO Then Error 28480
    
    'Carrega as Séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then Error 19164

    'Carrega a Combobox Ordenados
    For iIndice = 0 To 3
        Ordenados.AddItem asOrdenacaoString(iIndice)
    Next

    Ordenados.ListIndex = 0
    
    Set objEventoLote = New AdmEvento
    Set objGrid1 = New AdmGrid
        
    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        'Inicialização da parte de contabilidade
        lErro = objContabil.Contabil_Inicializa_Contabilidade3(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_FATURAMENTO)
        If lErro <> SUCESSO Then Error 59408
        
        lErro = objContabil.Contabil_Gera_Cabecalho_Automatico
        If lErro <> SUCESSO Then Error 59411
        
    Else
        
        CTBDataContabil.Enabled = False
        LabelDataContabil.Enabled = False
    
    End If
    
    lErro = CF2(Me, "GNF_Form_Load")
    If lErro <> SUCESSO Then Error 59411
    
    Call CarregaList_Regioes
    
    If gobjFAT.iExibeBtnExportaPV = MARCADO Then
        BotaoExportar.Visible = True
    End If
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 28480, 19164, 59408, 59411

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155898)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Carrega_Serie() As Long
'Carrega a combo de Série

Dim lErro As Long
Dim colSerie As New colSerie
Dim objSerie As ClassSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then Error 19165

    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = Err

    Select Case Err

        Case 19165

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155899)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objGrid = Nothing
    
    Set objEventoPedidoDe = Nothing
    Set objEventoPedidoAte = Nothing
    Set objEventoClienteDe = Nothing
    Set objEventoClienteAte = Nothing
    Set objEventoCodigoViagem = Nothing
    Set objEventoVendedor = Nothing

    'variaveis auxiliares à contabilizacao
    Set gobjContabAutomatica = Nothing
    Set gobjNFiscal = Nothing
    Set gobjPedidoVenda = Nothing
    Set gcolAlmoxFilial = Nothing
    
    Set objEventoLote = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing
        
    Set gobjGeracaoNFiscal = Nothing

End Sub

Public Sub GeraNFiscal_Click()

Dim iClick As Integer

    iAlterado = REGISTRO_ALTERADO

    'Verifica se é alguma linha válida
    If GridPedido.Row > objGrid.iLinhasExistentes Then Exit Sub

    'Verifica se está selecionando ou desselecionando
    If Len(Trim(GridPedido.TextMatrix(GridPedido.Row, iGrid_GeraNFiscal_Col))) > 0 Then
        iClick = StrParaInt(GridPedido.TextMatrix(GridPedido.Row, iGrid_GeraNFiscal_Col)) = MARCADO
    End If

    If iClick = True Then
        gobjGeracaoNFiscal.colNFiscalInfo(GridPedido.Row).iMarcada = MARCADO
    Else
        gobjGeracaoNFiscal.colNFiscalInfo(GridPedido.Row).iMarcada = DESMARCADO
    End If

    Exit Sub

End Sub

Public Sub GeraNFiscal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub GeraNFiscal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub GeraNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = GeraNFiscal
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub GridPedido_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Public Sub GridPedido_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Public Sub GridPedido_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Public Sub GridPedido_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Public Sub GridPedido_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Public Sub GridPedido_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Public Sub GridPedido_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)

End Sub

Public Sub GridPedido_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Public Sub GridPedido_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Sub LabelClienteAte_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche ClienteAte com o cliente da tela
    If Len(Trim(ClienteAte.Text)) > 0 Then objCliente.lCodigo = CLng(ClienteAte.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteAte)

End Sub

Public Sub LabelClienteDe_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection

    'Preenche ClienteDe com o cliente da tela
    If Len(Trim(ClienteDe.Text)) > 0 Then objCliente.lCodigo = CLng(ClienteDe.Text)

    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteDe)

End Sub

Public Sub LabelPedidoAte_Click()

Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim colSelecao As Collection

    'Preenche PedidoAte com o pedido da tela
    If Len(Trim(PedidoFinal.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoFinal.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedidoVendaLista
    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoAte)

End Sub

Public Sub LabelPedidoDe_Click()

Dim objPedidoDeVenda As New ClassPedidoDeVenda
Dim colSelecao As Collection

    'Preenche PedidoDe com o pedido da tela
    If Len(Trim(PedidoInicial.Text)) > 0 Then objPedidoDeVenda.lCodigo = CLng(PedidoInicial.Text)

    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    'Chama Tela PedidoVendaLista
    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoDeVenda, objEventoPedidoDe)

End Sub

Private Sub objEventoClienteAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente
Dim bCancel As Boolean

On Error GoTo Erro_objEventoClienteAte_evSelecao

    If Not ClienteAte.Enabled Then Exit Sub

    Set objCliente = obj1

    ClienteAte.Text = CStr(objCliente.lCodigo)

    'Chama o Validate de ClienteAte
    Call ClienteAte_Validate(bCancel)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoClienteAte_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155900)

    End Select

    Exit Sub

End Sub

Private Sub objEventoClienteDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCliente As ClassCliente

On Error GoTo Erro_objEventoClienteDe_evSelecao

    Set objCliente = obj1

    'Coloca o código do cliente em cliente de
    ClienteDe.Text = CStr(objCliente.lCodigo)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoClienteDe_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155901)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPedidoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda
Dim bCancel As Boolean

On Error GoTo Erro_objEventoPedidoAte_evSelecao

    Set objPedidoDeVenda = obj1

    PedidoFinal.Text = CStr(objPedidoDeVenda.lCodigo)

    'Chama o Validate de PedidoFinal
    Call PedidoFinal_Validate(bCancel)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoPedidoAte_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155902)

    End Select

    Exit Sub

End Sub

Private Sub objEventoPedidoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPedidoDeVenda As ClassPedidoDeVenda

On Error GoTo Erro_objEventoPedidoDe_evSelecao

    Set objPedidoDeVenda = obj1

    PedidoInicial.Text = CStr(objPedidoDeVenda.lCodigo)

    iAlterado = 0

    Me.Show

    Exit Sub

Erro_objEventoPedidoDe_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155903)

    End Select

    Exit Sub

End Sub

Public Sub Ordenados_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Ordenados_Click()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Ordenados_Click

    If Ordenados.ListIndex = -1 Then Exit Sub

    'Verifica se a coleção de NFiscal está vazia
    If gobjGeracaoNFiscal.colNFiscalInfo.Count = 0 Then Exit Sub

    'Passa a Ordenaçao escolhida para o Obj
    gobjGeracaoNFiscal.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
        
    'Verifica se tem seleção e Preenche o Grid se Tiver
    lErro = Traz_Pedidos_Selecionados()
    If lErro <> SUCESSO And lErro <> 51429 Then Error 58024
    If lErro = 51429 Then Error 51430
    
    Exit Sub

Erro_Ordenados_Click:

    Select Case Err

        Case 28476, 58024
        
        Case 51430
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PEDIDOS_VENDA_ENCONTRADOS", Err)
    
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155904)

    End Select

    Exit Sub

End Sub

Function Traz_Pedidos_Selecionados() As Long

Dim lErro As Long

On Error GoTo Erro_Traz_Pedidos_Selecionados

    'Limpa a coleção de NFiscais
    Set gobjGeracaoNFiscal = New ClassGeracaoNFiscal

    '???? O grid não estáva sendo limpo quando as opções do tab seleção eram alteradas. E quando ele não achava ninguém no BD com as novas características a msg de que não encontrou ninguém era exibida mas apareciam registros no grid.
    'Limpa o GridPedido
    Call Grid_Limpa(objGrid)
    
    lErro = Move_TabSelecao_Memoria
    If lErro <> SUCESSO Then Error 51448
    
    'Verifica se foi feita alguma seleção
    If gobjGeracaoNFiscal.sRegioes = "" And gobjGeracaoNFiscal.iVendedorInicial = 0 And gobjGeracaoNFiscal.iVendedorFinal = 0 And ExibeTodos.Value = 0 And Len(Trim(PedidoInicial.Text)) = 0 And Len(Trim(PedidoFinal.Text)) = 0 And Len(Trim(ClienteDe.Text)) = 0 And Len(Trim(ClienteAte.Text)) = 0 And _
        Len(Trim(DataEmissaoDe.ClipText)) = 0 And Len(Trim(DataEmissaoAte.ClipText)) = 0 And Len(Trim(DataEntregaDe.ClipText)) = 0 And Len(Trim(DataEntregaAte.ClipText)) = 0 And Len(Trim(CodigoViagem.Text)) = 0 Then Exit Function
    
    '???? Marquei a opção de exibir todos e ele carregou o grid no 2º tab. Voltei ao  tab seleção e desmarquei a opção de exibir todos. Ao coltar p\ o 2º tab o grid está vazio, mas a coleção de pedidos ainda está carregada com os pedidos anteriormente selecionados
    'Preenche a Coleção de NFiscais
    lErro = CF("GeracaoNFiscal_ObterPedidos", gobjGeracaoNFiscal)
    If lErro <> SUCESSO And lErro <> 58166 Then Error 58023
    If lErro = 58166 Then Error 51428

    'Preenche o GridPedido
    Call Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)

    Traz_Pedidos_Selecionados = SUCESSO
    
    Exit Function
    
Erro_Traz_Pedidos_Selecionados:
    
    Traz_Pedidos_Selecionados = Err
    
    Select Case Err
    
    Case 58023, 51428, 51448 'Tratado na rotina chamada
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155905)

    End Select

    Exit Function

End Function

'Private Sub Ordenados_Validate(Cancel As Boolean)
'
'Dim lErro As Long
'Dim iIndice As Integer
'Dim iCodigo As Integer
'
'On Error GoTo Erro_Ordenados_Validate
'
'    'Verifica se a coleção de NFiscal está vazia
'    If gobjGeracaoNFiscal.colNFiscalInfo.Count <> 0 Then
'
'        'Verifica se foi preenchida a ComboBox Ordenados
'        If Len(Trim(Ordenados.Text)) = 0 Then Exit Sub
'
'        'Verifica se está preenchida com o ítem selecionado na ComboBox Ordenados
'        If Ordenados.Text = Ordenados.List(Ordenados.ListIndex) Then Exit Sub
'
'        'Verifica se existe o ítem na List da Combo. Se existir seleciona.
'        lErro = Combo_Seleciona(Ordenados, iCodigo)
'        If lErro <> SUCESSO And lErro <> 6731 Then Error 28477
'
'        'Não existe o ítem com a STRING na List da ComboBox
'        If lErro = 6731 Then Error 28478
'
'        'Passa a Ordenaçao escolhida para o Obj
'        For iIndice = 0 To 4
'
'            If Ordenados.Text = asOrdenacaoString(iIndice) Then gobjGeracaoNFiscal.sOrdenacao = asOrdenacao(iIndice)
'
'        Next
'
'        'Limpa a coleção de NFiscais
'        If Not (gobjGeracaoNFiscal.colNFiscalInfo Is Nothing) Then
'
'            Do While gobjGeracaoNFiscal.colNFiscalInfo.Count <> 0
'
'                gobjGeracaoNFiscal.colNFiscalInfo.Remove (1)
'
'            Loop
'
'        End If
'
'        'Preenche a Coleção de NFiscais
'        lErro = CF("GeracaoNFiscal_ObterPedidos",gobjGeracaoNFiscal)
'        If lErro <> SUCESSO Then Error 28479
'
'        'Limpa o GridPedido
'        Call Grid_Limpa(objGrid)
'        objGrid.iLinhasExistentes = 0
'
'        'Preenche o GridPedido
'        Call Grid_Pedido_Preenche(gobjGeracaoNFiscal.colNFiscalInfo)
'
'    End If
'
'    Exit Sub
'
'Erro_Ordenados_Validate:

'    Cancel = True

'
'    Select Case Err
'
'        Case 28477, 28479
'
'        Case 28478
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDENACAO_NAO_ENCONTRADA", Err)
'
'        Case Else
'             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155906)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub PedidoFinal_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PedidoFinal_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(PedidoFinal, iAlterado)
    iTabPrincipalAlterado = iTabAux

End Sub

Public Sub PedidoFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoFinal_Validate

    If Len(Trim(PedidoFinal.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(PedidoFinal.Text)
        If lErro <> SUCESSO Then Error 58007
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(PedidoInicial.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then Error 58008
        End If
            
        objPedidoVenda.lCodigo = CLng(PedidoFinal.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then Error 58009
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then Error 58010
        
    End If
       
    Exit Sub

Erro_PedidoFinal_Validate:
            
    Cancel = True

    Select Case Err
    
        Case 58007, 58009
            
        Case 58008
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)
            
        Case 58010
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedidoVenda.lCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155907)

    End Select

    Exit Sub

End Sub

Public Sub PedidoInicial_Change()

    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PedidoInicial_GotFocus()
Dim iTabAux As Integer
    
    iTabAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(PedidoInicial, iAlterado)
    iTabPrincipalAlterado = iTabAux
        
End Sub

Public Sub PedidoInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda

On Error GoTo Erro_PedidoInicial_Validate

    If Len(Trim(PedidoInicial.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(PedidoInicial.Text)
        If lErro <> SUCESSO Then Error 58003
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(PedidoFinal.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(PedidoInicial.Text) > CLng(PedidoFinal.Text) Then Error 58004
        End If
            
        objPedidoVenda.lCodigo = CLng(PedidoInicial.Text)
        objPedidoVenda.iFilialEmpresa = giFilialEmpresa
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
        If lErro <> SUCESSO And lErro <> 26509 Then Error 58005
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then Error 58006
        
    End If
       
    Exit Sub

Erro_PedidoInicial_Validate:
    
    Cancel = True

    Select Case Err
            
        Case 58003, 58005
        
        Case 58004
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)

        Case 58006
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO", Err, objPedidoVenda.lCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155908)

    End Select

    Exit Sub

End Sub

Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    'Verifica se foi preenchida
    If Len(Trim(Serie.Text)) = 0 Then Exit Sub
    
    'Verifica se foi selecionada
    If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub

    'Tenta selecionar a serie
    lErro = Combo_Item_Igual(Serie)
    If lErro <> SUCESSO And lErro <> 12253 Then Error 44194
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then Error 44195

    Exit Sub

Erro_Serie_Validate:

    Cancel = True


    Select Case Err

        Case 44194

        Case 44195
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155909)

    End Select

End Sub


Public Sub TabStrip1_Click()
        
Dim lErro As Long

On Error GoTo Erro_TabStrip1_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStrip1.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

        'Torna Frame de Pedido visível
        Frame1(TabStrip1.SelectedItem.Index).Visible = True
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStrip1.SelectedItem.Index

        'Se Frame selecionado foi o de Pedido
        If TabStrip1.SelectedItem.Index = TAB_Pedidos Then
            If iTabPrincipalAlterado = REGISTRO_ALTERADO Then
                lErro = Trata_TabPedidos()
                If lErro <> SUCESSO And lErro <> 51429 Then Error 31382
                If lErro <> SUCESSO Then Error 51433
            End If
        
        End If
        
        Select Case iFrameAtual
        
            Case TAB_Selecao
                Parent.HelpContextID = IDH_GERACAO_NFISCAL_SELECAO
                
            Case TAB_Pedidos
                Parent.HelpContextID = IDH_GERACAO_NFISCAL_PEDIDOS
                        
        End Select

    End If

    Exit Sub

Erro_TabStrip1_Click:

    Select Case Err

        Case 31382

        Case 51433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEM_PEDIDOS_VENDA_ENCONTRADOS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155910)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a DataEmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 28462

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case Err

        Case 28462

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155911)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a DataEmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 28461

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case Err

        Case 28461

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155912)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 28459

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case Err

        Case 28459

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155913)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 28460

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case Err

        Case 28460

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155914)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_DownClick

    'Diminui a DataEntregaAte em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 28467

    Exit Sub

Erro_UpDownEntregaAte_DownClick:

    Select Case Err

        Case 28467

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155915)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaAte_UpClick

    'Aumenta a DataEntregaAte em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 28468

    Exit Sub

Erro_UpDownEntregaAte_UpClick:

    Select Case Err

        Case 28468

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155916)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_DownClick

    'Diminui a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 28465

    Exit Sub

Erro_UpDownEntregaDe_DownClick:

    Select Case Err

        Case 28465

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155917)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEntregaDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEntregaDe_UpClick

    'Aumenta a DataEntregaDe em 1 dia
    lErro = Data_Up_Down_Click(DataEntregaDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 28466

    Exit Sub

Erro_UpDownEntregaDe_UpClick:

    Select Case Err

        Case 28466

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155918)

    End Select

    Exit Sub

End Sub

Private Function Grid_Pedido_Preenche(colNFiscalInfo As Collection) As Long
'Preenche o Grid Pedido com os dados de colNFiscalInfo

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim objNFiscalInfo As ClassNFiscalInfo
Dim objTransportadora As New ClassTransportadora
Dim objFilialEmpresa As New AdmFiliais
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim bAchou As Boolean

On Error GoTo Erro_Grid_Pedido_Preenche

    'Se o número de NFiscal for maior que o número de linhas do Grid
    If colNFiscalInfo.Count + 1 > GridPedido.Rows Then
    
        'If colNFiscalInfo.Count > NUM_MAXIMO_PARCELAS Then Error 19167

        'Altera o número de linhas do Grid de acordo com o número de NFiscal
        GridPedido.Rows = colNFiscalInfo.Count + 1

        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGrid)

    End If

    iLinha = 0

    'Percorre todas as NFiscais da Coleção
    For Each objNFiscalInfo In colNFiscalInfo

        iLinha = iLinha + 1

        'Passa para a tela os dados da NFiscal em questão
        GridPedido.TextMatrix(iLinha, iGrid_GeraNFiscal_Col) = objNFiscalInfo.iMarcada
        GridPedido.TextMatrix(iLinha, iGrid_Pedido_Col) = objNFiscalInfo.lPedido
        GridPedido.TextMatrix(iLinha, iGrid_Cliente_Col) = objNFiscalInfo.lCliente
        GridPedido.TextMatrix(iLinha, iGrid_NomeRed_Col) = objNFiscalInfo.sClienteNomeReduzido
        
        If objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_POR_BLOQUEIO Then
            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_BLOQUEIO
        ElseIf objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_POR_FALTAESTOQUE Then
            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_FALTA_ESTOQUE
        ElseIf objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_OUTROS Then
            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_OUTROS
        ElseIf objNFiscalInfo.iMotivoNaoGerada = MOTIVO_NAOGERADA_POR_BLOQUEIO_CREDITO Then
            GridPedido.TextMatrix(iLinha, iGrid_Motivo_Col) = MOTIVO_NAOGERADA_DESCRICAO_BLOQUEIO_CREDITO
        End If
        
        ' Se a transportadora foi informada
        If objNFiscalInfo.iCodTransp > 0 Then
        
            objTransportadora.iCodigo = objNFiscalInfo.iCodTransp
            'Lê a transportadora
            lErro = CF("Transportadora_Le", objTransportadora)
            If lErro <> SUCESSO And lErro <> 19250 Then Error 51360
            If lErro <> SUCESSO Then Error 51361
            'Coloca a transportadora no grid
            objNFiscalInfo.sNomeRedTransp = objTransportadora.sNomeReduzido
            GridPedido.TextMatrix(iLinha, iGrid_TransPortadora_Col) = objNFiscalInfo.iCodTransp & SEPARADOR & objNFiscalInfo.sNomeRedTransp
        
        End If
    
        
        If objNFiscalInfo.dtEmissaoPedido <> DATA_NULA And objNFiscalInfo.dtEmissaoPedido <> 0 Then GridPedido.TextMatrix(iLinha, iGrid_Emissao_Col) = Format(objNFiscalInfo.dtEmissaoPedido, "dd/mm/yyyy")
        If objNFiscalInfo.dtEntregaPedido <> DATA_NULA And objNFiscalInfo.dtEntregaPedido Then GridPedido.TextMatrix(iLinha, iGrid_Entrega_Col) = Format(objNFiscalInfo.dtEntregaPedido, "dd/mm/yyyy")
        GridPedido.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objNFiscalInfo.dValorTotal, "Standard")
        GridPedido.TextMatrix(iLinha, iGrid_Filial_Col) = objNFiscalInfo.iFilialCliente
        GridPedido.TextMatrix(iLinha, iGrid_Estado_Col) = objNFiscalInfo.sSiglaEstadoEntrega
        GridPedido.TextMatrix(iLinha, iGrid_Cidade_Col) = objNFiscalInfo.sCidadeEntrega
        GridPedido.TextMatrix(iLinha, iGrid_Bairro_Col) = objNFiscalInfo.sBairroEntrega
        
        bAchou = False
        'Verifica se a FilialEmpresa do pedido já foi lida
        For Each objCodigoNome In colCodigoNome
            If objCodigoNome.iCodigo = objNFiscalInfo.iFilialEmpresa Then
                objFilialEmpresa.iCodFilial = objNFiscalInfo.iFilialEmpresa
                objFilialEmpresa.sNome = objCodigoNome.sNome
                bAchou = True
                Exit For
            End If
        Next
        'Se ainda não foi lida
        If Not bAchou Then
            
            objFilialEmpresa.iCodFilial = objNFiscalInfo.iFilialEmpresa
            objFilialEmpresa.lCodEmpresa = glEmpresa
            'Lê a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then Error 51434
            If lErro <> SUCESSO Then Error 51435
            'Adiciona na coleção das filiais já lidas a filial lida agora
            colCodigoNome.Add objFilialEmpresa.iCodFilial, objFilialEmpresa.sNome
        
        End If
        '???? O campo empresa do grid de pedidos é na verdade o campo filialempresa. Colocar na forma "codigo-nomered" e alterar o título.
        'Coloca no grid o código-Nome da Filial do Pedido
        GridPedido.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

    Next

    Call Grid_Refresh_Checkbox(objGrid)

    'Passa para o Obj o número de NFiscais passados pela Coleção
    objGrid.iLinhasExistentes = colNFiscalInfo.Count
    
    Grid_Pedido_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Pedido_Preenche:

    Grid_Pedido_Preenche = Err
    
    Select Case Err
    
        Case 19167
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", Err, colNFiscalInfo.Count, NUM_MAXIMO_PARCELAS)
        
        Case 51434
        
        Case 51435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", Err, objFilialEmpresa.iCodFilial)
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155919)

    End Select
    
    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 28481

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 28481
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155920)

    End Select

    Exit Function

End Function

Private Function Trata_TabPedidos() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_TabPedidos

    If Ordenados.ListIndex = -1 Then

        Ordenados.ListIndex = 0

    Else
        
        'Verifica se tem seleção e Preenche o Grid de acordo com a ordenação
        lErro = Traz_Pedidos_Selecionados()
        If lErro <> SUCESSO And lErro <> 51428 Then Error 58025
        If lErro = 51428 Then Error 51429
        
    End If

    iTabPrincipalAlterado = 0

    Exit Function

Erro_Trata_TabPedidos:

    Trata_TabPedidos = Err

    Select Case Err

        Case 58025, 51429

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155921)

    End Select

    Exit Function

End Function

Private Function Move_TabSelecao_Memoria() As Long

Dim lErro As Long, iNRegiao As Integer, sRegioes As String, sRegiao As String, iIndice As Integer

On Error GoTo Erro_Move_TabSelecao_Memoria

    '???? Quando eu trago os clientes pelo browse antes de selecionar os pedidos ele não verifica se o cliente De é anterior ao Cliente Até

    gobjGeracaoNFiscal.iTodosOsPedidos = ExibeTodos.Value

    'Recolhe a data de emissão De
    gobjGeracaoNFiscal.dtEmissaoDe = StrParaDate(DataEmissaoDe.Text)
    'Recolhe a Data de emissão Até
    gobjGeracaoNFiscal.dtEmissaoAte = StrParaDate(DataEmissaoAte.Text)

    'Se a Data de emissão De e Até foram preenchidas
    If gobjGeracaoNFiscal.dtEmissaoDe <> DATA_NULA And gobjGeracaoNFiscal.dtEmissaoAte <> DATA_NULA Then
        'Verifica se Data Emissão De é anterior a Data de Emissão Até
        If gobjGeracaoNFiscal.dtEmissaoAte < gobjGeracaoNFiscal.dtEmissaoDe Then Error 51444
    End If

    'Recolhe as data de entrega DE e ATÉ
    gobjGeracaoNFiscal.dtEntregaDe = StrParaDate(DataEntregaDe.Text)
    gobjGeracaoNFiscal.dtEntregaAte = StrParaDate(DataEntregaAte.Text)
    gobjGeracaoNFiscal.lViagem = StrParaLong(CodigoViagem.Text)

    'Se a Data de entrega De e Até foram preenchidas
    If gobjGeracaoNFiscal.dtEntregaDe <> DATA_NULA And gobjGeracaoNFiscal.dtEntregaAte <> DATA_NULA Then
        'Verifica se Data Entrega De é anterior a Data de Entrega Até
        If gobjGeracaoNFiscal.dtEntregaAte < gobjGeracaoNFiscal.dtEntregaDe Then Error 51445
    End If
    
    'Recolhe Pedido De e Até
    gobjGeracaoNFiscal.lPedidosDe = StrParaLong(PedidoInicial.Text)
    gobjGeracaoNFiscal.lPedidosAte = StrParaLong(PedidoFinal.Text)

    'Se PedidoFinal e PedidoInicial estão preenchidos
    If gobjGeracaoNFiscal.lPedidosDe <> 0 And gobjGeracaoNFiscal.lPedidosAte <> 0 Then
        'Verifica se Data Pedido De é menor que pedido Até
        If gobjGeracaoNFiscal.lPedidosAte < gobjGeracaoNFiscal.lPedidosDe Then Error 51446
    End If
    
    'Recolhe o Cliente De e o ATé
    gobjGeracaoNFiscal.lClientesDe = StrParaLong(ClienteDe.Text)
    gobjGeracaoNFiscal.lClientesAte = StrParaLong(ClienteAte.Text)
    
    'Se ClienteAté e ClienteDe estão preenchidos
    If gobjGeracaoNFiscal.lClientesDe <> 0 And gobjGeracaoNFiscal.lClientesAte <> 0 Then
        'Verifica se Cliente De é menor que Cliente Até
        If gobjGeracaoNFiscal.lClientesAte < gobjGeracaoNFiscal.lClientesDe Then Error 51447
    End If
    
    'Sairam os campos Filiais
    gobjGeracaoNFiscal.iFilialPedidoDe = 0
    gobjGeracaoNFiscal.iFilialPedidoAte = 0
    
    gobjGeracaoNFiscal.sOrdenacao = asOrdenacao(Ordenados.ListIndex)
    
    gobjGeracaoNFiscal.iVendedorInicial = Codigo_Extrai(VendedorInicial.Text)
    gobjGeracaoNFiscal.iVendedorFinal = Codigo_Extrai(VendedorFinal.Text)
    
    If gobjGeracaoNFiscal.iVendedorInicial <> 0 And gobjGeracaoNFiscal.iVendedorFinal <> 0 And gobjGeracaoNFiscal.iVendedorFinal < gobjGeracaoNFiscal.iVendedorInicial Then Error 51448
    
    iNRegiao = 0
    'Percorre toda a Lista
    For iIndice = 0 To ListRegioes.ListCount - 1
        If ListRegioes.Selected(iIndice) = True Then
            sRegiao = Codigo_Extrai(ListRegioes.List(iIndice))
            If iNRegiao = 0 Then
                sRegioes = sRegiao
            Else
                sRegioes = sRegioes + "," + sRegiao
            End If
            iNRegiao = iNRegiao + 1
        End If
    Next
    
    'se todas as regioes foram marcadas entao nao precisa filtrar
    If iNRegiao = ListRegioes.ListCount Then sRegioes = ""
    
    gobjGeracaoNFiscal.sRegioes = sRegioes
    
    Move_TabSelecao_Memoria = SUCESSO

    Exit Function
    
Erro_Move_TabSelecao_Memoria:

    Move_TabSelecao_Memoria = Err

    Select Case Err
    
        Case 51444
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAODE_MAIOR_DATAEMISSAOATE", Err, DataEmissaoDe.Text, DataEmissaoAte.Text)
        
        Case 51445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAENTREGADE_MAIOR_DATAENTREGAATE", Err, DataEntregaDe.Text, DataEntregaAte.Text)
        
        Case 51446
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PEDIDOINICIAL_MAIOR_PEDIDOFINAL", Err)
        
        Case 51447
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTEDE_MAIOR_CLIENTEATE", Err)
        
        Case 51448
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_INICIAL_MAIOR", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155922)
    
    End Select

    Exit Function

End Function

Public Sub DataSaida_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataSaida_Validate

    'Verifica se a Data de Saida foi digitada
    If Len(Trim(DataSaida.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataSaida.Text)
    If lErro <> SUCESSO Then Error 31385

    Exit Sub

Erro_DataSaida_Validate:

    Cancel = True


    Select Case Err

        Case 31385

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155923)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_DownClick

    lErro = Data_Up_Down_Click(DataSaida, DIMINUI_DATA)
    If lErro Then Error 31384

    Exit Sub

Erro_UpDownSaida_DownClick:

    Select Case Err

        Case 31384

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155924)

    End Select

    Exit Sub

End Sub

Public Sub UpDownSaida_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownSaida_UpClick

    lErro = Data_Up_Down_Click(DataSaida, AUMENTA_DATA)
    If lErro Then Error 31383

    Exit Sub

Erro_UpDownSaida_UpClick:

    Select Case Err

        Case 31383

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155925)

    End Select

    Exit Sub

End Sub


Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_GERACAO_NFISCAL_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Geração de Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GeracaoNFiscal"
    
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
        
        If Me.ActiveControl Is VendedorInicial Then
            Call LabelVendedorDe_Click
        ElseIf Me.ActiveControl Is VendedorFinal Then
            Call LabelVendedorAte_Click
        ElseIf Me.ActiveControl Is Pedido Then
            Call BotaoPedido_Click
        ElseIf Me.ActiveControl Is PedidoInicial Then
            Call LabelPedidoDe_Click
        ElseIf Me.ActiveControl Is PedidoFinal Then
            Call LabelPedidoAte_Click
        ElseIf Me.ActiveControl Is ClienteDe Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteAte Then
            Call LabelClienteAte_Click
        End If
    
    End If

End Sub


Private Sub NotaFiscal_Imprime()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio

    If gobjGeracaoNFiscal.iImprime = IMPRIME_NOTA_FISCAL Then
    
        'se estiver gerando uma nota fiscal fatura
        If gobjGeracaoNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV Then
                        
            Call objRelatorio.Rel_Menu_Executar("Emissão das Notas Fiscais Fatura", Serie.Text)
        
        Else 'gerando uma nf que nao seja fatura
            Call objRelatorio.Rel_Menu_Executar("Emissão das Notas Fiscais", Serie.Text)
        
        End If
        
    End If


End Sub

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim dQuantidadeConvertida As Double
Dim sProdutoFormatado As String, sProdutoTela As String
Dim iPreenchido As Integer
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim bEncontrouProduto As Boolean
Dim iLinha2 As Integer, objCodigoNome As New AdmlCodigoNome
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sContaMascarada As String, sAlmoxNomeRed As String
Dim iAlmoxPadrao As Integer, objCliente As New ClassCliente
Dim bEncontrouQuant As Boolean, objFilialCliente As New ClassFilialCliente
Dim bEncontrouQuant2 As Boolean, objItem As ClassItemNF, objAlocacao As ClassItemNFAlocacao
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sDocInfo As String
Dim objProduto As New ClassProduto
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case ESCANINHO_CUSTO
            
            For Each objItem In gobjNFiscal.ColItensNF
                    
                For Each objAlocacao In objItem.colAlocacoes
                    objMnemonicoValor.colValor.Add ESCANINHO_NOSSO
                Next
            
            Next
        
        Case ESCANINHO_CUSTO_CONSIG
            
            For Each objItem In gobjNFiscal.ColItensNF

                For Each objAlocacao In objItem.colAlocacoes
                    objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO
                Next
            
            Next
        
        Case QUANT_ALOCADA_CONSIG
            
            For Each objItem In gobjNFiscal.ColItensNF
            
                For Each objAlocacao In objItem.colAlocacoes
                    
                    'Define o produto que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
                    lErro = CF("Produto_Formata", objItem.sProduto, sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 79012
                    
                    objItemMovEstoque.sProduto = sProdutoFormatado
                    
                    'Define o almoxarifado que será passado como parâmetro para MovEstoque_Le_QuantVendConsig
                    objItemMovEstoque.iAlmoxarifado = objAlocacao.iAlmoxarifado
                    
                    'Define o tipo de movimento, o DocOrigem e o TipoNumIntoDocOrigem que serão passados como parâmetros para MovEstoque_Le_QuantVendConsig
                                        
                    objItemMovEstoque.iTipoMov = MOV_EST_NF_VENDA_MAT_CONSIG
                    
                    'Define a sigla do DocInfo que será passado como parâmetro para a função MovEstoque_Le_QuantVendConsig
                    Select Case gobjGeracaoNFiscal.iTipoNFiscal
                        Case DOCINFO_NFISFVPV
                            sDocInfo = "NFISFVPV"
                        
                        Case DOCINFO_NFISVPV
                            sDocInfo = "NFISVPV"
                    End Select
                    
                    objItemMovEstoque.sDocOrigem = sDocInfo & " " & Serie.Text & " " & gobjNFiscal.lNumNotaFiscal
                    objItemMovEstoque.iTipoNumIntDocOrigem = MOVEST_TIPONUMINTDOCORIGEM_ITEMNFISCAL
                    
                    'Verifica se MovEstoque_Le_QuantVendConsig não encontrou erro
                    lErro = CF("MovEstoque_Le_QuantVendConsig", objItemMovEstoque)
                    If lErro <> SUCESSO And lErro <> 79003 Then gError 79010
                
                    'se não conseguiu encontrar mov. estoque para os parametros em questao ==> quantidade é zerada
                    If lErro = 79003 Then objItemMovEstoque.dQuantidade = 0
                
                    'Passa para o mnemônico o valor encontrado por MovEstoque_Le_QuantVendConsig
                    objMnemonicoValor.colValor.Add objItemMovEstoque.dQuantidade
                    
                    Next
                
                Next
        
        Case CTACONTABILEST1 'parametros: produto no formato da tela do grid de itens, produto no formato da tela do grid de alocacoes e nome reduzido do almoxarifado
            
            For Each objItem In gobjNFiscal.ColItensNF
                
                For Each objAlocacao In objItem.colAlocacoes
                
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                            
                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))

                        'Lê o Nome Reduzido do Almoxarifado
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then Error 59388

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                            
                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then Error 59390

                            objEstoqueProduto.sProduto = sProdutoFormatado
                            
                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then Error 59391

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_RetornaContaTela(objEstoqueProduto.sContaContabil, sContaMascarada)
                                If lErro <> SUCESSO Then Error 64225

                                objMnemonicoValor.colValor.Add sContaMascarada
                            Else
                                objMnemonicoValor.colValor.Add ""
                            End If
                        Else
                            objMnemonicoValor.colValor.Add ""
                        End If
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                
                Next
                
            Next
                    
        Case CODIGO1
            objMnemonicoValor.colValor.Add gobjPedidoVenda.lCodigo
        
        Case QUANT_ESTOQUE
        
            For Each objItem In gobjNFiscal.ColItensNF
            
                lErro = CF("UMEstoque_Conversao", objItem.sProduto, objItem.sUnidadeMed, objItem.dQuantidade, dQuantidadeConvertida)
                If lErro <> SUCESSO Then Error 64214
                
                objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
            Next
            
        Case ALMOX1
        
            If gcolAlmoxFilial.Count = 0 Then
            
                lErro = CF("Almoxarifados_Le_FilialEmpresa", gobjNFiscal.iFilialEmpresa, gcolAlmoxFilial)
                If lErro <> SUCESSO Then Error 59423
                
            End If
            
            For Each objItem In gobjNFiscal.ColItensNF
                
                For Each objAlocacao In objItem.colAlocacoes
                
                    sAlmoxNomeRed = ""
                    
                    For Each objAlmoxarifado In gcolAlmoxFilial
                    
                        If objAlocacao.iAlmoxarifado = objAlmoxarifado.iCodigo Then
                        
                            sAlmoxNomeRed = objAlmoxarifado.sNomeReduzido
                            Exit For
                            
                        End If
                        
                    Next
                    
                    objMnemonicoValor.colValor.Add sAlmoxNomeRed
                    
                Next
                
            Next
            
        Case DATA_EMISSAO
            objMnemonicoValor.colValor.Add gobjNFiscal.dtDataEmissao
        
        Case DATA_SAIDA
            objMnemonicoValor.colValor.Add gobjNFiscal.dtDataSaida
            
        Case DESCONTO1
            For Each objItem In gobjNFiscal.ColItensNF
            
                objMnemonicoValor.colValor.Add objItem.dValorDesconto
            
            Next
            
        Case DESCRICAO_ITEM
            For Each objItem In gobjNFiscal.ColItensNF
            
                objMnemonicoValor.colValor.Add objItem.sDescricaoItem
            
            Next
            
        Case FILIAL1
            
            objFilialCliente.lCodCliente = gobjNFiscal.lCliente
            objFilialCliente.iCodFilial = gobjNFiscal.iFilialCli
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then Error 59392
            If lErro <> SUCESSO Then Error 59393
            
            objMnemonicoValor.colValor.Add objFilialCliente.sNome
            
        Case FILIAL_CONTA
            
            objFilialCliente.lCodCliente = gobjNFiscal.lCliente
            objFilialCliente.iCodFilial = gobjNFiscal.iFilialCli
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO And lErro <> 12567 Then Error 59392
            If lErro <> SUCESSO Then Error 59393
                
            If objFilialCliente.sContaContabil <> "" Then
            
                lErro = Mascara_RetornaContaTela(objFilialCliente.sContaContabil, sContaMascarada)
                If lErro <> SUCESSO Then Error 59392
            
            Else
            
                sContaMascarada = ""
                
            End If
            
            objMnemonicoValor.colValor.Add sContaMascarada
        
        Case CLIENTE1
        
            objCodigoNome.lCodigo = gobjNFiscal.lCliente
            
            lErro = CF("Cliente_Le_NomeRed", objCodigoNome)
            If lErro <> SUCESSO And lErro <> 12553 Then Error 59394
            If lErro <> SUCESSO Then Error 59395
                        
            objMnemonicoValor.colValor.Add objCodigoNome.sNome
            
        Case NATUREZA_OP
            objMnemonicoValor.colValor.Add gobjNFiscal.sNaturezaOp

        Case NFISCAL1
            objMnemonicoValor.colValor.Add gobjNFiscal.lNumNotaFiscal

        Case PRODUTO1
            For Each objItem In gobjNFiscal.ColItensNF
            
                lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
                If lErro <> SUCESSO Then Error 59397
                
                objMnemonicoValor.colValor.Add sProdutoTela
            
            Next
            
        Case QUANTIDADE1
            For Each objItem In gobjNFiscal.ColItensNF
            
                objMnemonicoValor.colValor.Add objItem.dQuantidade
            
            Next
            
        Case Serie1
            objMnemonicoValor.colValor.Add gobjNFiscal.sSerie
            
        Case UNIDADE_MED
            For Each objItem In gobjNFiscal.ColItensNF
            
                objMnemonicoValor.colValor.Add objItem.sUnidadeMed
            
            Next

        Case VALOR_TOTAL
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorTotal

        Case PRECO_UNITARIO
            For Each objItem In gobjNFiscal.ColItensNF
            
                objMnemonicoValor.colValor.Add objItem.dPrecoUnitario
            
            Next
            
        Case PRECO_TOTAL
            For Each objItem In gobjNFiscal.ColItensNF
            
                objMnemonicoValor.colValor.Add objItem.dValorTotal
            
            Next
            
        Case ICMS
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dICMSValor

        Case ICMS_DIFAL_REMETENTE
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dICMSInterestVlrUFRemet

        Case ICMS_DIFAL_DESTINO
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dICMSInterestVlrUFDest

        Case ICMS_FCP
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dICMSInterestVlrFCPUFDest
        
        Case ICMSSUBST
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dICMSSubstValor

        Case VALOR_FRETE
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorFrete

        Case VALOR_SEGURO
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorSeguro

        Case VALOR_DESPESAS
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorOutrasDespesas

        Case IPI
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dIPIValor

        Case VALOR_DESCONTO
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorDesconto

        Case ISS_VALOR
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dISSValor

        Case ISS_INCLUSO
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.iISSIncluso <> 0

        Case VALOR_IRRF
            objMnemonicoValor.colValor.Add gobjNFiscal.objTributacaoNF.dIRRFValor
            
        Case VALOR_PRODUTOS
            objMnemonicoValor.colValor.Add gobjNFiscal.dValorProdutos
                    
        Case PRODUTO_ALMOX
            For Each objItem In gobjNFiscal.ColItensNF
                
                For Each objAlocacao In objItem.colAlocacoes
                
                    lErro = Mascara_RetornaProdutoTela(objItem.sProduto, sProdutoTela)
                    If lErro <> SUCESSO Then Error 59409
                    
                    objMnemonicoValor.colValor.Add sProdutoTela
                
                Next
                
            Next
            
        Case QUANT_ALOCADA

            For Each objItem In gobjNFiscal.ColItensNF
                
                If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                
                    For Each objAlocacao In objItem.colAlocacoes
                    
                        objMnemonicoValor.colValor.Add objAlocacao.dQuantidade
                    
                    Next
                
                Else
                
                    For Each objAlocacao In objItem.colAlocacoes
                    
                        objMnemonicoValor.colValor.Add 0
                    
                    Next
                
                End If
                
            Next
            
        Case UNIDADE_MED_EST
            For Each objItem In gobjNFiscal.ColItensNF
                
                For Each objAlocacao In objItem.colAlocacoes
                
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                        objMnemonicoValor.colValor.Add objItem.sUMEstoque
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                
                Next
                
            Next
            
        Case CONTACONTABILAPLIC
        
            For Each objItem In gobjNFiscal.ColItensNF
                
                'Lê o produto
                objProduto.sCodigo = objItem.sProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 59423
    
                If Len(Trim(objProduto.sContaContabil)) <> 0 Then
                
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 59423

                    objMnemonicoValor.colValor.Add sContaMascarada
                
                Else
                
                    objMnemonicoValor.colValor.Add ""
                
                End If
            
            Next
        
        Case CONTACONTABILPROD
        
            For Each objItem In gobjNFiscal.ColItensNF
                
                'Lê o produto
                objProduto.sCodigo = objItem.sProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 59423
    
                If Len(Trim(objProduto.sContaContabilProducao)) <> 0 Then
                
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabilProducao, sContaMascarada)
                    If lErro <> SUCESSO Then gError 59423

                    objMnemonicoValor.colValor.Add sContaMascarada
                
                Else
                
                    objMnemonicoValor.colValor.Add ""
                
                End If
        
            Next
        
        Case CONTACONTABILCATPROD
                
            For Each objItem In gobjNFiscal.ColItensNF
                
                'Lê o produto
                objProduto.sCodigo = objItem.sProduto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 59423
    
                objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                
                objProdutoCategoria.sProduto = objItem.sProduto
                
                Set objCatProdItem = New ClassCategoriaProdutoItem
                
                lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                If lErro <> SUCESSO And lErro <> 183548 Then gError 59423
                
                If lErro = SUCESSO Then
                    
                    Set objPlanoConta = New ClassPlanoConta
                
                    'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                    lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                    If lErro <> SUCESSO And lErro <> 5451 Then gError 59423
                
                    If lErro = SUCESSO Then
                
                        sContaMascarada = String(STRING_CONTA, 0)
    
                        lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                        If lErro <> SUCESSO Then gError 183799
    
                        objMnemonicoValor.colValor.Add sContaMascarada
                        
                    Else
                
                        objMnemonicoValor.colValor.Add ""
                
                    End If
                
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
                    
            Next
        
        Case Else
            Error 59389

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 59389
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 59388, 59390, 59391, 59392, 59394, 59397, 59409, 59410, 59423
        
        Case 59393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_CADASTRADA", Err, objFilialCliente.iCodFilial, objFilialCliente.lCodCliente)
        
        Case 59395
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", Err, gobjNFiscal.lCliente)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155926)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de nota fiscal e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, objItem As ClassItemNF, iNumAlocacoes As Integer

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjNFiscal = vParams(0)
    Set gobjPedidoVenda = vParams(1)

    'percorre itens otendo qtde de alocacoes da nf como um todo
    For Each objItem In gobjNFiscal.ColItensNF
        iNumAlocacoes = iNumAlocacoes + objItem.colAlocacoes.Count
    Next
    
    GridAlocacao.Tag = iNumAlocacoes
    GridItens.Tag = gobjNFiscal.ColItensNF.Count

    'obter numero de Doc
    lErro = CF("Voucher_Automatico_Trans", gobjNFiscal.iFilialEmpresa, giExercicio, giPeriodo, MODULO_FATURAMENTO, lDoc)
    If lErro <> SUCESSO Then Error 59398
    
    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro(Me, IIf(gobjNFiscal.iTipoNFiscal = DOCINFO_NFISFVPV, "NFiscalFaturaPedido", "NFiscalPedido"), gobjNFiscal.lNumIntDoc, gobjNFiscal.lCliente, gobjNFiscal.iFilialCli, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, gobjNFiscal.iFilialEmpresa, gobjGeracaoNFiscal.iLoteContabil, gobjNFiscal.lNumNotaFiscal)
    If lErro <> SUCESSO Then Error 59399

    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 59398, 59399
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155927)
     
    End Select
     
    Exit Function

End Function

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

Public Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Public Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Public Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Public Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Public Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Public Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'Traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Function GeracaoNF_Prepara_CTB() As Long
'prepara informacoes necessarias para a contabilizacao

Dim lErro As Long, objPeriodo As New ClassPeriodo
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_GeracaoNF_Prepara_CTB

    If (gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO) Then
    
        'Verifica se a data contabil está preenchida
        If Len(CTBDataContabil.ClipText) = 0 Then gError 59400

        gobjGeracaoNFiscal.objTelaAtualizacao = Me
        gobjGeracaoNFiscal.dtContabil = CDate(CTBDataContabil.Text)
        
        If gobjGeracaoNFiscal.dtContabil <> gdtDataAtual Then
    
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATA_CONTABIL_DIFERE_TRANSACAO", gobjGeracaoNFiscal.dtContabil, gdtDataAtual)
    
            If vbMsgRes = vbNo Then gError 92041
    
        End If
        
        'se o lote estiver preenchido ==> não pode ser com o valor zero
'        If Len(CTBLote.ClipText) > 0 And giTipoVersao = VERSAO_FULL Then
        If Len(CTBLote.ClipText) > 0 Then
            gobjGeracaoNFiscal.iLoteContabil = CInt(CTBLote.ClipText)
        Else
            'se não estiver preenchido o lote ==> atualizacao imediata e o valor do lote será zero internamente
            gobjGeracaoNFiscal.iLoteContabil = 0
        End If
    
        'Coloca o periodo relativo a data na tela
        lErro = CF("Periodo_Le", gobjGeracaoNFiscal.dtContabil, objPeriodo)
        If lErro <> SUCESSO Then gError 59401
        
        giPeriodo = objPeriodo.iPeriodo
        giExercicio = objPeriodo.iExercicio

    End If
    
    GeracaoNF_Prepara_CTB = SUCESSO
     
    Exit Function
    
Erro_GeracaoNF_Prepara_CTB:

    GeracaoNF_Prepara_CTB = gErr
     
    Select Case gErr
          
        Case 59400
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CONTABIL_NAO_PREENCHIDA", gErr)
        
        Case 59401, 92041
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155928)
     
    End Select
     
    Exit Function

End Function

Public Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

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

Public Sub BotaoImprimirPI_Click()

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objNFiscalInfo As New ClassNFiscalInfo

On Error GoTo Erro_BotaoImprimirPI_Click

    If objGrid.iLinhasExistentes = 0 Then Exit Sub
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se nenhuma linha válida estiver com o foco
    If GridPedido.Row < 1 Or GridPedido.Row > objGrid.iLinhasExistentes Then gError 132443
    
    'Passa a linha do Grid para o Obj
    Set objNFiscalInfo = gobjGeracaoNFiscal.colNFiscalInfo.Item(GridPedido.Row)

    'Passa os dados do NFiscal para o Obj
    objPedidoVenda.iFilialEmpresa = objNFiscalInfo.iFilialEmpresa
    objPedidoVenda.lCodigo = objNFiscalInfo.lPedido
   
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 132440

    If lErro = 26509 Then gError 132441
    
    lErro = objRelatorio.ExecutarDireto("Pedido Interno", "PedidoVenda >= @NPEDVENDINIC E PedidoVenda <= @NPEDVENDFIM", 1, "PedVend4", "NPEDVENDINIC", objPedidoVenda.lCodigo, "NPEDVENDFIM", objPedidoVenda.lCodigo)
    If lErro <> SUCESSO Then gError 132442

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoImprimirPI_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 132440, 132442
        
        Case 132441
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_NAO_CADASTRADA", gErr, objPedidoVenda.lCodigo)

        Case 132443
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_SELECIONADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 176083)

    End Select

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

End Sub

Function GNF_Form_Load(ByVal objCT As Object) As Long
    GNF_Form_Load = SUCESSO
End Function


Private Sub objEventoCodigoViagem_evSelecao(obj1 As Object)

Dim lErro As Long

On Error GoTo Erro_objEventoCodigoViagem_evSelecao

    If obj1.lCodigo <> 0 Then
        CodigoViagem.PromptInclude = False
        CodigoViagem.Text = CStr(obj1.lCodigo)
        CodigoViagem.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoCodigoViagem_evSelecao:

    Select Case gErr

        Case 205311

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205312)

    End Select

    Exit Sub

End Sub

Public Sub LabelViagem_Click()

Dim lErro As Long
'Dim objMapaDeEntrega As New ClassMapaDeEntrega
Dim colSelecao As New Collection

On Error GoTo Erro_LabelViagem_Click

    Call Chama_Tela("MapaDeEntregaLista", colSelecao, Nothing, objEventoCodigoViagem)

    Exit Sub

Erro_LabelViagem_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205313)

    End Select

    Exit Sub

End Sub

Public Sub CodigoViagem_GotFocus()
Dim iFrameAux As Integer
    
    iFrameAux = iTabPrincipalAlterado
    Call MaskEdBox_TrataGotFocus(CodigoViagem, iAlterado)
    iTabPrincipalAlterado = iFrameAux

End Sub

Public Sub CodigoViagem_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CodigoViagem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodigoViagem_Validate

    'Testa para ver se a viagem existe

    Exit Sub

Erro_CodigoViagem_Validate:

    Cancel = True

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143404)

    End Select

    Exit Sub

End Sub

Private Function Viagem_Imprime(ByVal lCodigo As Long) As Long

Dim lErro As Long
Dim objRelatorio As New AdmRelatorio
Dim objMapa As New ClassMapaDeEntrega

On Error GoTo Erro_Viagem_Imprime

    If lCodigo > 0 Then

        'Transforma o ponteiro do mouse em ampulheta
        GL_objMDIForm.MousePointer = vbHourglass
        
        'Guarda no obj o código do orçamento passado como parâmetro
        objMapa.lCodigo = lCodigo
        objMapa.iFilialEmpresa = giFilialEmpresa
        
        'Lê os dados do orçamento para verificar se o mesmo existe no BD
        lErro = CF("MapaDeEntrega_Le", objMapa)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
        'Se não encontrou => erro, pois não é possível imprimir um orçamento inexistente
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 209068
        
        lErro = objRelatorio.ExecutarDireto("Romaneio de Entrega", "Codigo = @NCODIGO", 1, "RomEnt", "NCODIGO", CStr(lCodigo))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Transforma o ponteiro do mouse em seta (padrão)
        GL_objMDIForm.MousePointer = vbDefault
        
    End If
    
    Viagem_Imprime = SUCESSO
    
    Exit Function

Erro_Viagem_Imprime:

    Viagem_Imprime = gErr
    
    Select Case gErr
           
        Case 209068
            Call Rotina_Erro(vbOKOnly, "ERRO_MAPADEENTREGA_NAO_CADASTRADO", gErr, objMapa.iFilialEmpresa, objMapa.lCodigo)
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209069)
    
    End Select
    
    'Transforma o ponteiro do mouse em seta (padrão)
    GL_objMDIForm.MousePointer = vbDefault

End Function

Public Sub LabelVendedorAte_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 0
    
    If Len(Trim(VendedorFinal.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorFinal.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Public Sub LabelVendedorDe_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection

    giVendedorInicial = 1
    
    If Len(Trim(VendedorInicial.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(VendedorInicial.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Public Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    'Preenche campo Vendedor
    If giVendedorInicial = 1 Then
        VendedorInicial.Text = CStr(objVendedor.iCodigo)
        Call VendedorInicial_Validate(bSGECancelDummy)
    Else
        VendedorFinal.Text = CStr(objVendedor.iCodigo)
        Call VendedorFinal_Validate(bSGECancelDummy)
    End If

    Me.Show

    Exit Sub

End Sub

Public Sub VendedorInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorInicial_Validate

    If Len(Trim(VendedorInicial.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorInicial, objVendedor, 0)
        If lErro <> SUCESSO Then gError 138367

    End If
    
    giVendedorInicial = 1
    
    Exit Sub

Erro_VendedorInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 138367

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167058)

    End Select

End Sub

Public Sub VendedorFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_VendedorFinal_Validate

    If Len(Trim(VendedorFinal.Text)) > 0 Then

        'Tenta ler o vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le2(VendedorFinal, objVendedor, 0)
        If lErro <> SUCESSO Then gError 138368

    End If
    
    giVendedorInicial = 0
 
    Exit Sub

Erro_VendedorFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 138368
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167059)

    End Select

End Sub

Private Function CarregaList_Regioes() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_CarregaList_Regioes
    
    'Preenche Combo Regiao
    Set colCodigoDescricao = New AdmColCodigoNome

    'Lê cada codigo e descricao da tabela RegioesVendas
    lErro = CF("Cod_Nomes_Le", "RegioesVendas", "Codigo", "Descricao", STRING_REGIAO_VENDA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 207090

    'preenche a ComboBox Regiao com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        ListRegioes.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        ListRegioes.ItemData(ListRegioes.NewIndex) = objCodigoDescricao.iCodigo
    Next

    CarregaList_Regioes = SUCESSO

    Exit Function

Erro_CarregaList_Regioes:

    CarregaList_Regioes = gErr

    Select Case gErr

        Case 207900

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 172566)

    End Select

    Exit Function

End Function

Public Sub BotaoMarcar_Click()
'marcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = True
    Next

End Sub

Public Sub BotaoDesmarcar_Click()
'desmarcar todos os itens da listbox
Dim iIndice As Integer

    For iIndice = 0 To ListRegioes.ListCount - 1
        ListRegioes.Selected(iIndice) = False
    Next

End Sub

Public Sub VendedorInicial_Change()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Public Sub VendedorFinal_Change()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ListRegioes_Click()
    iAlterado = REGISTRO_ALTERADO
    iTabPrincipalAlterado = REGISTRO_ALTERADO
End Sub

Public Sub BotaoExportar_Click()

Dim lErro As Long
Dim objNFiscalInfo As New ClassNFiscalInfo

On Error GoTo Erro_BotaoExportar_Click

    GL_objMDIForm.MousePointer = vbHourglass

    For Each objNFiscalInfo In gobjGeracaoNFiscal.colNFiscalInfo
    
        If objNFiscalInfo.iMarcada = MARCADO Then

            lErro = CF("PedidoVenda_Exporta", objNFiscalInfo.iFilialEmpresa, objNFiscalInfo.lPedido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If

    Next

    Call Rotina_Aviso(vbOKOnly, "AVISO_ARQUIVO_GERADO")

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExportar_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 204899)

    End Select
    
    Exit Sub

End Sub
