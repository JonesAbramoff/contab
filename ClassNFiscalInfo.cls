VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFiscalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsSerie As String
Private mvarsSiglaDoc As String 'tipo da NF
Private mvarlNumero As Long
Private mvarlPedido As Long
Private mvarlCliente As Long
Private mvarsClienteNomeReduzido As String
Private mvariFilialCliente As Integer
Private mvardtEmissaoPedido As Date
Private mvardtEntregaPedido As Date
Private mvardtDataSaidaNF As Date
Private mvardtEmissao As Date
Private mvardValorTotal As Double
Private mvariMarcada As Integer
Private mvardISSValor As Double
Private mvarlNumIntDoc As Long
Private mvarsCidadeEntrega As String
Private mvarsSiglaEstadoEntrega As String
Private mvarsBairroEntrega As String
Private mvarlClienteBenef As Long
Private mvariFilialCliBenef As Integer
Private mvarlFornecedorBenef As Long
Private mvariFilialFornBenef As Integer
Private mvarsNomeRedTransp As String
Private mvariCodTransp As Integer
Private mvariMotivoNaoGerada As Integer 'local copy

Private mvariCodTranspRedesp As Integer
Private mvariDetPagFrete As Integer

Public Property Let iDetPagFrete(ByVal vData As Integer)
    mvariDetPagFrete = vData
End Property

Public Property Get iDetPagFrete() As Integer
    iDetPagFrete = mvariDetPagFrete
End Property

Public Property Let iCodTranspRedesp(ByVal vData As Integer)
    mvariCodTranspRedesp = vData
End Property

Public Property Get iCodTranspRedesp() As Integer
    iCodTranspRedesp = mvariCodTranspRedesp
End Property


Public Property Let iMotivoNaoGerada(ByVal vData As Integer)
    mvariMotivoNaoGerada = vData
End Property


Public Property Get iMotivoNaoGerada() As Integer
    iMotivoNaoGerada = mvariMotivoNaoGerada
End Property

Public Property Let iCodTransp(ByVal vData As Integer)
Attribute iCodTransp.VB_Description = "Código da Transportadora."
    mvariCodTransp = vData
End Property


Public Property Get iCodTransp() As Integer
    iCodTransp = mvariCodTransp
End Property



Public Property Let sNomeRedTransp(ByVal vData As String)
Attribute sNomeRedTransp.VB_Description = "NomeReduzido da Transportadora."
    mvarsNomeRedTransp = vData
End Property


Public Property Get sNomeRedTransp() As String
    sNomeRedTransp = mvarsNomeRedTransp
End Property




Public Property Let iFilialFornBenef(ByVal vData As Integer)
Attribute iFilialFornBenef.VB_Description = "Filial do Fornecedor que vai beneficiar os nossos produtos."
    mvariFilialFornBenef = vData
End Property

Public Property Get iFilialFornBenef() As Integer
    iFilialFornBenef = mvariFilialFornBenef
End Property

Public Property Let lFornecedorBenef(ByVal vData As Long)
Attribute lFornecedorBenef.VB_Description = "Código do Fornecedor que vai beneficiar os nossos produtos."
    mvarlFornecedorBenef = vData
End Property

Public Property Get lFornecedorBenef() As Long
    lFornecedorBenef = mvarlFornecedorBenef
End Property

Public Property Let iFilialCliBenef(ByVal vData As Integer)
Attribute iFilialCliBenef.VB_Description = "Código da Filial do Cliente cujos produtos serão beneficiados por nós."
    mvariFilialCliBenef = vData
End Property

Public Property Get iFilialCliBenef() As Integer
    iFilialCliBenef = mvariFilialCliBenef
End Property

Public Property Let lClienteBenef(ByVal vData As Long)
Attribute lClienteBenef.VB_Description = "Código do Cliente cujos produtos serão beneficiados por nós."
    mvarlClienteBenef = vData
End Property

Public Property Get lClienteBenef() As Long
    lClienteBenef = mvarlClienteBenef
End Property

Public Property Let sBairroEntrega(ByVal vData As String)
    mvarsBairroEntrega = vData
End Property

Public Property Get sBairroEntrega() As String
    sBairroEntrega = mvarsBairroEntrega
End Property

Public Property Let sSiglaEstadoEntrega(ByVal vData As String)
Attribute sSiglaEstadoEntrega.VB_Description = "Sigla do Estado onde será feita a entrega do pedido"
    mvarsSiglaEstadoEntrega = vData
End Property

Public Property Get sSiglaEstadoEntrega() As String
    sSiglaEstadoEntrega = mvarsSiglaEstadoEntrega
End Property

Public Property Let sCidadeEntrega(ByVal vData As String)
Attribute sCidadeEntrega.VB_Description = "Cidade onde será feita a entrega do pedido"
    mvarsCidadeEntrega = vData
End Property

Public Property Get sCidadeEntrega() As String
    sCidadeEntrega = mvarsCidadeEntrega
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dISSValor(ByVal vData As Double)
    mvardISSValor = vData
End Property
Public Property Get dISSValor() As Double
    dISSValor = mvardISSValor
End Property

Public Property Let dtEmissao(ByVal vData As Date)
    mvardtEmissao = vData
End Property

Public Property Get dtEmissao() As Date
    dtEmissao = mvardtEmissao
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let sSiglaDoc(ByVal vData As String)
    mvarsSiglaDoc = vData
End Property

Public Property Get sSiglaDoc() As String
    sSiglaDoc = mvarsSiglaDoc
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let iMarcada(ByVal vData As Integer)
    mvariMarcada = vData
End Property

Public Property Get iMarcada() As Integer
    iMarcada = mvariMarcada
End Property

Public Property Let dtDataSaidaNF(ByVal vData As Date)
    mvardtDataSaidaNF = vData
End Property

Public Property Get dtDataSaidaNF() As Date
    dtDataSaidaNF = mvardtDataSaidaNF
End Property

Public Property Let dtEntregaPedido(ByVal vData As Date)
    mvardtEntregaPedido = vData
End Property

Public Property Get dtEntregaPedido() As Date
    dtEntregaPedido = mvardtEntregaPedido
End Property

Public Property Let dtEmissaoPedido(ByVal vData As Date)
    mvardtEmissaoPedido = vData
End Property

Public Property Get dtEmissaoPedido() As Date
    dtEmissaoPedido = mvardtEmissaoPedido
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let sClienteNomeReduzido(ByVal vData As String)
    mvarsClienteNomeReduzido = vData
End Property

Public Property Get sClienteNomeReduzido() As String
    sClienteNomeReduzido = mvarsClienteNomeReduzido
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lPedido(ByVal vData As Long)
    mvarlPedido = vData
End Property

Public Property Get lPedido() As Long
    lPedido = mvarlPedido
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property



