VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlCodigo As Long
Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvardtDataVenda As Date
Private mvardQuantidade As Double
Private mvarsLote As String
Private mvariFilialOP As Integer
Private mvarsSerie As String
Private mvarlNumNotaFiscal As Long
Private mvarlFornecedor As Long
Private mvariFilialFornecedor As Integer
Private mvarlTipoGarantia As Long
Private mvariGarantiaTotal As Integer
Private mvariGarantiaTotalPrazo As Integer
Private mvarcolNumSerie As New Collection
Private mvarcolProduto As New Collection
Private mvarsServico As String
Private mvariAtivo As Integer
Private mvarobjTela As Object
Private mvarlCliFabr As Long
Private mvariFilialCliFabr As Integer

Public Property Let lCliFabr(ByVal vData As Long)
    mvarlCliFabr = vData
End Property

Public Property Get lCliFabr() As Long
    lCliFabr = mvarlCliFabr
End Property

Public Property Let iFilialCliFabr(ByVal vData As Integer)
    mvariFilialCliFabr = vData
End Property

Public Property Get iFilialCliFabr() As Integer
    iFilialCliFabr = mvariFilialCliFabr
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let sServico(ByVal vData As String)
    mvarsServico = vData
End Property

Public Property Get sServico() As String
    sServico = mvarsServico
End Property

Public Property Set colProduto(ByVal vData As Collection)
    Set mvarcolProduto = vData
End Property

Public Property Get colProduto() As Collection
    Set colProduto = mvarcolProduto
End Property

Public Property Set colNumSerie(ByVal vData As Collection)
    Set mvarcolNumSerie = vData
End Property

Public Property Get colNumSerie() As Collection
    Set colNumSerie = mvarcolNumSerie
End Property

Public Property Let iGarantiaTotalPrazo(ByVal vData As Integer)
    mvariGarantiaTotalPrazo = vData
End Property

Public Property Get iGarantiaTotalPrazo() As Integer
    iGarantiaTotalPrazo = mvariGarantiaTotalPrazo
End Property

Public Property Let iGarantiaTotal(ByVal vData As Integer)
    mvariGarantiaTotal = vData
End Property

Public Property Get iGarantiaTotal() As Integer
    iGarantiaTotal = mvariGarantiaTotal
End Property

Public Property Let iFilialFornecedor(ByVal vData As Integer)
    mvariFilialFornecedor = vData
End Property

Public Property Get iFilialFornecedor() As Integer
    iFilialFornecedor = mvariFilialFornecedor
End Property

Public Property Let lTipoGarantia(ByVal vData As Long)
    mvarlTipoGarantia = vData
End Property

Public Property Get lTipoGarantia() As Long
    lTipoGarantia = mvarlTipoGarantia
End Property

Public Property Let lNumNotaFiscal(ByVal vData As Long)
    mvarlNumNotaFiscal = vData
End Property

Public Property Get lNumNotaFiscal() As Long
    lNumNotaFiscal = mvarlNumNotaFiscal
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let iFilialOP(ByVal vData As Integer)
    mvariFilialOP = vData
End Property

Public Property Get iFilialOP() As Integer
    iFilialOP = mvariFilialOP
End Property

Public Property Let sLote(ByVal vData As String)
    mvarsLote = vData
End Property

Public Property Get sLote() As String
    sLote = mvarsLote
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dtDataVenda(ByVal vData As Date)
    mvardtDataVenda = vData
End Property

Public Property Get dtDataVenda() As Date
    dtDataVenda = mvardtDataVenda
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property


