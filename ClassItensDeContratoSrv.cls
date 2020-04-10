VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItensDeContratoSrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlCodigo As Long
Private mvariFilialEmpresa As Integer
Private mvarlNumIntItemContrato As Long
Private mvardQuantidade As Double
Private mvarsLote As String
Private mvariFilialOP As Integer
Private mvariGarantiaTotal As Integer
Private mvarlTipoGarantia As Long
Private mvardtDataContratoIni As Date
Private mvardtDataContratoFim As Date
Private mvarsProduto As String
Private mvarsServico As String
Private mvarcolNumSerie As New Collection
Private mvarcolProduto As New Collection
Private mvarsCodigoContrato As String
Private mvarobjTela As Object

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let sCodigoContrato(ByVal vData As String)
    mvarsCodigoContrato = vData
End Property

Public Property Get sCodigoContrato() As String
    sCodigoContrato = mvarsCodigoContrato
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

Public Property Let sServico(ByVal vData As String)
    mvarsServico = vData
End Property

Public Property Get sServico() As String
    sServico = mvarsServico
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let dtDataContratoFim(ByVal vData As Date)
    mvardtDataContratoFim = vData
End Property

Public Property Get dtDataContratoFim() As Date
    dtDataContratoFim = mvardtDataContratoFim
End Property

Public Property Let dtDataContratoIni(ByVal vData As Date)
    mvardtDataContratoIni = vData
End Property

Public Property Get dtDataContratoIni() As Date
    dtDataContratoIni = mvardtDataContratoIni
End Property

Public Property Let lTipoGarantia(ByVal vData As Long)
    mvarlTipoGarantia = vData
End Property

Public Property Get lTipoGarantia() As Long
    lTipoGarantia = mvarlTipoGarantia
End Property

Public Property Let iGarantiaTotal(ByVal vData As Integer)
    mvariGarantiaTotal = vData
End Property

Public Property Get iGarantiaTotal() As Integer
    iGarantiaTotal = mvariGarantiaTotal
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

Public Property Let lNumIntItemContrato(ByVal vData As Long)
    mvarlNumIntItemContrato = vData
End Property

Public Property Get lNumIntItemContrato() As Long
    lNumIntItemContrato = mvarlNumIntItemContrato
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



