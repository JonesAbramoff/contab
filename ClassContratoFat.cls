VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContratoFat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCondPagto As Integer
Private mvariFilCli As Integer
Private mvarlCliente As Long
Private mvardtDataRef As Date
Private mvarsSerie As String
Private mvarsNaturezaOp As String
Private mvarcolItens As New Collection
Private mvarsContrato As String
Private mvariTipoTributacao As Integer
Private mvariTipoErro As Integer
Private mvariFilialEmpresa As Integer
Private mvarlNumNotaFiscal As Long
Private mvardValor As Double
Private mvarlNumIntDocCobranca As Long
Private mvarlFornecedor As Long
Private mvariFilialFornecedor As Integer
Private mvariRecibo As Integer

Public Property Let iRecibo(ByVal vData As Integer)
    mvariRecibo = vData
End Property

Public Property Get iRecibo() As Integer
    iRecibo = mvariRecibo
End Property

Public Property Let iFilialFornecedor(ByVal vData As Integer)
    mvariFilialFornecedor = vData
End Property

Public Property Get iFilialFornecedor() As Integer
    iFilialFornecedor = mvariFilialFornecedor
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let lNumIntDocCobranca(ByVal vData As Long)
    mvarlNumIntDocCobranca = vData
End Property

Public Property Get lNumIntDocCobranca() As Long
    lNumIntDocCobranca = mvarlNumIntDocCobranca
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lNumNotaFiscal(ByVal vData As Long)
    mvarlNumNotaFiscal = vData
End Property

Public Property Get lNumNotaFiscal() As Long
    lNumNotaFiscal = mvarlNumNotaFiscal
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iTipoErro(ByVal vData As Integer)
    mvariTipoErro = vData
End Property

Public Property Get iTipoErro() As Integer
    iTipoErro = mvariTipoErro
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
    sContrato = mvarsContrato
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let sSerie(ByVal vData As String)
    mvarsSerie = vData
End Property

Public Property Get sSerie() As String
    sSerie = mvarsSerie
End Property

Public Property Let dtDataRef(ByVal vData As Date)
    mvardtDataRef = vData
End Property

Public Property Get dtDataRef() As Date
    dtDataRef = mvardtDataRef
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilCli(ByVal vData As Integer)
    mvariFilCli = vData
End Property

Public Property Get iFilCli() As Integer
    iFilCli = mvariFilCli
End Property

Public Property Let iCondPagto(ByVal vData As Integer)
    mvariCondPagto = vData
End Property

Public Property Get iCondPagto() As Integer
    iCondPagto = mvariCondPagto
End Property



