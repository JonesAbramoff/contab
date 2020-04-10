VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPFatAporteCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarobjCredito As Object
Private mvardSaldoAtual As Double
Private mvardSaldoAnt As Double
Private mvarlCliente As Long
Private mvardtDataVenc As Date
Private mvardValor As Double
Private mvarlCodigo As Long
Private mvarlNumIntDocPagto As Long
Private mvarlNumIntDocAporte As Long
Private mvarlNumIntDocCredito As Long
Private mvariTipoPagtoAporte As Integer
Private mvariNumParcelas As Integer
Private mvariParcelaAporte As Integer

Public Property Let iParcelaAporte(ByVal vData As Integer)
    mvariParcelaAporte = vData
End Property

Public Property Get iParcelaAporte() As Integer
    iParcelaAporte = mvariParcelaAporte
End Property

Public Property Let iNumParcelas(ByVal vData As Integer)
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let iTipoPagtoAporte(ByVal vData As Integer)
    mvariTipoPagtoAporte = vData
End Property

Public Property Get iTipoPagtoAporte() As Integer
    iTipoPagtoAporte = mvariTipoPagtoAporte
End Property

Public Property Let lNumIntDocCredito(ByVal vData As Long)
    mvarlNumIntDocCredito = vData
End Property

Public Property Get lNumIntDocCredito() As Long
    lNumIntDocCredito = mvarlNumIntDocCredito
End Property

Public Property Let lNumIntDocAporte(ByVal vData As Long)
    mvarlNumIntDocAporte = vData
End Property

Public Property Get lNumIntDocAporte() As Long
    lNumIntDocAporte = mvarlNumIntDocAporte
End Property

Public Property Let lNumIntDocPagto(ByVal vData As Long)
    mvarlNumIntDocPagto = vData
End Property

Public Property Get lNumIntDocPagto() As Long
    lNumIntDocPagto = mvarlNumIntDocPagto
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtDataVenc(ByVal vData As Date)
    mvardtDataVenc = vData
End Property

Public Property Get dtDataVenc() As Date
    dtDataVenc = mvardtDataVenc
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let dSaldoAtual(ByVal vData As Double)
    mvardSaldoAtual = vData
End Property

Public Property Get dSaldoAtual() As Double
    dSaldoAtual = mvardSaldoAtual
End Property

Public Property Set objCredito(ByVal vData As Object)
    Set mvarobjCredito = vData
End Property

Public Property Get objCredito() As Object
    Set objCredito = mvarobjCredito
End Property

Public Property Let dSaldoAnt(ByVal vData As Double)
    mvardSaldoAnt = vData
End Property

Public Property Get dSaldoAnt() As Double
    dSaldoAnt = mvardSaldoAnt
End Property



