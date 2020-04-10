VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPFatAporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarobjAportePagto As Object
Private mvardSaldoAtual As Double
Private mvardSaldoAnt As Double
Private mvarlCliente As Long
Private mvariProxParcela As Integer
Private mvarlCodigo As Long
Private mvariMoeda As Integer
Private mvariTipoPagtoAporte As Integer
Private mvardPrevisaoUSS As Double
Private mvardRealizadoUSS As Double
Private mvardRealizadoRS As Double
Private mvardtPrevDataDe As Date
Private mvardtPrevDataAte As Date
Private mvarlNumVouAntMeta As Long
Private mvardtDataVouAntMeta As Date
Private mvardRealizadoAntMetaUSS As Double
Private mvardRealizadoAntMetaRS As Double
Private mvariNumParcelas As Integer

Public Property Let iNumParcelas(ByVal vData As Integer)
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let dRealizadoAntMetaRS(ByVal vData As Double)
    mvardRealizadoAntMetaRS = vData
End Property

Public Property Get dRealizadoAntMetaRS() As Double
    dRealizadoAntMetaRS = mvardRealizadoAntMetaRS
End Property

Public Property Let dRealizadoAntMetaUSS(ByVal vData As Double)
    mvardRealizadoAntMetaUSS = vData
End Property

Public Property Get dRealizadoAntMetaUSS() As Double
    dRealizadoAntMetaUSS = mvardRealizadoAntMetaUSS
End Property

Public Property Let dtDataVouAntMeta(ByVal vData As Date)
    mvardtDataVouAntMeta = vData
End Property

Public Property Get dtDataVouAntMeta() As Date
    dtDataVouAntMeta = mvardtDataVouAntMeta
End Property

Public Property Let lNumVouAntMeta(ByVal vData As Long)
    mvarlNumVouAntMeta = vData
End Property

Public Property Get lNumVouAntMeta() As Long
    lNumVouAntMeta = mvarlNumVouAntMeta
End Property

Public Property Let dtPrevDataDe(ByVal vData As Date)
    mvardtPrevDataDe = vData
End Property

Public Property Get dtPrevDataDe() As Date
    dtPrevDataDe = mvardtPrevDataDe
End Property

Public Property Let dtPrevDataAte(ByVal vData As Date)
    mvardtPrevDataAte = vData
End Property

Public Property Get dtPrevDataAte() As Date
    dtPrevDataAte = mvardtPrevDataAte
End Property

Public Property Let dRealizadoRS(ByVal vData As Double)
    mvardRealizadoRS = vData
End Property

Public Property Get dRealizadoRS() As Double
    dRealizadoRS = mvardRealizadoRS
End Property

Public Property Let dRealizadoUSS(ByVal vData As Double)
    mvardRealizadoUSS = vData
End Property

Public Property Get dRealizadoUSS() As Double
    dRealizadoUSS = mvardRealizadoUSS
End Property

Public Property Let dPrevisaoUSS(ByVal vData As Double)
    mvardPrevisaoUSS = vData
End Property

Public Property Get dPrevisaoUSS() As Double
    dPrevisaoUSS = mvardPrevisaoUSS
End Property

Public Property Let iTipoPagtoAporte(ByVal vData As Integer)
    mvariTipoPagtoAporte = vData
End Property

Public Property Get iTipoPagtoAporte() As Integer
    iTipoPagtoAporte = mvariTipoPagtoAporte
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iProxParcela(ByVal vData As Integer)
    mvariProxParcela = vData
End Property

Public Property Get iProxParcela() As Integer
    iProxParcela = mvariProxParcela
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

Public Property Let dSaldoAnt(ByVal vData As Double)
    mvardSaldoAnt = vData
End Property

Public Property Get dSaldoAnt() As Double
    dSaldoAnt = mvardSaldoAnt
End Property

Public Property Set objAportePagto(ByVal vData As Object)
    Set mvarobjAportePagto = vData
End Property

Public Property Get objAportePagto() As Object
    Set objAportePagto = mvarobjAportePagto
End Property





