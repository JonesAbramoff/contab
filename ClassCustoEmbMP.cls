VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCustoEmbMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvardtDataAtualizacao As Date
Private mvardCusto As Double
Private mvardAliquotaICMS As Double
Private mvariCondicaoPagto As Integer
Private mvardFretePorKG As Double

Private mvariCondicaoPagtoInf As Integer
Private mvariAliquotaICMSInf As Integer
Private mvariFretePorKGInf As Integer
Private mvarcolAux As New Collection
Private mvardtDataUltCalc As Date
Private mvardCustoUltCalc As Double
Private mvardtDataCalcAnt As Date
Private mvardCustoCalcAnt As Double

Public Property Let dCustoCalcAnt(ByVal vData As Double)
    mvardCustoCalcAnt = vData
End Property

Public Property Get dCustoCalcAnt() As Double
    dCustoCalcAnt = mvardCustoCalcAnt
End Property

Public Property Let dtDataCalcAnt(ByVal vData As Date)
    mvardtDataCalcAnt = vData
End Property

Public Property Get dtDataCalcAnt() As Date
    dtDataCalcAnt = mvardtDataCalcAnt
End Property

Public Property Let dCustoUltCalc(ByVal vData As Double)
    mvardCustoUltCalc = vData
End Property

Public Property Get dCustoUltCalc() As Double
    dCustoUltCalc = mvardCustoUltCalc
End Property

Public Property Let dtDataUltCalc(ByVal vData As Date)
    mvardtDataUltCalc = vData
End Property

Public Property Get dtDataUltCalc() As Date
    dtDataUltCalc = mvardtDataUltCalc
End Property

Public Property Set colAux(ByVal vData As Object)
    Set mvarcolAux = vData
End Property

Public Property Get colAux() As Collection
    Set colAux = mvarcolAux
End Property

Public Property Let iFretePorKGInf(ByVal vData As Integer)
    mvariFretePorKGInf = vData
End Property

Public Property Get iFretePorKGInf() As Integer
    iFretePorKGInf = mvariFretePorKGInf
End Property

Public Property Let iAliquotaICMSInf(ByVal vData As Integer)
    mvariAliquotaICMSInf = vData
End Property

Public Property Get iAliquotaICMSInf() As Integer
    iAliquotaICMSInf = mvariAliquotaICMSInf
End Property

Public Property Let iCondicaoPagtoInf(ByVal vData As Integer)
    mvariCondicaoPagtoInf = vData
End Property

Public Property Get iCondicaoPagtoInf() As Integer
    iCondicaoPagtoInf = mvariCondicaoPagtoInf
End Property

Public Property Let dFretePorKg(ByVal vData As Double)
    mvardFretePorKG = vData
End Property

Public Property Get dFretePorKg() As Double
    dFretePorKg = mvardFretePorKG
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dCusto(ByVal vData As Double)
    mvardCusto = vData
End Property

Public Property Get dCusto() As Double
    dCusto = mvardCusto
End Property

Public Property Let dtDataAtualizacao(ByVal vData As Date)
    mvardtDataAtualizacao = vData
End Property

Public Property Get dtDataAtualizacao() As Date
    dtDataAtualizacao = mvardtDataAtualizacao
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


