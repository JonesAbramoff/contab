VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVRelComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariEmpresa As Integer
Private mvarsVendedor As String
Private mvarlCliente As Long
Private mvarsNomeCliente As String
Private mvarsNomeProduto As String
Private mvardComissPerc As Double
Private mvardComissao As Double
Private mvarlNumVou As Long
Private mvardtDataEmi As Date
Private mvardtDataFat As Date
Private mvardValorBase As Double
Private mvariAporte As Integer
Private mvarlNumIntRel As Integer

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iAporte(ByVal vData As Integer)
    mvariAporte = vData
End Property

Public Property Get iAporte() As Integer
    iAporte = mvariAporte
End Property

Public Property Let dValorBase(ByVal vData As Double)
    mvardValorBase = vData
End Property

Public Property Get dValorBase() As Double
    dValorBase = mvardValorBase
End Property

Public Property Let dtDataFat(ByVal vData As Date)
    mvardtDataFat = vData
End Property

Public Property Get dtDataFat() As Date
    dtDataFat = mvardtDataFat
End Property

Public Property Let dtDataEmi(ByVal vData As Date)
    mvardtDataEmi = vData
End Property

Public Property Get dtDataEmi() As Date
    dtDataEmi = mvardtDataEmi
End Property

Public Property Let lNumVou(ByVal vData As Long)
    mvarlNumVou = vData
End Property

Public Property Get lNumVou() As Long
    lNumVou = mvarlNumVou
End Property

Public Property Let dComissao(ByVal vData As Double)
    mvardComissao = vData
End Property

Public Property Get dComissao() As Double
    dComissao = mvardComissao
End Property

Public Property Let dComissPerc(ByVal vData As Double)
    mvardComissPerc = vData
End Property

Public Property Get dComissPerc() As Double
    dComissPerc = mvardComissPerc
End Property

Public Property Let sNomeProduto(ByVal vData As String)
    mvarsNomeProduto = vData
End Property

Public Property Get sNomeProduto() As String
    sNomeProduto = mvarsNomeProduto
End Property

Public Property Let sNomeCliente(ByVal vData As String)
    mvarsNomeCliente = vData
End Property

Public Property Get sNomeCliente() As String
    sNomeCliente = mvarsNomeCliente
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let sVendedor(ByVal vData As String)
    mvarsVendedor = vData
End Property

Public Property Get sVendedor() As String
    sVendedor = mvarsVendedor
End Property

Public Property Let iEmpresa(ByVal vData As Integer)
    mvariEmpresa = vData
End Property

Public Property Get iEmpresa() As Integer
    iEmpresa = mvariEmpresa
End Property

