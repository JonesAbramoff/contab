VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRecebimentoCarneParc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsCarneNumero As String
Private mvariParcelaNumero As Integer
Private mvarlParcelaNumIntDoc As Long
Private mvardtParcelaVencimento As Date
Private mvardParcelaValor As Double
Private mvardParcelaDesconto As Double
Private mvardParcelaJuros As Double
Private mvardParcelaMulta As Double
Private mvarlCliente As Long
Private mvarlNumIntDoc As Long
Private mvardtDataBaixa As Date
Private mvariStatus As Integer

Public Property Let lNumIntDoc(ByVal vdata As Long)
    mvarlNumIntDoc = vdata
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lCliente(ByVal vdata As Long)
    mvarlCliente = vdata
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iStatus(ByVal vdata As Integer)
    mvariStatus = vdata
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iParcelaNumero(ByVal vdata As Integer)
    mvariParcelaNumero = vdata
End Property

Public Property Get iParcelaNumero() As Integer
    iParcelaNumero = mvariParcelaNumero
End Property

Public Property Let sCarneNumero(ByVal vdata As String)
    mvarsCarneNumero = vdata
End Property

Public Property Get sCarneNumero() As String
    sCarneNumero = mvarsCarneNumero
End Property

Public Property Let lParcelaNumIntDoc(ByVal vdata As Long)
    mvarlParcelaNumIntDoc = vdata
End Property

Public Property Get lParcelaNumIntDoc() As Long
    lParcelaNumIntDoc = mvarlParcelaNumIntDoc
End Property

Public Property Let dtParcelaVencimento(ByVal vdata As Date)
    mvardtParcelaVencimento = vdata
End Property

Public Property Get dtParcelaVencimento() As Date
    dtParcelaVencimento = mvardtParcelaVencimento
End Property

Public Property Let dtDataBaixa(ByVal vdata As Date)
    mvardtDataBaixa = vdata
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dParcelaDesconto(ByVal vdata As Double)
    mvardParcelaDesconto = vdata
End Property

Public Property Get dParcelaDesconto() As Double
    dParcelaDesconto = mvardParcelaDesconto
End Property

Public Property Let dParcelaValor(ByVal vdata As Double)
    mvardParcelaValor = vdata
End Property

Public Property Get dParcelaValor() As Double
    dParcelaValor = mvardParcelaValor
End Property

Public Property Let dParcelaMulta(ByVal vdata As Double)
    mvardParcelaMulta = vdata
End Property

Public Property Get dParcelaMulta() As Double
    dParcelaMulta = mvardParcelaMulta
End Property

Public Property Let dParcelaJuros(ByVal vdata As Double)
    mvardParcelaJuros = vdata
End Property

Public Property Get dParcelaJuros() As Double
    dParcelaJuros = mvardParcelaJuros
End Property



