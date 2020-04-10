VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCarneParcelasImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlCodCliente As Long
Private mvarsNomeCliente As String
Private mvarsCPFCGCCliente As String
Private mvarsCodCarne As String
Private mvariParcelaNumero As Integer
Private mvardtDataRefCarne As Date
Private mvardtDataVencParcela As Date
Private mvardParcelaValor As Double
Private mvardDesconto As Double
Private mvardMulta As Double
Private mvardJuros As Double
Private mvarobjBaixasParcCarne As ClassBaixasParcCarne
Private mvarlNumIntDocParc As Long
Private mvardValorBaixado As Double

Public Property Let dValorBaixado(ByVal vdata As Double)
    mvardValorBaixado = vdata
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let lNumIntDocParc(ByVal vdata As Long)
    mvarlNumIntDocParc = vdata
End Property

Public Property Get lNumIntDocParc() As Long
    lNumIntDocParc = mvarlNumIntDocParc
End Property

Public Property Set objBaixasParcCarne(ByVal vdata As ClassBaixasParcCarne)
    Set mvarobjBaixasParcCarne = vdata
End Property

Public Property Get objBaixasParcCarne() As ClassBaixasParcCarne
    Set objBaixasParcCarne = mvarobjBaixasParcCarne
End Property

Public Property Let dJuros(ByVal vdata As Double)
    mvardJuros = vdata
End Property

Public Property Get dJuros() As Double
    dJuros = mvardJuros
End Property

Public Property Let dMulta(ByVal vdata As Double)
    mvardMulta = vdata
End Property

Public Property Get dMulta() As Double
    dMulta = mvardMulta
End Property

Public Property Let dDesconto(ByVal vdata As Double)
    mvardDesconto = vdata
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let sCodCarne(ByVal vdata As String)
    mvarsCodCarne = vdata
End Property

Public Property Get sCodCarne() As String
    sCodCarne = mvarsCodCarne
End Property

Public Property Let sCPFCGCCliente(ByVal vdata As String)
    mvarsCPFCGCCliente = vdata
End Property

Public Property Get sCPFCGCCliente() As String
    sCPFCGCCliente = mvarsCPFCGCCliente
End Property

Public Property Let sNomeCliente(ByVal vdata As String)
    mvarsNomeCliente = vdata
End Property

Public Property Get sNomeCliente() As String
    sNomeCliente = mvarsNomeCliente
End Property

Public Property Let dtDataVencParcela(ByVal vdata As Date)
    mvardtDataVencParcela = vdata
End Property

Public Property Get dtDataVencParcela() As Date
    dtDataVencParcela = mvardtDataVencParcela
End Property

Public Property Let dtDataRefCarne(ByVal vdata As Date)
    mvardtDataRefCarne = vdata
End Property

Public Property Get dtDataRefCarne() As Date
    dtDataRefCarne = mvardtDataRefCarne
End Property

Public Property Let lCodCliente(ByVal vdata As Long)
    mvarlCodCliente = vdata
End Property

Public Property Get lCodCliente() As Long
    lCodCliente = mvarlCodCliente
End Property

Public Property Let iParcelaNumero(ByVal vdata As Integer)
    mvariParcelaNumero = vdata
End Property

Public Property Get iParcelaNumero() As Integer
    iParcelaNumero = mvariParcelaNumero
End Property

Public Property Let dParcelaValor(ByVal vdata As Double)
    mvardParcelaValor = vdata
End Property

Public Property Get dParcelaValor() As Double
    dParcelaValor = mvardParcelaValor
End Property

