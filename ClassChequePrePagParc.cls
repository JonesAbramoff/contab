VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChequePrePagParc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntCheque As Long
Private mvarlNumIntParcela As Long
Private mvardValorPago As Double
Private mvardValorBaixado As Double
Private mvardValorJuros As Double
Private mvardValorMulta As Double
Private mvardValorDesconto As Double
Private mvarobjTituloPag As New ClassTituloPagar
Private mvarobjParcelaPag As New ClassParcelaPagar

Public Property Set objParcelaPag(ByVal vData As ClassParcelaPagar)
    Set mvarobjParcelaPag = vData
End Property

Public Property Get objParcelaPag() As ClassParcelaPagar
    Set objParcelaPag = mvarobjParcelaPag
End Property

Public Property Set objTituloPag(ByVal vData As ClassTituloPagar)
    Set mvarobjTituloPag = vData
End Property

Public Property Get objTituloPag() As ClassTituloPagar
    Set objTituloPag = mvarobjTituloPag
End Property

Public Property Let lNumIntCheque(ByVal vData As Long)
    mvarlNumIntCheque = vData
End Property

Public Property Get lNumIntCheque() As Long
    lNumIntCheque = mvarlNumIntCheque
End Property

Public Property Let lNumIntParcela(ByVal vData As Long)
    mvarlNumIntParcela = vData
End Property

Public Property Get lNumIntParcela() As Long
    lNumIntParcela = mvarlNumIntParcela
End Property

Public Property Let dValorPago(ByVal vData As Double)
    mvardValorPago = vData
End Property

Public Property Get dValorPago() As Double
    dValorPago = mvardValorPago
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let dValorMulta(ByVal vData As Double)
    mvardValorMulta = vData
End Property

Public Property Get dValorMulta() As Double
    dValorMulta = mvardValorMulta
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property
