VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixasParcCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvarlNumIntBaixa As Long
Private mvarlNumIntParcelaCarne As Long
Private mvardValorBaixado As Double
Private mvardValorDesconto As Double
Private mvardValorMulta As Double
Private mvardValorJuros As Double
Private mvarsCodCarne As String
Private mvarlNumIntDoc As Long

Public Property Let lNumIntDoc(ByVal vdata As Long)
    mvarlNumIntDoc = vdata
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sCodCarne(ByVal vdata As String)
    mvarsCodCarne = vdata
End Property

Public Property Get sCodCarne() As String
    sCodCarne = mvarsCodCarne
End Property

Public Property Let lNumIntBaixa(ByVal vdata As Long)
    mvarlNumIntBaixa = vdata
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property

Public Property Let lNumIntParcelaCarne(ByVal vdata As Long)
    mvarlNumIntParcelaCarne = vdata
End Property

Public Property Get lNumIntParcelaCarne() As Long
    lNumIntParcelaCarne = mvarlNumIntParcelaCarne
End Property

Public Property Let dValorDesconto(ByVal vdata As Double)
    mvardValorDesconto = vdata
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorBaixado(ByVal vdata As Double)
    mvardValorBaixado = vdata
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dValorMulta(ByVal vdata As Double)
    mvardValorMulta = vdata
End Property

Public Property Get dValorMulta() As Double
    dValorMulta = mvardValorMulta
End Property

Public Property Let dValorJuros(ByVal vdata As Double)
    mvardValorJuros = vdata
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let iFilialEmpresa(ByVal vdata As Integer)
    mvariFilialEmpresa = vdata
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

