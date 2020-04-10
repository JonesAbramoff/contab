VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassBaixaParcRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntBaixa As Long
Private mvarlNumIntParcela As Long
Private mvariSequencial As Integer
Private mvariStatus As Integer
Private mvardValorMulta As Double
Private mvardValorJuros As Double
Private mvardValorDesconto As Double
Private mvardValorBaixado As Double
Private mvariCobrador As Integer
Private mvarlNumIntDoc As Long
Private mvardValorRecebido As Double

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dValorRecebido(ByVal vData As Double)
    mvardValorRecebido = vData
End Property

Public Property Get dValorRecebido() As Double
    dValorRecebido = mvardValorRecebido
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
Attribute dValorBaixado.VB_Description = "quanto deduz do saldo da parcela"
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
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

Public Property Let iStatus(ByVal vData As Integer)
Attribute iStatus.VB_Description = "1:lancado, 2:excluido"
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iSequencial(ByVal vData As Integer)
Attribute iSequencial.VB_Description = "uma parcela pode ter varias baixas parciais"
    mvariSequencial = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let lNumIntParcela(ByVal vData As Long)
Attribute lNumIntParcela.VB_Description = "referencia a tabela de parcelas a receber"
    mvarlNumIntParcela = vData
End Property

Public Property Get lNumIntParcela() As Long
    lNumIntParcela = mvarlNumIntParcela
End Property

Public Property Let lNumIntBaixa(ByVal vData As Long)
Attribute lNumIntBaixa.VB_Description = "chave primaria"
    mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property



