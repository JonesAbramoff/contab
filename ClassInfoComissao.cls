VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassInfoComissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntCom As Long
Private mvariTipoTitulo As Integer
Private mvarlNumIntDoc As Long
Private mvariCodVendedor As Integer
Private mvardtDataBaixa As Date
Private mvardPercentual As Double
Private mvardValorBase As Double
Private mvardValor As Double
Private mvariStatus As Integer
Private mvariFilialEmpresa As Integer
Private mvarsVendedorNomeRed As String
Private mvardtDataGeracao As Date

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

Public Property Let sVendedorNomeRed(ByVal vData As String)
    mvarsVendedorNomeRed = vData
End Property

Public Property Get sVendedorNomeRed() As String
    sVendedorNomeRed = mvarsVendedorNomeRed
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dValorBase(ByVal vData As Double)
    mvardValorBase = vData
End Property

Public Property Get dValorBase() As Double
    dValorBase = mvardValorBase
End Property

Public Property Let dPercentual(ByVal vData As Double)
    mvardPercentual = vData
End Property

Public Property Get dPercentual() As Double
    dPercentual = mvardPercentual
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let iCodVendedor(ByVal vData As Integer)
    mvariCodVendedor = vData
End Property

Public Property Get iCodVendedor() As Integer
    iCodVendedor = mvariCodVendedor
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iTipoTitulo(ByVal vData As Integer)
    mvariTipoTitulo = vData
End Property

Public Property Get iTipoTitulo() As Integer
    iTipoTitulo = mvariTipoTitulo
End Property

Public Property Let lNumIntCom(ByVal vData As Long)
    mvarlNumIntCom = vData
End Property

Public Property Get lNumIntCom() As Long
    lNumIntCom = mvarlNumIntCom
End Property



