VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassComisInpalPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvariVendedor As Integer
Private mvariTecnico As Integer
Private mvardPercComissaoEmissao As Double
Private mvardPercComissaoBaixa As Double
Private mvariComissaoSobreTotal As Integer
Private mvariComissaoFrete As Integer
Private mvariComissaoDesp As Integer
Private mvariComissaoIPI As Integer
Private mvariComissaoSeguro As Integer
Private mvardAjudaCusto As Double
Private mvariTipoAjudaCusto As Integer
Private mvarcolComissoesInpalRegras As New Collection

Public Property Set colComissoesInpalRegras(ByVal vData As Object)
    Set mvarcolComissoesInpalRegras = vData
End Property

Public Property Get colComissoesInpalRegras() As Collection
    Set colComissoesInpalRegras = mvarcolComissoesInpalRegras
End Property


Public Property Let iTipoAjudaCusto(ByVal vData As Integer)
    mvariTipoAjudaCusto = vData
End Property

Public Property Get iTipoAjudaCusto() As Integer
    iTipoAjudaCusto = mvariTipoAjudaCusto
End Property

Public Property Let dAjudaCusto(ByVal vData As Double)
    mvardAjudaCusto = vData
End Property

Public Property Get dAjudaCusto() As Double
    dAjudaCusto = mvardAjudaCusto
End Property

Public Property Let iComissaoSeguro(ByVal vData As Integer)
    mvariComissaoSeguro = vData
End Property

Public Property Get iComissaoSeguro() As Integer
    iComissaoSeguro = mvariComissaoSeguro
End Property

Public Property Let iComissaoIPI(ByVal vData As Integer)
    mvariComissaoIPI = vData
End Property

Public Property Get iComissaoIPI() As Integer
    iComissaoIPI = mvariComissaoIPI
End Property

Public Property Let iComissaoDesp(ByVal vData As Integer)
    mvariComissaoDesp = vData
End Property

Public Property Get iComissaoDesp() As Integer
    iComissaoDesp = mvariComissaoDesp
End Property

Public Property Let iComissaoFrete(ByVal vData As Integer)
    mvariComissaoFrete = vData
End Property

Public Property Get iComissaoFrete() As Integer
    iComissaoFrete = mvariComissaoFrete
End Property

Public Property Let iComissaoSobreTotal(ByVal vData As Integer)
    mvariComissaoSobreTotal = vData
End Property

Public Property Get iComissaoSobreTotal() As Integer
    iComissaoSobreTotal = mvariComissaoSobreTotal
End Property

Public Property Let dPercComissaoBaixa(ByVal vData As Double)
    mvardPercComissaoBaixa = vData
End Property

Public Property Get dPercComissaoBaixa() As Double
    dPercComissaoBaixa = mvardPercComissaoBaixa
End Property

Public Property Let dPercComissaoEmissao(ByVal vData As Double)
    mvardPercComissaoEmissao = vData
End Property

Public Property Get dPercComissaoEmissao() As Double
    dPercComissaoEmissao = mvardPercComissaoEmissao
End Property

Public Property Let iTecnico(ByVal vData As Integer)
    mvariTecnico = vData
End Property

Public Property Get iTecnico() As Integer
    iTecnico = mvariTecnico
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property
