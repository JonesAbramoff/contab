VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTipoVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvardPercComissao As Double
Private mvardPercComissaoBaixa As Double
Private mvardPercComissaoEmissao As Double
Private mvariComissaoSobreTotal As Integer
Private mvariComissaoFrete As Integer
Private mvariComissaoICM As Integer
Private mvariComissaoIPI As Integer
Private mvarsDescricao As String
Private mvariComissaoSeguro As Integer

Public Property Let iComissaoSeguro(ByVal vData As Integer)
    mvariComissaoSeguro = vData
End Property

Public Property Get iComissaoSeguro() As Integer
    iComissaoSeguro = mvariComissaoSeguro
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iComissaoIPI(ByVal vData As Integer)
    mvariComissaoIPI = vData
End Property

Public Property Get iComissaoIPI() As Integer
    iComissaoIPI = mvariComissaoIPI
End Property

Public Property Let iComissaoICM(ByVal vData As Integer)
    mvariComissaoICM = vData
End Property

Public Property Get iComissaoICM() As Integer
    iComissaoICM = mvariComissaoICM
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

Public Property Let dPercComissaoEmissao(ByVal vData As Double)
    mvardPercComissaoEmissao = vData
End Property

Public Property Get dPercComissaoEmissao() As Double
    dPercComissaoEmissao = mvardPercComissaoEmissao
End Property

Public Property Let dPercComissaoBaixa(ByVal vData As Double)
    mvardPercComissaoBaixa = vData
End Property

Public Property Get dPercComissaoBaixa() As Double
    dPercComissaoBaixa = mvardPercComissaoBaixa
End Property

Public Property Let dPercComissao(ByVal vData As Double)
    mvardPercComissao = vData
End Property

Public Property Get dPercComissao() As Double
    dPercComissao = mvardPercComissao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



