VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivRegESItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocRegES As Long
Private mvariNumItem As Integer
Private mvarsCFOP As String
Private mvarlNumIntCadProd As Long
Private mvardQuantidade As Double
Private mvardValorProduto As Double
Private mvardValorDescontoDespAcess As Double
Private mvardValorICMSBase As Double
Private mvardValorICMSSubstBase As Double
Private mvardValorIPI As Double
Private mvardAliquotaICMS As Double
Private mvariTipoTribICMS As Integer 'local copy
Private mvariTipoTribIPI As Integer 'local copy
Private mvardValorICMS As Double 'local copy
Private mvardValorIPIBase As Double 'local copy
Private mvardAliquotaIPI As Double 'local copy
Private mvardRedBaseICMS As Double 'local copy
Private mvardRedBaseIPI As Double 'local copy
Private mvariTipoTributacao As Integer 'local copy
Private mvardAliquotaSubst As Double 'local copy
Private mvardValorSubst As Double 'local copy

Public dValorICMSCredito As Double

Public Property Let dValorSubst(ByVal vData As Double)
    mvardValorSubst = vData
End Property

Public Property Get dValorSubst() As Double
    dValorSubst = mvardValorSubst
End Property

Public Property Let dAliquotaSubst(ByVal vData As Double)
    mvardAliquotaSubst = vData
End Property

Public Property Get dAliquotaSubst() As Double
    dAliquotaSubst = mvardAliquotaSubst
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let dRedBaseIPI(ByVal vData As Double)
    mvardRedBaseIPI = vData
End Property

Public Property Get dRedBaseIPI() As Double
    dRedBaseIPI = mvardRedBaseIPI
End Property

Public Property Let dRedBaseICMS(ByVal vData As Double)
    mvardRedBaseICMS = vData
End Property

Public Property Get dRedBaseICMS() As Double
    dRedBaseICMS = mvardRedBaseICMS
End Property

Public Property Let dAliquotaIPI(ByVal vData As Double)
    mvardAliquotaIPI = vData
End Property

Public Property Get dAliquotaIPI() As Double
    dAliquotaIPI = mvardAliquotaIPI
End Property

Public Property Let dValorIPIBase(ByVal vData As Double)
    mvardValorIPIBase = vData
End Property

Public Property Get dValorIPIBase() As Double
    dValorIPIBase = mvardValorIPIBase
End Property

Public Property Let dValorICMS(ByVal vData As Double)
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
End Property

Public Property Let iTipoTribIPI(ByVal vData As Integer)
    mvariTipoTribIPI = vData
End Property

Public Property Get iTipoTribIPI() As Integer
    iTipoTribIPI = mvariTipoTribIPI
End Property

Public Property Let iTipoTribICMS(ByVal vData As Integer)
    mvariTipoTribICMS = vData
End Property

Public Property Get iTipoTribICMS() As Integer
    iTipoTribICMS = mvariTipoTribICMS
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
Attribute dAliquotaICMS.VB_Description = "Aliquota de ICMS do Item NF"
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dValorIPI(ByVal vData As Double)
Attribute dValorIPI.VB_Description = "Valor IPI do Item NF"
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorICMSSubstBase(ByVal vData As Double)
Attribute dValorICMSSubstBase.VB_Description = "Valor ICMS Base Subst. do Item NF"
    mvardValorICMSSubstBase = vData
End Property

Public Property Get dValorICMSSubstBase() As Double
    dValorICMSSubstBase = mvardValorICMSSubstBase
End Property

Public Property Let dValorICMSBase(ByVal vData As Double)
Attribute dValorICMSBase.VB_Description = "Valor ICMS Base do Item NF"
    mvardValorICMSBase = vData
End Property

Public Property Get dValorICMSBase() As Double
    dValorICMSBase = mvardValorICMSBase
End Property

Public Property Let dValorDescontoDespAcess(ByVal vData As Double)
Attribute dValorDescontoDespAcess.VB_Description = "valor de frete, seguro, outras ou desconto a nivel de item incluindo desconto global rateado"
    mvardValorDescontoDespAcess = vData
End Property

Public Property Get dValorDescontoDespAcess() As Double
    dValorDescontoDespAcess = mvardValorDescontoDespAcess
End Property

Public Property Let dValorProduto(ByVal vData As Double)
Attribute dValorProduto.VB_Description = "Valor Total dos Produtos"
    mvardValorProduto = vData
End Property

Public Property Get dValorProduto() As Double
    dValorProduto = mvardValorProduto
End Property

Public Property Let dQuantidade(ByVal vData As Double)
Attribute dQuantidade.VB_Description = "Quantidade do Item NF"
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let lNumIntCadProd(ByVal vData As Long)
Attribute lNumIntCadProd.VB_Description = "Número Interno do CadProdutos a que ficou associado este Item NF"
    mvarlNumIntCadProd = vData
End Property

Public Property Get lNumIntCadProd() As Long
    lNumIntCadProd = mvarlNumIntCadProd
End Property

Public Property Let sCFOP(ByVal vData As String)
Attribute sCFOP.VB_Description = "Natureza de Operação do Item da NF"
    mvarsCFOP = vData
End Property

Public Property Get sCFOP() As String
    sCFOP = mvarsCFOP
End Property

Public Property Let iNumItem(ByVal vData As Integer)
Attribute iNumItem.VB_Description = "numero de ordem do item na nf"
    mvariNumItem = vData
End Property

Public Property Get iNumItem() As Integer
    iNumItem = mvariNumItem
End Property

Public Property Let lNumIntDocRegES(ByVal vData As Long)
Attribute lNumIntDocRegES.VB_Description = "Número Interno do Reg ES a que ficou associado este item NF"
    mvarlNumIntDocRegES = vData
End Property

Public Property Get lNumIntDocRegES() As Long
    lNumIntDocRegES = mvarlNumIntDocRegES
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Número Interno"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
