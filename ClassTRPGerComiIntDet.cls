VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRPGerComiIntDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocGerComi As Long
Private mvariSeq As Integer
Private mvarlNumIntDocComi As Long
Private mvardValorBase As Double
Private mvardValorComissao As Double
Private mvariVendedor As Integer
Private mvardtDataGeracao As Date
Private mvarsNomeReduzidoVendedor As String
Private mvardPercComissao As Double

Public Property Let dPercComissao(ByVal vData As Double)
    mvardPercComissao = vData
End Property

Public Property Get dPercComissao() As Double
    dPercComissao = mvardPercComissao
End Property

Public Property Let sNomeReduzidoVendedor(ByVal vData As String)
    mvarsNomeReduzidoVendedor = vData
End Property

Public Property Get sNomeReduzidoVendedor() As String
    sNomeReduzidoVendedor = mvarsNomeReduzidoVendedor
End Property

Public Property Let dtDataGeracao(ByVal vData As Date)
    mvardtDataGeracao = vData
End Property

Public Property Get dtDataGeracao() As Date
    dtDataGeracao = mvardtDataGeracao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocGerComi(ByVal vData As Long)
    mvarlNumIntDocGerComi = vData
End Property

Public Property Get lNumIntDocGerComi() As Long
    lNumIntDocGerComi = mvarlNumIntDocGerComi
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lNumIntDocComi(ByVal vData As Long)
    mvarlNumIntDocComi = vData
End Property

Public Property Get lNumIntDocComi() As Long
    lNumIntDocComi = mvarlNumIntDocComi
End Property

Public Property Let dValorBase(ByVal vData As Double)
    mvardValorBase = vData
End Property

Public Property Get dValorBase() As Double
    dValorBase = mvardValorBase
End Property

Public Property Let dValorComissao(ByVal vData As Double)
    mvardValorComissao = vData
End Property

Public Property Get dValorComissao() As Double
    dValorComissao = mvardValorComissao
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property
