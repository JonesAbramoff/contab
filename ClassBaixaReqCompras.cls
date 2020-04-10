VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaReqCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariSoResiduais As Integer
Private mvarlRequisicaoDe As Long
Private mvarlRequisicaoAte As Long
Private mvarsCclDe As String
Private mvarsCclAte As String
Private mvarlRequisitanteDe As Long
Private mvarlRequisitanteAte As Long
Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvardtDataLimiteDe As Date
Private mvardtDataLimiteAte As Date
Private mvarcolReqComprasInfo As New Collection
Private mvarsOrdenacao As String
Private mvardResiduo As Double

Public Property Let dResiduo(ByVal vData As Double)
Attribute dResiduo.VB_Description = "testando"
    mvardResiduo = vData
End Property

Public Property Get dResiduo() As Double
    dResiduo = mvardResiduo
End Property

Public Property Let sOrdenacao(ByVal vData As String)
    mvarsOrdenacao = vData
End Property

Public Property Get sOrdenacao() As String
    sOrdenacao = mvarsOrdenacao
End Property

Public Property Set colReqComprasInfo(ByVal vData As Object)
Attribute colReqComprasInfo.VB_Description = "coleção "
    Set mvarcolReqComprasInfo = vData
End Property

Public Property Get colReqComprasInfo() As Collection
    Set colReqComprasInfo = mvarcolReqComprasInfo
End Property

Public Property Let dtDataLimiteAte(ByVal vData As Date)
    mvardtDataLimiteAte = vData
End Property

Public Property Get dtDataLimiteAte() As Date
    dtDataLimiteAte = mvardtDataLimiteAte
End Property

Public Property Let dtDataLimiteDe(ByVal vData As Date)
    mvardtDataLimiteDe = vData
End Property

Public Property Get dtDataLimiteDe() As Date
    dtDataLimiteDe = mvardtDataLimiteDe
End Property

Public Property Let dtDataAte(ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte() As Date
    dtDataAte = mvardtDataAte
End Property

Public Property Let dtDataDe(ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe() As Date
    dtDataDe = mvardtDataDe
End Property

Public Property Let lRequisitanteAte(ByVal vData As Long)
    mvarlRequisitanteAte = vData
End Property

Public Property Get lRequisitanteAte() As Long
    lRequisitanteAte = mvarlRequisitanteAte
End Property

Public Property Let lRequisitanteDe(ByVal vData As Long)
    mvarlRequisitanteDe = vData
End Property

Public Property Get lRequisitanteDe() As Long
    lRequisitanteDe = mvarlRequisitanteDe
End Property

Public Property Let sCclAte(ByVal vData As String)
    mvarsCclAte = vData
End Property

Public Property Get sCclAte() As String
    sCclAte = mvarsCclAte
End Property

Public Property Let sCclDe(ByVal vData As String)
    mvarsCclDe = vData
End Property

Public Property Get sCclDe() As String
    sCclDe = mvarsCclDe
End Property

Public Property Let lRequisicaoAte(ByVal vData As Long)
    mvarlRequisicaoAte = vData
End Property

Public Property Get lRequisicaoAte() As Long
    lRequisicaoAte = mvarlRequisicaoAte
End Property

Public Property Let lRequisicaoDe(ByVal vData As Long)
    mvarlRequisicaoDe = vData
End Property

Public Property Get lRequisicaoDe() As Long
    lRequisicaoDe = mvarlRequisicaoDe
End Property

Public Property Let iSoResiduais(ByVal vData As Integer)
    mvariSoResiduais = vData
End Property

Public Property Get iSoResiduais() As Integer
    iSoResiduais = mvariSoResiduais
End Property
