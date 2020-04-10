VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassReqCompraEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardtDataDe As Date
Private mvardtDataAte As Date
Private mvarlCodigoDe As Long
Private mvarlCodigoAte As Long
Private mvarcolRequisicao As New Collection
Private mvardtDataLimiteDe As Date
Private mvardtDataLimiteAte As Date
Private mvardtDataLimCDe As Date
Private mvardtDataLimCAte As Date
Private mvariTipo As Integer

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
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

Public Property Let dtDataLimCAte(ByVal vData As Date)
    mvardtDataLimCAte = vData
End Property

Public Property Get dtDataLimCAte() As Date
    dtDataLimCAte = mvardtDataLimCAte
End Property

Public Property Let dtDataLimCDe(ByVal vData As Date)
    mvardtDataLimCDe = vData
End Property

Public Property Get dtDataLimCDe() As Date
    dtDataLimCDe = mvardtDataLimCDe
End Property

Public Property Set colRequisicao(ByVal vData As Collection)
    Set mvarcolRequisicao = vData
End Property

Public Property Get colRequisicao() As Collection
    Set colRequisicao = mvarcolRequisicao
End Property

Public Property Let lCodigoAte(ByVal vData As Long)
    mvarlCodigoAte = vData
End Property

Public Property Get lCodigoAte() As Long
    lCodigoAte = mvarlCodigoAte
End Property

Public Property Let lCodigoDe(ByVal vData As Long)
    mvarlCodigoDe = vData
End Property

Public Property Get lCodigoDe() As Long
    lCodigoDe = mvarlCodigoDe
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



