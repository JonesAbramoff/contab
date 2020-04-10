VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelAcompPRJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvariSeq As Integer
Private mvarlNumIntDocPRJ As Long
Private mvarsProjeto As String
Private mvarsDescricaoPRJ As String
Private mvarlCliente As Long
Private mvarsNomeCliente As String
Private mvarlNumIntDocEtapa As Long
Private mvarsEtapa As String
Private mvarsDescricaoEtapa As String
Private mvardCustoPrev As Double
Private mvardCustoReal As Double
Private mvardPercCompPRJ As Double
Private mvardPercCompEtapa As Double
Private mvardtDataIniPrevEtapa As Date
Private mvardtDataFimPrevEtapa As Date
Private mvardtDataIniRealEtapa As Date
Private mvardtDataFimRealEtapa As Date
Private mvardtDataIniPrevPRJ As Date
Private mvardtDataFimPrevPRJ As Date
Private mvardtDataIniRealPRJ As Date
Private mvardtDataFimRealPRJ As Date
Private mvarcolDet As New Collection

Public Property Set colDet(ByVal vData As Collection)
    Set mvarcolDet = vData
End Property

Public Property Get colDet() As Collection
    Set colDet = mvarcolDet
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Let sProjeto(ByVal vData As String)
    mvarsProjeto = vData
End Property

Public Property Get sProjeto() As String
    sProjeto = mvarsProjeto
End Property

Public Property Let sDescricaoPRJ(ByVal vData As String)
    mvarsDescricaoPRJ = vData
End Property

Public Property Get sDescricaoPRJ() As String
    sDescricaoPRJ = mvarsDescricaoPRJ
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let sNomeCliente(ByVal vData As String)
    mvarsNomeCliente = vData
End Property

Public Property Get sNomeCliente() As String
    sNomeCliente = mvarsNomeCliente
End Property

Public Property Let lNumIntDocEtapa(ByVal vData As Long)
    mvarlNumIntDocEtapa = vData
End Property

Public Property Get lNumIntDocEtapa() As Long
    lNumIntDocEtapa = mvarlNumIntDocEtapa
End Property

Public Property Let sEtapa(ByVal vData As String)
    mvarsEtapa = vData
End Property

Public Property Get sEtapa() As String
    sEtapa = mvarsEtapa
End Property

Public Property Let sDescricaoEtapa(ByVal vData As String)
    mvarsDescricaoEtapa = vData
End Property

Public Property Get sDescricaoEtapa() As String
    sDescricaoEtapa = mvarsDescricaoEtapa
End Property

Public Property Let dCustoPrev(ByVal vData As Double)
    mvardCustoPrev = vData
End Property

Public Property Get dCustoPrev() As Double
    dCustoPrev = mvardCustoPrev
End Property

Public Property Let dCustoReal(ByVal vData As Double)
    mvardCustoReal = vData
End Property

Public Property Get dCustoReal() As Double
    dCustoReal = mvardCustoReal
End Property

Public Property Let dPercCompPRJ(ByVal vData As Double)
    mvardPercCompPRJ = vData
End Property

Public Property Get dPercCompPRJ() As Double
    dPercCompPRJ = mvardPercCompPRJ
End Property

Public Property Let dPercCompEtapa(ByVal vData As Double)
    mvardPercCompEtapa = vData
End Property

Public Property Get dPercCompEtapa() As Double
    dPercCompEtapa = mvardPercCompEtapa
End Property

Public Property Let dtDataIniPrevEtapa(ByVal vData As Date)
    mvardtDataIniPrevEtapa = vData
End Property

Public Property Get dtDataIniPrevEtapa() As Date
    dtDataIniPrevEtapa = mvardtDataIniPrevEtapa
End Property

Public Property Let dtDataFimPrevEtapa(ByVal vData As Date)
    mvardtDataFimPrevEtapa = vData
End Property

Public Property Get dtDataFimPrevEtapa() As Date
    dtDataFimPrevEtapa = mvardtDataFimPrevEtapa
End Property

Public Property Let dtDataIniRealEtapa(ByVal vData As Date)
    mvardtDataIniRealEtapa = vData
End Property

Public Property Get dtDataIniRealEtapa() As Date
    dtDataIniRealEtapa = mvardtDataIniRealEtapa
End Property

Public Property Let dtDataFimRealEtapa(ByVal vData As Date)
    mvardtDataFimRealEtapa = vData
End Property

Public Property Get dtDataFimRealEtapa() As Date
    dtDataFimRealEtapa = mvardtDataFimRealEtapa
End Property

Public Property Let dtDataIniPrevPRJ(ByVal vData As Date)
    mvardtDataIniPrevPRJ = vData
End Property

Public Property Get dtDataIniPrevPRJ() As Date
    dtDataIniPrevPRJ = mvardtDataIniPrevPRJ
End Property

Public Property Let dtDataFimPrevPRJ(ByVal vData As Date)
    mvardtDataFimPrevPRJ = vData
End Property

Public Property Get dtDataFimPrevPRJ() As Date
    dtDataFimPrevPRJ = mvardtDataFimPrevPRJ
End Property

Public Property Let dtDataIniRealPRJ(ByVal vData As Date)
    mvardtDataIniRealPRJ = vData
End Property

Public Property Get dtDataIniRealPRJ() As Date
    dtDataIniRealPRJ = mvardtDataIniRealPRJ
End Property

Public Property Let dtDataFimRealPRJ(ByVal vData As Date)
    mvardtDataFimRealPRJ = vData
End Property

Public Property Get dtDataFimRealPRJ() As Date
    dtDataFimRealPRJ = mvardtDataFimRealPRJ
End Property
