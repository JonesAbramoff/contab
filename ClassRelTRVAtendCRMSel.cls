VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelTRVAtendCRMSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntRel As Long
Private mvardtFiltroDataDe As Date
Private mvardtFiltroDataAte As Date
Private mvardtDataGer As Date
Private mvardHoraGer As Double
Private mvarsUsuGer As String
Private mvarcolItens As New Collection
Private mvariSoCallCenter As Integer

Public Property Let iSoCallCenter(ByVal vData As Integer)
    mvariSoCallCenter = vData
End Property

Public Property Get iSoCallCenter() As Integer
    iSoCallCenter = mvariSoCallCenter
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let dtFiltroDataDe(ByVal vData As Date)
    mvardtFiltroDataDe = vData
End Property

Public Property Get dtFiltroDataDe() As Date
    dtFiltroDataDe = mvardtFiltroDataDe
End Property

Public Property Let dtFiltroDataAte(ByVal vData As Date)
    mvardtFiltroDataAte = vData
End Property

Public Property Get dtFiltroDataAte() As Date
    dtFiltroDataAte = mvardtFiltroDataAte
End Property

Public Property Let dtDataGer(ByVal vData As Date)
    mvardtDataGer = vData
End Property

Public Property Get dtDataGer() As Date
    dtDataGer = mvardtDataGer
End Property

Public Property Let dHoraGer(ByVal vData As Double)
    mvardHoraGer = vData
End Property

Public Property Get dHoraGer() As Double
    dHoraGer = mvardHoraGer
End Property

Public Property Let sUsuGer(ByVal vData As String)
    mvarsUsuGer = vData
End Property

Public Property Get sUsuGer() As String
    sUsuGer = mvarsUsuGer
End Property
