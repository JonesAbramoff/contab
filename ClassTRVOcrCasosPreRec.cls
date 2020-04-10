VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasosPreRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocOcrCaso As Long
Private mvardValor As Double
Private mvardtData As Date
Private mvarlNumIntDocTitRecReembolso As Long
Private mvarlNumFatTitRecReembolso As Long
Private mvarsDescricao As String
Private mvardtDataPagto As Date

Public Property Let lNumIntDocOcrCaso(ByVal vData As Long)
    mvarlNumIntDocOcrCaso = vData
End Property

Public Property Get lNumIntDocOcrCaso() As Long
    lNumIntDocOcrCaso = mvarlNumIntDocOcrCaso
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lNumIntDocTitRecReembolso(ByVal vData As Long)
    mvarlNumIntDocTitRecReembolso = vData
End Property

Public Property Get lNumIntDocTitRecReembolso() As Long
    lNumIntDocTitRecReembolso = mvarlNumIntDocTitRecReembolso
End Property

Public Property Let lNumFatTitRecReembolso(ByVal vData As Long)
    mvarlNumFatTitRecReembolso = vData
End Property

Public Property Get lNumFatTitRecReembolso() As Long
    lNumFatTitRecReembolso = mvarlNumFatTitRecReembolso
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dtDataPagto(ByVal vData As Date)
    mvardtDataPagto = vData
End Property

Public Property Get dtDataPagto() As Date
    dtDataPagto = mvardtDataPagto
End Property

Private Sub Class_Initialize()
    mvardtDataPagto = DATA_NULA
End Sub
