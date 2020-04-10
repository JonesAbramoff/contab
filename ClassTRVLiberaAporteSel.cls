VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVLiberaAporteSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarcolAportes As New Collection
Private mvardtDataPagtoDe As Date
Private mvarlCodigoDe As Long
Private mvardtDataPagtoAte As Date
Private mvarlCodigoAte As Long

Public Property Let lCodigoDe(ByVal vData As Long)
    mvarlCodigoDe = vData
End Property

Public Property Get lCodigoDe() As Long
    lCodigoDe = mvarlCodigoDe
End Property

Public Property Let lCodigoAte(ByVal vData As Long)
    mvarlCodigoAte = vData
End Property

Public Property Get lCodigoAte() As Long
    lCodigoAte = mvarlCodigoAte
End Property

Public Property Let dtDataPagtoDe(ByVal vData As Date)
    mvardtDataPagtoDe = vData
End Property

Public Property Get dtDataPagtoDe() As Date
    dtDataPagtoDe = mvardtDataPagtoDe
End Property

Public Property Let dtDataPagtoAte(ByVal vData As Date)
    mvardtDataPagtoAte = vData
End Property

Public Property Get dtDataPagtoAte() As Date
    dtDataPagtoAte = mvardtDataPagtoAte
End Property


Public Property Set colAportes(ByVal vData As Collection)
    Set mvarcolAportes = vData
End Property

Public Property Get colAportes() As Collection
    Set colAportes = mvarcolAportes
End Property
