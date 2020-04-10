VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPedCompraEnvio"
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
Private mvarcolPedidos As New Collection

Public Property Set colPedidos(ByVal vData As Collection)
    Set mvarcolPedidos = vData
End Property

Public Property Get colPedidos() As Collection
    Set colPedidos = mvarcolPedidos
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



