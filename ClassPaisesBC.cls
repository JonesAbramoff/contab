VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassPaisesBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariCodBacen AS Integer
Private mvarsPais AS String
Private mvariTribFav AS Integer

Public Property Let iCodBacen (ByVal vData As Integer)
    mvariCodBacen = vData
End Property

Public Property Get iCodBacen () AS Integer
    iCodBacen= mvariCodBacen
End Property

Public Property Let sPais (ByVal vData As String)
    mvarsPais = vData
End Property

Public Property Get sPais () AS String
    sPais= mvarsPais
End Property

Public Property Let iTribFav (ByVal vData As Integer)
    mvariTribFav = vData
End Property

Public Property Get iTribFav () AS Integer
    iTribFav= mvariTribFav
End Property
