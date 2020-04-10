VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassTRVCliDataCallCenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCliente AS Long
Private mvardtDataDe AS Date
Private mvardtDataAte AS Date

Public Property Let lCliente (ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente () AS Long
    lCliente= mvarlCliente
End Property

Public Property Let dtDataDe (ByVal vData As Date)
    mvardtDataDe = vData
End Property

Public Property Get dtDataDe () AS Date
    dtDataDe= mvardtDataDe
End Property

Public Property Let dtDataAte (ByVal vData As Date)
    mvardtDataAte = vData
End Property

Public Property Get dtDataAte () AS Date
    dtDataAte= mvardtDataAte
End Property
