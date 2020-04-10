VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassDERegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDE AS Long
Private mvarsNumRegistro AS String
Private mvardtDataRegistro AS Date

Public Property Let lNumIntDE (ByVal vData As Long)
    mvarlNumIntDE = vData
End Property

Public Property Get lNumIntDE () AS Long
    lNumIntDE= mvarlNumIntDE
End Property

Public Property Let sNumRegistro (ByVal vData As String)
    mvarsNumRegistro = vData
End Property

Public Property Get sNumRegistro () AS String
    sNumRegistro= mvarsNumRegistro
End Property

Public Property Let dtDataRegistro (ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro () AS Date
    dtDataRegistro= mvardtDataRegistro
End Property
