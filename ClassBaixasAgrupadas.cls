VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassBaixasAgrupadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc AS Long
Private mvardtDataBaixa AS Date
Private mvarsUsuario AS String

Public Property Let lNumIntDoc (ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc () AS Long
    lNumIntDoc= mvarlNumIntDoc
End Property

Public Property Let dtDataBaixa (ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa () AS Date
    dtDataBaixa= mvardtDataBaixa
End Property

Public Property Let sUsuario (ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario () AS String
    sUsuario= mvarsUsuario
End Property
