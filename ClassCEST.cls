VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassCEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarsCodigo AS String
Private mvarsNCM AS String
Private mvarsDescricao AS String

Public Property Let sCodigo (ByVal vData As String)
    mvarsCodigo = vData
End Property

Public Property Get sCodigo () AS String
    sCodigo= mvarsCodigo
End Property

Public Property Let sNCM (ByVal vData As String)
    mvarsNCM = vData
End Property

Public Property Get sNCM () AS String
    sNCM= mvarsNCM
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property
