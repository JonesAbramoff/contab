VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassCertificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlCodigo AS Long
Private mvarsDescricao AS String
Private mvarsSigla AS String
Private mvarlValidade AS Long

Public Property Let lCodigo (ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo () AS Long
    lCodigo= mvarlCodigo
End Property

Public Property Let sDescricao (ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao () AS String
    sDescricao= mvarsDescricao
End Property

Public Property Let sSigla (ByVal vData As String)
    mvarsSigla = vData
End Property

Public Property Get sSigla () AS String
    sSigla= mvarsSigla
End Property

Public Property Let lValidade (ByVal vData As Long)
    mvarlValidade = vData
End Property

Public Property Get lValidade () AS Long
    lValidade= mvarlValidade
End Property
