VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMapaDeEntregaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocMapa As Long
Private mvarlNumIntDoc As Long
Private mvarlSeq As Long

Public Property Let lNumIntDocMapa(ByVal vData As Long)
    mvarlNumIntDocMapa = vData
End Property

Public Property Get lNumIntDocMapa() As Long
    lNumIntDocMapa = mvarlNumIntDocMapa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lSeq(ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq() As Long
    lSeq = mvarlSeq
End Property
