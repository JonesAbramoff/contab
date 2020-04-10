VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Attribute VB_Name = "ClassMapaDeEntregaNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level", "; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocMapa AS Long
Private mvarlNumIntDocNF AS Long
Private mvarlSeq AS Long

Public Property Let lNumIntDocMapa (ByVal vData As Long)
    mvarlNumIntDocMapa = vData
End Property

Public Property Get lNumIntDocMapa () AS Long
    lNumIntDocMapa= mvarlNumIntDocMapa
End Property

Public Property Let lNumIntDocNF (ByVal vData As Long)
    mvarlNumIntDocNF = vData
End Property

Public Property Get lNumIntDocNF () AS Long
    lNumIntDocNF= mvarlNumIntDocNF
End Property

Public Property Let lSeq (ByVal vData As Long)
    mvarlSeq = vData
End Property

Public Property Get lSeq () AS Long
    lSeq= mvarlSeq
End Property
