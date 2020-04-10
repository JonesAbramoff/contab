VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemRCItemPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardQuantidade As Double
Private mvardQuantRecebida As Double
Private mvarlItemPC As Long
Private mvarlItemRC As Long
Public Property Let lItemRC(ByVal vData As Long)
    mvarlItemRC = vData
End Property


Public Property Get lItemRC() As Long
    lItemRC = mvarlItemRC
End Property



Public Property Let lItemPC(ByVal vData As Long)
    mvarlItemPC = vData
End Property


Public Property Get lItemPC() As Long
    lItemPC = mvarlItemPC
End Property



Public Property Let dQuantRecebida(ByVal vData As Double)
    mvardQuantRecebida = vData
End Property


Public Property Get dQuantRecebida() As Double
    dQuantRecebida = mvardQuantRecebida
End Property



Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property



