VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemNFItemRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlItemReqCompra As Long
Private mvardQuantidade As Double
Private mvarlItemNFiscal As Long
Private mvarlNumIntDocItemPC As Long

Public Property Let lNumIntDocItemPC(ByVal vData As Long)
    mvarlNumIntDocItemPC = vData
End Property

Public Property Get lNumIntDocItemPC() As Long
    lNumIntDocItemPC = mvarlNumIntDocItemPC
End Property

Public Property Let lItemNFiscal(ByVal vData As Long)
    mvarlItemNFiscal = vData
End Property

Public Property Get lItemNFiscal() As Long
    lItemNFiscal = mvarlItemNFiscal
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let lItemReqCompra(ByVal vData As Long)
Attribute lItemReqCompra.VB_Description = "NumInt do ItemReqCompra"
    mvarlItemReqCompra = vData
End Property

Public Property Get lItemReqCompra() As Long
    lItemReqCompra = mvarlItemReqCompra
End Property



