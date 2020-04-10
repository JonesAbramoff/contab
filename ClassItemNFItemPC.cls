VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemNFItemPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvardQuantidade As Double 'local copy
Private mvarlItemPedCompra As Long 'local copy
Private mvarlItemNFiscal As Long 'local copy
Private mvardTaxa As Double

Public Property Let lItemNFiscal(ByVal vData As Long)
    mvarlItemNFiscal = vData
End Property

Public Property Get lItemNFiscal() As Long
    lItemNFiscal = mvarlItemNFiscal
End Property

Public Property Let lItemPedCompra(ByVal vData As Long)
Attribute lItemPedCompra.VB_Description = "NumInt do ItemPedCompra"
    mvarlItemPedCompra = vData
End Property

Public Property Get lItemPedCompra() As Long
    lItemPedCompra = mvarlItemPedCompra
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dTaxa(ByVal vData As Double)
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property
