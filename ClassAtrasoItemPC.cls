VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAtrasoItemPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlAtraso As Long
Private mvardValorEntregue As Double
Public Property Let dValorEntregue(ByVal vData As Double)
Attribute dValorEntregue.VB_Description = "Valor entregue do Item de Pedido Compras."
    mvardValorEntregue = vData
End Property


Public Property Get dValorEntregue() As Double
    dValorEntregue = mvardValorEntregue
End Property



Public Property Let lAtraso(ByVal vData As Long)
Attribute lAtraso.VB_Description = "Dias de atraso."
    mvarlAtraso = vData
End Property


Public Property Get lAtraso() As Long
    lAtraso = mvarlAtraso
End Property



