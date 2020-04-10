VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemPVItemRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlItemPV As Long 'local copy
Private mvarlItemRC As Long 'local copy
Private mvardQuantidade As Long 'local copy
Public Property Let dQuantidade(ByVal vData As Long)
Attribute dQuantidade.VB_Description = "Na Unidade de Medida do Item de Pedido de Venda"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let lItemRC(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lItemRC = 5
    mvarlItemRC = vData
End Property


Public Property Get lItemRC() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lItemRC
    lItemRC = mvarlItemRC
End Property



Public Property Let lItemPV(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lItemPV = 5
    mvarlItemPV = vData
End Property


Public Property Get lItemPV() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lItemPV
    lItemPV = mvarlItemPV
End Property



