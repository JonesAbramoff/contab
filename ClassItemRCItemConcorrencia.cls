VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemRCItemConcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlItemReqCompra As Long 'local copy
Private mvardQuantidade As Double 'local copy
Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let lItemReqCompra(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lItemReqCompra = 5
    mvarlItemReqCompra = vData
End Property


Public Property Get lItemReqCompra() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lItemReqCompra
    lItemReqCompra = mvarlItemReqCompra
End Property



