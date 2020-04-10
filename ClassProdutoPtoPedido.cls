VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoPtoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarobjProduto As ClassProduto 'local copy
Private mvarobjProdutoFilial As ClassProdutoFilial 'local copy
Private mvarobjEstoqueProduto As ClassEstoqueProduto 'local copy
Public Property Set objEstoqueProduto(ByVal vData As ClassEstoqueProduto)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objEstoqueProduto = Form1
    Set mvarobjEstoqueProduto = vData
End Property


Public Property Get objEstoqueProduto() As ClassEstoqueProduto
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objEstoqueProduto
    Set objEstoqueProduto = mvarobjEstoqueProduto
End Property



Public Property Set objProdutoFilial(ByVal vData As ClassProdutoFilial)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objProdutoFilial = Form1
    Set mvarobjProdutoFilial = vData
End Property


Public Property Get objProdutoFilial() As ClassProdutoFilial
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objProdutoFilial
    Set objProdutoFilial = mvarobjProdutoFilial
End Property



Public Property Set objProduto(ByVal vData As ClassProduto)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objProduto = Form1
    Set mvarobjProduto = vData
End Property


Public Property Get objProduto() As ClassProduto
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objProduto
    Set objProduto = mvarobjProduto
End Property



