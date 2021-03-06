VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoProdutoForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarobjCotacaoProduto As ClassCotacaoProduto 'local copy
Private mvarobjFornecedorProdutoFF As ClassFornecedorProdutoFF 'local copy
Public Property Set objFornecedorProdutoFF(ByVal vData As ClassFornecedorProdutoFF)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objFornedorProdutoFF = Form1
    Set mvarobjFornecedorProdutoFF = vData
End Property


Public Property Get objFornecedorProdutoFF() As ClassFornecedorProdutoFF
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objFornedorProdutoFF
    Set objFornecedorProdutoFF = mvarobjFornecedorProdutoFF
End Property



Public Property Set objCotacaoProduto(ByVal vData As ClassCotacaoProduto)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objCotacaoProduto = Form1
    Set mvarobjCotacaoProduto = vData
End Property


Public Property Get objCotacaoProduto() As ClassCotacaoProduto
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objCotacaoProduto
    Set objCotacaoProduto = mvarobjCotacaoProduto
End Property



