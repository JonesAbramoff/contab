VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoPtoPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodProdutoDe As String 'local copy
Private mvarsCodProdutoAte As String 'local copy
Private mvarcolTipoProduto As Collection 'local copy
Private mvarcolProdutoPtoPedido As Collection 'local copy
Private mvariExibePPZerado As Integer

Public Property Let iExibePPZerado(ByVal vData As Integer)
    mvariExibePPZerado = vData
End Property

Public Property Get iExibePPZerado() As Integer
    iExibePPZerado = mvariExibePPZerado
End Property

Public Property Set colProdutoPtoPedido(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colProdutoPtoPedido = Form1
    Set mvarcolProdutoPtoPedido = vData
End Property


Public Property Get colProdutoPtoPedido() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colProdutoPtoPedido
    Set colProdutoPtoPedido = mvarcolProdutoPtoPedido
End Property



Public Property Set colTipoProduto(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colTipoProduto = Form1
    Set mvarcolTipoProduto = vData
End Property


Public Property Get colTipoProduto() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colTipoProduto
    Set colTipoProduto = mvarcolTipoProduto
End Property



Public Property Let sCodProdutoAte(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodProdutoAte = 5
    mvarsCodProdutoAte = vData
End Property


Public Property Get sCodProdutoAte() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodProdutoAte
    sCodProdutoAte = mvarsCodProdutoAte
End Property



Public Property Let sCodProdutoDe(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodProdutoDe = 5
    mvarsCodProdutoDe = vData
End Property


Public Property Get sCodProdutoDe() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodProdutoDe
    sCodProdutoDe = mvarsCodProdutoDe
End Property



