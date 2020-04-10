VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassSubstProdutoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsProdutoSubstituto As String 'local copy
Private mvarsProduto As String 'local copy
Private mvarcolOutrosProdutosNF As New Collection 'local copy
Public Property Set colOutrosProdutosNF(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colOutrosProdutosNF = Form1
    Set mvarcolOutrosProdutosNF = vData
End Property


Public Property Get colOutrosProdutosNF() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colOutrosProdutosNF
    Set colOutrosProdutosNF = mvarcolOutrosProdutosNF
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property



Public Property Let sProdutoSubstituto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoSubstituto = 5
    mvarsProdutoSubstituto = vData
End Property


Public Property Get sProdutoSubstituto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoSubstituto
    sProdutoSubstituto = mvarsProdutoSubstituto
End Property



