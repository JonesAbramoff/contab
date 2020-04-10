VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoKitProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsProduto As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUnidadeMed As String 'local copy
Private mvariComposicao As Integer 'local copy
Private mvariControleEstoque As Integer 'local copy
Private mvarsSiglaUMEstoque As String 'local copy
Private mvariClasseUM As Integer 'local copy
Private mvarsVersao As String 'local copy
Private mvarsProdutoRaiz As String 'local copy
Private mvariSeq As Integer
Private mvardPercentualPerda As Double

Public Property Let dPercentualPerda(ByVal vData As Double)
    mvardPercentualPerda = vData
End Property

Public Property Get dPercentualPerda() As Double
    dPercentualPerda = mvardPercentualPerda
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sProdutoRaiz(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoRaiz = 5
    mvarsProdutoRaiz = vData
End Property


Public Property Get sProdutoRaiz() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoRaiz
    sProdutoRaiz = mvarsProdutoRaiz
End Property



Public Property Let sVersao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sVersao = 5
    mvarsVersao = vData
End Property


Public Property Get sVersao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sVersao
    sVersao = mvarsVersao
End Property



Public Property Let iClasseUM(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iClasseUM = 5
    mvariClasseUM = vData
End Property


Public Property Get iClasseUM() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iClasseUM
    iClasseUM = mvariClasseUM
End Property



Public Property Let sSiglaUMEstoque(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaUMEstoque = 5
    mvarsSiglaUMEstoque = vData
End Property


Public Property Get sSiglaUMEstoque() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaUMEstoque
    sSiglaUMEstoque = mvarsSiglaUMEstoque
End Property



Public Property Let iControleEstoque(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iControleEstoque = 5
    mvariControleEstoque = vData
End Property


Public Property Get iControleEstoque() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iControleEstoque
    iControleEstoque = mvariControleEstoque
End Property



Public Property Let iComposicao(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iComposicao = 5
    mvariComposicao = vData
End Property


Public Property Get iComposicao() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iComposicao
    iComposicao = mvariComposicao
End Property



Public Property Let sUnidadeMed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnidadeMed = 5
    mvarsUnidadeMed = vData
End Property


Public Property Get sUnidadeMed() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnidadeMed
    sUnidadeMed = mvarsUnidadeMed
End Property



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



