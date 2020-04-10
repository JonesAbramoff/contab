VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlCotacao As Long 'local copy
Private mvarsProduto As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUM As String 'local copy
Private mvarlFornecedor As Long 'local copy
Private mvariFilial As Integer 'local copy
Private mvarcolItemReqCompras As New Collection 'local copy
Private mvarcolItemFormPreco As New Collection 'local copy
Private mvariEscolhido As Integer 'local copy
Private mvariSelecionado As Integer 'local copy
Private mvarlLinkProv As Long

Public Property Let lLinkProv(ByVal vData As Long)
    mvarlLinkProv = vData
End Property

Public Property Get lLinkProv() As Long
    lLinkProv = mvarlLinkProv
End Property

Public Property Let iSelecionado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSelecionado = 5
    mvariSelecionado = vData
End Property


Public Property Get iSelecionado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSelecionado
    iSelecionado = mvariSelecionado
End Property



Public Property Let iEscolhido(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iEscolhido = 5
    mvariEscolhido = vData
End Property


Public Property Get iEscolhido() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iEscolhido
    iEscolhido = mvariEscolhido
End Property



Public Property Set colItemReqCompras(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colItemReqCompras = Form1
    Set mvarcolItemReqCompras = vData
End Property


Public Property Get colItemReqCompras() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colItemReqCompras
    Set colItemReqCompras = mvarcolItemReqCompras
End Property

Public Property Set colItemFormPreco(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colItemReqCompras = Form1
    Set mvarcolItemFormPreco = vData
End Property


Public Property Get colItemFormPreco() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colItemReqCompras
    Set colItemFormPreco = mvarcolItemFormPreco
End Property



Public Property Let iFilial(ByVal vData As Integer)
Attribute iFilial.VB_Description = "Filial do Fornecedor determinado."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilial = 5
    mvariFilial = vData
End Property


Public Property Get iFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilial
    iFilial = mvariFilial
End Property



Public Property Let lFornecedor(ByVal vData As Long)
Attribute lFornecedor.VB_Description = "Fornecedor determinado."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornecedor = 5
    mvarlFornecedor = vData
End Property


Public Property Get lFornecedor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornecedor
    lFornecedor = mvarlFornecedor
End Property



Public Property Let sUM(ByVal vData As String)
Attribute sUM.VB_Description = "Unidade de Medida da quantidade."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUM = 5
    mvarsUM = vData
End Property


Public Property Get sUM() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUM
    sUM = mvarsUM
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
Attribute sProduto.VB_Description = "Código do Produto."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property



Public Property Let lCotacao(ByVal vData As Long)
Attribute lCotacao.VB_Description = "Número interno da cotação."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCotacao = 5
    mvarlCotacao = vData
End Property


Public Property Get lCotacao() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCotacao
    lCotacao = mvarlCotacao
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property



