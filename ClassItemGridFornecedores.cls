VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemGridFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsProduto As String 'local copy
Private mvarsDescProduto As String 'local copy
Private mvarsFornecedor As String 'local copy
Private mvarsFilialForn As String 'local copy
Private mvarsUltimaCotacao As String 'local copy
Private mvarsTipoFrete As String 'local copy
Private mvarsDataUltimaCotacao As String 'local copy
Private mvarsDataUltimaCompra As String 'local copy
Private mvarsPrazoEntrega As String 'local copy
Private mvarsUMCompra As String 'local copy
Private mvarsQuantPedida As String 'local copy
Private mvarsQuantRecebida As String 'local copy
Private mvarsCondicaoPagto As String 'local copy
Private mvarsSaldoTitulos As String 'local copy
Private mvarsObservacao As String 'local copy
Private mvarlNumIntFornecedorProdutoFF As Long 'local copy
Private mvarsExclusivo As String 'local copy
Private mvarsEscolhido As String 'local copy
Private mvariSelecionado As Integer 'local copy
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



Public Property Let sEscolhido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEscolhido = 5
    mvarsEscolhido = vData
End Property


Public Property Get sEscolhido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEscolhido
    sEscolhido = mvarsEscolhido
End Property



Public Property Let sExclusivo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sExclusivo = 5
    mvarsExclusivo = vData
End Property


Public Property Get sExclusivo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sExclusivo
    sExclusivo = mvarsExclusivo
End Property



Public Property Let lNumIntFornecedorProdutoFF(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntFornecedorProdutoFF = 5
    mvarlNumIntFornecedorProdutoFF = vData
End Property


Public Property Get lNumIntFornecedorProdutoFF() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntFornecedorProdutoFF
    lNumIntFornecedorProdutoFF = mvarlNumIntFornecedorProdutoFF
End Property



Public Property Let sObservacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property


Public Property Get sObservacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property



Public Property Let sSaldoTitulos(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSaldoTitulos = 5
    mvarsSaldoTitulos = vData
End Property


Public Property Get sSaldoTitulos() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSaldoTitulos
    sSaldoTitulos = mvarsSaldoTitulos
End Property



Public Property Let sCondicaoPagto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCondicaoPagto = 5
    mvarsCondicaoPagto = vData
End Property


Public Property Get sCondicaoPagto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCondicaoPagto
    sCondicaoPagto = mvarsCondicaoPagto
End Property



Public Property Let sQuantRecebida(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sQuantRecebida = 5
    mvarsQuantRecebida = vData
End Property


Public Property Get sQuantRecebida() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sQuantRecebida
    sQuantRecebida = mvarsQuantRecebida
End Property



Public Property Let sQuantPedida(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sQuantPedida = 5
    mvarsQuantPedida = vData
End Property


Public Property Get sQuantPedida() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sQuantPedida
    sQuantPedida = mvarsQuantPedida
End Property



Public Property Let sUMCompra(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMCompra = 5
    mvarsUMCompra = vData
End Property


Public Property Get sUMCompra() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMCompra
    sUMCompra = mvarsUMCompra
End Property



Public Property Let sPrazoEntrega(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sPrazoEntrega = 5
    mvarsPrazoEntrega = vData
End Property


Public Property Get sPrazoEntrega() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sPrazoEntrega
    sPrazoEntrega = mvarsPrazoEntrega
End Property



Public Property Let sDataUltimaCompra(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDataUltimaCompra = 5
    mvarsDataUltimaCompra = vData
End Property


Public Property Get sDataUltimaCompra() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDataUltimaCompra
    sDataUltimaCompra = mvarsDataUltimaCompra
End Property



Public Property Let sDataUltimaCotacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDataUltimaCotacao = 5
    mvarsDataUltimaCotacao = vData
End Property


Public Property Get sDataUltimaCotacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDataUltimaCotacao
    sDataUltimaCotacao = mvarsDataUltimaCotacao
End Property



Public Property Let sTipoFrete(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTipoFrete = 5
    mvarsTipoFrete = vData
End Property


Public Property Get sTipoFrete() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTipoFrete
    sTipoFrete = mvarsTipoFrete
End Property



Public Property Let sUltimaCotacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUltimaCotacao = 5
    mvarsUltimaCotacao = vData
End Property


Public Property Get sUltimaCotacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUltimaCotacao
    sUltimaCotacao = mvarsUltimaCotacao
End Property



Public Property Let sFilialForn(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFilialForn = 5
    mvarsFilialForn = vData
End Property


Public Property Get sFilialForn() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFilialForn
    sFilialForn = mvarsFilialForn
End Property



Public Property Let sFornecedor(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFornecedor = 5
    mvarsFornecedor = vData
End Property


Public Property Get sFornecedor() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFornecedor
    sFornecedor = mvarsFornecedor
End Property



Public Property Let sDescProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescProduto = 5
    mvarsDescProduto = vData
End Property


Public Property Get sDescProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescProduto
    sDescProduto = mvarsDescProduto
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



