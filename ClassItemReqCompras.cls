VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemReqCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarsProduto As String
Private mvarsDescProduto As String
Private mvariStatus As Integer
Private mvardQuantidade As Double
Private mvardQuantPedida As Double
Private mvardQuantRecebida As Double
Private mvardQuantCancelada As Double
Private mvarsUM As String
Private mvarsCcl As String
Private mvariAlmoxarifado As Integer
Private mvarsContaContabil As String
Private mvariCreditaICMS As Integer
Private mvariCreditaIPI As Integer
Private mvarlObservacao As Long
Private mvarsObservacao As String
Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvariExclusivo As Integer
Private mvardQuantNaConcorrencia As Double
Private mvardQuantNoPedido As Double
Private mvardQuantNoPedidoRecebida As Double
Private mvardQuantNaCotacao As Double
Private mvarlReqCompra As Long
Private mvardQuantCotar As Double
Private mvariTipoTributacao As Integer
Private mvardQuantComprar As Double
Private mvariSelecionado As Integer
Private mvariItem As Integer
Private mvariTipoProduto As Integer
Private mvardQuantCotarAnterior As Double
Private mvarcolItemPV As New Collection
Private mvariTempoRessup As Integer

Public Property Let iTempoRessup(ByVal vData As Integer)
    mvariTempoRessup = vData
End Property

Public Property Get iTempoRessup() As Integer
    iTempoRessup = mvariTempoRessup
End Property

Public Property Set colItemPV(ByVal vData As Collection)
    Set mvarcolItemPV = vData
End Property

Public Property Get colItemPV() As Collection
    Set colItemPV = mvarcolItemPV
End Property

Public Property Let dQuantCotarAnterior(ByVal vData As Double)
    mvardQuantCotarAnterior = vData
End Property

Public Property Get dQuantCotarAnterior() As Double
    dQuantCotarAnterior = mvardQuantCotarAnterior
End Property



Public Property Let iTipoProduto(ByVal vData As Integer)
    mvariTipoProduto = vData
End Property


Public Property Get iTipoProduto() As Integer
    iTipoProduto = mvariTipoProduto
End Property



Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property


Public Property Get iItem() As Integer
    iItem = mvariItem
End Property


Public Property Let iSelecionado(ByVal vData As Integer)
    mvariSelecionado = vData
End Property


Public Property Get iSelecionado() As Integer
    iSelecionado = mvariSelecionado
End Property



Public Property Let dQuantComprar(ByVal vData As Double)
    mvardQuantComprar = vData
End Property


Public Property Get dQuantComprar() As Double
    dQuantComprar = mvardQuantComprar
End Property



Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property


Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property


Public Property Let dQuantCotar(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantCotar = 5
    mvardQuantCotar = vData
End Property


Public Property Get dQuantCotar() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantCotar
    dQuantCotar = mvardQuantCotar
End Property




Public Property Let lReqCompra(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lReqCompra = 5
    mvarlReqCompra = vData
End Property


Public Property Get lReqCompra() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lReqCompra
    lReqCompra = mvarlReqCompra
End Property



Public Property Let dQuantNaCotacao(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantNaCotacao = 5
    mvardQuantNaCotacao = vData
End Property


Public Property Get dQuantNaCotacao() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantNaCotacao
    dQuantNaCotacao = mvardQuantNaCotacao
End Property



Public Property Let dQuantNoPedidoRecebida(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantNoPedidoRecebida = 5
    mvardQuantNoPedidoRecebida = vData
End Property


Public Property Get dQuantNoPedidoRecebida() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantNoPedidoRecebida
    dQuantNoPedidoRecebida = mvardQuantNoPedidoRecebida
End Property



Public Property Let dQuantNoPedido(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantNoPedido = 5
    mvardQuantNoPedido = vData
End Property


Public Property Get dQuantNoPedido() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantNoPedido
    dQuantNoPedido = mvardQuantNoPedido
End Property



Public Property Let dQuantNaConcorrencia(ByVal vData As Double)
Attribute dQuantNaConcorrencia.VB_Description = "Quantidade na concorrência."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantNaConcorrencia = 5
    mvardQuantNaConcorrencia = vData
End Property


Public Property Get dQuantNaConcorrencia() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantNaConcorrencia
    dQuantNaConcorrencia = mvardQuantNaConcorrencia
End Property




Public Property Let iExclusivo(ByVal vData As Integer)
    mvariExclusivo = vData
End Property

Public Property Get iExclusivo() As Integer
    iExclusivo = mvariExclusivo
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let sObservacao(ByVal vData As String)

'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property


Public Property Get sObservacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property



Public Property Let lObservacao(ByVal vData As Long)

'Syntax: X.lObservacao = 5
    mvarlObservacao = vData
End Property


Public Property Get lObservacao() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lObservacao
    lObservacao = mvarlObservacao
End Property



Public Property Let iCreditaIPI(ByVal vData As Integer)

'Syntax: X.iCreditaIPI = 5
    mvariCreditaIPI = vData
End Property


Public Property Get iCreditaIPI() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCreditaIPI
    iCreditaIPI = mvariCreditaIPI
End Property



Public Property Let iCreditaICMS(ByVal vData As Integer)

'Syntax: X.iCreditaICMS = 5
    mvariCreditaICMS = vData
End Property


Public Property Get iCreditaICMS() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCreditaICMS
    iCreditaICMS = mvariCreditaICMS
End Property



Public Property Let sContaContabil(ByVal vData As String)

'Syntax: X.sContaContabil = 5
    mvarsContaContabil = vData
End Property


Public Property Get sContaContabil() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaContabil
    sContaContabil = mvarsContaContabil
End Property



Public Property Let iAlmoxarifado(ByVal vData As Integer)

'Syntax: X.iAlmoxarifado = 5
    mvariAlmoxarifado = vData
End Property


Public Property Get iAlmoxarifado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAlmoxarifado
    iAlmoxarifado = mvariAlmoxarifado
End Property



Public Property Let sCcl(ByVal vData As String)

'Syntax: X.sCcl = 5
    mvarsCcl = vData
End Property


Public Property Get sCcl() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCcl
    sCcl = mvarsCcl
End Property



Public Property Let sUM(ByVal vData As String)

'Syntax: X.sUM = 5
    mvarsUM = vData
End Property


Public Property Get sUM() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUM
    sUM = mvarsUM
End Property



Public Property Let dQuantCancelada(ByVal vData As Double)

'Syntax: X.dQuantCancelada = 5
    mvardQuantCancelada = vData
End Property


Public Property Get dQuantCancelada() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantCancelada
    dQuantCancelada = mvardQuantCancelada
End Property



Public Property Let dQuantRecebida(ByVal vData As Double)

'Syntax: X.dQuantRecebida = 5
    mvardQuantRecebida = vData
End Property


Public Property Get dQuantRecebida() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantRecebida
    dQuantRecebida = mvardQuantRecebida
End Property



Public Property Let dQuantPedida(ByVal vData As Double)

'Syntax: X.dQuantPedida = 5
    mvardQuantPedida = vData
End Property


Public Property Get dQuantPedida() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantPedida
    dQuantPedida = mvardQuantPedida
End Property



Public Property Let dQuantidade(ByVal vData As Double)

'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let iStatus(ByVal vData As Integer)

'Syntax: X.iStatus = 5
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iStatus
    iStatus = mvariStatus
End Property



Public Property Let sDescProduto(ByVal vData As String)

'Syntax: X.sDescProduto = 5
    mvarsDescProduto = vData
End Property


Public Property Get sDescProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescProduto
    sDescProduto = mvarsDescProduto
End Property



Public Property Let sProduto(ByVal vData As String)

'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property


Public Property Let lNumIntDoc(ByVal vData As Long)

'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property



