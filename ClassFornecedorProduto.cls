VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFornecedorProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlFornecedor As Long
Private mvarsProduto As String
Private mvarsProdutoFornecedor As String
Private mvardLoteMinimo As Double
Private mvardLoteEconomico As Double
Private mvardQuantPedAbertos As Double
Private mvariTempoMedio As Integer
Private mvardQuantPedida As Double
Private mvardQuantRecebida As Double
Private mvardValor As Double
Private mvardtDataPedido As Date
Private mvardtDataReceb As Date

Public Property Let dtDataReceb(ByVal vData As Date)
    mvardtDataReceb = vData
End Property

Public Property Get dtDataReceb() As Date
    dtDataReceb = mvardtDataReceb
End Property

Public Property Let dtDataPedido(ByVal vData As Date)
    mvardtDataPedido = vData
End Property

Public Property Get dtDataPedido() As Date
    dtDataPedido = mvardtDataPedido
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dQuantRecebida(ByVal vData As Double)
    mvardQuantRecebida = vData
End Property

Public Property Get dQuantRecebida() As Double
    dQuantRecebida = mvardQuantRecebida
End Property

Public Property Let dQuantPedida(ByVal vData As Double)
    mvardQuantPedida = vData
End Property

Public Property Get dQuantPedida() As Double
    dQuantPedida = mvardQuantPedida
End Property

Public Property Let iTempoMedio(ByVal vData As Integer)
    mvariTempoMedio = vData
End Property

Public Property Get iTempoMedio() As Integer
    iTempoMedio = mvariTempoMedio
End Property

Public Property Let dQuantPedAbertos(ByVal vData As Double)
    mvardQuantPedAbertos = vData
End Property

Public Property Get dQuantPedAbertos() As Double
    dQuantPedAbertos = mvardQuantPedAbertos
End Property

Public Property Let dLoteEconomico(ByVal vData As Double)
    mvardLoteEconomico = vData
End Property

Public Property Get dLoteEconomico() As Double
    dLoteEconomico = mvardLoteEconomico
End Property

Public Property Let dLoteMinimo(ByVal vData As Double)
    mvardLoteMinimo = vData
End Property

Public Property Get dLoteMinimo() As Double
    dLoteMinimo = mvardLoteMinimo
End Property

Public Property Let sProdutoFornecedor(ByVal vData As String)
    mvarsProdutoFornecedor = vData
End Property

Public Property Get sProdutoFornecedor() As String
    sProdutoFornecedor = mvarsProdutoFornecedor
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property



