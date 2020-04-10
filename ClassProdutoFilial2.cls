VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Private mvariFilialEmpresa As Integer
Private mvarsProduto As String
Private mvariAlmoxarifado As Integer
Private mvarlFornecedor As Long
Private mvariFilialForn As Integer
Private mvariVisibilidadeAlmoxarifados As Integer
Private mvardEstoqueSeguranca As Double
Private mvariESCalculado As Integer
Private mvardEstoqueMaximo As Double
Private mvariTemPtoPedido As Integer
Private mvardPontoPedido As Double
Private mvariPPCalculado As Integer
Private mvarsClasseABC As String
Private mvardLoteEconomico As Double
Private mvariIntRessup As Integer
Private mvariTempoRessup As Integer
Private mvariTRCalculado As Integer
Private mvardTempoRessupMax As Double
Private mvardConsumoMedio As Double
Private mvariCMCalculado As Integer
Private mvardConsumoMedioMax As Double
Private mvariMesesConsumoMedio As Integer
Private mvariControleEstoque As Integer
Private mvariTabelaPreco As Integer
Private mvardQuantPedida As Double
Private mvarsSituacaoTribECF As String
Private mvarsICMSAliquota As String
Private mvarsDescricao As String
Private mvarsCategoriaItem As String
Private mvardLoteMinimo As Double
Private mvariProdNaFilial As Integer
Private mvardDescontoItem As Double
Private mvardDescontoValor As Double

Public Property Let dDescontoValor(ByVal vData As Double)
    mvardDescontoValor = vData
End Property

Public Property Get dDescontoValor() As Double
    dDescontoValor = mvardDescontoValor
End Property

Public Property Let dDescontoItem(ByVal vData As Double)
    mvardDescontoItem = vData
End Property

Public Property Get dDescontoItem() As Double
    dDescontoItem = mvardDescontoItem
End Property

Public Property Let iProdNaFilial(ByVal vData As Integer)
    mvariProdNaFilial = vData
End Property

Public Property Get iProdNaFilial() As Integer
    iProdNaFilial = mvariProdNaFilial
End Property

Public Property Let sCategoriaItem(ByVal vData As String)
    mvarsCategoriaItem = vData
End Property

Public Property Get sCategoriaItem() As String
    sCategoriaItem = mvarsCategoriaItem
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sICMSAliquota(ByVal vData As String)
    mvarsICMSAliquota = vData
End Property

Public Property Get sICMSAliquota() As String
    sICMSAliquota = mvarsICMSAliquota
End Property

Public Property Let sSituacaoTribECF(ByVal vData As String)
    mvarsSituacaoTribECF = vData
End Property

Public Property Get sSituacaoTribECF() As String
    sSituacaoTribECF = mvarsSituacaoTribECF
End Property

Public Property Let dQuantPedida(ByVal vData As Double)
    mvardQuantPedida = vData
End Property

Public Property Get dQuantPedida() As Double
    dQuantPedida = mvardQuantPedida
End Property

Public Property Let iTabelaPreco(ByVal vData As Double)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Double
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let iControleEstoque(ByVal vData As Integer)
    mvariControleEstoque = vData
End Property

Public Property Get iControleEstoque() As Integer
    iControleEstoque = mvariControleEstoque
End Property

Public Property Let iMesesConsumoMedio(ByVal vData As Integer)
    mvariMesesConsumoMedio = vData
End Property

Public Property Get iMesesConsumoMedio() As Integer
    iMesesConsumoMedio = mvariMesesConsumoMedio
End Property

Public Property Let dConsumoMedioMax(ByVal vData As Double)
    mvardConsumoMedioMax = vData
End Property

Public Property Get dConsumoMedioMax() As Double
    dConsumoMedioMax = mvardConsumoMedioMax
End Property

Public Property Let iCMCalculado(ByVal vData As Integer)
    mvariCMCalculado = vData
End Property

Public Property Get iCMCalculado() As Integer
    iCMCalculado = mvariCMCalculado
End Property

Public Property Let dConsumoMedio(ByVal vData As Double)
    mvardConsumoMedio = vData
End Property

Public Property Get dConsumoMedio() As Double
    dConsumoMedio = mvardConsumoMedio
End Property

Public Property Let dTempoRessupMax(ByVal vData As Double)
    mvardTempoRessupMax = vData
End Property

Public Property Get dTempoRessupMax() As Double
    dTempoRessupMax = mvardTempoRessupMax
End Property

Public Property Let iTRCalculado(ByVal vData As Integer)
    mvariTRCalculado = vData
End Property

Public Property Get iTRCalculado() As Integer
    iTRCalculado = mvariTRCalculado
End Property

Public Property Let iTempoRessup(ByVal vData As Integer)
    mvariTempoRessup = vData
End Property

Public Property Get iTempoRessup() As Integer
    iTempoRessup = mvariTempoRessup
End Property

Public Property Let iIntRessup(ByVal vData As Integer)
    mvariIntRessup = vData
End Property

Public Property Get iIntRessup() As Integer
    iIntRessup = mvariIntRessup
End Property

Public Property Let dLoteEconomico(ByVal vData As Double)
    mvardLoteEconomico = vData
End Property

Public Property Get dLoteEconomico() As Double
    dLoteEconomico = mvardLoteEconomico
End Property

Public Property Let sClasseABC(ByVal vData As String)
    mvarsClasseABC = vData
End Property

Public Property Get sClasseABC() As String
    sClasseABC = mvarsClasseABC
End Property

Public Property Let iPPCalculado(ByVal vData As Integer)
    mvariPPCalculado = vData
End Property

Public Property Get iPPCalculado() As Integer
    iPPCalculado = mvariPPCalculado
End Property

Public Property Let dPontoPedido(ByVal vData As Double)
    mvardPontoPedido = vData
End Property

Public Property Get dPontoPedido() As Double
    dPontoPedido = mvardPontoPedido
End Property

Public Property Let iTemPtoPedido(ByVal vData As Integer)
    mvariTemPtoPedido = vData
End Property

Public Property Get iTemPtoPedido() As Integer
    iTemPtoPedido = mvariTemPtoPedido
End Property

Public Property Let dEstoqueMaximo(ByVal vData As Double)
    mvardEstoqueMaximo = vData
End Property

Public Property Get dEstoqueMaximo() As Double
    dEstoqueMaximo = mvardEstoqueMaximo
End Property

Public Property Let iESCalculado(ByVal vData As Integer)
    mvariESCalculado = vData
End Property

Public Property Get iESCalculado() As Integer
    iESCalculado = mvariESCalculado
End Property

Public Property Let dEstoqueSeguranca(ByVal vData As Double)
    mvardEstoqueSeguranca = vData
End Property

Public Property Get dEstoqueSeguranca() As Double
    dEstoqueSeguranca = mvardEstoqueSeguranca
End Property

Public Property Let iVisibilidadeAlmoxarifados(ByVal vData As Integer)
    mvariVisibilidadeAlmoxarifados = vData
End Property

Public Property Get iVisibilidadeAlmoxarifados() As Integer
    iVisibilidadeAlmoxarifados = mvariVisibilidadeAlmoxarifados
End Property

Public Property Let iFilialForn(ByVal vData As Integer)
    mvariFilialForn = vData
End Property

Public Property Get iFilialForn() As Integer
    iFilialForn = mvariFilialForn
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iAlmoxarifado(ByVal vData As Integer)
    mvariAlmoxarifado = vData
End Property

Public Property Get iAlmoxarifado() As Integer
    iAlmoxarifado = mvariAlmoxarifado
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dLoteMinimo(ByVal vData As Double)
    mvardLoteMinimo = vData
End Property

Public Property Get dLoteMinimo() As Double
    dLoteMinimo = mvardLoteMinimo
End Property

