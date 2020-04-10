VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTipoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvardLimiteCredito As Double
Private mvariCodMensagem As Integer
Private mvardDesconto As Double
Private mvarsObservacao As String
Private mvariTabelaPreco As Integer
Private mvarsContaContabil As String
Private mvariVendedor As Integer
Private mvardComissaoVendas As Double
Private mvariRegiao As Integer
Private mvariFreqVisitas As Integer
Private mvariCodCobrador As Integer
Private mvariCodTransportadora As Integer
Private mvariPadraoCobranca As Integer
Private mvariCondicaoPagto As Integer
Private mvarcolCategoriaItem As New Collection

Public Property Set colCategoriaItem(ByVal vData As Object)
    Set mvarcolCategoriaItem = vData
End Property

Public Property Get colCategoriaItem() As Collection
    Set colCategoriaItem = mvarcolCategoriaItem
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let iPadraoCobranca(ByVal vData As Integer)
    mvariPadraoCobranca = vData
End Property

Public Property Get iPadraoCobranca() As Integer
    iPadraoCobranca = mvariPadraoCobranca
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
End Property

Public Property Let iCodCobrador(ByVal vData As Integer)
    mvariCodCobrador = vData
End Property

Public Property Get iCodCobrador() As Integer
    iCodCobrador = mvariCodCobrador
End Property

Public Property Let iFreqVisitas(ByVal vData As Integer)
    mvariFreqVisitas = vData
End Property

Public Property Get iFreqVisitas() As Integer
    iFreqVisitas = mvariFreqVisitas
End Property

Public Property Let iRegiao(ByVal vData As Integer)
    mvariRegiao = vData
End Property

Public Property Get iRegiao() As Integer
    iRegiao = mvariRegiao
End Property

Public Property Let dComissaoVendas(ByVal vData As Double)
    mvardComissaoVendas = vData
End Property

Public Property Get dComissaoVendas() As Double
    dComissaoVendas = mvardComissaoVendas
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dDesconto(ByVal vData As Double)
    mvardDesconto = vData
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let iCodMensagem(ByVal vData As Integer)
    mvariCodMensagem = vData
End Property

Public Property Get iCodMensagem() As Integer
    iCodMensagem = mvariCodMensagem
End Property

Public Property Let dLimiteCredito(ByVal vData As Double)
    mvardLimiteCredito = vData
End Property

Public Property Get dLimiteCredito() As Double
    dLimiteCredito = mvardLimiteCredito
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



