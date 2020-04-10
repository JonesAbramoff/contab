VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPedidoCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvarsContato As String
Private mvardtDataEmissao As Date
Private mvardtData As Date
Private mvardtDataValidade As Date
Private mvariTipoFrete As Integer
Private mvariStatus As Integer
Private mvarcolItens As New Collection
Private mvarlCotacao As Long 'serve para pegar o Destino
Private mvariCondPagtoPrazo As Integer
Private mvardtDataBaixa As Date
Private mvarsCondPagto As String
Private mvarsNomeFornecedor As String
Private mvarsFilialFornecedor As String
Private mvarsNomeFilEmp As String
Private mvardtDataFixa As Date

Public Property Let dtDataFixa(ByVal vData As Date)
    mvardtDataFixa = vData
End Property

Public Property Get dtDataFixa() As Date
    dtDataFixa = mvardtDataFixa
End Property

Public Property Let sNomeFilEmp(ByVal vData As String)
    mvarsNomeFilEmp = vData
End Property

Public Property Get sNomeFilEmp() As String
    sNomeFilEmp = mvarsNomeFilEmp
End Property

Public Property Let sFilialFornecedor(ByVal vData As String)
    mvarsFilialFornecedor = vData
End Property

Public Property Get sFilialFornecedor() As String
    sFilialFornecedor = mvarsFilialFornecedor
End Property

Public Property Let sNomeFornecedor(ByVal vData As String)
    mvarsNomeFornecedor = vData
End Property

Public Property Get sNomeFornecedor() As String
    sNomeFornecedor = mvarsNomeFornecedor
End Property

Public Property Let sCondPagto(ByVal vData As String)
    mvarsCondPagto = vData
End Property

Public Property Get sCondPagto() As String
    sCondPagto = mvarsCondPagto
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let iCondPagtoPrazo(ByVal vData As Integer)
Attribute iCondPagtoPrazo.VB_Description = "Condição de pagamento a prazo"
    mvariCondPagtoPrazo = vData
End Property

Public Property Get iCondPagtoPrazo() As Integer
    iCondPagtoPrazo = mvariCondPagtoPrazo
End Property

Public Property Let lCotacao(ByVal vData As Long)
    mvarlCotacao = vData
End Property

Public Property Get lCotacao() As Long
    lCotacao = mvarlCotacao
End Property

Public Property Set colItens(ByVal vData As Object)
Attribute colItens.VB_Description = "Ítens de Pedido de Cotação."
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let iTipoFrete(ByVal vData As Integer)
    mvariTipoFrete = vData
End Property

Public Property Get iTipoFrete() As Integer
    iTipoFrete = mvariTipoFrete
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
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

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
