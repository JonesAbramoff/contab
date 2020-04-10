VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPagAntecipBaixado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntBaixa As Long
Private mvarlNumIntPag As Long
Private mvardValor As Double
Private mvardValorBaixado As Double
Private mvarsNomeReduzido As String
Private mvarlFornecedor As Long
Private mvariFilialFornecedor As Integer
Private mvardtDataMovimento As Date
Private mvariCodConta As Integer
Private mvariTipoMeioPagto As Integer
Private mvarlNumMovto As Long
Private mvarsContaCorrenteNome As String

Public Property Let sContaCorrenteNome(ByVal vData As String)
    mvarsContaCorrenteNome = vData
End Property

Public Property Get sContaCorrenteNome() As String
    sContaCorrenteNome = mvarsContaCorrenteNome
End Property

Public Property Let lNumMovto(ByVal vData As Long)
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let dtDataMovimento(ByVal vData As Date)
    mvardtDataMovimento = vData
End Property

Public Property Get dtDataMovimento() As Date
    dtDataMovimento = mvardtDataMovimento
End Property

Public Property Let iFilialFornecedor(ByVal vData As Integer)
    mvariFilialFornecedor = vData
End Property

Public Property Get iFilialFornecedor() As Integer
    iFilialFornecedor = mvariFilialFornecedor
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let lNumIntPag(ByVal vData As Long)
    mvarlNumIntPag = vData
End Property

Public Property Get lNumIntPag() As Long
    lNumIntPag = mvarlNumIntPag
End Property

Public Property Let lNumIntBaixa(ByVal vData As Long)
    mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property



