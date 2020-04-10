VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBaixaAntecipCredForn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlFornecedor As Long
Private mvariFilialFornecedor As Integer
Private mvariTipoBaixa As Integer
Private mvardtDataEmissaoDe As Date
Private mvardtDataEmissaoAte As Date
Private mvariCtaCorrenteSelecionada As Integer
Private mvarlNumCreditoDe As Long
Private mvarlNumCreditoAte As Long
Private mvarsTipoCredito As String
Private mvarcolAntecipCredFornecedor As New Collection
Private mvariCtaCorrenteTodas As Integer
Private mvariTipoCreditoTodos As Integer
Private mvarlNumIntBaixa As Long
Private mvardtDataBaixa As Date
Private mvariCtaCorrenteBaixa As Integer
Private mvardValorBaixa As Double
Private mvarsHistoricoBaixa As String
Private mvarlNumMovtoCCI As Long

Public Property Let lNumMovtoCCI(ByVal vData As Long)
    mvarlNumMovtoCCI = vData
End Property

Public Property Get lNumMovtoCCI() As Long
    lNumMovtoCCI = mvarlNumMovtoCCI
End Property

Public Property Get colAntecipCredFornecedor() As Collection
    Set colAntecipCredFornecedor = mvarcolAntecipCredFornecedor
End Property

Public Property Set colAntecipCredFornecedor(ByVal vData As Collection)
    Set mvarcolAntecipCredFornecedor = vData
End Property

Public Property Let sTipoCredito(ByVal vData As String)
    mvarsTipoCredito = vData
End Property

Public Property Get sTipoCredito() As String
    sTipoCredito = mvarsTipoCredito
End Property

Public Property Let lNumCreditoAte(ByVal vData As Long)
    mvarlNumCreditoAte = vData
End Property

Public Property Get lNumCreditoAte() As Long
    lNumCreditoAte = mvarlNumCreditoAte
End Property

Public Property Let lNumCreditoDe(ByVal vData As Long)
    mvarlNumCreditoDe = vData
End Property

Public Property Get lNumCreditoDe() As Long
    lNumCreditoDe = mvarlNumCreditoDe
End Property

Public Property Let iCtaCorrenteSelecionada(ByVal vData As Integer)
    mvariCtaCorrenteSelecionada = vData
End Property

Public Property Get iCtaCorrenteSelecionada() As Integer
    iCtaCorrenteSelecionada = mvariCtaCorrenteSelecionada
End Property

Public Property Let dtDataEmissaoAte(ByVal vData As Date)
    mvardtDataEmissaoAte = vData
End Property

Public Property Get dtDataEmissaoAte() As Date
    dtDataEmissaoAte = mvardtDataEmissaoAte
End Property

Public Property Let dtDataEmissaoDe(ByVal vData As Date)
    mvardtDataEmissaoDe = vData
End Property

Public Property Get dtDataEmissaoDe() As Date
    dtDataEmissaoDe = mvardtDataEmissaoDe
End Property

Public Property Let iTipoBaixa(ByVal vData As Integer)
    mvariTipoBaixa = vData
End Property

Public Property Get iTipoBaixa() As Integer
    iTipoBaixa = mvariTipoBaixa
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

Public Property Let iCtaCorrenteTodas(ByVal vData As Integer)
    mvariCtaCorrenteTodas = vData
End Property

Public Property Get iCtaCorrenteTodas() As Integer
    iCtaCorrenteTodas = mvariCtaCorrenteTodas
End Property

Public Property Let iTipoCreditoTodos(ByVal vData As Integer)
    mvariTipoCreditoTodos = vData
End Property

Public Property Get iTipoCreditoTodos() As Integer
    iTipoCreditoTodos = mvariTipoCreditoTodos
End Property

Public Property Let iCtaCorrenteBaixa(ByVal vData As Integer)
    mvariCtaCorrenteBaixa = vData
End Property

Public Property Get iCtaCorrenteBaixa() As Integer
    iCtaCorrenteBaixa = mvariCtaCorrenteBaixa
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dValorBaixa(ByVal vData As Double)
    mvardValorBaixa = vData
End Property

Public Property Get dValorBaixa() As Double
    dValorBaixa = mvardValorBaixa
End Property

Public Property Let sHistoricoBaixa(ByVal vData As String)
    mvarsHistoricoBaixa = vData
End Property

Public Property Get sHistoricoBaixa() As String
    sHistoricoBaixa = mvarsHistoricoBaixa
End Property

Public Property Let lNumIntBaixa(ByVal vData As Long)
    mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property

