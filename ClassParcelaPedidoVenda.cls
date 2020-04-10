VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassParcelaPedidoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardValor As Double
Private mvardtDataVencimento As Date
Private mvariNumParcela As Integer
Private mvariDesconto1Codigo As Integer
Private mvardtDesconto1Ate As Date
Private mvardDesconto1Valor As Double
Private mvariDesconto2Codigo As Integer
Private mvardtDesconto2Ate As Date
Private mvardDesconto2Valor As Double
Private mvardtDesconto3Ate As Date
Private mvardDesconto3Valor As Double
Private mvariDesconto3Codigo As Integer
Private mvariTipoPagto As Integer
Private mvariCodConta As Integer
Private mvardtDataCredito As Date
Private mvardtDataEmissaoCheque As Date
Private mvariBancoCheque As Integer
Private mvarsAgenciaCheque As String
Private mvarsContaCorrenteCheque As String
Private mvarlNumeroCheque As Long
Private mvardtDataDepositoCheque As Date
Private mvariAdmMeioPagto As Integer
Private mvariParcelamento As Integer
Private mvarsNumeroCartao As String
Private mvardtValidadeCartao As Date
Private mvarsAprovacaoCartao As String
Private mvardtDataTransacaoCartao As Date

Public Property Let dtDataTransacaoCartao(ByVal vData As Date)
    mvardtDataTransacaoCartao = vData
End Property

Public Property Get dtDataTransacaoCartao() As Date
    dtDataTransacaoCartao = mvardtDataTransacaoCartao
End Property

Public Property Let sAprovacaoCartao(ByVal vData As String)
    mvarsAprovacaoCartao = vData
End Property

Public Property Get sAprovacaoCartao() As String
    sAprovacaoCartao = mvarsAprovacaoCartao
End Property

Public Property Let dtValidadeCartao(ByVal vData As Date)
    mvardtValidadeCartao = vData
End Property

Public Property Get dtValidadeCartao() As Date
    dtValidadeCartao = mvardtValidadeCartao
End Property

Public Property Let sNumeroCartao(ByVal vData As String)
    mvarsNumeroCartao = vData
End Property

Public Property Get sNumeroCartao() As String
    sNumeroCartao = mvarsNumeroCartao
End Property

Public Property Let iParcelamento(ByVal vData As Integer)
    mvariParcelamento = vData
End Property

Public Property Get iParcelamento() As Integer
    iParcelamento = mvariParcelamento
End Property

Public Property Let iAdmMeioPagto(ByVal vData As Integer)
    mvariAdmMeioPagto = vData
End Property

Public Property Get iAdmMeioPagto() As Integer
    iAdmMeioPagto = mvariAdmMeioPagto
End Property

Public Property Let dtDataDepositoCheque(ByVal vData As Date)
    mvardtDataDepositoCheque = vData
End Property

Public Property Get dtDataDepositoCheque() As Date
    dtDataDepositoCheque = mvardtDataDepositoCheque
End Property

Public Property Let lNumeroCheque(ByVal vData As Long)
    mvarlNumeroCheque = vData
End Property

Public Property Get lNumeroCheque() As Long
    lNumeroCheque = mvarlNumeroCheque
End Property

Public Property Let sContaCorrenteCheque(ByVal vData As String)
    mvarsContaCorrenteCheque = vData
End Property

Public Property Get sContaCorrenteCheque() As String
    sContaCorrenteCheque = mvarsContaCorrenteCheque
End Property

Public Property Let sAgenciaCheque(ByVal vData As String)
    mvarsAgenciaCheque = vData
End Property

Public Property Get sAgenciaCheque() As String
    sAgenciaCheque = mvarsAgenciaCheque
End Property

Public Property Let iBancoCheque(ByVal vData As Integer)
    mvariBancoCheque = vData
End Property

Public Property Get iBancoCheque() As Integer
    iBancoCheque = mvariBancoCheque
End Property

Public Property Let dtDataEmissaoCheque(ByVal vData As Date)
    mvardtDataEmissaoCheque = vData
End Property

Public Property Get dtDataEmissaoCheque() As Date
    dtDataEmissaoCheque = mvardtDataEmissaoCheque
End Property

Public Property Let dtDataCredito(ByVal vData As Date)
    mvardtDataCredito = vData
End Property

Public Property Get dtDataCredito() As Date
    dtDataCredito = mvardtDataCredito
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iTipoPagto(ByVal vData As Integer)
    mvariTipoPagto = vData
End Property

Public Property Get iTipoPagto() As Integer
    iTipoPagto = mvariTipoPagto
End Property

Public Property Let iDesconto3Codigo(ByVal vData As Integer)
    mvariDesconto3Codigo = vData
End Property

Public Property Get iDesconto3Codigo() As Integer
    iDesconto3Codigo = mvariDesconto3Codigo
End Property

Public Property Let dDesconto3Valor(ByVal vData As Double)
    mvardDesconto3Valor = vData
End Property

Public Property Get dDesconto3Valor() As Double
    dDesconto3Valor = mvardDesconto3Valor
End Property

Public Property Let dtDesconto3Ate(ByVal vData As Date)
    mvardtDesconto3Ate = vData
End Property

Public Property Get dtDesconto3Ate() As Date
    dtDesconto3Ate = mvardtDesconto3Ate
End Property

Public Property Let dDesconto2Valor(ByVal vData As Double)
    mvardDesconto2Valor = vData
End Property

Public Property Get dDesconto2Valor() As Double
    dDesconto2Valor = mvardDesconto2Valor
End Property

Public Property Let dtDesconto2Ate(ByVal vData As Date)
    mvardtDesconto2Ate = vData
End Property

Public Property Get dtDesconto2Ate() As Date
    dtDesconto2Ate = mvardtDesconto2Ate
End Property

Public Property Let iDesconto2Codigo(ByVal vData As Integer)
    mvariDesconto2Codigo = vData
End Property

Public Property Get iDesconto2Codigo() As Integer
    iDesconto2Codigo = mvariDesconto2Codigo
End Property

Public Property Let dDesconto1Valor(ByVal vData As Double)
    mvardDesconto1Valor = vData
End Property

Public Property Get dDesconto1Valor() As Double
    dDesconto1Valor = mvardDesconto1Valor
End Property

Public Property Let dtDesconto1Ate(ByVal vData As Date)
    mvardtDesconto1Ate = vData
End Property

Public Property Get dtDesconto1Ate() As Date
    dtDesconto1Ate = mvardtDesconto1Ate
End Property

Public Property Let iDesconto1Codigo(ByVal vData As Integer)
    mvariDesconto1Codigo = vData
End Property

Public Property Get iDesconto1Codigo() As Integer
    iDesconto1Codigo = mvariDesconto1Codigo
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property


Private Sub Class_Initialize()
        mvardtDataDepositoCheque = DATA_NULA
        mvardtDataEmissaoCheque = DATA_NULA
        mvardtValidadeCartao = DATA_NULA
End Sub

