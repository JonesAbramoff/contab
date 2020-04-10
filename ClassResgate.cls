VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassResgate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigoAplicacao As Long
Private mvariSeqResgate As Integer
Private mvarlNumMovto As Long
Private mvardValorResgatado As Double
Private mvardRendimentos As Double
Private mvardValorIRRF As Double
Private mvardDescontos As Double
Private mvardSaldoAnterior As Double
Private mvardValorCreditado As Double
Private mvardtDataResgatePrevista As Date
Private mvardValorResgatePrevisto As Double
Private mvardTaxaPrevista As Double
Private mvarsHistorico As String
Private mvariTipoMeioPagto As Integer
Private mvarlNumero As Long
Private mvarsNumRefExterna As String
Private mvardtDataMovimento As Date
Private mvariStatus As Integer

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dtDataMovimento(ByVal vData As Date)
    mvardtDataMovimento = vData
End Property

Public Property Get dtDataMovimento() As Date
    dtDataMovimento = mvardtDataMovimento
End Property

Public Property Let sNumRefExterna(ByVal vData As String)
    mvarsNumRefExterna = vData
End Property

Public Property Get sNumRefExterna() As String
    sNumRefExterna = mvarsNumRefExterna
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let dTaxaPrevista(ByVal vData As Double)
    mvardTaxaPrevista = vData
End Property

Public Property Get dTaxaPrevista() As Double
    dTaxaPrevista = mvardTaxaPrevista
End Property

Public Property Let dValorResgatePrevisto(ByVal vData As Double)
    mvardValorResgatePrevisto = vData
End Property

Public Property Get dValorResgatePrevisto() As Double
    dValorResgatePrevisto = mvardValorResgatePrevisto
End Property

Public Property Let dtDataResgatePrevista(ByVal vData As Date)
    mvardtDataResgatePrevista = vData
End Property

Public Property Get dtDataResgatePrevista() As Date
    dtDataResgatePrevista = mvardtDataResgatePrevista
End Property

Public Property Let dValorCreditado(ByVal vData As Double)
    mvardValorCreditado = vData
End Property

Public Property Get dValorCreditado() As Double
    dValorCreditado = mvardValorCreditado
End Property

Public Property Let dSaldoAnterior(ByVal vData As Double)
Attribute dSaldoAnterior.VB_Description = "Valor do Saldo antes do resgate"
    mvardSaldoAnterior = vData
End Property

Public Property Get dSaldoAnterior() As Double
    dSaldoAnterior = mvardSaldoAnterior
End Property

Public Property Let dDescontos(ByVal vData As Double)
Attribute dDescontos.VB_Description = "valor dos descontos do tipo IOF sobre o valor resgatado"
    mvardDescontos = vData
End Property

Public Property Get dDescontos() As Double
    dDescontos = mvardDescontos
End Property

Public Property Let dValorIRRF(ByVal vData As Double)
Attribute dValorIRRF.VB_Description = "Valor do imposto de renda retido na fonte."
    mvardValorIRRF = vData
End Property

Public Property Get dValorIRRF() As Double
    dValorIRRF = mvardValorIRRF
End Property

Public Property Let dRendimentos(ByVal vData As Double)
Attribute dRendimentos.VB_Description = "valor acrescido ao saldo da aplicacao entre este resgate e o anterior ou a data de aplicacao, se for o 1o resgate."
    mvardRendimentos = vData
End Property

Public Property Get dRendimentos() As Double
    dRendimentos = mvardRendimentos
End Property

Public Property Let dValorResgatado(ByVal vData As Double)
Attribute dValorResgatado.VB_Description = "valor do resgate, incluindo os rendimentos, antes dos descontos"
    mvardValorResgatado = vData
End Property

Public Property Get dValorResgatado() As Double
    dValorResgatado = mvardValorResgatado
End Property

Public Property Let lNumMovto(ByVal vData As Long)
Attribute lNumMovto.VB_Description = "identifica o deposito em MovtosContaCorrente"
    mvarlNumMovto = vData
End Property

Public Property Get lNumMovto() As Long
    lNumMovto = mvarlNumMovto
End Property

Public Property Let iSeqResgate(ByVal vData As Integer)
Attribute iSeqResgate.VB_Description = "sequencial do resgate na aplicacao"
    mvariSeqResgate = vData
End Property

Public Property Get iSeqResgate() As Integer
    iSeqResgate = mvariSeqResgate
End Property

Public Property Let lCodigoAplicacao(ByVal vData As Long)
Attribute lCodigoAplicacao.VB_Description = "Identificador da Aplicacao"
    mvarlCodigoAplicacao = vData
End Property

Public Property Get lCodigoAplicacao() As Long
    lCodigoAplicacao = mvarlCodigoAplicacao
End Property



