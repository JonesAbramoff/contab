VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoBaixaRecCancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntBaixaParcRec As Long
Private mvarlNumIntCancelamento As Long
Private mvarlNumIntTitulo As Long
Private mvarlNumIntBaixa As Long
Private mvardValorDesconto As Double
Private mvardValorMulta As Double
Private mvardValorBaixado As Double
Private mvardValorJuros As Double
Private mvarlNumIntParcela As Long
Private mvardtDataBaixa As Date
Private mvardtDataCredito As Date
Private mvardtDataContabilBaixa As Date
Private mvardtDataRegistroBaixa As Date
Private mvariMotivoBaixa As Integer
Private mvariStatusBaixaRec As Integer
Private mvarlNumIntDebRecCli As Long
Private mvarlNumMovCta As Long
Private mvarsHistoricoBaixa As String
Private mvarsContaCorrente As String
Private mvardValorTipoBaixa As Double
Private mvardtDataEmissaoDebito As Date
Private mvariTipoMovCCI As Integer
Private mvarlNumDebito As Long
Private mvarsHistoricoMovCCI As String
Private mvariFilialEmpresa As Integer
Private mvarlNumDocumento As Long
Private mvarsSiglaDocumento As String
Private mvariNumParcela As Integer
Private mvardValorParcela As Double
Private mvariSequencial As Integer
Private mvardValorDescontoCanc As Double
Private mvardValorJurosCanc As Double
Private mvardValorMultaCanc As Double
Private mvardValorBaixadoCanc As Double
Private mvardValorTotalCanc As Double
Private mvariStatusBaixaParcRec As Integer
Private mvardValorMovCCI As Double
Private mvarlNumIntRecAntecip As Long
Private mvardSaldoDebito As Double
Private mvardValorDebito As Double
Private mvardtDataCancelamento As Date
Private mvardValorPago As Double
Private mvariCodConta As Integer
Private mvariTipoMeioPagto As Integer
Private mvariItemSelecionado As Integer
Private mvariCobrador As Integer
Private mvariCarteiraCobrador As Integer

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property

Public Property Let iCarteiraCobrador(ByVal vData As Integer)
    mvariCarteiraCobrador = vData
End Property

Public Property Get iCarteiraCobrador() As Integer
    iCarteiraCobrador = mvariCarteiraCobrador
End Property

Public Property Let dtDataCancelamento(ByVal vData As Date)
    mvardtDataCancelamento = vData
End Property

Public Property Get dtDataCancelamento() As Date
    dtDataCancelamento = mvardtDataCancelamento
End Property

Public Property Let dValorDebito(ByVal vData As Double)
    mvardValorDebito = vData
End Property

Public Property Get dValorDebito() As Double
    dValorDebito = mvardValorDebito
End Property

Public Property Let dSaldoDebito(ByVal vData As Double)
    mvardSaldoDebito = vData
End Property

Public Property Get dSaldoDebito() As Double
    dSaldoDebito = mvardSaldoDebito
End Property

Public Property Let lNumIntRecAntecip(ByVal vData As Long)
    mvarlNumIntRecAntecip = vData
End Property

Public Property Get lNumIntRecAntecip() As Long
    lNumIntRecAntecip = mvarlNumIntRecAntecip
End Property

Public Property Let dValorMovCCI(ByVal vData As Double)
    mvardValorMovCCI = vData
End Property

Public Property Get dValorMovCCI() As Double
    dValorMovCCI = mvardValorMovCCI
End Property

Public Property Let iStatusBaixaParcRec(ByVal vData As Integer)
    mvariStatusBaixaParcRec = vData
End Property

Public Property Get iStatusBaixaParcRec() As Integer
    iStatusBaixaParcRec = mvariStatusBaixaParcRec
End Property

Public Property Let dValorTotalCanc(ByVal vData As Double)
    mvardValorTotalCanc = vData
End Property

Public Property Get dValorTotalCanc() As Double
    dValorTotalCanc = mvardValorTotalCanc
End Property

Public Property Let dValorBaixadoCanc(ByVal vData As Double)
    mvardValorBaixadoCanc = vData
End Property

Public Property Get dValorBaixadoCanc() As Double
    dValorBaixadoCanc = mvardValorBaixadoCanc
End Property

Public Property Let dValorMultaCanc(ByVal vData As Double)
    mvardValorMultaCanc = vData
End Property

Public Property Get dValorMultaCanc() As Double
    dValorMultaCanc = mvardValorMultaCanc
End Property

Public Property Let dValorJurosCanc(ByVal vData As Double)
    mvardValorJurosCanc = vData
End Property

Public Property Get dValorJurosCanc() As Double
    dValorJurosCanc = mvardValorJurosCanc
End Property

Public Property Let dValorDescontoCanc(ByVal vData As Double)
    mvardValorDescontoCanc = vData
End Property

Public Property Get dValorDescontoCanc() As Double
    dValorDescontoCanc = mvardValorDescontoCanc
End Property

Public Property Let iSequencial(ByVal vData As Integer)
    mvariSequencial = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let dValorParcela(ByVal vData As Double)
    mvardValorParcela = vData
End Property

Public Property Get dValorParcela() As Double
    dValorParcela = mvardValorParcela
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let sSiglaDocumento(ByVal vData As String)
    mvarsSiglaDocumento = vData
End Property

Public Property Get sSiglaDocumento() As String
    sSiglaDocumento = mvarsSiglaDocumento
End Property

Public Property Let lNumIntTitulo(ByVal vData As Long)
    mvarlNumIntTitulo = vData
End Property

Public Property Get lNumIntTitulo() As Long
    lNumIntTitulo = mvarlNumIntTitulo
End Property

Public Property Let lNumDocumento(ByVal vData As Long)
    mvarlNumDocumento = vData
End Property

Public Property Get lNumDocumento() As Long
    lNumDocumento = mvarlNumDocumento
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let sHistoricoMovCCI(ByVal vData As String)
    mvarsHistoricoMovCCI = vData
End Property

Public Property Get sHistoricoMovCCI() As String
    sHistoricoMovCCI = mvarsHistoricoMovCCI
End Property

Public Property Let lNumDebito(ByVal vData As Long)
    mvarlNumDebito = vData
End Property

Public Property Get lNumDebito() As Long
    lNumDebito = mvarlNumDebito
End Property

Public Property Let iTipoMovCCI(ByVal vData As Integer)
    mvariTipoMovCCI = vData
End Property

Public Property Get iTipoMovCCI() As Integer
    iTipoMovCCI = mvariTipoMovCCI
End Property

Public Property Let dtDataEmissaoDebito(ByVal vData As Date)
    mvardtDataEmissaoDebito = vData
End Property

Public Property Get dtDataEmissaoDebito() As Date
    dtDataEmissaoDebito = mvardtDataEmissaoDebito
End Property

Public Property Let dValorTipoBaixa(ByVal vData As Double)
    mvardValorTipoBaixa = vData
End Property

Public Property Get dValorTipoBaixa() As Double
    dValorTipoBaixa = mvardValorTipoBaixa
End Property

Public Property Let sContaCorrente(ByVal vData As String)
    mvarsContaCorrente = vData
End Property

Public Property Get sContaCorrente() As String
    sContaCorrente = mvarsContaCorrente
End Property

Public Property Let sHistoricoBaixa(ByVal vData As String)
    mvarsHistoricoBaixa = vData
End Property

Public Property Get sHistoricoBaixa() As String
    sHistoricoBaixa = mvarsHistoricoBaixa
End Property

Public Property Let lNumMovCta(ByVal vData As Long)
    mvarlNumMovCta = vData
End Property

Public Property Get lNumMovCta() As Long
    lNumMovCta = mvarlNumMovCta
End Property

Public Property Let lNumIntDebRecCli(ByVal vData As Long)
    mvarlNumIntDebRecCli = vData
End Property

Public Property Get lNumIntDebRecCli() As Long
    lNumIntDebRecCli = mvarlNumIntDebRecCli
End Property

Public Property Let iStatusBaixaRec(ByVal vData As Integer)
    mvariStatusBaixaRec = vData
End Property

Public Property Get iStatusBaixaRec() As Integer
    iStatusBaixaRec = mvariStatusBaixaRec
End Property

Public Property Let iMotivoBaixa(ByVal vData As Integer)
    mvariMotivoBaixa = vData
End Property

Public Property Get iMotivoBaixa() As Integer
    iMotivoBaixa = mvariMotivoBaixa
End Property

Public Property Let dtDataRegistroBaixa(ByVal vData As Date)
    mvardtDataRegistroBaixa = vData
End Property

Public Property Get dtDataRegistroBaixa() As Date
    dtDataRegistroBaixa = mvardtDataRegistroBaixa
End Property

Public Property Let dtDataContabilBaixa(ByVal vData As Date)
    mvardtDataContabilBaixa = vData
End Property

Public Property Get dtDataContabilBaixa() As Date
    dtDataContabilBaixa = mvardtDataContabilBaixa
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dtDataCredito(ByVal vData As Date)
    mvardtDataCredito = vData
End Property

Public Property Get dtDataCredito() As Date
    dtDataCredito = mvardtDataCredito
End Property

Public Property Let lNumIntParcela(ByVal vData As Long)
    mvarlNumIntParcela = vData
End Property

Public Property Get lNumIntParcela() As Long
    lNumIntParcela = mvarlNumIntParcela
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dValorMulta(ByVal vData As Double)
    mvardValorMulta = vData
End Property

Public Property Get dValorMulta() As Double
    dValorMulta = mvardValorMulta
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let lNumIntBaixa(ByVal vData As Long)
    mvarlNumIntBaixa = vData
End Property

Public Property Get lNumIntBaixa() As Long
    lNumIntBaixa = mvarlNumIntBaixa
End Property

Public Property Let lNumIntBaixaParcRec(ByVal vData As Long)
    mvarlNumIntBaixaParcRec = vData
End Property

Public Property Get lNumIntBaixaParcRec() As Long
    lNumIntBaixaParcRec = mvarlNumIntBaixaParcRec
End Property

Public Property Let lNumIntCancelamento(ByVal vData As Long)
    mvarlNumIntCancelamento = vData
End Property

Public Property Get lNumIntCancelamento() As Long
    lNumIntCancelamento = mvarlNumIntCancelamento
End Property

Public Property Let dValorPago(ByVal vData As Double)
    mvardValorPago = vData
End Property

Public Property Get dValorPago() As Double
    dValorPago = mvardValorPago
End Property

Public Property Let iCodConta(ByVal vData As Integer)
    mvariCodConta = vData
End Property

Public Property Get iCodConta() As Integer
    iCodConta = mvariCodConta
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let iItemSelecionado(ByVal vData As Integer)
    mvariItemSelecionado = vData
End Property

Public Property Get iItemSelecionado() As Integer
    iItemSelecionado = mvariItemSelecionado
End Property

Private Sub Class_Initialize()
    mvardtDataCredito = DATA_NULA
End Sub
