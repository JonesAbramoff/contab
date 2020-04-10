VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassDetRetCobr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariCodCarteiraNoBanco As Integer
Private mvariCodOcorrencia As Integer
Private mvardtDataCredito As Date
Private mvardtDataVencimento As Date
Private mvardtDataBaixa As Date
Private mvardValorTitulo As Double
Private mvardValorRecebido As Double
Private mvardValorJuros As Double
Private mvardValorDesconto As Double
Private mvardValorTarifa As Double
Private mvarsSeuNumero As String
Private mvarsNossoNumero As String
Private mvarsUsoDaEmpresa As String
Private mvarlNumSequencial As Long
Private mvariTipoCritica As Integer
Private mvardValorEntregue As Double
Private mvardValorJurosCliente As Double

Private mvariCodOcorrencia1 As Integer
Private mvariCodOcorrencia2 As Integer
Private mvariCodOcorrencia3 As Integer
Private mvariCodOcorrencia4 As Integer
Private mvariCodOcorrencia5 As Integer
Private mvarlNumIntDoc As Long
Private mvarlNumIntParc As Long
Private mvariTipoOcorrencia As Integer
Private mvarsCodLiquidacao As String
Private mvardValorCreditado As Double

Public Property Let dValorCreditado(ByVal vData As Double)
    mvardValorCreditado = vData
End Property

Public Property Get dValorCreditado() As Double
    dValorCreditado = mvardValorCreditado
End Property

Public Property Let sCodLiquidacao(ByVal vData As String)
    mvarsCodLiquidacao = vData
End Property

Public Property Get sCodLiquidacao() As String
    sCodLiquidacao = mvarsCodLiquidacao
End Property

Public Property Let iTipoOcorrencia(ByVal vData As Integer)
    mvariTipoOcorrencia = vData
End Property

Public Property Get iTipoOcorrencia() As Integer
    iTipoOcorrencia = mvariTipoOcorrencia
End Property

Public Property Let lNumIntParc(ByVal vData As Long)
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc() As Long
    lNumIntParc = mvarlNumIntParc
End Property

Public Property Let iCodOcorrencia1(ByVal vData As Integer)
    mvariCodOcorrencia1 = vData
End Property

Public Property Get iCodOcorrencia1() As Integer
    iCodOcorrencia1 = mvariCodOcorrencia1
End Property

Public Property Let iCodOcorrencia2(ByVal vData As Integer)
    mvariCodOcorrencia2 = vData
End Property

Public Property Get iCodOcorrencia2() As Integer
    iCodOcorrencia2 = mvariCodOcorrencia2
End Property

Public Property Let iCodOcorrencia3(ByVal vData As Integer)
    mvariCodOcorrencia3 = vData
End Property

Public Property Get iCodOcorrencia3() As Integer
    iCodOcorrencia3 = mvariCodOcorrencia3
End Property

Public Property Let iCodOcorrencia4(ByVal vData As Integer)
    mvariCodOcorrencia4 = vData
End Property

Public Property Get iCodOcorrencia4() As Integer
    iCodOcorrencia4 = mvariCodOcorrencia4
End Property

Public Property Let iCodOcorrencia5(ByVal vData As Integer)
    mvariCodOcorrencia5 = vData
End Property

Public Property Get iCodOcorrencia5() As Integer
    iCodOcorrencia5 = mvariCodOcorrencia5
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let dValorJurosCliente(ByVal vData As Double)
    mvardValorJurosCliente = vData
End Property

Public Property Get dValorJurosCliente() As Double
    dValorJurosCliente = mvardValorJurosCliente
End Property

Public Property Let dValorEntregue(ByVal vData As Double)
    mvardValorEntregue = vData
End Property

Public Property Get dValorEntregue() As Double
    dValorEntregue = mvardValorEntregue
End Property

Public Property Let iTipoCritica(ByVal vData As Integer)
    mvariTipoCritica = vData
End Property

Public Property Get iTipoCritica() As Integer
    iTipoCritica = mvariTipoCritica
End Property

Public Property Let dValorTitulo(ByVal vData As Double)
    mvardValorTitulo = vData
End Property

Public Property Get dValorTitulo() As Double
    dValorTitulo = mvardValorTitulo
End Property

Public Property Let lNumSequencial(ByVal vData As Long)
    mvarlNumSequencial = vData
End Property

Public Property Get lNumSequencial() As Long
    lNumSequencial = mvarlNumSequencial
End Property

Public Property Let sSeuNumero(ByVal vData As String)
    mvarsSeuNumero = vData
End Property

Public Property Get sSeuNumero() As String
    sSeuNumero = mvarsSeuNumero
End Property

Public Property Let sUsoDaEmpresa(ByVal vData As String)
    mvarsUsoDaEmpresa = vData
End Property

Public Property Get sUsoDaEmpresa() As String
    sUsoDaEmpresa = mvarsUsoDaEmpresa
End Property

Public Property Let iCodCarteiraNoBanco(ByVal vData As Integer)
    mvariCodCarteiraNoBanco = vData
End Property

Public Property Get iCodCarteiraNoBanco() As Integer
    iCodCarteiraNoBanco = mvariCodCarteiraNoBanco
End Property

Public Property Let iCodOcorrencia(ByVal vData As Integer)
    mvariCodOcorrencia = vData
End Property

Public Property Get iCodOcorrencia() As Integer
    iCodOcorrencia = mvariCodOcorrencia
End Property

Public Property Let sNossoNumero(ByVal vData As String)
    mvarsNossoNumero = vData
End Property

Public Property Get sNossoNumero() As String
    sNossoNumero = mvarsNossoNumero
End Property

Public Property Let dValorRecebido(ByVal vData As Double)
    mvardValorRecebido = vData
End Property

Public Property Get dValorRecebido() As Double
    dValorRecebido = mvardValorRecebido
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorTarifa(ByVal vData As Double)
    mvardValorTarifa = vData
End Property

Public Property Get dValorTarifa() As Double
    dValorTarifa = mvardValorTarifa
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
     dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
     dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let dtDataCredito(ByVal vData As Date)
    mvardtDataCredito = vData
End Property

Public Property Get dtDataCredito() As Date
     dtDataCredito = mvardtDataCredito
End Property
