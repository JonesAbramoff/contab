VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCarteiraCobrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariCobrador As Integer
Private mvariCodCarteiraCobranca As Integer
Private mvarsContaContabil As String
Private mvariDesativada As Integer
Private mvariDiasDeRetencao As Integer
Private mvardTaxaCobranca As Double
Private mvardTaxaDesconto As Double
Private mvarsContaDuplDescontadas As String
Private mvarlQuantidadeAtual As Long
Private mvarlQuantidadeAtualBanco As Long
Private mvardSaldoAtual As Double
Private mvardSaldoAtualBanco As Double
Private mvarsFaixaNossoNumeroInicial As String
Private mvarsFaixaNossoNumeroFinal As String
Private mvarsFaixaNossoNumeroProx As String
Private mvarsNomeNoBanco As String
Private mvarsCodCarteiraNoBanco As String
Private mvariNumCarteiraNoBanco As Integer
Private mvariComRegistro As Integer
Private mvariImprimeBoleta As Integer
Private mvariGeraNossoNumero As Integer
Private mvariFormPreImp As Integer

Public Property Let iFormPreImp(ByVal vData As Integer)
    mvariFormPreImp = vData
End Property

Public Property Get iFormPreImp() As Integer
    iFormPreImp = mvariFormPreImp
End Property

Public Property Let iGeraNossoNumero(ByVal vData As Integer)
    mvariGeraNossoNumero = vData
End Property

Public Property Get iGeraNossoNumero() As Integer
    iGeraNossoNumero = mvariGeraNossoNumero
End Property

Public Property Let iImprimeBoleta(ByVal vData As Integer)
    mvariImprimeBoleta = vData
End Property

Public Property Get iImprimeBoleta() As Integer
    iImprimeBoleta = mvariImprimeBoleta
End Property
Public Property Let iComRegistro(ByVal vData As Integer)
    mvariComRegistro = vData
End Property

Public Property Get iComRegistro() As Integer
    iComRegistro = mvariComRegistro
End Property
Public Property Let iNumCarteiraNoBanco(ByVal vData As Integer)
    mvariNumCarteiraNoBanco = vData
End Property

Public Property Get iNumCarteiraNoBanco() As Integer
    iNumCarteiraNoBanco = mvariNumCarteiraNoBanco
End Property

Public Property Let sCodCarteiraNoBanco(ByVal vData As String)
    mvarsCodCarteiraNoBanco = vData
End Property

Public Property Get sCodCarteiraNoBanco() As String
    sCodCarteiraNoBanco = mvarsCodCarteiraNoBanco
End Property

Public Property Let sNomeNoBanco(ByVal vData As String)
    mvarsNomeNoBanco = vData
End Property

Public Property Get sNomeNoBanco() As String
    sNomeNoBanco = mvarsNomeNoBanco
End Property

Public Property Let sFaixaNossoNumeroProx(ByVal vData As String)
Attribute sFaixaNossoNumeroProx.VB_Description = "prox # a ser usado da faixa liberada pelo banco"
    mvarsFaixaNossoNumeroProx = vData
End Property

Public Property Get sFaixaNossoNumeroProx() As String
    sFaixaNossoNumeroProx = mvarsFaixaNossoNumeroProx
End Property

Public Property Let sFaixaNossoNumeroFinal(ByVal vData As String)
Attribute sFaixaNossoNumeroFinal.VB_Description = "ultimo # da faixa liberada pelo banco"
    mvarsFaixaNossoNumeroFinal = vData
End Property

Public Property Get sFaixaNossoNumeroFinal() As String
    sFaixaNossoNumeroFinal = mvarsFaixaNossoNumeroFinal
End Property

Public Property Let sFaixaNossoNumeroInicial(ByVal vData As String)
Attribute sFaixaNossoNumeroInicial.VB_Description = "primeiro # da faixa liberada pelo banco"
    mvarsFaixaNossoNumeroInicial = vData
End Property

Public Property Get sFaixaNossoNumeroInicial() As String
    sFaixaNossoNumeroInicial = mvarsFaixaNossoNumeroInicial
End Property

Public Property Let dSaldoAtualBanco(ByVal vData As Double)
Attribute dSaldoAtualBanco.VB_Description = "saldo das parcelas nesta carteira segundo arquivo retorno do bco"
    mvardSaldoAtualBanco = vData
End Property

Public Property Get dSaldoAtualBanco() As Double
    dSaldoAtualBanco = mvardSaldoAtualBanco
End Property

Public Property Let dSaldoAtual(ByVal vData As Double)
Attribute dSaldoAtual.VB_Description = "saldo das parcelas nesta carteira  (estatistico)"
    mvardSaldoAtual = vData
End Property

Public Property Get dSaldoAtual() As Double
    dSaldoAtual = mvardSaldoAtual
End Property

Public Property Let lQuantidadeAtualBanco(ByVal vData As Long)
Attribute lQuantidadeAtualBanco.VB_Description = "de parcelas nesta carteira (estatistico) segundo arquivo retorno do bco"
    mvarlQuantidadeAtualBanco = vData
End Property

Public Property Get lQuantidadeAtualBanco() As Long
    lQuantidadeAtualBanco = mvarlQuantidadeAtualBanco
End Property

Public Property Let lQuantidadeAtual(ByVal vData As Long)
Attribute lQuantidadeAtual.VB_Description = "de parcelas nesta carteira (estatistico)"
    mvarlQuantidadeAtual = vData
End Property

Public Property Get lQuantidadeAtual() As Long
    lQuantidadeAtual = mvarlQuantidadeAtual
End Property

Public Property Let sContaDuplDescontadas(ByVal vData As String)
Attribute sContaDuplDescontadas.VB_Description = "conta contabil que registra as duplicatas descontadas pelo bco"
    mvarsContaDuplDescontadas = vData
End Property

Public Property Get sContaDuplDescontadas() As String
    sContaDuplDescontadas = mvarsContaDuplDescontadas
End Property

Public Property Let dTaxaDesconto(ByVal vData As Double)
    mvardTaxaDesconto = vData
End Property

Public Property Get dTaxaDesconto() As Double
    dTaxaDesconto = mvardTaxaDesconto
End Property

Public Property Let dTaxaCobranca(ByVal vData As Double)
    mvardTaxaCobranca = vData
End Property

Public Property Get dTaxaCobranca() As Double
    dTaxaCobranca = mvardTaxaCobranca
End Property

Public Property Let iDiasDeRetencao(ByVal vData As Integer)
    mvariDiasDeRetencao = vData
End Property

Public Property Get iDiasDeRetencao() As Integer
    iDiasDeRetencao = mvariDiasDeRetencao
End Property

Public Property Let iDesativada(ByVal vData As Integer)
Attribute iDesativada.VB_Description = "0 ou 1"
    mvariDesativada = vData
End Property

Public Property Get iDesativada() As Integer
    iDesativada = mvariDesativada
End Property

Public Property Let sContaContabil(ByVal vData As String)
Attribute sContaContabil.VB_Description = "conta de ctas a receber correspondente a esta carteira. Em cobranca descontada registra os titulos enviados ao bco p/desconto."
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let iCodCarteiraCobranca(ByVal vData As Integer)
Attribute iCodCarteiraCobranca.VB_Description = "chave p/CarteirasCobranca"
    mvariCodCarteiraCobranca = vData
End Property

Public Property Get iCodCarteiraCobranca() As Integer
    iCodCarteiraCobranca = mvariCodCarteiraCobranca
End Property

Public Property Let iCobrador(ByVal vData As Integer)
    mvariCobrador = vData
End Property

Public Property Get iCobrador() As Integer
    iCobrador = mvariCobrador
End Property



