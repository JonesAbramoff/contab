VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassInfoParcRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvariMarcada As Integer
Private mvarsNomeRedCliente As String
Private mvarsRazaoSocialCli As String
Private mvariFilialCliente As Integer
Private mvarlCliente As Long
Private mvarlNumTitulo As Long
Private mvariNumParcela As Integer
Private mvarlNumIntParc As Long
Private mvardValor As Double
Private mvardValorOriginal As Double
Private mvardValorJuros As Double
Private mvardValorMulta As Double
Private mvardValorDesconto As Double
Private mvardtVencimento As Date
Private mvariPadraoCobranca As Integer
Private mvarsSiglaDocumento As String
Private mvariFilialEmpresa As Integer
Private mvarsCobradorNomeRed As String
Private mvarsCartCobrDesc As String
Private mvariCobrador As Integer
Private mvariCarteiraCobrador As Integer
Private mvardtDataVencimentoReal As Date
Private mvarlIdImpressaoBoleto As Long
Private mvardtDataImpressaoBoleto As Date
Private mvardtDataBaixa As Date
Private mvardValorBaixado As Double
Private mvariSequencial As Integer
'Janaina
Private mvardValorReceber As Double
Private mvardSaldoParcela As Double
'Janaina
Private mvardtDataEmissao As Date
Private mvarsNossoNumero As String

Public Property Let sNossoNumero(ByVal vData As String)
    mvarsNossoNumero = vData
End Property

Public Property Get sNossoNumero() As String
    sNossoNumero = mvarsNossoNumero
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

'Janaina
Public Property Let dValorReceber(ByVal vData As Double)
    mvardValorReceber = vData
'Janaina
End Property
'Janaina
Public Property Get dValorReceber() As Double
    dValorReceber = mvardValorReceber
'Janaina
End Property
'Janaina
Public Property Let dSaldoParcela(ByVal vData As Double)
    mvardSaldoParcela = vData
'Janaina
End Property
'Janaina
Public Property Get dSaldoParcela() As Double
    dSaldoParcela = mvardSaldoParcela
'Janaina
End Property

Public Property Let iSequencial(ByVal vData As Integer)
    mvariSequencial = vData
End Property

Public Property Get iSequencial() As Integer
    iSequencial = mvariSequencial
End Property

Public Property Let dValorBaixado(ByVal vData As Double)
    mvardValorBaixado = vData
End Property

Public Property Get dValorBaixado() As Double
    dValorBaixado = mvardValorBaixado
End Property

Public Property Let dtDataBaixa(ByVal vData As Date)
    mvardtDataBaixa = vData
End Property

Public Property Get dtDataBaixa() As Date
    dtDataBaixa = mvardtDataBaixa
End Property

Public Property Let dtDataImpressaoBoleto(ByVal vData As Date)
    mvardtDataImpressaoBoleto = vData
End Property

Public Property Get dtDataImpressaoBoleto() As Date
    dtDataImpressaoBoleto = mvardtDataImpressaoBoleto
End Property

Public Property Let lIdImpressaoBoleto(ByVal vData As Long)
    mvarlIdImpressaoBoleto = vData
End Property

Public Property Get lIdImpressaoBoleto() As Long
    lIdImpressaoBoleto = mvarlIdImpressaoBoleto
End Property

Public Property Let dtDataVencimentoReal(ByVal vData As Date)
    mvardtDataVencimentoReal = vData
End Property

Public Property Get dtDataVencimentoReal() As Date
    dtDataVencimentoReal = mvardtDataVencimentoReal
End Property

Public Property Let sCartCobrDesc(ByVal vData As String)
    mvarsCartCobrDesc = vData
End Property

Public Property Get sCartCobrDesc() As String
    sCartCobrDesc = mvarsCartCobrDesc
End Property

Public Property Let sCobradorNomeRed(ByVal vData As String)
    mvarsCobradorNomeRed = vData
End Property

Public Property Get sCobradorNomeRed() As String
    sCobradorNomeRed = mvarsCobradorNomeRed
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iPadraoCobranca(ByVal vData As Integer)
    mvariPadraoCobranca = vData
End Property

Public Property Get iPadraoCobranca() As Integer
    iPadraoCobranca = mvariPadraoCobranca
End Property

Public Property Let dtVencimento(ByVal vData As Date)
    mvardtVencimento = vData
End Property

Public Property Get dtVencimento() As Date
    dtVencimento = mvardtVencimento
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dValorOriginal(ByVal vData As Double)
    mvardValorOriginal = vData
End Property

Public Property Get dValorOriginal() As Double
    dValorOriginal = mvardValorOriginal
End Property

Public Property Let dValorJuros(ByVal vData As Double)
    mvardValorJuros = vData
End Property

Public Property Get dValorJuros() As Double
    dValorJuros = mvardValorJuros
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

Public Property Let lNumIntParc(ByVal vData As Long)
    mvarlNumIntParc = vData
End Property

Public Property Get lNumIntParc() As Long
    lNumIntParc = mvarlNumIntParc
End Property

Public Property Let iNumParcela(ByVal vData As Integer)
    mvariNumParcela = vData
End Property

Public Property Get iNumParcela() As Integer
    iNumParcela = mvariNumParcela
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iFilialCliente(ByVal vData As Integer)
    mvariFilialCliente = vData
End Property

Public Property Get iFilialCliente() As Integer
    iFilialCliente = mvariFilialCliente
End Property

Public Property Let sRazaoSocialCli(ByVal vData As String)
    mvarsRazaoSocialCli = vData
End Property

Public Property Get sRazaoSocialCli() As String
    sRazaoSocialCli = mvarsRazaoSocialCli
End Property

Public Property Let sNomeRedCliente(ByVal vData As String)
    mvarsNomeRedCliente = vData
End Property

Public Property Get sNomeRedCliente() As String
    sNomeRedCliente = mvarsNomeRedCliente
End Property

Public Property Let iMarcada(ByVal vData As Integer)
    mvariMarcada = vData
End Property

Public Property Get iMarcada() As Integer
    iMarcada = mvariMarcada
End Property

Public Property Let sSiglaDocumento(ByVal vData As String)
    mvarsSiglaDocumento = vData
End Property

Public Property Get sSiglaDocumento() As String
    sSiglaDocumento = mvarsSiglaDocumento
End Property

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


