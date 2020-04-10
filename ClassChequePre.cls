VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassChequePre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlNumIntCheque As Long
Private mvariFilialEmpresa As Integer
Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvariBanco As Integer
Private mvarsAgencia As String
Private mvarsContaCorrente As String
Private mvarlNumero As Long
Private mvardtDataDeposito As Date
Private mvardValor As Double
Private mvarlNumBordero As Long
Private mvariChequeSel As Integer
Private mvarsCPFCGC As String
Private mvarlNumMovtoCaixa As Long
Private mvariAprovado As Integer
Private mvariNaoEspecificado As Integer
Private mvarlNumBorderoLoja As Long
Private mvarlSequencialLoja As Long
Private mvarlSequencialBack As Long
Private mvariFilialEmpresaLoja As Integer
Private mvarlSequencial As Long 'este campo só existe na view
Private mvarlCupomFiscal As Long 'estes campos só existem na tela
Private mvariECF As Integer 'estes campos só existem na tela
Private mvarlNumMovtoSangria As Long
Private mvariStatus As Integer
Private mvariTipoBordero As Integer
Private mvarsCarne As String
Private mvarlNumIntChequeBord As Long
Private mvariCaixa As Integer
Private mvarlSequencialCaixa As Long
Private mvarlNumIntExt As Long
Private mvarlNumBorderoLojaBanco As Long
Private mvariLocalizacao As Integer
Private mvardtDataEmissao As Date

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let iLocalizacao(ByVal vData As Integer)
    mvariLocalizacao = vData
End Property

Public Property Get iLocalizacao() As Integer
    iLocalizacao = mvariLocalizacao
End Property

Public Property Let lNumBorderoLojaBanco(ByVal vData As Long)
    mvarlNumBorderoLojaBanco = vData
End Property

Public Property Get lNumBorderoLojaBanco() As Long
    lNumBorderoLojaBanco = mvarlNumBorderoLojaBanco
End Property

Public Property Let lNumIntExt(ByVal vData As Long)
    mvarlNumIntExt = vData
End Property

Public Property Get lNumIntExt() As Long
    lNumIntExt = mvarlNumIntExt
End Property


Public Property Let lSequencialCaixa(ByVal vData As Long)
    mvarlSequencialCaixa = vData
End Property

Public Property Get lSequencialCaixa() As Long
    lSequencialCaixa = mvarlSequencialCaixa
End Property

Public Property Let iCaixa(ByVal vData As Integer)
    mvariCaixa = vData
End Property

Public Property Get iCaixa() As Integer
    iCaixa = mvariCaixa
End Property

Public Property Let lNumIntChequeBord(ByVal vData As Long)
    mvarlNumIntChequeBord = vData
End Property

Public Property Get lNumIntChequeBord() As Long
    lNumIntChequeBord = mvarlNumIntChequeBord
End Property

Public Property Let sCarne(ByVal vData As String)
    mvarsCarne = vData
End Property

Public Property Get sCarne() As String
    sCarne = mvarsCarne
End Property

Public Property Let iTipoBordero(ByVal vData As Integer)
    mvariTipoBordero = vData
End Property

Public Property Get iTipoBordero() As Integer
    iTipoBordero = mvariTipoBordero
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumMovtoSangria(ByVal vData As Long)
    mvarlNumMovtoSangria = vData
End Property

Public Property Get lNumMovtoSangria() As Long
    lNumMovtoSangria = mvarlNumMovtoSangria
End Property

Public Property Let iECF(ByVal vData As Integer)
    mvariECF = vData
End Property

Public Property Get iECF() As Integer
    iECF = mvariECF
End Property

Public Property Let lCupomFiscal(ByVal vData As Long)
    mvarlCupomFiscal = vData
End Property

Public Property Get lCupomFiscal() As Long
    lCupomFiscal = mvarlCupomFiscal
End Property
    
Public Property Let lSequencial(ByVal vData As Long)
    mvarlSequencial = vData
End Property

Public Property Get lSequencial() As Long
    lSequencial = mvarlSequencial
End Property

Public Property Let iFilialEmpresaLoja(ByVal vData As Integer)
    mvariFilialEmpresaLoja = vData
End Property

Public Property Get iFilialEmpresaLoja() As Integer
    iFilialEmpresaLoja = mvariFilialEmpresaLoja
End Property

Public Property Let lSequencialBack(ByVal vData As Long)
    mvarlSequencialBack = vData
End Property

Public Property Get lSequencialBack() As Long
    lSequencialBack = mvarlSequencialBack
End Property

Public Property Let lSequencialLoja(ByVal vData As Long)
    mvarlSequencialLoja = vData
End Property

Public Property Get lSequencialLoja() As Long
    lSequencialLoja = mvarlSequencialLoja
End Property

Public Property Let lNumBorderoLoja(ByVal vData As Long)
    mvarlNumBorderoLoja = vData
End Property

Public Property Get lNumBorderoLoja() As Long
    lNumBorderoLoja = mvarlNumBorderoLoja
End Property

Public Property Let iNaoEspecificado(ByVal vData As Integer)
    mvariNaoEspecificado = vData
End Property

Public Property Get iNaoEspecificado() As Integer
    iNaoEspecificado = mvariNaoEspecificado
End Property

Public Property Let iAprovado(ByVal vData As Integer)
    mvariAprovado = vData
End Property

Public Property Get iAprovado() As Integer
    iAprovado = mvariAprovado
End Property

Public Property Let lNumMovtoCaixa(ByVal vData As Long)
    mvarlNumMovtoCaixa = vData
End Property

Public Property Get lNumMovtoCaixa() As Long
    lNumMovtoCaixa = mvarlNumMovtoCaixa
End Property

Public Property Let sCPFCGC(ByVal vData As String)
    mvarsCPFCGC = vData
End Property

Public Property Get sCPFCGC() As String
    sCPFCGC = mvarsCPFCGC
End Property

Public Property Let iChequeSel(ByVal vData As Integer)
    mvariChequeSel = vData
End Property

Public Property Get iChequeSel() As Integer
    iChequeSel = mvariChequeSel
End Property

Public Property Let lNumBordero(ByVal vData As Long)
    mvarlNumBordero = vData
End Property

Public Property Get lNumBordero() As Long
    lNumBordero = mvarlNumBordero
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dtDataDeposito(ByVal vData As Date)
    mvardtDataDeposito = vData
End Property

Public Property Get dtDataDeposito() As Date
    dtDataDeposito = mvardtDataDeposito
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let sContaCorrente(ByVal vData As String)
    mvarsContaCorrente = vData
End Property

Public Property Get sContaCorrente() As String
    sContaCorrente = mvarsContaCorrente
End Property

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lNumIntCheque(ByVal vData As Long)
    mvarlNumIntCheque = vData
End Property

Public Property Get lNumIntCheque() As Long
    lNumIntCheque = mvarlNumIntCheque
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property


