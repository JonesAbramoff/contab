VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTituloReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colParcelaReceber"
Option Explicit

'####################################################
'Inserido por Wagner
Public objInfoUsu As Object
'####################################################

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvarsSiglaDocumento As String
Private mvarlNumTitulo As Long
Private mvardtDataEmissao As Date
Private mvariStatus As Integer
Private mvardSaldo As Double
Private mvariNumParcelas As Integer
Private mvardValor As Double
Private mvardValorIRRF As Double
Private mvardValorISS As Double
Private mvardISSRetido As Double
Private mvardValorINSS As Double
Private mvariINSSRetido As Integer
Private mvardPercJurosDiario As Double
Private mvardPercMulta As Double
Private mvarsObservacao As String
Private mvariCondicaoPagto As Integer
Private mvardtDataRegistro As Date
Private mvarcolParcelaReceber As New colParcelaReceber
Private mvarcolComissoesEmissao As New Collection
Private mvariEspecie As Integer
Private mvariExclusao As Integer 'indica que está ocorrendo a exclusao fisica de uma nota fiscal vinculada ao tituto
Private mvardPISRetido As Double
Private mvardCOFINSRetido As Double
Private mvardCSLLRetido As Double

Private mvariMoeda As Integer
Private mvarsCcl As String
Private mvarlCpoGenerico1 As Long
Private mvariReajustePeriodicidade As Integer
Private mvardtReajusteBase As Date
Private mvardtReajustadoAte As Date
Private mvarsNatureza As String

Private mvarobjProjetoInfo As Object

Private mvariBaixaTela As Integer
Private mvariParcelaTela As Integer
Private mvardtDataEstorno As Date
Private mvardValorDesconto As Double

Public iAtivo As Integer
Public dComissao As Double

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dtDataEstorno(ByVal vData As Date)
    mvardtDataEstorno = vData
End Property

Public Property Get dtDataEstorno() As Date
    dtDataEstorno = mvardtDataEstorno
End Property

Public Property Let iParcelaTela(ByVal vData As Integer)
    mvariParcelaTela = vData
End Property

Public Property Get iParcelaTela() As Integer
    iParcelaTela = mvariParcelaTela
End Property

Public Property Let iBaixaTela(ByVal vData As Integer)
    mvariBaixaTela = vData
End Property

Public Property Get iBaixaTela() As Integer
    iBaixaTela = mvariBaixaTela
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Let dtReajusteBase(ByVal vData As Date)
    mvardtReajusteBase = vData
End Property

Public Property Get dtReajusteBase() As Date
    dtReajusteBase = mvardtReajusteBase
End Property

Public Property Let dtReajustadoAte(ByVal vData As Date)
    mvardtReajustadoAte = vData
End Property

Public Property Get dtReajustadoAte() As Date
    dtReajustadoAte = mvardtReajustadoAte
End Property

Public Property Let lCpoGenerico1(ByVal vData As Long)
    mvarlCpoGenerico1 = vData
End Property

Public Property Get lCpoGenerico1() As Long
    lCpoGenerico1 = mvarlCpoGenerico1
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let iReajustePeriodicidade(ByVal vData As Integer)
    mvariReajustePeriodicidade = vData
End Property

Public Property Get iReajustePeriodicidade() As Integer
    iReajustePeriodicidade = mvariReajustePeriodicidade
End Property

Public Property Let dCSLLRetido(ByVal vData As Double)
    mvardCSLLRetido = vData
End Property

Public Property Get dCSLLRetido() As Double
    dCSLLRetido = mvardCSLLRetido
End Property

Public Property Let dCOFINSRetido(ByVal vData As Double)
    mvardCOFINSRetido = vData
End Property

Public Property Get dCOFINSRetido() As Double
    dCOFINSRetido = mvardCOFINSRetido
End Property

Public Property Let dPISRetido(ByVal vData As Double)
    mvardPISRetido = vData
End Property

Public Property Get dPISRetido() As Double
    dPISRetido = mvardPISRetido
End Property

Public Property Let iExclusao(ByVal vData As Integer)
    mvariExclusao = vData
End Property

Public Property Get iExclusao() As Integer
    iExclusao = mvariExclusao
End Property

Public Property Let iEspecie(ByVal vData As Integer)
    mvariEspecie = vData
End Property

Public Property Get iEspecie() As Integer
    iEspecie = mvariEspecie
End Property

Public Property Set colComissoesEmissao(ByVal vData As Object)
    Set mvarcolComissoesEmissao = vData
End Property

Public Property Get colComissoesEmissao() As Collection
    Set colComissoesEmissao = mvarcolComissoesEmissao
End Property

Public Property Set colParcelaReceber(ByVal vData As Object)
    Set mvarcolParcelaReceber = vData
End Property

Public Property Get colParcelaReceber() As colParcelaReceber
    Set colParcelaReceber = mvarcolParcelaReceber
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let iNumParcelas(ByVal vData As Integer)
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let dSaldo(ByVal vData As Double)
    mvardSaldo = vData
End Property

Public Property Get dSaldo() As Double
    dSaldo = mvardSaldo
End Property

Public Property Let sSiglaDocumento(ByVal vData As String)
    mvarsSiglaDocumento = vData
End Property

Public Property Get sSiglaDocumento() As String
    sSiglaDocumento = mvarsSiglaDocumento
End Property

Public Property Let dISSRetido(ByVal vData As Double)
    mvardISSRetido = vData
End Property

Public Property Get dISSRetido() As Double
    dISSRetido = mvardISSRetido
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let dValorISS(ByVal vData As Double)
    mvardValorISS = vData
End Property

Public Property Get dValorISS() As Double
    dValorISS = mvardValorISS
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dPercMulta(ByVal vData As Double)
    mvardPercMulta = vData
End Property

Public Property Get dPercMulta() As Double
    dPercMulta = mvardPercMulta
End Property

Public Property Let dPercJurosDiario(ByVal vData As Double)
    mvardPercJurosDiario = vData
End Property

Public Property Get dPercJurosDiario() As Double
    dPercJurosDiario = mvardPercJurosDiario
End Property

Public Property Let lNumTitulo(ByVal vData As Long)
    mvarlNumTitulo = vData
End Property

Public Property Get lNumTitulo() As Long
    lNumTitulo = mvarlNumTitulo
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dValorIRRF(ByVal vData As Double)
    mvardValorIRRF = vData
End Property

Public Property Get dValorIRRF() As Double
    dValorIRRF = mvardValorIRRF
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let dValorINSS(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dValorINSS() As Double
    dValorINSS = mvardValorINSS
End Property

Public Property Let iINSSRetido(ByVal vData As Integer)
    mvariINSSRetido = vData
End Property

Public Property Get iINSSRetido() As Integer
    iINSSRetido = mvariINSSRetido
End Property

Private Sub Class_Initialize()
    mvardtDataEmissao = DATA_NULA
    mvardtReajustadoAte = DATA_NULA
    mvardtReajusteBase = DATA_NULA
    mvardtDataEstorno = DATA_NULA
End Sub
