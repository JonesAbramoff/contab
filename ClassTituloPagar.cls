VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTituloPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardtDataEmissao As Date
Private mvariFilial As Integer
Private mvarlNumIntDoc As Long
Private mvardOutrasDespesas As Double
Private mvardValorFrete As Double
Private mvardValorProdutos As Double
Private mvardValorSeguro As Double
Private mvardValorTotal As Double
Private mvarlFornecedor As Long
Private mvariStatus As Integer
Private mvardValorIRRF As Double
Private mvardValorICMS As Double
Private mvardValorICMSSubst As Double
Private mvardValorIPI As Double
Private mvardValorINSS As Double
Private mvarsObservacao As String
Private mvarlNumTitulo As Long
Private mvariFilialEmpresa As Integer
Private mvariCreditoIPI As Integer
Private mvarsSiglaDocumento As String
Private mvardSaldo As Double
Private mvariNumParcelas As Integer
Private mvariCreditoICMS As Integer
Private mvarlNumPedCompra As Long
Private mvariFilialPedCompra As Integer
Private mvariINSSRetido As Integer
Private mvardtDataRegistro As Date
Private mvariCondicaoPagto As Integer
Private mvarcolParcelaPagar As New colParcelaPagar
Private mvariExclusao As Integer 'indica que está ocorrendo a exclusao fisica de uma nota fiscal vinculada ao tituto
Private mvardPISRetido As Double
Private mvardCOFINSRetido As Double
Private mvardCSLLRetido As Double
Private mvariMoeda As Integer
Private mvardTaxaMoeda As Double
Private mvarsHistorico As String
Private mvardISSRetido As Double

'####################################################
'Inserido por Wagner
Public objInfoUsu As Object
'####################################################

Private mvarobjProjetoInfo As Object
Private mvarobjAnexos As Object

Private mvarsNatureza As String
Private mvariBaixaTela As Integer
Private mvariParcelaTela As Integer
Private mvardtDataEstorno As Date
Private mvarsCcl As String

Public Property Set objAnexos(ByVal vData As Object)
    Set mvarobjAnexos = vData
End Property

Public Property Get objAnexos() As Object
    Set objAnexos = mvarobjAnexos
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
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

Public Property Let dISSRetido(ByVal vData As Double)
    mvardISSRetido = vData
End Property

Public Property Get dISSRetido() As Double
    dISSRetido = mvardISSRetido
End Property

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Let sHistorico(ByVal vData As String)
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dTaxaMoeda(ByVal vData As Double)
    mvardTaxaMoeda = vData
End Property

Public Property Get dTaxaMoeda() As Double
    dTaxaMoeda = mvardTaxaMoeda
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

Public Property Set colParcelaPagar(ByVal vData As Object)
    Set mvarcolParcelaPagar = vData
End Property

Public Property Get colParcelaPagar() As colParcelaPagar
    Set colParcelaPagar = mvarcolParcelaPagar
End Property

Public Property Let dtDataRegistro(ByVal vData As Date)
    mvardtDataRegistro = vData
End Property

Public Property Get dtDataRegistro() As Date
    dtDataRegistro = mvardtDataRegistro
End Property

Public Property Let iINSSRetido(ByVal vData As Integer)
    mvariINSSRetido = vData
End Property

Public Property Get iINSSRetido() As Integer
    iINSSRetido = mvariINSSRetido
End Property

Public Property Let iFilialPedCompra(ByVal vData As Integer)
    mvariFilialPedCompra = vData
End Property

Public Property Get iFilialPedCompra() As Integer
    iFilialPedCompra = mvariFilialPedCompra
End Property

Public Property Let lNumPedCompra(ByVal vData As Long)
    mvarlNumPedCompra = vData
End Property

Public Property Get lNumPedCompra() As Long
    lNumPedCompra = mvarlNumPedCompra
End Property





Public Property Let iCreditoICMS(ByVal vData As Integer)
    mvariCreditoICMS = vData
End Property

Public Property Get iCreditoICMS() As Integer
    iCreditoICMS = mvariCreditoICMS
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

Public Property Let iCreditoIPI(ByVal vData As Integer)
    mvariCreditoIPI = vData
End Property

Public Property Get iCreditoIPI() As Integer
    iCreditoIPI = mvariCreditoIPI
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
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

Public Property Let dValorINSS(ByVal vData As Double)
    mvardValorINSS = vData
End Property

Public Property Get dValorINSS() As Double
    dValorINSS = mvardValorINSS
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorICMSSubst(ByVal vData As Double)
    mvardValorICMSSubst = vData
End Property

Public Property Get dValorICMSSubst() As Double
    dValorICMSSubst = mvardValorICMSSubst
End Property

Public Property Let dValorICMS(ByVal vData As Double)
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
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

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dOutrasDespesas(ByVal vData As Double)
    mvardOutrasDespesas = vData
End Property

Public Property Get dOutrasDespesas() As Double
    dOutrasDespesas = mvardOutrasDespesas
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

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Private Sub Class_Initialize()
    mvardtDataEstorno = DATA_NULA
End Sub
