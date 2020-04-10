VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNFsPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlFornecedor As Long
Private mvariFilial As Integer
Private mvarlNumNotaFiscal As Long
Private mvardtDataEmissao As Date
Private mvariStatus As Integer
Private mvarlNumIntTitPag As Long
Private mvardtDataVencimento As Date
Private mvardValorTotal As Double
Private mvardValorSeguro As Double
Private mvardValorFrete As Double
Private mvardOutrasDespesas As Double
Private mvardValorProdutos As Double
Private mvardValorIRRF As Double
Private mvardValorICMS As Double
Private mvardValorICMSSubst As Double
Private mvariCreditoICMS As Integer
Private mvardValorIPI As Double
Private mvariCreditoIPI As Integer
Private mvariExclusao As Integer 'indica que está ocorrendo a exclusao fisica de uma nota fiscal vinculada ao tituto
Private mvardPISRetido As Double
Private mvardCOFINSRetido As Double
Private mvardCSLLRetido As Double
Private mvardTaxaMoeda As Double
Private mvarsHistorico As String
Private mvardISSRetido As Double
Private mvardValorINSS As Double
Private mvariINSSRetido As Integer

'####################################################
'Inserido por Wagner
Public objInfoUsu As Object
'####################################################
Private mvarobjProjetoInfo As Object
Private mvardtDataEstorno As Date

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

Public Property Let dISSRetido(ByVal vData As Double)
    mvardISSRetido = vData
End Property

Public Property Get dISSRetido() As Double
    dISSRetido = mvardISSRetido
End Property

Public Property Let dtDataEstorno(ByVal vData As Date)
    mvardtDataEstorno = vData
End Property

Public Property Get dtDataEstorno() As Date
    dtDataEstorno = mvardtDataEstorno
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

Public Property Let iCreditoIPI(ByVal vData As Integer)
Attribute iCreditoIPI.VB_Description = "0  ou  1  -> determina se o valorIPI será um crédito de IPI"
    mvariCreditoIPI = vData
End Property

Public Property Get iCreditoIPI() As Integer
    iCreditoIPI = mvariCreditoIPI
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let iCreditoICMS(ByVal vData As Integer)
Attribute iCreditoICMS.VB_Description = "0  ou  1  -> determina se ValorICMS+ValorICMSSubst gerará um crédito de ICMS"
    mvariCreditoICMS = vData
End Property

Public Property Get iCreditoICMS() As Integer
    iCreditoICMS = mvariCreditoICMS
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
Attribute dValorIRRF.VB_Description = "IR fonte. Subtrai do ValorTotal para obter o valor líquido a pagar"
    mvardValorIRRF = vData
End Property

Public Property Get dValorIRRF() As Double
    dValorIRRF = mvardValorIRRF
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Let dOutrasDespesas(ByVal vData As Double)
    mvardOutrasDespesas = vData
End Property

Public Property Get dOutrasDespesas() As Double
    dOutrasDespesas = mvardOutrasDespesas
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorTotal(ByVal vData As Double)
Attribute dValorTotal.VB_Description = "Valor da Nota (não é o valor líquido a pagar que é igual a esse menos o ValorIRRF)"
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dtDataVencimento(ByVal vData As Date)
    mvardtDataVencimento = vData
End Property

Public Property Get dtDataVencimento() As Date
    dtDataVencimento = mvardtDataVencimento
End Property

Public Property Let lNumIntTitPag(ByVal vData As Long)
Attribute lNumIntTitPag.VB_Description = "No. da fatura (se estiver zerado a NF ainda nao esta associada a uma fatura)"
    mvarlNumIntTitPag = vData
End Property

Public Property Get lNumIntTitPag() As Long
    lNumIntTitPag = mvarlNumIntTitPag
End Property

Public Property Let iStatus(ByVal vData As Integer)
Attribute iStatus.VB_Description = "STATUS_SUSPENSO , STATUS_ABERTO , STATUS_EXCLUIDO"
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let lNumNotaFiscal(ByVal vData As Long)
    mvarlNumNotaFiscal = vData
End Property

Public Property Get lNumNotaFiscal() As Long
    lNumNotaFiscal = mvarlNumNotaFiscal
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lFornecedor(ByVal vData As Long)
    mvarlFornecedor = vData
End Property

Public Property Get lFornecedor() As Long
    lFornecedor = mvarlFornecedor
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Private Sub Class_Initialize()
    mvardtDataEstorno = DATA_NULA
End Sub
