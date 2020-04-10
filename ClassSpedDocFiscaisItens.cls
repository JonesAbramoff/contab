VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSpedDocFiscaisItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntSped As Long
Private mvariTipoDoc As Integer
Private mvarlNumIntDoc As Long
Private mvariItem As Integer
Private mvarsRegistro As String
Private mvarsProduto As String
Private mvarsDescricao As String
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvarsMovFisica As String
Private mvarsICMSCST As String
Private mvarsCFOP As String
Private mvardICMSAliquota As Double
Private mvardICMSSTAliquota As Double
Private mvarsIPICST As String
Private mvarsIPICodEnq As String
Private mvardIPIAliquota As Double
Private mvarsPISCST As String
Private mvardPISAliquota As Double
Private mvardPISBCQtd As Double
Private mvardPISAliquotaQtd As Double
Private mvarsCofinsCST As String
Private mvardCofinsAliquota As Double
Private mvardCofinsBCQtd As Double
Private mvardCofinsAliquotaQtd As Double
Private mvardValorDesconto As Double
Private mvardValorFrete As Double
Private mvardValorSeguro As Double
Private mvardValorDespesas As Double
Private mvardValorServico As Double
Private mvardValorNaoTributado As Double
Private mvardValorBaseICMS As Double
Private mvardValorICMS As Double
Private mvardValorBaseICMSST As Double
Private mvardValorICMSST As Double
Private mvardValorBaseIPI As Double
Private mvardValorIPI As Double
Private mvardValorBasePIS As Double
Private mvardValorPIS As Double
Private mvardValorBaseCofins As Double
Private mvardValorCofins As Double
Private mvardValorPisRetido As Double
Private mvardValorPisST As Double
Private mvardValorCofinsRetido As Double
Private mvardValorCofinsST As Double
Private mvardValorISS As Double
Private mvardPrecoUnitario As Double
Private mvardPrecoTotal As Double
Private mvarsNatBCCred As String
Private mvardValorRecebido As Double
Private mvardValorNoCR As Double
Private mvardValorRecContrPrev As Double
Private mvardAliquotaContrPrev As Double
Private mvarsCodAtividadeTab511 As String
Private mvarsNCM As String

'EFD ICMS 2019
Public sCodISSServ As String
Public dValorBaseISS As Double
Public dAliquotaISS As Double

Public Property Let sNCM(ByVal vData As String)
    mvarsNCM = vData
End Property

Public Property Get sNCM() As String
    sNCM = mvarsNCM
End Property

Public Property Let sCodAtividadeTab511(ByVal vData As String)
    mvarsCodAtividadeTab511 = vData
End Property

Public Property Get sCodAtividadeTab511() As String
    sCodAtividadeTab511 = mvarsCodAtividadeTab511
End Property

Public Property Let dAliquotaContrPrev(ByVal vData As Double)
    mvardAliquotaContrPrev = vData
End Property

Public Property Get dAliquotaContrPrev() As Double
    dAliquotaContrPrev = mvardAliquotaContrPrev
End Property

Public Property Let dValorRecContrPrev(ByVal vData As Double)
    mvardValorRecContrPrev = vData
End Property

Public Property Get dValorRecContrPrev() As Double
    dValorRecContrPrev = mvardValorRecContrPrev
End Property

Public Property Let dValorNoCR(ByVal vData As Double)
    mvardValorNoCR = vData
End Property

Public Property Get dValorNoCR() As Double
    dValorNoCR = mvardValorNoCR
End Property

Public Property Let dValorRecebido(ByVal vData As Double)
    mvardValorRecebido = vData
End Property

Public Property Get dValorRecebido() As Double
    dValorRecebido = mvardValorRecebido
End Property

Public Property Let sNatBCCred(ByVal vData As String)
    mvarsNatBCCred = vData
End Property

Public Property Get sNatBCCred() As String
    sNatBCCred = mvarsNatBCCred
End Property

Public Property Let dPrecoTotal(ByVal vData As Double)
    mvardPrecoTotal = vData
End Property

Public Property Get dPrecoTotal() As Double
    dPrecoTotal = mvardPrecoTotal
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let lNumIntSped(ByVal vData As Long)
    mvarlNumIntSped = vData
End Property

Public Property Get lNumIntSped() As Long
    lNumIntSped = mvarlNumIntSped
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let sRegistro(ByVal vData As String)
    mvarsRegistro = vData
End Property

Public Property Get sRegistro() As String
    sRegistro = mvarsRegistro
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sMovFisica(ByVal vData As String)
    mvarsMovFisica = vData
End Property

Public Property Get sMovFisica() As String
    sMovFisica = mvarsMovFisica
End Property

Public Property Let sICMSCST(ByVal vData As String)
    mvarsICMSCST = vData
End Property

Public Property Get sICMSCST() As String
    sICMSCST = mvarsICMSCST
End Property

Public Property Let sCFOP(ByVal vData As String)
    mvarsCFOP = vData
End Property

Public Property Get sCFOP() As String
    sCFOP = mvarsCFOP
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
    mvardICMSAliquota = vData
End Property

Public Property Get dICMSAliquota() As Double
    dICMSAliquota = mvardICMSAliquota
End Property

Public Property Let dICMSSTAliquota(ByVal vData As Double)
    mvardICMSSTAliquota = vData
End Property

Public Property Get dICMSSTAliquota() As Double
    dICMSSTAliquota = mvardICMSSTAliquota
End Property

Public Property Let sIPICST(ByVal vData As String)
    mvarsIPICST = vData
End Property

Public Property Get sIPICST() As String
    sIPICST = mvarsIPICST
End Property

Public Property Let sIPICodEnq(ByVal vData As String)
    mvarsIPICodEnq = vData
End Property

Public Property Get sIPICodEnq() As String
    sIPICodEnq = mvarsIPICodEnq
End Property

Public Property Let dIPIAliquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIAliquota() As Double
    dIPIAliquota = mvardIPIAliquota
End Property

Public Property Let sPISCST(ByVal vData As String)
    mvarsPISCST = vData
End Property

Public Property Get sPISCST() As String
    sPISCST = mvarsPISCST
End Property

Public Property Let dPISAliquota(ByVal vData As Double)
    mvardPISAliquota = vData
End Property

Public Property Get dPISAliquota() As Double
    dPISAliquota = mvardPISAliquota
End Property

Public Property Let dPISBCQtd(ByVal vData As Double)
    mvardPISBCQtd = vData
End Property

Public Property Get dPISBCQtd() As Double
    dPISBCQtd = mvardPISBCQtd
End Property

Public Property Let dPISAliquotaQtd(ByVal vData As Double)
    mvardPISAliquotaQtd = vData
End Property

Public Property Get dPISAliquotaQtd() As Double
    dPISAliquotaQtd = mvardPISAliquotaQtd
End Property

Public Property Let sCofinsCST(ByVal vData As String)
    mvarsCofinsCST = vData
End Property

Public Property Get sCofinsCST() As String
    sCofinsCST = mvarsCofinsCST
End Property

Public Property Let dCOFINSAliquota(ByVal vData As Double)
    mvardCofinsAliquota = vData
End Property

Public Property Get dCOFINSAliquota() As Double
    dCOFINSAliquota = mvardCofinsAliquota
End Property

Public Property Let dCofinsBCQtd(ByVal vData As Double)
    mvardCofinsBCQtd = vData
End Property

Public Property Get dCofinsBCQtd() As Double
    dCofinsBCQtd = mvardCofinsBCQtd
End Property

Public Property Let dCofinsAliquotaQtd(ByVal vData As Double)
    mvardCofinsAliquotaQtd = vData
End Property

Public Property Get dCofinsAliquotaQtd() As Double
    dCofinsAliquotaQtd = mvardCofinsAliquotaQtd
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
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

Public Property Let dValorDespesas(ByVal vData As Double)
    mvardValorDespesas = vData
End Property

Public Property Get dValorDespesas() As Double
    dValorDespesas = mvardValorDespesas
End Property

Public Property Let dValorServico(ByVal vData As Double)
    mvardValorServico = vData
End Property

Public Property Get dValorServico() As Double
    dValorServico = mvardValorServico
End Property

Public Property Let dValorNaoTributado(ByVal vData As Double)
    mvardValorNaoTributado = vData
End Property

Public Property Get dValorNaoTributado() As Double
    dValorNaoTributado = mvardValorNaoTributado
End Property

Public Property Let dValorBaseICMS(ByVal vData As Double)
    mvardValorBaseICMS = vData
End Property

Public Property Get dValorBaseICMS() As Double
    dValorBaseICMS = mvardValorBaseICMS
End Property

Public Property Let dValorICMS(ByVal vData As Double)
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
End Property

Public Property Let dValorBaseICMSST(ByVal vData As Double)
    mvardValorBaseICMSST = vData
End Property

Public Property Get dValorBaseICMSST() As Double
    dValorBaseICMSST = mvardValorBaseICMSST
End Property

Public Property Let dValorICMSST(ByVal vData As Double)
    mvardValorICMSST = vData
End Property

Public Property Get dValorICMSST() As Double
    dValorICMSST = mvardValorICMSST
End Property

Public Property Let dValorBaseIPI(ByVal vData As Double)
    mvardValorBaseIPI = vData
End Property

Public Property Get dValorBaseIPI() As Double
    dValorBaseIPI = mvardValorBaseIPI
End Property

Public Property Let dValorIPI(ByVal vData As Double)
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorBasePIS(ByVal vData As Double)
    mvardValorBasePIS = vData
End Property

Public Property Get dValorBasePIS() As Double
    dValorBasePIS = mvardValorBasePIS
End Property

Public Property Let dValorPIS(ByVal vData As Double)
    mvardValorPIS = vData
End Property

Public Property Get dValorPIS() As Double
    dValorPIS = mvardValorPIS
End Property

Public Property Let dValorBaseCofins(ByVal vData As Double)
    mvardValorBaseCofins = vData
End Property

Public Property Get dValorBaseCofins() As Double
    dValorBaseCofins = mvardValorBaseCofins
End Property

Public Property Let dValorCofins(ByVal vData As Double)
    mvardValorCofins = vData
End Property

Public Property Get dValorCofins() As Double
    dValorCofins = mvardValorCofins
End Property

Public Property Let dValorPisRetido(ByVal vData As Double)
    mvardValorPisRetido = vData
End Property

Public Property Get dValorPisRetido() As Double
    dValorPisRetido = mvardValorPisRetido
End Property

Public Property Let dValorPisST(ByVal vData As Double)
    mvardValorPisST = vData
End Property

Public Property Get dValorPisST() As Double
    dValorPisST = mvardValorPisST
End Property

Public Property Let dValorCofinsRetido(ByVal vData As Double)
    mvardValorCofinsRetido = vData
End Property

Public Property Get dValorCofinsRetido() As Double
    dValorCofinsRetido = mvardValorCofinsRetido
End Property

Public Property Let dValorCofinsST(ByVal vData As Double)
    mvardValorCofinsST = vData
End Property

Public Property Get dValorCofinsST() As Double
    dValorCofinsST = mvardValorCofinsST
End Property

Public Property Let dValorISS(ByVal vData As Double)
    mvardValorISS = vData
End Property

Public Property Get dValorISS() As Double
    dValorISS = mvardValorISS
End Property
