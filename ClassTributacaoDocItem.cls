VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTributacaoDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

Private mvarobjDoc As Object

Public sCListServNFe As String

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvariTipoDoc As Integer
Private mvarlNumIntDocItem As Long
Private mvarlNumIntDoc As Long
Private mvariItem As Integer
Private mvarsProduto As String
Private mvarsProdutoDescricao As String
Private mvarsUnidadeMed As String
Private mvardQuantidade As Double
Private mvardPrecoTotal As Double
Private mvardPrecoUnitario As Double
Private mvardPrecoUnitarioMoeda As Double
Private mvardDescontoGrid As Double
Private mvarsNaturezaOp As String
Private mvariTipoTributacao As Integer
Private mvariIPITipo As Integer
Private mvarsIPICodProduto As String
Private mvardIPIBaseCalculo As Double
Private mvardIPIPercRedBase As Double
Private mvardIPIAliquota As Double
Private mvardIPIValor As Double
Private mvardIPICredito As Double
Private mvariICMSTipo As Integer
Private mvardICMSBase As Double
Private mvardICMSPercRedBase As Double
Private mvardICMSAliquota As Double
Private mvardICMSValor As Double
Private mvardICMSCredito As Double
Private mvardICMSSubstBase As Double
Private mvardICMSSubstAliquota As Double
Private mvardICMSSubstValor As Double
Private mvardPISCredito As Double
Private mvardCOFINSCredito As Double
Private mvardICMSAliquotaAdicaoDI As Double
Private mvardICMSPercRedBaseAdicaoDI As Double
Private mvardPISAliquotaAdicaoDI As Double
Private mvardCOFINSAliquotaAdicaoDI As Double
Private mvardDespImpICMSBase As Double
Private mvardDespImpICMSValor As Double
Private mvardDespImpICMSCredito As Double
Private mvardICMSSubstPercRedBase As Double
Private mvardICMSSubstPercMVA As Double
Private mvariPISTipo As Integer
Private mvariPISTipoCalculo As Integer
Private mvardPISBase As Double
Private mvardPISAliquota As Double
Private mvardPISAliquotaValor As Double
Private mvardPISQtde As Double
Private mvardPISValor As Double
Private mvariPISSTTipoCalculo As Integer
Private mvardPISSTBase As Double
Private mvardPISSTAliquota As Double
Private mvardPISSTAliquotaValor As Double
Private mvardPISSTQtde As Double
Private mvardPISSTValor As Double
Private mvariCOFINSTipo As Integer
Private mvariCOFINSTipoCalculo As Integer
Private mvardCOFINSBase As Double
Private mvardCOFINSAliquota As Double
Private mvardCOFINSAliquotaValor As Double
Private mvardCOFINSQtde As Double
Private mvardCOFINSValor As Double
Private mvariCOFINSSTTipoCalculo As Integer
Private mvardCOFINSSTBase As Double
Private mvardCOFINSSTAliquota As Double
Private mvardCOFINSSTAliquotaValor As Double
Private mvardCOFINSSTQtde As Double
Private mvardCOFINSSTValor As Double
Private mvarsCST As String
Private mvarsISSQN As String
Private mvardISSBase As Double
Private mvardISSAliquota As Double
Private mvarsISSCidadeIBGE As String
Private mvardISSValor As Double
Private mvardIIBase As Double
Private mvardIIDespAduaneira As Double
Private mvardIIIOF As Double
Private mvardIIValor As Double
Private mvariICMSBaseModalidade As Integer
Private mvariICMSSubstBaseModalidade As Integer
Private mvarsIPIEnquadramentoClasse As String
Private mvarsIPIEnquadramentoCodigo As String
Private mvarsIPICNPJProdutor As String
Private mvarsIPISeloCodigo As String
Private mvarlIPISeloQtde As Long
Private mvariIPITipoCalculo As Integer
Private mvardIPIUnidadePadraoQtde As Double
Private mvardIPIUnidadePadraoValor As Double
Private mvardValorFreteItem As Double
Private mvardValorSeguroItem As Double
Private mvardValorOutrasDespesasItem As Double
Private mvardValorDescontoItem As Double
Private mvariOrigemMercadoria As Integer
Private mvarsFCI As String
Private mvariExTIPI As Integer
Private mvarsGenero As String
Private mvariProdutoEspecifico As Integer
Private mvarsEAN As String
Private mvarsEANTrib As String
Private mvardQtdTrib As Double
Private mvarsUMTrib As String
Private mvardValorUnitTrib As Double
Private mvariNaturezaOpManual As Integer
Private mvariTipoTributacaoManual As Integer
Private mvariIPITipoManual As Integer
Private mvariIPICodProdutoManual As Integer
Private mvariIPIBaseCalculoManual As Integer
Private mvariIPIPercRedBaseManual As Integer
Private mvariIPIAliquotaManual As Integer
Private mvariIPIValorManual As Integer
Private mvariIPICreditoManual As Integer
Private mvariICMSTipoManual As Integer
Private mvariICMSBaseManual As Integer
Private mvariICMSPercRedBaseManual As Integer
Private mvariICMSAliquotaManual As Integer
Private mvariICMSValorManual As Integer
Private mvariICMSCreditoManual As Integer
Private mvariICMSSubstBaseManual As Integer
Private mvariICMSSubstAliquotaManual As Integer
Private mvariICMSSubstValorManual As Integer
Private mvariPISCreditoManual As Integer
Private mvariCOFINSCreditoManual As Integer
Private mvariICMSAliquotaAdicaoDIManual As Integer
Private mvariICMSPercRedBaseAdicaoDIManual As Integer
Private mvariPISAliquotaAdicaoDIManual As Integer
Private mvariCOFINSAliquotaAdicaoDIManual As Integer
Private mvariDespImpICMSBaseManual As Integer
Private mvariDespImpICMSValorManual As Integer
Private mvariDespImpICMSCreditoManual As Integer
Private mvariICMSSubstPercRedBaseManual As Integer
Private mvariICMSSubstPercMVAManual As Integer
Private mvariPISTipoManual As Integer
Private mvariPISTipoCalculoManual As Integer
Private mvariPISBaseManual As Integer
Private mvariPISAliquotaManual As Integer
Private mvariPISAliquotaValorManual As Integer
Private mvariPISQtdeManual As Integer
Private mvariPISValorManual As Integer
Private mvariPISSTTipoCalculoManual As Integer
Private mvariPISSTBaseManual As Integer
Private mvariPISSTAliquotaManual As Integer
Private mvariPISSTAliquotaValorManual As Integer
Private mvariPISSTQtdeManual As Integer
Private mvariPISSTValorManual As Integer
Private mvariCOFINSTipoManual As Integer
Private mvariCOFINSTipoCalculoManual As Integer
Private mvariCOFINSBaseManual As Integer
Private mvariCOFINSAliquotaManual As Integer
Private mvariCOFINSAliquotaValorManual As Integer
Private mvariCOFINSQtdeManual As Integer
Private mvariCOFINSValorManual As Integer
Private mvariCOFINSSTTipoCalculoManual As Integer
Private mvariCOFINSSTBaseManual As Integer
Private mvariCOFINSSTAliquotaManual As Integer
Private mvariCOFINSSTAliquotaValorManual As Integer
Private mvariCOFINSSTQtdeManual As Integer
Private mvariCOFINSSTValorManual As Integer
Private mvariCSTManual As Integer
Private mvariISSQNManual As Integer
Private mvariISSBaseManual As Integer
Private mvariISSAliquotaManual As Integer
Private mvariISSCidadeIBGEManual As Integer
Private mvariISSValorManual As Integer
Private mvariIIBaseManual As Integer
Private mvariIIDespAduaneiraManual As Integer
Private mvariIIIOFManual As Integer
Private mvariIIValorManual As Integer
Private mvariICMSBaseModalidadeManual As Integer
Private mvariICMSSubstBaseModalidadeManual As Integer
Private mvariIPIEnquadramentoClasseManual As Integer
Private mvariIPIEnquadramentoCodigoManual As Integer
Private mvariIPICNPJProdutorManual As Integer
Private mvariIPISeloCodigoManual As Integer
Private mvariIPISeloQtdeManual As Integer
Private mvariIPITipoCalculoManual As Integer
Private mvariIPIUnidadePadraoQtdeManual As Integer
Private mvariIPIUnidadePadraoValorManual As Integer
Private mvariValorFreteItemManual As Integer
Private mvariValorSeguroItemManual As Integer
Private mvariValorOutrasDespesasItemManual As Integer
Private mvariValorDescontoItemManual As Integer
Private mvariOrigemMercadoriaManual As Integer
Private mvariFCIManual As Integer
Private mvardValorAux As Double

Private mvardICMSSTCobrAntBase As Double
Private mvariICMSSTCobrAntBaseManual As Integer
Private mvardICMSSTCobrAntValor As Double
Private mvariICMSSTCobrAntValorManual As Integer

Private mvardICMSpercBaseOperacaoPropria As Double
Private mvariICMSpercBaseOperacaoPropriaManual As Integer
Private mvarsICMSUFDevidoST As String
Private mvariICMSUFDevidoSTManual As Integer
Private mvardICMSvBCSTRet As Double
Private mvariICMSvBCSTRetManual As Integer
Private mvardICMSvICMSSTRet As Double
Private mvariICMSvICMSSTRetManual As Integer
Private mvardICMSvBCSTDest As Double
Private mvariICMSvBCSTDestManual As Integer
Private mvardICMSvICMSSTDest As Double
Private mvariICMSvICMSSTDestManual As Integer
Private mvardICMSpCredSN As Double
Private mvariICMSpCredSNManual As Integer
Private mvardICMSvCredSN As Double
Private mvariICMSvCredSNManual As Integer
Private mvardICMSValorIsento As Double
Private mvariICMSValorIsentoManual As Integer
Private mvariICMSMotivo As Integer
Private mvariICMSMotivoManual As Integer
Private mvariICMSSimplesTipo As Integer
Private mvariICMSSimplesTipoManual As Integer
Private mvariISSTipo As Integer
Private mvariISSTipoManual As Integer
Private mvariRegimeTributario As Integer
Private mvarsISSCST As String
Private mvarsCSOSN As String
Private mvarsNatBCCred As String
Private mvariNatBCCredManual As Integer
Private mvariRegimeTributarioManual As Integer
Private mvariTotTribManual As Integer
Private mvardTotTrib As Double
Private mvariTotTribTipo As Integer
Private mvariTotTribTipoManual As Integer

'nfe 3.10
Public dICMS51ValorOp As Double
Public dICMSPercDifer As Double
Public dICMSValorDif As Double
Public iICMS51ValorOpManual As Integer
Public iICMSPercDiferManual As Integer
Public iICMSValorDifManual As Integer
Public dISSValorDeducao As Double
Public dISSValorOutrasRet As Double
Public dISSValorDescIncond As Double
Public dISSValorDescCond As Double
Public dISSValorRet As Double
Public iISSIndExigibilidade As Integer
Public sISSCodServ As String
Public sISSMunicIncidImp As String
Public iISSCodPais As Integer
Public sISSNumProcesso As String
Public iISSIndIncentivo As Integer
Public iISSValorDeducaoManual As Integer
Public iISSValorOutrasRetManual As Integer
Public iISSValorDescIncondManual As Integer
Public iISSValorDescCondManual As Integer
Public iISSValorRetManual As Integer
Public iISSIndExigibilidadeManual As Integer
Public iISSCodServManual As Integer
Public iISSMunicIncidImpManual As Integer
Public iISSCodPaisManual As Integer
Public iISSNumProcessoManual As Integer
Public iISSIndIncentivoManual As Integer

'IBPTax 0.0.9
Public dTotTribAliqFed As Double
Public dTotTribAliqEst As Double
Public dTotTribAliqMunic As Double
Public iTotTribAliqFedManual As Integer
Public iTotTribAliqEstManual As Integer
Public iTotTribAliqMunicManual As Integer

'NT 2015.003 (ICMS Interestadual)
Public sCEST As String
Public dICMSInterestBCUFDest As Double
Public iICMSInterestBCUFDestManual As Integer
Public dICMSInterestPercFCPUFDest As Double
Public iICMSInterestPercFCPUFDestManual As Integer
Public dICMSInterestAliqUFDest As Double
Public iICMSInterestAliqUFDestManual As Integer
Public dICMSInterestAliq As Double
Public iICMSInterestAliqManual As Integer
Public dICMSInterestPercPartilha As Double
Public iICMSInterestPercPartilhaManual As Integer
Public dICMSInterestVlrFCPUFDest As Double
Public iICMSInterestVlrFCPUFDestManual As Integer
Public dICMSInterestVlrUFDest As Double
Public iICMSInterestVlrUFDestManual As Integer
Public dICMSInterestVlrUFRemet As Double
Public iICMSInterestVlrUFRemetManual As Integer

'UA. Tributos Devolvidos (para o item da NF-e)
Public dpDevol As Double
Public ipDevolManual As Integer
Public dIPIVlrDevolvido As Double
Public iIPIVlrDevolvidoManual As Integer

'nfe 4.00
Public dICMSInterestBCFCPUFDest As Double
Public iICMSInterestBCFCPUFDestManual As Integer
Public dICMSvBCFCPST As Double
Public iICMSvBCFCPSTManual As Integer
Public dICMSpFCPST As Double
Public iICMSpFCPSTManual As Integer
Public dICMSvFCPST As Double
Public iICMSvFCPSTManual As Integer
Public dICMSvBCFCPSTRet As Double
Public iICMSvBCFCPSTRetManual As Integer
Public dICMSpFCPSTRet As Double
Public iICMSpFCPSTRetManual As Integer
Public dICMSvFCPSTRet As Double
Public iICMSvFCPSTRetManual As Integer
Public dICMSSTCobrAntAliquota As Double
Public iICMSSTCobrAntAliquotaManual As Integer
Public dICMSvBCFCP As Double
Public iICMSvBCFCPManual As Integer
Public dICMSpFCP As Double
Public iICMSpFCPManual As Integer
Public dICMSvFCP As Double
Public iICMSvFCPManual As Integer
Public scBenef As String 'Sempre Manual então ficará sem o flag
Public sindEscala As String 'Sempre Automático então ficará sem o flag
Public sCNPJFab As String 'Sempre Automático então ficará sem o flag

'EC 52/2017
Public iICMSSTBaseDupla As Integer
Public iICMSSTBaseDuplaManual As Integer

'EC 87/2015 http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/orientacao/orientacao_002_2016.pdf
Public iICMSInterestBaseDupla As Integer
Public iICMSInterestBaseDuplaManual As Integer
Public icBenefManual As Integer

Public Property Let iTotTribTipo(ByVal vData As Integer)
    mvariTotTribTipo = vData
End Property

Public Property Get iTotTribTipo() As Integer
    iTotTribTipo = mvariTotTribTipo
End Property

Public Property Let iTotTribTipoManual(ByVal vData As Integer)
    mvariTotTribTipoManual = vData
End Property

Public Property Get iTotTribTipoManual() As Integer
    iTotTribTipoManual = mvariTotTribTipoManual
End Property

Public Property Let dTotTrib(ByVal vData As Double)
    mvardTotTrib = vData
End Property

Public Property Get dTotTrib() As Double
    dTotTrib = mvardTotTrib
End Property

Public Property Let iTotTribManual(ByVal vData As Integer)
    mvariTotTribManual = vData
End Property

Public Property Get iTotTribManual() As Integer
    iTotTribManual = mvariTotTribManual
End Property

Public Property Let iNatBCCredManual(ByVal vData As Integer)
    mvariNatBCCredManual = vData
End Property

Public Property Get iNatBCCredManual() As Integer
    iNatBCCredManual = mvariNatBCCredManual
End Property

Public Property Let sNatBCCred(ByVal vData As String)
    mvarsNatBCCred = vData
End Property

Public Property Get sNatBCCred() As String
    sNatBCCred = mvarsNatBCCred
End Property

Public Property Let sCSOSN(ByVal vData As String)
    mvarsCSOSN = vData
End Property

Public Property Get sCSOSN() As String
    sCSOSN = mvarsCSOSN
End Property

Public Property Let sISSCST(ByVal vData As String)
    mvarsISSCST = vData
End Property

Public Property Get sISSCST() As String
    sISSCST = mvarsISSCST
End Property

Public Property Let iICMSpercBaseOperacaoPropriaManual(ByVal vData As Integer)
    mvariICMSpercBaseOperacaoPropriaManual = vData
End Property

Public Property Get iICMSpercBaseOperacaoPropriaManual() As Integer
    iICMSpercBaseOperacaoPropriaManual = mvariICMSpercBaseOperacaoPropriaManual
End Property

Public Property Let dICMSpercBaseOperacaoPropria(ByVal vData As Double)
    mvardICMSpercBaseOperacaoPropria = vData
End Property

Public Property Get dICMSpercBaseOperacaoPropria() As Double
    dICMSpercBaseOperacaoPropria = mvardICMSpercBaseOperacaoPropria
End Property

Public Property Let iICMSUFDevidoSTManual(ByVal vData As Integer)
    mvariICMSUFDevidoSTManual = vData
End Property

Public Property Get iICMSUFDevidoSTManual() As Integer
    iICMSUFDevidoSTManual = mvariICMSUFDevidoSTManual
End Property

Public Property Let sICMSUFDevidoST(ByVal vData As String)
    mvarsICMSUFDevidoST = vData
End Property

Public Property Get sICMSUFDevidoST() As String
    sICMSUFDevidoST = mvarsICMSUFDevidoST
End Property

Public Property Let iICMSvBCSTRetManual(ByVal vData As Integer)
    mvariICMSvBCSTRetManual = vData
End Property

Public Property Get iICMSvBCSTRetManual() As Integer
    iICMSvBCSTRetManual = mvariICMSvBCSTRetManual
End Property

Public Property Let dICMSvBCSTRet(ByVal vData As Double)
    mvardICMSvBCSTRet = vData
End Property

Public Property Get dICMSvBCSTRet() As Double
    dICMSvBCSTRet = mvardICMSvBCSTRet
End Property

Public Property Let iICMSvICMSSTRetManual(ByVal vData As Integer)
    mvariICMSvICMSSTRetManual = vData
End Property

Public Property Get iICMSvICMSSTRetManual() As Integer
    iICMSvICMSSTRetManual = mvariICMSvICMSSTRetManual
End Property

Public Property Let dICMSvICMSSTRet(ByVal vData As Double)
    mvardICMSvICMSSTRet = vData
End Property

Public Property Get dICMSvICMSSTRet() As Double
    dICMSvICMSSTRet = mvardICMSvICMSSTRet
End Property

Public Property Let iICMSvBCSTDestManual(ByVal vData As Integer)
    mvariICMSvBCSTDestManual = vData
End Property

Public Property Get iICMSvBCSTDestManual() As Integer
    iICMSvBCSTDestManual = mvariICMSvBCSTDestManual
End Property

Public Property Let dICMSvBCSTDest(ByVal vData As Double)
    mvardICMSvBCSTDest = vData
End Property

Public Property Get dICMSvBCSTDest() As Double
    dICMSvBCSTDest = mvardICMSvBCSTDest
End Property

Public Property Let iICMSvICMSSTDestManual(ByVal vData As Integer)
    mvariICMSvICMSSTDestManual = vData
End Property

Public Property Get iICMSvICMSSTDestManual() As Integer
    iICMSvICMSSTDestManual = mvariICMSvICMSSTDestManual
End Property

Public Property Let dICMSvICMSSTDest(ByVal vData As Double)
    mvardICMSvICMSSTDest = vData
End Property

Public Property Get dICMSvICMSSTDest() As Double
    dICMSvICMSSTDest = mvardICMSvICMSSTDest
End Property

Public Property Let iICMSpCredSNManual(ByVal vData As Integer)
    mvariICMSpCredSNManual = vData
End Property

Public Property Get iICMSpCredSNManual() As Integer
    iICMSpCredSNManual = mvariICMSpCredSNManual
End Property

Public Property Let dICMSpCredSN(ByVal vData As Double)
    mvardICMSpCredSN = vData
End Property

Public Property Get dICMSpCredSN() As Double
    dICMSpCredSN = mvardICMSpCredSN
End Property

Public Property Let iICMSvCredSNManual(ByVal vData As Integer)
    mvariICMSvCredSNManual = vData
End Property

Public Property Get iICMSvCredSNManual() As Integer
    iICMSvCredSNManual = mvariICMSvCredSNManual
End Property

Public Property Let dICMSvCredSN(ByVal vData As Double)
    mvardICMSvCredSN = vData
End Property

Public Property Get dICMSvCredSN() As Double
    dICMSvCredSN = mvardICMSvCredSN
End Property

Public Property Let iICMSValorIsentoManual(ByVal vData As Integer)
    mvariICMSValorIsentoManual = vData
End Property

Public Property Get iICMSValorIsentoManual() As Integer
    iICMSValorIsentoManual = mvariICMSValorIsentoManual
End Property

Public Property Let dICMSValorIsento(ByVal vData As Double)
    mvardICMSValorIsento = vData
End Property

Public Property Get dICMSValorIsento() As Double
    dICMSValorIsento = mvardICMSValorIsento
End Property

Public Property Let iICMSMotivoManual(ByVal vData As Integer)
    mvariICMSMotivoManual = vData
End Property

Public Property Get iICMSMotivoManual() As Integer
    iICMSMotivoManual = mvariICMSMotivoManual
End Property

Public Property Let iICMSMotivo(ByVal vData As Integer)
    mvariICMSMotivo = vData
End Property

Public Property Get iICMSMotivo() As Integer
    iICMSMotivo = mvariICMSMotivo
End Property

Public Property Let iICMSSimplesTipoManual(ByVal vData As Integer)
    mvariICMSSimplesTipoManual = vData
End Property

Public Property Get iICMSSimplesTipoManual() As Integer
    iICMSSimplesTipoManual = mvariICMSSimplesTipoManual
End Property

Public Property Let iICMSSimplesTipo(ByVal vData As Integer)
    mvariICMSSimplesTipo = vData
End Property

Public Property Get iICMSSimplesTipo() As Integer
    iICMSSimplesTipo = mvariICMSSimplesTipo
End Property

Public Property Let iISSTipoManual(ByVal vData As Integer)
    mvariISSTipoManual = vData
End Property

Public Property Get iISSTipoManual() As Integer
    iISSTipoManual = mvariISSTipoManual
End Property

Public Property Let iISSTipo(ByVal vData As Integer)
    mvariISSTipo = vData
End Property

Public Property Get iISSTipo() As Integer
    iISSTipo = mvariISSTipo
End Property

Public Property Let iRegimeTributario(ByVal vData As Integer)
    mvariRegimeTributario = vData
End Property

Public Property Get iRegimeTributario() As Integer
    iRegimeTributario = mvariRegimeTributario
End Property

Public Property Let iRegimeTributarioManual(ByVal vData As Integer)
    mvariRegimeTributarioManual = vData
End Property

Public Property Get iRegimeTributarioManual() As Integer
    iRegimeTributarioManual = mvariRegimeTributarioManual
End Property

Public Property Let iICMSSTCobrAntValorManual(ByVal vData As Integer)
    mvariICMSSTCobrAntValorManual = vData
End Property

Public Property Get iICMSSTCobrAntValorManual() As Integer
    iICMSSTCobrAntValorManual = mvariICMSSTCobrAntValorManual
End Property

Public Property Let dICMSSTCobrAntValor(ByVal vData As Double)
    mvardICMSSTCobrAntValor = vData
End Property

Public Property Get dICMSSTCobrAntValor() As Double
    dICMSSTCobrAntValor = mvardICMSSTCobrAntValor
End Property

Public Property Let iICMSSTCobrAntBaseManual(ByVal vData As Integer)
    mvariICMSSTCobrAntBaseManual = vData
End Property

Public Property Get iICMSSTCobrAntBaseManual() As Integer
    iICMSSTCobrAntBaseManual = mvariICMSSTCobrAntBaseManual
End Property

Public Property Let dICMSSTCobrAntBase(ByVal vData As Double)
    mvardICMSSTCobrAntBase = vData
End Property

Public Property Get dICMSSTCobrAntBase() As Double
    dICMSSTCobrAntBase = mvardICMSSTCobrAntBase
End Property

Public Property Let dValorAux(ByVal vData As Double)
    mvardValorAux = vData
End Property

Public Property Get dValorAux() As Double
    dValorAux = mvardValorAux
End Property

Public Property Let iTipoDoc(ByVal vData As Integer)
    mvariTipoDoc = vData
End Property

Public Property Get iTipoDoc() As Integer
    iTipoDoc = mvariTipoDoc
End Property

Public Property Get iTipo() As Integer
    iTipo = 0 'ITEM_TIPO_NORMAL
End Property

Public Property Let lNumIntDocItem(ByVal vData As Long)
    mvarlNumIntDocItem = vData
End Property

Public Property Get lNumIntDocItem() As Long
    lNumIntDocItem = mvarlNumIntDocItem
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

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sProdutoDescricao(ByVal vData As String)
    mvarsProdutoDescricao = vData
End Property

Public Property Get sProdutoDescricao() As String
    sProdutoDescricao = mvarsProdutoDescricao
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
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

Public Property Let dPrecoUnitarioMoeda(ByVal vData As Double)
    mvardPrecoUnitarioMoeda = vData
End Property

Public Property Get dPrecoUnitarioMoeda() As Double
    dPrecoUnitarioMoeda = mvardPrecoUnitarioMoeda
End Property

Public Property Let dDescontoGrid(ByVal vData As Double)
    mvardDescontoGrid = vData
End Property

Public Property Get dDescontoGrid() As Double
    dDescontoGrid = mvardDescontoGrid
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let iIPITipo(ByVal vData As Integer)
    mvariIPITipo = vData
End Property

Public Property Get iIPITipo() As Integer
    iIPITipo = mvariIPITipo
End Property

Public Property Let sIPICodProduto(ByVal vData As String)
    mvarsIPICodProduto = vData
End Property

Public Property Get sIPICodProduto() As String
    sIPICodProduto = mvarsIPICodProduto
End Property

Public Property Let dIPIBaseCalculo(ByVal vData As Double)
    mvardIPIBaseCalculo = vData
End Property

Public Property Get dIPIBaseCalculo() As Double
    dIPIBaseCalculo = mvardIPIBaseCalculo
End Property

Public Property Let dIPIPercRedBase(ByVal vData As Double)
    mvardIPIPercRedBase = vData
End Property

Public Property Get dIPIPercRedBase() As Double
    dIPIPercRedBase = mvardIPIPercRedBase
End Property

Public Property Let dIPIAliquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIAliquota() As Double
    dIPIAliquota = mvardIPIAliquota
End Property

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let dIPICredito(ByVal vData As Double)
    mvardIPICredito = vData
End Property

Public Property Get dIPICredito() As Double
    dIPICredito = mvardIPICredito
End Property

Public Property Let iICMSTipo(ByVal vData As Integer)
    mvariICMSTipo = vData
End Property

Public Property Get iICMSTipo() As Integer
    iICMSTipo = mvariICMSTipo
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let dICMSPercRedBase(ByVal vData As Double)
    mvardICMSPercRedBase = vData
End Property

Public Property Get dICMSPercRedBase() As Double
    dICMSPercRedBase = mvardICMSPercRedBase
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
    mvardICMSAliquota = vData
End Property

Public Property Get dICMSAliquota() As Double
    dICMSAliquota = mvardICMSAliquota
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let dICMSCredito(ByVal vData As Double)
    mvardICMSCredito = vData
End Property

Public Property Get dICMSCredito() As Double
    dICMSCredito = mvardICMSCredito
End Property

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
End Property

Public Property Let dICMSSubstAliquota(ByVal vData As Double)
    mvardICMSSubstAliquota = vData
End Property

Public Property Get dICMSSubstAliquota() As Double
    dICMSSubstAliquota = mvardICMSSubstAliquota
End Property

Public Property Let dICMSSubstValor(ByVal vData As Double)
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let dPISCredito(ByVal vData As Double)
    mvardPISCredito = vData
End Property

Public Property Get dPISCredito() As Double
    dPISCredito = mvardPISCredito
End Property

Public Property Let dCOFINSCredito(ByVal vData As Double)
    mvardCOFINSCredito = vData
End Property

Public Property Get dCOFINSCredito() As Double
    dCOFINSCredito = mvardCOFINSCredito
End Property

Public Property Let dICMSAliquotaAdicaoDI(ByVal vData As Double)
    mvardICMSAliquotaAdicaoDI = vData
End Property

Public Property Get dICMSAliquotaAdicaoDI() As Double
    dICMSAliquotaAdicaoDI = mvardICMSAliquotaAdicaoDI
End Property

Public Property Let dICMSPercRedBaseAdicaoDI(ByVal vData As Double)
    mvardICMSPercRedBaseAdicaoDI = vData
End Property

Public Property Get dICMSPercRedBaseAdicaoDI() As Double
    dICMSPercRedBaseAdicaoDI = mvardICMSPercRedBaseAdicaoDI
End Property

Public Property Let dPISAliquotaAdicaoDI(ByVal vData As Double)
    mvardPISAliquotaAdicaoDI = vData
End Property

Public Property Get dPISAliquotaAdicaoDI() As Double
    dPISAliquotaAdicaoDI = mvardPISAliquotaAdicaoDI
End Property

Public Property Let dCOFINSAliquotaAdicaoDI(ByVal vData As Double)
    mvardCOFINSAliquotaAdicaoDI = vData
End Property

Public Property Get dCOFINSAliquotaAdicaoDI() As Double
    dCOFINSAliquotaAdicaoDI = mvardCOFINSAliquotaAdicaoDI
End Property

Public Property Let dDespImpICMSBase(ByVal vData As Double)
    mvardDespImpICMSBase = vData
End Property

Public Property Get dDespImpICMSBase() As Double
    dDespImpICMSBase = mvardDespImpICMSBase
End Property

Public Property Let dDespImpICMSValor(ByVal vData As Double)
    mvardDespImpICMSValor = vData
End Property

Public Property Get dDespImpICMSValor() As Double
    dDespImpICMSValor = mvardDespImpICMSValor
End Property

Public Property Let dDespImpICMSCredito(ByVal vData As Double)
    mvardDespImpICMSCredito = vData
End Property

Public Property Get dDespImpICMSCredito() As Double
    dDespImpICMSCredito = mvardDespImpICMSCredito
End Property

Public Property Let dICMSSubstPercRedBase(ByVal vData As Double)
    mvardICMSSubstPercRedBase = vData
End Property

Public Property Get dICMSSubstPercRedBase() As Double
    dICMSSubstPercRedBase = mvardICMSSubstPercRedBase
End Property

Public Property Let dICMSSubstPercMVA(ByVal vData As Double)
    mvardICMSSubstPercMVA = vData
End Property

Public Property Get dICMSSubstPercMVA() As Double
    dICMSSubstPercMVA = mvardICMSSubstPercMVA
End Property

Public Property Let iPISTipo(ByVal vData As Integer)
    mvariPISTipo = vData
End Property

Public Property Get iPISTipo() As Integer
    iPISTipo = mvariPISTipo
End Property

Public Property Let iPISTipoCalculo(ByVal vData As Integer)
    mvariPISTipoCalculo = vData
End Property

Public Property Get iPISTipoCalculo() As Integer
    iPISTipoCalculo = mvariPISTipoCalculo
End Property

Public Property Let dPISBase(ByVal vData As Double)
    mvardPISBase = vData
End Property

Public Property Get dPISBase() As Double
    dPISBase = mvardPISBase
End Property

Public Property Let dPISAliquota(ByVal vData As Double)
    mvardPISAliquota = vData
End Property

Public Property Get dPISAliquota() As Double
    dPISAliquota = mvardPISAliquota
End Property

Public Property Let dPISAliquotaValor(ByVal vData As Double)
    mvardPISAliquotaValor = vData
End Property

Public Property Get dPISAliquotaValor() As Double
    dPISAliquotaValor = mvardPISAliquotaValor
End Property

Public Property Let dPISQtde(ByVal vData As Double)
    mvardPISQtde = vData
End Property

Public Property Get dPISQtde() As Double
    dPISQtde = mvardPISQtde
End Property

Public Property Let dPISValor(ByVal vData As Double)
    mvardPISValor = vData
End Property

Public Property Get dPISValor() As Double
    dPISValor = mvardPISValor
End Property

Public Property Let iPISSTTipoCalculo(ByVal vData As Integer)
    mvariPISSTTipoCalculo = vData
End Property

Public Property Get iPISSTTipoCalculo() As Integer
    iPISSTTipoCalculo = mvariPISSTTipoCalculo
End Property

Public Property Let dPISSTBase(ByVal vData As Double)
    mvardPISSTBase = vData
End Property

Public Property Get dPISSTBase() As Double
    dPISSTBase = mvardPISSTBase
End Property

Public Property Let dPISSTAliquota(ByVal vData As Double)
    mvardPISSTAliquota = vData
End Property

Public Property Get dPISSTAliquota() As Double
    dPISSTAliquota = mvardPISSTAliquota
End Property

Public Property Let dPISSTAliquotaValor(ByVal vData As Double)
    mvardPISSTAliquotaValor = vData
End Property

Public Property Get dPISSTAliquotaValor() As Double
    dPISSTAliquotaValor = mvardPISSTAliquotaValor
End Property

Public Property Let dPISSTQtde(ByVal vData As Double)
    mvardPISSTQtde = vData
End Property

Public Property Get dPISSTQtde() As Double
    dPISSTQtde = mvardPISSTQtde
End Property

Public Property Let dPISSTValor(ByVal vData As Double)
    mvardPISSTValor = vData
End Property

Public Property Get dPISSTValor() As Double
    dPISSTValor = mvardPISSTValor
End Property

Public Property Let iCOFINSTipo(ByVal vData As Integer)
    mvariCOFINSTipo = vData
End Property

Public Property Get iCOFINSTipo() As Integer
    iCOFINSTipo = mvariCOFINSTipo
End Property

Public Property Let iCOFINSTipoCalculo(ByVal vData As Integer)
    mvariCOFINSTipoCalculo = vData
End Property

Public Property Get iCOFINSTipoCalculo() As Integer
    iCOFINSTipoCalculo = mvariCOFINSTipoCalculo
End Property

Public Property Let dCOFINSBase(ByVal vData As Double)
    mvardCOFINSBase = vData
End Property

Public Property Get dCOFINSBase() As Double
    dCOFINSBase = mvardCOFINSBase
End Property

Public Property Let dCOFINSAliquota(ByVal vData As Double)
    mvardCOFINSAliquota = vData
End Property

Public Property Get dCOFINSAliquota() As Double
    dCOFINSAliquota = mvardCOFINSAliquota
End Property

Public Property Let dCOFINSAliquotaValor(ByVal vData As Double)
    mvardCOFINSAliquotaValor = vData
End Property

Public Property Get dCOFINSAliquotaValor() As Double
    dCOFINSAliquotaValor = mvardCOFINSAliquotaValor
End Property

Public Property Let dCOFINSQtde(ByVal vData As Double)
    mvardCOFINSQtde = vData
End Property

Public Property Get dCOFINSQtde() As Double
    dCOFINSQtde = mvardCOFINSQtde
End Property

Public Property Let dCOFINSValor(ByVal vData As Double)
    mvardCOFINSValor = vData
End Property

Public Property Get dCOFINSValor() As Double
    dCOFINSValor = mvardCOFINSValor
End Property

Public Property Let iCOFINSSTTipoCalculo(ByVal vData As Integer)
    mvariCOFINSSTTipoCalculo = vData
End Property

Public Property Get iCOFINSSTTipoCalculo() As Integer
    iCOFINSSTTipoCalculo = mvariCOFINSSTTipoCalculo
End Property

Public Property Let dCOFINSSTBase(ByVal vData As Double)
    mvardCOFINSSTBase = vData
End Property

Public Property Get dCOFINSSTBase() As Double
    dCOFINSSTBase = mvardCOFINSSTBase
End Property

Public Property Let dCOFINSSTAliquota(ByVal vData As Double)
    mvardCOFINSSTAliquota = vData
End Property

Public Property Get dCOFINSSTAliquota() As Double
    dCOFINSSTAliquota = mvardCOFINSSTAliquota
End Property

Public Property Let dCOFINSSTAliquotaValor(ByVal vData As Double)
    mvardCOFINSSTAliquotaValor = vData
End Property

Public Property Get dCOFINSSTAliquotaValor() As Double
    dCOFINSSTAliquotaValor = mvardCOFINSSTAliquotaValor
End Property

Public Property Let dCOFINSSTQtde(ByVal vData As Double)
    mvardCOFINSSTQtde = vData
End Property

Public Property Get dCOFINSSTQtde() As Double
    dCOFINSSTQtde = mvardCOFINSSTQtde
End Property

Public Property Let dCOFINSSTValor(ByVal vData As Double)
    mvardCOFINSSTValor = vData
End Property

Public Property Get dCOFINSSTValor() As Double
    dCOFINSSTValor = mvardCOFINSSTValor
End Property

Public Property Let sCST(ByVal vData As String)
    mvarsCST = vData
End Property

Public Property Get sCST() As String
    sCST = mvarsCST
End Property

Public Property Let sISSQN(ByVal vData As String)
    mvarsISSQN = vData
End Property

Public Property Get sISSQN() As String
    sISSQN = mvarsISSQN
End Property

Public Property Let dISSBase(ByVal vData As Double)
    mvardISSBase = vData
End Property

Public Property Get dISSBase() As Double
    dISSBase = mvardISSBase
End Property

Public Property Let dISSAliquota(ByVal vData As Double)
    mvardISSAliquota = vData
End Property

Public Property Get dISSAliquota() As Double
    dISSAliquota = mvardISSAliquota
End Property

Public Property Let sISSCidadeIBGE(ByVal vData As String)
    mvarsISSCidadeIBGE = vData
End Property

Public Property Get sISSCidadeIBGE() As String
    sISSCidadeIBGE = mvarsISSCidadeIBGE
End Property

Public Property Let dISSValor(ByVal vData As Double)
    mvardISSValor = vData
End Property

Public Property Get dISSValor() As Double
    dISSValor = mvardISSValor
End Property

Public Property Let dIIBase(ByVal vData As Double)
    mvardIIBase = vData
End Property

Public Property Get dIIBase() As Double
    dIIBase = mvardIIBase
End Property

Public Property Let dIIDespAduaneira(ByVal vData As Double)
    mvardIIDespAduaneira = vData
End Property

Public Property Get dIIDespAduaneira() As Double
    dIIDespAduaneira = mvardIIDespAduaneira
End Property

Public Property Let dIIIOF(ByVal vData As Double)
    mvardIIIOF = vData
End Property

Public Property Get dIIIOF() As Double
    dIIIOF = mvardIIIOF
End Property

Public Property Let dIIValor(ByVal vData As Double)
    mvardIIValor = vData
End Property

Public Property Get dIIValor() As Double
    dIIValor = mvardIIValor
End Property

Public Property Let iICMSBaseModalidade(ByVal vData As Integer)
    mvariICMSBaseModalidade = vData
End Property

Public Property Get iICMSBaseModalidade() As Integer
    iICMSBaseModalidade = mvariICMSBaseModalidade
End Property

Public Property Let iICMSSubstBaseModalidade(ByVal vData As Integer)
    mvariICMSSubstBaseModalidade = vData
End Property

Public Property Get iICMSSubstBaseModalidade() As Integer
    iICMSSubstBaseModalidade = mvariICMSSubstBaseModalidade
End Property

Public Property Let sIPIEnquadramentoClasse(ByVal vData As String)
    mvarsIPIEnquadramentoClasse = vData
End Property

Public Property Get sIPIEnquadramentoClasse() As String
    sIPIEnquadramentoClasse = mvarsIPIEnquadramentoClasse
End Property

Public Property Let sIPIEnquadramentoCodigo(ByVal vData As String)
    mvarsIPIEnquadramentoCodigo = vData
End Property

Public Property Get sIPIEnquadramentoCodigo() As String
    sIPIEnquadramentoCodigo = mvarsIPIEnquadramentoCodigo
End Property

Public Property Let sIPICNPJProdutor(ByVal vData As String)
    mvarsIPICNPJProdutor = vData
End Property

Public Property Get sIPICNPJProdutor() As String
    sIPICNPJProdutor = mvarsIPICNPJProdutor
End Property

Public Property Let sIPISeloCodigo(ByVal vData As String)
    mvarsIPISeloCodigo = vData
End Property

Public Property Get sIPISeloCodigo() As String
    sIPISeloCodigo = mvarsIPISeloCodigo
End Property

Public Property Let lIPISeloQtde(ByVal vData As Long)
    mvarlIPISeloQtde = vData
End Property

Public Property Get lIPISeloQtde() As Long
    lIPISeloQtde = mvarlIPISeloQtde
End Property

Public Property Let iIPITipoCalculo(ByVal vData As Integer)
    mvariIPITipoCalculo = vData
End Property

Public Property Get iIPITipoCalculo() As Integer
    iIPITipoCalculo = mvariIPITipoCalculo
End Property

Public Property Let dIPIUnidadePadraoQtde(ByVal vData As Double)
    mvardIPIUnidadePadraoQtde = vData
End Property

Public Property Get dIPIUnidadePadraoQtde() As Double
    dIPIUnidadePadraoQtde = mvardIPIUnidadePadraoQtde
End Property

Public Property Let dIPIUnidadePadraoValor(ByVal vData As Double)
    mvardIPIUnidadePadraoValor = vData
End Property

Public Property Get dIPIUnidadePadraoValor() As Double
    dIPIUnidadePadraoValor = mvardIPIUnidadePadraoValor
End Property

Public Property Let dValorFreteItem(ByVal vData As Double)
    mvardValorFreteItem = vData
End Property

Public Property Get dValorFreteItem() As Double
    dValorFreteItem = mvardValorFreteItem
End Property

Public Property Let dValorSeguroItem(ByVal vData As Double)
    mvardValorSeguroItem = vData
End Property

Public Property Get dValorSeguroItem() As Double
    dValorSeguroItem = mvardValorSeguroItem
End Property

Public Property Let dValorOutrasDespesasItem(ByVal vData As Double)
    mvardValorOutrasDespesasItem = vData
End Property

Public Property Get dValorOutrasDespesasItem() As Double
    dValorOutrasDespesasItem = mvardValorOutrasDespesasItem
End Property

Public Property Let dValorDescontoItem(ByVal vData As Double)
    mvardValorDescontoItem = vData
End Property

Public Property Get dValorDescontoItem() As Double
    dValorDescontoItem = mvardValorDescontoItem
End Property

Public Property Let iOrigemMercadoria(ByVal vData As Integer)
    mvariOrigemMercadoria = vData
End Property

Public Property Get iOrigemMercadoria() As Integer
    iOrigemMercadoria = mvariOrigemMercadoria
End Property

Public Property Let sFCI(ByVal vData As String)
    mvarsFCI = vData
End Property

Public Property Get sFCI() As String
    sFCI = mvarsFCI
End Property

Public Property Let iExTIPI(ByVal vData As Integer)
    mvariExTIPI = vData
End Property

Public Property Get iExTIPI() As Integer
    iExTIPI = mvariExTIPI
End Property

Public Property Let sGenero(ByVal vData As String)
    mvarsGenero = vData
End Property

Public Property Get sGenero() As String
    sGenero = mvarsGenero
End Property

Public Property Let iProdutoEspecifico(ByVal vData As Integer)
    mvariProdutoEspecifico = vData
End Property

Public Property Get iProdutoEspecifico() As Integer
    iProdutoEspecifico = mvariProdutoEspecifico
End Property

Public Property Let sEAN(ByVal vData As String)
    mvarsEAN = vData
End Property

Public Property Get sEAN() As String
    sEAN = mvarsEAN
End Property

Public Property Let sEANTrib(ByVal vData As String)
    mvarsEANTrib = vData
End Property

Public Property Get sEANTrib() As String
    sEANTrib = mvarsEANTrib
End Property

Public Property Let dQtdTrib(ByVal vData As Double)
    mvardQtdTrib = vData
End Property

Public Property Get dQtdTrib() As Double
    dQtdTrib = mvardQtdTrib
End Property

Public Property Let sUMTrib(ByVal vData As String)
    mvarsUMTrib = vData
End Property

Public Property Get sUMTrib() As String
    sUMTrib = mvarsUMTrib
End Property

Public Property Let dValorUnitTrib(ByVal vData As Double)
    mvardValorUnitTrib = vData
End Property

Public Property Get dValorUnitTrib() As Double
    dValorUnitTrib = mvardValorUnitTrib
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iIPITipoManual(ByVal vData As Integer)
    mvariIPITipoManual = vData
End Property

Public Property Get iIPITipoManual() As Integer
    iIPITipoManual = mvariIPITipoManual
End Property

Public Property Let iIPICodProdutoManual(ByVal vData As Integer)
    mvariIPICodProdutoManual = vData
End Property

Public Property Get iIPICodProdutoManual() As Integer
    iIPICodProdutoManual = mvariIPICodProdutoManual
End Property

Public Property Let iIPIBaseCalculoManual(ByVal vData As Integer)
    mvariIPIBaseCalculoManual = vData
End Property

Public Property Get iIPIBaseCalculoManual() As Integer
    iIPIBaseCalculoManual = mvariIPIBaseCalculoManual
End Property

Public Property Let iIPIPercRedBaseManual(ByVal vData As Integer)
    mvariIPIPercRedBaseManual = vData
End Property

Public Property Get iIPIPercRedBaseManual() As Integer
    iIPIPercRedBaseManual = mvariIPIPercRedBaseManual
End Property

Public Property Let iIPIAliquotaManual(ByVal vData As Integer)
    mvariIPIAliquotaManual = vData
End Property

Public Property Get iIPIAliquotaManual() As Integer
    iIPIAliquotaManual = mvariIPIAliquotaManual
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let iIPICreditoManual(ByVal vData As Integer)
    mvariIPICreditoManual = vData
End Property

Public Property Get iIPICreditoManual() As Integer
    iIPICreditoManual = mvariIPICreditoManual
End Property

Public Property Let iICMSTipoManual(ByVal vData As Integer)
    mvariICMSTipoManual = vData
End Property

Public Property Get iICMSTipoManual() As Integer
    iICMSTipoManual = mvariICMSTipoManual
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let iICMSPercRedBaseManual(ByVal vData As Integer)
    mvariICMSPercRedBaseManual = vData
End Property

Public Property Get iICMSPercRedBaseManual() As Integer
    iICMSPercRedBaseManual = mvariICMSPercRedBaseManual
End Property

Public Property Let iICMSAliquotaManual(ByVal vData As Integer)
    mvariICMSAliquotaManual = vData
End Property

Public Property Get iICMSAliquotaManual() As Integer
    iICMSAliquotaManual = mvariICMSAliquotaManual
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let iICMSCreditoManual(ByVal vData As Integer)
    mvariICMSCreditoManual = vData
End Property

Public Property Get iICMSCreditoManual() As Integer
    iICMSCreditoManual = mvariICMSCreditoManual
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let iICMSSubstAliquotaManual(ByVal vData As Integer)
    mvariICMSSubstAliquotaManual = vData
End Property

Public Property Get iICMSSubstAliquotaManual() As Integer
    iICMSSubstAliquotaManual = mvariICMSSubstAliquotaManual
End Property

Public Property Let iICMSSubstValorManual(ByVal vData As Integer)
    mvariICMSSubstValorManual = vData
End Property

Public Property Get iICMSSubstValorManual() As Integer
    iICMSSubstValorManual = mvariICMSSubstValorManual
End Property

Public Property Let iPISCreditoManual(ByVal vData As Integer)
    mvariPISCreditoManual = vData
End Property

Public Property Get iPISCreditoManual() As Integer
    iPISCreditoManual = mvariPISCreditoManual
End Property

Public Property Let iCOFINSCreditoManual(ByVal vData As Integer)
    mvariCOFINSCreditoManual = vData
End Property

Public Property Get iCOFINSCreditoManual() As Integer
    iCOFINSCreditoManual = mvariCOFINSCreditoManual
End Property

Public Property Let iICMSAliquotaAdicaoDIManual(ByVal vData As Integer)
    mvariICMSAliquotaAdicaoDIManual = vData
End Property

Public Property Get iICMSAliquotaAdicaoDIManual() As Integer
    iICMSAliquotaAdicaoDIManual = mvariICMSAliquotaAdicaoDIManual
End Property

Public Property Let iICMSPercRedBaseAdicaoDIManual(ByVal vData As Integer)
    mvariICMSPercRedBaseAdicaoDIManual = vData
End Property

Public Property Get iICMSPercRedBaseAdicaoDIManual() As Integer
    iICMSPercRedBaseAdicaoDIManual = mvariICMSPercRedBaseAdicaoDIManual
End Property

Public Property Let iPISAliquotaAdicaoDIManual(ByVal vData As Integer)
    mvariPISAliquotaAdicaoDIManual = vData
End Property

Public Property Get iPISAliquotaAdicaoDIManual() As Integer
    iPISAliquotaAdicaoDIManual = mvariPISAliquotaAdicaoDIManual
End Property

Public Property Let iCOFINSAliquotaAdicaoDIManual(ByVal vData As Integer)
    mvariCOFINSAliquotaAdicaoDIManual = vData
End Property

Public Property Get iCOFINSAliquotaAdicaoDIManual() As Integer
    iCOFINSAliquotaAdicaoDIManual = mvariCOFINSAliquotaAdicaoDIManual
End Property

Public Property Let iDespImpICMSBaseManual(ByVal vData As Integer)
    mvariDespImpICMSBaseManual = vData
End Property

Public Property Get iDespImpICMSBaseManual() As Integer
    iDespImpICMSBaseManual = mvariDespImpICMSBaseManual
End Property

Public Property Let iDespImpICMSValorManual(ByVal vData As Integer)
    mvariDespImpICMSValorManual = vData
End Property

Public Property Get iDespImpICMSValorManual() As Integer
    iDespImpICMSValorManual = mvariDespImpICMSValorManual
End Property

Public Property Let iDespImpICMSCreditoManual(ByVal vData As Integer)
    mvariDespImpICMSCreditoManual = vData
End Property

Public Property Get iDespImpICMSCreditoManual() As Integer
    iDespImpICMSCreditoManual = mvariDespImpICMSCreditoManual
End Property

Public Property Let iICMSSubstPercRedBaseManual(ByVal vData As Integer)
    mvariICMSSubstPercRedBaseManual = vData
End Property

Public Property Get iICMSSubstPercRedBaseManual() As Integer
    iICMSSubstPercRedBaseManual = mvariICMSSubstPercRedBaseManual
End Property

Public Property Let iICMSSubstPercMVAManual(ByVal vData As Integer)
    mvariICMSSubstPercMVAManual = vData
End Property

Public Property Get iICMSSubstPercMVAManual() As Integer
    iICMSSubstPercMVAManual = mvariICMSSubstPercMVAManual
End Property

Public Property Let iPISTipoManual(ByVal vData As Integer)
    mvariPISTipoManual = vData
End Property

Public Property Get iPISTipoManual() As Integer
    iPISTipoManual = mvariPISTipoManual
End Property

Public Property Let iPISTipoCalculoManual(ByVal vData As Integer)
    mvariPISTipoCalculoManual = vData
End Property

Public Property Get iPISTipoCalculoManual() As Integer
    iPISTipoCalculoManual = mvariPISTipoCalculoManual
End Property

Public Property Let iPISBaseManual(ByVal vData As Integer)
    mvariPISBaseManual = vData
End Property

Public Property Get iPISBaseManual() As Integer
    iPISBaseManual = mvariPISBaseManual
End Property

Public Property Let iPISAliquotaManual(ByVal vData As Integer)
    mvariPISAliquotaManual = vData
End Property

Public Property Get iPISAliquotaManual() As Integer
    iPISAliquotaManual = mvariPISAliquotaManual
End Property

Public Property Let iPISAliquotaValorManual(ByVal vData As Integer)
    mvariPISAliquotaValorManual = vData
End Property

Public Property Get iPISAliquotaValorManual() As Integer
    iPISAliquotaValorManual = mvariPISAliquotaValorManual
End Property

Public Property Let iPISQtdeManual(ByVal vData As Integer)
    mvariPISQtdeManual = vData
End Property

Public Property Get iPISQtdeManual() As Integer
    iPISQtdeManual = mvariPISQtdeManual
End Property

Public Property Let iPISValorManual(ByVal vData As Integer)
    mvariPISValorManual = vData
End Property

Public Property Get iPISValorManual() As Integer
    iPISValorManual = mvariPISValorManual
End Property

Public Property Let iPISSTTipoCalculoManual(ByVal vData As Integer)
    mvariPISSTTipoCalculoManual = vData
End Property

Public Property Get iPISSTTipoCalculoManual() As Integer
    iPISSTTipoCalculoManual = mvariPISSTTipoCalculoManual
End Property

Public Property Let iPISSTBaseManual(ByVal vData As Integer)
    mvariPISSTBaseManual = vData
End Property

Public Property Get iPISSTBaseManual() As Integer
    iPISSTBaseManual = mvariPISSTBaseManual
End Property

Public Property Let iPISSTAliquotaManual(ByVal vData As Integer)
    mvariPISSTAliquotaManual = vData
End Property

Public Property Get iPISSTAliquotaManual() As Integer
    iPISSTAliquotaManual = mvariPISSTAliquotaManual
End Property

Public Property Let iPISSTAliquotaValorManual(ByVal vData As Integer)
    mvariPISSTAliquotaValorManual = vData
End Property

Public Property Get iPISSTAliquotaValorManual() As Integer
    iPISSTAliquotaValorManual = mvariPISSTAliquotaValorManual
End Property

Public Property Let iPISSTQtdeManual(ByVal vData As Integer)
    mvariPISSTQtdeManual = vData
End Property

Public Property Get iPISSTQtdeManual() As Integer
    iPISSTQtdeManual = mvariPISSTQtdeManual
End Property

Public Property Let iPISSTValorManual(ByVal vData As Integer)
    mvariPISSTValorManual = vData
End Property

Public Property Get iPISSTValorManual() As Integer
    iPISSTValorManual = mvariPISSTValorManual
End Property

Public Property Let iCOFINSTipoManual(ByVal vData As Integer)
    mvariCOFINSTipoManual = vData
End Property

Public Property Get iCOFINSTipoManual() As Integer
    iCOFINSTipoManual = mvariCOFINSTipoManual
End Property

Public Property Let iCOFINSTipoCalculoManual(ByVal vData As Integer)
    mvariCOFINSTipoCalculoManual = vData
End Property

Public Property Get iCOFINSTipoCalculoManual() As Integer
    iCOFINSTipoCalculoManual = mvariCOFINSTipoCalculoManual
End Property

Public Property Let iCOFINSBaseManual(ByVal vData As Integer)
    mvariCOFINSBaseManual = vData
End Property

Public Property Get iCOFINSBaseManual() As Integer
    iCOFINSBaseManual = mvariCOFINSBaseManual
End Property

Public Property Let iCOFINSAliquotaManual(ByVal vData As Integer)
    mvariCOFINSAliquotaManual = vData
End Property

Public Property Get iCOFINSAliquotaManual() As Integer
    iCOFINSAliquotaManual = mvariCOFINSAliquotaManual
End Property

Public Property Let iCOFINSAliquotaValorManual(ByVal vData As Integer)
    mvariCOFINSAliquotaValorManual = vData
End Property

Public Property Get iCOFINSAliquotaValorManual() As Integer
    iCOFINSAliquotaValorManual = mvariCOFINSAliquotaValorManual
End Property

Public Property Let iCOFINSQtdeManual(ByVal vData As Integer)
    mvariCOFINSQtdeManual = vData
End Property

Public Property Get iCOFINSQtdeManual() As Integer
    iCOFINSQtdeManual = mvariCOFINSQtdeManual
End Property

Public Property Let iCOFINSValorManual(ByVal vData As Integer)
    mvariCOFINSValorManual = vData
End Property

Public Property Get iCOFINSValorManual() As Integer
    iCOFINSValorManual = mvariCOFINSValorManual
End Property

Public Property Let iCOFINSSTTipoCalculoManual(ByVal vData As Integer)
    mvariCOFINSSTTipoCalculoManual = vData
End Property

Public Property Get iCOFINSSTTipoCalculoManual() As Integer
    iCOFINSSTTipoCalculoManual = mvariCOFINSSTTipoCalculoManual
End Property

Public Property Let iCOFINSSTBaseManual(ByVal vData As Integer)
    mvariCOFINSSTBaseManual = vData
End Property

Public Property Get iCOFINSSTBaseManual() As Integer
    iCOFINSSTBaseManual = mvariCOFINSSTBaseManual
End Property

Public Property Let iCOFINSSTAliquotaManual(ByVal vData As Integer)
    mvariCOFINSSTAliquotaManual = vData
End Property

Public Property Get iCOFINSSTAliquotaManual() As Integer
    iCOFINSSTAliquotaManual = mvariCOFINSSTAliquotaManual
End Property

Public Property Let iCOFINSSTAliquotaValorManual(ByVal vData As Integer)
    mvariCOFINSSTAliquotaValorManual = vData
End Property

Public Property Get iCOFINSSTAliquotaValorManual() As Integer
    iCOFINSSTAliquotaValorManual = mvariCOFINSSTAliquotaValorManual
End Property

Public Property Let iCOFINSSTQtdeManual(ByVal vData As Integer)
    mvariCOFINSSTQtdeManual = vData
End Property

Public Property Get iCOFINSSTQtdeManual() As Integer
    iCOFINSSTQtdeManual = mvariCOFINSSTQtdeManual
End Property

Public Property Let iCOFINSSTValorManual(ByVal vData As Integer)
    mvariCOFINSSTValorManual = vData
End Property

Public Property Get iCOFINSSTValorManual() As Integer
    iCOFINSSTValorManual = mvariCOFINSSTValorManual
End Property

Public Property Let iCSTManual(ByVal vData As Integer)
    mvariCSTManual = vData
End Property

Public Property Get iCSTManual() As Integer
    iCSTManual = mvariCSTManual
End Property

Public Property Let iISSQNManual(ByVal vData As Integer)
    mvariISSQNManual = vData
End Property

Public Property Get iISSQNManual() As Integer
    iISSQNManual = mvariISSQNManual
End Property

Public Property Let iISSBaseManual(ByVal vData As Integer)
    mvariISSBaseManual = vData
End Property

Public Property Get iISSBaseManual() As Integer
    iISSBaseManual = mvariISSBaseManual
End Property

Public Property Let iISSAliquotaManual(ByVal vData As Integer)
    mvariISSAliquotaManual = vData
End Property

Public Property Get iISSAliquotaManual() As Integer
    iISSAliquotaManual = mvariISSAliquotaManual
End Property

Public Property Let iISSCidadeIBGEManual(ByVal vData As Integer)
    mvariISSCidadeIBGEManual = vData
End Property

Public Property Get iISSCidadeIBGEManual() As Integer
    iISSCidadeIBGEManual = mvariISSCidadeIBGEManual
End Property

Public Property Let iISSValorManual(ByVal vData As Integer)
    mvariISSValorManual = vData
End Property

Public Property Get iISSValorManual() As Integer
    iISSValorManual = mvariISSValorManual
End Property

Public Property Let iIIBaseManual(ByVal vData As Integer)
    mvariIIBaseManual = vData
End Property

Public Property Get iIIBaseManual() As Integer
    iIIBaseManual = mvariIIBaseManual
End Property

Public Property Let iIIDespAduaneiraManual(ByVal vData As Integer)
    mvariIIDespAduaneiraManual = vData
End Property

Public Property Get iIIDespAduaneiraManual() As Integer
    iIIDespAduaneiraManual = mvariIIDespAduaneiraManual
End Property

Public Property Let iIIIOFManual(ByVal vData As Integer)
    mvariIIIOFManual = vData
End Property

Public Property Get iIIIOFManual() As Integer
    iIIIOFManual = mvariIIIOFManual
End Property

Public Property Let iIIValorManual(ByVal vData As Integer)
    mvariIIValorManual = vData
End Property

Public Property Get iIIValorManual() As Integer
    iIIValorManual = mvariIIValorManual
End Property

Public Property Let iICMSBaseModalidadeManual(ByVal vData As Integer)
    mvariICMSBaseModalidadeManual = vData
End Property

Public Property Get iICMSBaseModalidadeManual() As Integer
    iICMSBaseModalidadeManual = mvariICMSBaseModalidadeManual
End Property

Public Property Let iICMSSubstBaseModalidadeManual(ByVal vData As Integer)
    mvariICMSSubstBaseModalidadeManual = vData
End Property

Public Property Get iICMSSubstBaseModalidadeManual() As Integer
    iICMSSubstBaseModalidadeManual = mvariICMSSubstBaseModalidadeManual
End Property

Public Property Let iIPIEnquadramentoClasseManual(ByVal vData As Integer)
    mvariIPIEnquadramentoClasseManual = vData
End Property

Public Property Get iIPIEnquadramentoClasseManual() As Integer
    iIPIEnquadramentoClasseManual = mvariIPIEnquadramentoClasseManual
End Property

Public Property Let iIPIEnquadramentoCodigoManual(ByVal vData As Integer)
    mvariIPIEnquadramentoCodigoManual = vData
End Property

Public Property Get iIPIEnquadramentoCodigoManual() As Integer
    iIPIEnquadramentoCodigoManual = mvariIPIEnquadramentoCodigoManual
End Property

Public Property Let iIPICNPJProdutorManual(ByVal vData As Integer)
    mvariIPICNPJProdutorManual = vData
End Property

Public Property Get iIPICNPJProdutorManual() As Integer
    iIPICNPJProdutorManual = mvariIPICNPJProdutorManual
End Property

Public Property Let iIPISeloCodigoManual(ByVal vData As Integer)
    mvariIPISeloCodigoManual = vData
End Property

Public Property Get iIPISeloCodigoManual() As Integer
    iIPISeloCodigoManual = mvariIPISeloCodigoManual
End Property

Public Property Let iIPISeloQtdeManual(ByVal vData As Integer)
    mvariIPISeloQtdeManual = vData
End Property

Public Property Get iIPISeloQtdeManual() As Integer
    iIPISeloQtdeManual = mvariIPISeloQtdeManual
End Property

Public Property Let iIPITipoCalculoManual(ByVal vData As Integer)
    mvariIPITipoCalculoManual = vData
End Property

Public Property Get iIPITipoCalculoManual() As Integer
    iIPITipoCalculoManual = mvariIPITipoCalculoManual
End Property

Public Property Let iIPIUnidadePadraoQtdeManual(ByVal vData As Integer)
    mvariIPIUnidadePadraoQtdeManual = vData
End Property

Public Property Get iIPIUnidadePadraoQtdeManual() As Integer
    iIPIUnidadePadraoQtdeManual = mvariIPIUnidadePadraoQtdeManual
End Property

Public Property Let iIPIUnidadePadraoValorManual(ByVal vData As Integer)
    mvariIPIUnidadePadraoValorManual = vData
End Property

Public Property Get iIPIUnidadePadraoValorManual() As Integer
    iIPIUnidadePadraoValorManual = mvariIPIUnidadePadraoValorManual
End Property

Public Property Let iValorFreteItemManual(ByVal vData As Integer)
    mvariValorFreteItemManual = vData
End Property

Public Property Get iValorFreteItemManual() As Integer
    iValorFreteItemManual = mvariValorFreteItemManual
End Property

Public Property Let iValorSeguroItemManual(ByVal vData As Integer)
    mvariValorSeguroItemManual = vData
End Property

Public Property Get iValorSeguroItemManual() As Integer
    iValorSeguroItemManual = mvariValorSeguroItemManual
End Property

Public Property Let iValorOutrasDespesasItemManual(ByVal vData As Integer)
    mvariValorOutrasDespesasItemManual = vData
End Property

Public Property Get iValorOutrasDespesasItemManual() As Integer
    iValorOutrasDespesasItemManual = mvariValorOutrasDespesasItemManual
End Property

Public Property Let iValorDescontoItemManual(ByVal vData As Integer)
    mvariValorDescontoItemManual = vData
End Property

Public Property Get iValorDescontoItemManual() As Integer
    iValorDescontoItemManual = mvariValorDescontoItemManual
End Property

Public Property Let iOrigemMercadoriaManual(ByVal vData As Integer)
    mvariOrigemMercadoriaManual = vData
End Property

Public Property Get iOrigemMercadoriaManual() As Integer
    iOrigemMercadoriaManual = mvariOrigemMercadoriaManual
End Property

Public Property Let iFCIManual(ByVal vData As Integer)
    mvariFCIManual = vData
End Property

Public Property Get iFCIManual() As Integer
    iFCIManual = mvariFCIManual
End Property

Public Property Set objDoc(ByVal vData As Object)
    Set mvarobjDoc = vData
End Property

Public Property Get objDoc() As Object
    Set objDoc = mvarobjDoc
End Property

Public Sub Copia(ByVal objTribDocItemOriginal As ClassTributacaoDocItem)
On Error GoTo Erro_Copia
    Call Copia2(objTribDocItemOriginal)
    Exit Sub
Erro_Copia:
    Select Case gErr
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202902)
    End Select
    Exit Sub
End Sub

Public Sub Copia2(objTribDocItemOriginal As Variant)

On Error GoTo Erro_Copia2

    sCListServNFe = objTribDocItemOriginal.sCListServNFe
    iTipoDoc = objTribDocItemOriginal.iTipoDoc
    lNumIntDocItem = objTribDocItemOriginal.lNumIntDocItem
    lNumIntDoc = objTribDocItemOriginal.lNumIntDoc
    iItem = objTribDocItemOriginal.iItem
    sProduto = objTribDocItemOriginal.sProduto
    sProdutoDescricao = objTribDocItemOriginal.sProdutoDescricao
    sUnidadeMed = objTribDocItemOriginal.sUnidadeMed
    dQuantidade = objTribDocItemOriginal.dQuantidade
    dPrecoTotal = objTribDocItemOriginal.dPrecoTotal
    dPrecoUnitario = objTribDocItemOriginal.dPrecoUnitario
    dPrecoUnitarioMoeda = objTribDocItemOriginal.dPrecoUnitarioMoeda
    dDescontoGrid = objTribDocItemOriginal.dDescontoGrid
    sNaturezaOp = objTribDocItemOriginal.sNaturezaOp
    iTipoTributacao = objTribDocItemOriginal.iTipoTributacao
    iIPITipo = objTribDocItemOriginal.iIPITipo
    sIPICodProduto = objTribDocItemOriginal.sIPICodProduto
    dIPIBaseCalculo = objTribDocItemOriginal.dIPIBaseCalculo
    dIPIPercRedBase = objTribDocItemOriginal.dIPIPercRedBase
    dIPIAliquota = objTribDocItemOriginal.dIPIAliquota
    dIPIValor = objTribDocItemOriginal.dIPIValor
    dIPICredito = objTribDocItemOriginal.dIPICredito
    iICMSTipo = objTribDocItemOriginal.iICMSTipo
    dICMSBase = objTribDocItemOriginal.dICMSBase
    dICMSPercRedBase = objTribDocItemOriginal.dICMSPercRedBase
    dICMSAliquota = objTribDocItemOriginal.dICMSAliquota
    dICMSValor = objTribDocItemOriginal.dICMSValor
    dICMSCredito = objTribDocItemOriginal.dICMSCredito
    dICMSSubstBase = objTribDocItemOriginal.dICMSSubstBase
    dICMSSubstAliquota = objTribDocItemOriginal.dICMSSubstAliquota
    dICMSSubstValor = objTribDocItemOriginal.dICMSSubstValor
    dPISCredito = objTribDocItemOriginal.dPISCredito
    dCOFINSCredito = objTribDocItemOriginal.dCOFINSCredito
    dICMSAliquotaAdicaoDI = objTribDocItemOriginal.dICMSAliquotaAdicaoDI
    dICMSPercRedBaseAdicaoDI = objTribDocItemOriginal.dICMSPercRedBaseAdicaoDI
    dPISAliquotaAdicaoDI = objTribDocItemOriginal.dPISAliquotaAdicaoDI
    dCOFINSAliquotaAdicaoDI = objTribDocItemOriginal.dCOFINSAliquotaAdicaoDI
    dDespImpICMSBase = objTribDocItemOriginal.dDespImpICMSBase
    dDespImpICMSValor = objTribDocItemOriginal.dDespImpICMSValor
    dDespImpICMSCredito = objTribDocItemOriginal.dDespImpICMSCredito
    dICMSSubstPercRedBase = objTribDocItemOriginal.dICMSSubstPercRedBase
    dICMSSubstPercMVA = objTribDocItemOriginal.dICMSSubstPercMVA
    iPISTipo = objTribDocItemOriginal.iPISTipo
    iPISTipoCalculo = objTribDocItemOriginal.iPISTipoCalculo
    dPISBase = objTribDocItemOriginal.dPISBase
    dPISAliquota = objTribDocItemOriginal.dPISAliquota
    dPISAliquotaValor = objTribDocItemOriginal.dPISAliquotaValor
    dPISQtde = objTribDocItemOriginal.dPISQtde
    dPISValor = objTribDocItemOriginal.dPISValor
    iPISSTTipoCalculo = objTribDocItemOriginal.iPISSTTipoCalculo
    dPISSTBase = objTribDocItemOriginal.dPISSTBase
    dPISSTAliquota = objTribDocItemOriginal.dPISSTAliquota
    dPISSTAliquotaValor = objTribDocItemOriginal.dPISSTAliquotaValor
    dPISSTQtde = objTribDocItemOriginal.dPISSTQtde
    dPISSTValor = objTribDocItemOriginal.dPISSTValor
    iCOFINSTipo = objTribDocItemOriginal.iCOFINSTipo
    iCOFINSTipoCalculo = objTribDocItemOriginal.iCOFINSTipoCalculo
    dCOFINSBase = objTribDocItemOriginal.dCOFINSBase
    dCOFINSAliquota = objTribDocItemOriginal.dCOFINSAliquota
    dCOFINSAliquotaValor = objTribDocItemOriginal.dCOFINSAliquotaValor
    dCOFINSQtde = objTribDocItemOriginal.dCOFINSQtde
    dCOFINSValor = objTribDocItemOriginal.dCOFINSValor
    iCOFINSSTTipoCalculo = objTribDocItemOriginal.iCOFINSSTTipoCalculo
    dCOFINSSTBase = objTribDocItemOriginal.dCOFINSSTBase
    dCOFINSSTAliquota = objTribDocItemOriginal.dCOFINSSTAliquota
    dCOFINSSTAliquotaValor = objTribDocItemOriginal.dCOFINSSTAliquotaValor
    dCOFINSSTQtde = objTribDocItemOriginal.dCOFINSSTQtde
    dCOFINSSTValor = objTribDocItemOriginal.dCOFINSSTValor
    sCST = objTribDocItemOriginal.sCST
    sISSQN = objTribDocItemOriginal.sISSQN
    dISSBase = objTribDocItemOriginal.dISSBase
    dISSAliquota = objTribDocItemOriginal.dISSAliquota
    sISSCidadeIBGE = objTribDocItemOriginal.sISSCidadeIBGE
    dISSValor = objTribDocItemOriginal.dISSValor
    dIIBase = objTribDocItemOriginal.dIIBase
    dIIDespAduaneira = objTribDocItemOriginal.dIIDespAduaneira
    dIIIOF = objTribDocItemOriginal.dIIIOF
    dIIValor = objTribDocItemOriginal.dIIValor
    iICMSBaseModalidade = objTribDocItemOriginal.iICMSBaseModalidade
    iICMSSubstBaseModalidade = objTribDocItemOriginal.iICMSSubstBaseModalidade
    sIPIEnquadramentoClasse = objTribDocItemOriginal.sIPIEnquadramentoClasse
    sIPIEnquadramentoCodigo = objTribDocItemOriginal.sIPIEnquadramentoCodigo
    sIPICNPJProdutor = objTribDocItemOriginal.sIPICNPJProdutor
    sIPISeloCodigo = objTribDocItemOriginal.sIPISeloCodigo
    lIPISeloQtde = objTribDocItemOriginal.lIPISeloQtde
    iIPITipoCalculo = objTribDocItemOriginal.iIPITipoCalculo
    dIPIUnidadePadraoQtde = objTribDocItemOriginal.dIPIUnidadePadraoQtde
    dIPIUnidadePadraoValor = objTribDocItemOriginal.dIPIUnidadePadraoValor
    dValorFreteItem = objTribDocItemOriginal.dValorFreteItem
    dValorSeguroItem = objTribDocItemOriginal.dValorSeguroItem
    dValorOutrasDespesasItem = objTribDocItemOriginal.dValorOutrasDespesasItem
    dValorDescontoItem = objTribDocItemOriginal.dValorDescontoItem
    iOrigemMercadoria = objTribDocItemOriginal.iOrigemMercadoria
    sFCI = objTribDocItemOriginal.sFCI
    iExTIPI = objTribDocItemOriginal.iExTIPI
    sGenero = objTribDocItemOriginal.sGenero
    iProdutoEspecifico = objTribDocItemOriginal.iProdutoEspecifico
    sEAN = objTribDocItemOriginal.sEAN
    sEANTrib = objTribDocItemOriginal.sEANTrib
    dQtdTrib = objTribDocItemOriginal.dQtdTrib
    sUMTrib = objTribDocItemOriginal.sUMTrib
    dValorUnitTrib = objTribDocItemOriginal.dValorUnitTrib
    iNaturezaOpManual = objTribDocItemOriginal.iNaturezaOpManual
    iTipoTributacaoManual = objTribDocItemOriginal.iTipoTributacaoManual
    iIPITipoManual = objTribDocItemOriginal.iIPITipoManual
    iIPICodProdutoManual = objTribDocItemOriginal.iIPICodProdutoManual
    iIPIBaseCalculoManual = objTribDocItemOriginal.iIPIBaseCalculoManual
    iIPIPercRedBaseManual = objTribDocItemOriginal.iIPIPercRedBaseManual
    iIPIAliquotaManual = objTribDocItemOriginal.iIPIAliquotaManual
    iIPIValorManual = objTribDocItemOriginal.iIPIValorManual
    iIPICreditoManual = objTribDocItemOriginal.iIPICreditoManual
    iICMSTipoManual = objTribDocItemOriginal.iICMSTipoManual
    iICMSBaseManual = objTribDocItemOriginal.iICMSBaseManual
    iICMSPercRedBaseManual = objTribDocItemOriginal.iICMSPercRedBaseManual
    iICMSAliquotaManual = objTribDocItemOriginal.iICMSAliquotaManual
    iICMSValorManual = objTribDocItemOriginal.iICMSValorManual
    iICMSCreditoManual = objTribDocItemOriginal.iICMSCreditoManual
    iICMSSubstBaseManual = objTribDocItemOriginal.iICMSSubstBaseManual
    iICMSSubstAliquotaManual = objTribDocItemOriginal.iICMSSubstAliquotaManual
    iICMSSubstValorManual = objTribDocItemOriginal.iICMSSubstValorManual
    iPISCreditoManual = objTribDocItemOriginal.iPISCreditoManual
    iCOFINSCreditoManual = objTribDocItemOriginal.iCOFINSCreditoManual
    iICMSAliquotaAdicaoDIManual = objTribDocItemOriginal.iICMSAliquotaAdicaoDIManual
    iICMSPercRedBaseAdicaoDIManual = objTribDocItemOriginal.iICMSPercRedBaseAdicaoDIManual
    iPISAliquotaAdicaoDIManual = objTribDocItemOriginal.iPISAliquotaAdicaoDIManual
    iCOFINSAliquotaAdicaoDIManual = objTribDocItemOriginal.iCOFINSAliquotaAdicaoDIManual
    iDespImpICMSBaseManual = objTribDocItemOriginal.iDespImpICMSBaseManual
    iDespImpICMSValorManual = objTribDocItemOriginal.iDespImpICMSValorManual
    iDespImpICMSCreditoManual = objTribDocItemOriginal.iDespImpICMSCreditoManual
    iICMSSubstPercRedBaseManual = objTribDocItemOriginal.iICMSSubstPercRedBaseManual
    iICMSSubstPercMVAManual = objTribDocItemOriginal.iICMSSubstPercMVAManual
    iPISTipoManual = objTribDocItemOriginal.iPISTipoManual
    iPISTipoCalculoManual = objTribDocItemOriginal.iPISTipoCalculoManual
    iPISBaseManual = objTribDocItemOriginal.iPISBaseManual
    iPISAliquotaManual = objTribDocItemOriginal.iPISAliquotaManual
    iPISAliquotaValorManual = objTribDocItemOriginal.iPISAliquotaValorManual
    iPISQtdeManual = objTribDocItemOriginal.iPISQtdeManual
    iPISValorManual = objTribDocItemOriginal.iPISValorManual
    iPISSTTipoCalculoManual = objTribDocItemOriginal.iPISSTTipoCalculoManual
    iPISSTBaseManual = objTribDocItemOriginal.iPISSTBaseManual
    iPISSTAliquotaManual = objTribDocItemOriginal.iPISSTAliquotaManual
    iPISSTAliquotaValorManual = objTribDocItemOriginal.iPISSTAliquotaValorManual
    iPISSTQtdeManual = objTribDocItemOriginal.iPISSTQtdeManual
    iPISSTValorManual = objTribDocItemOriginal.iPISSTValorManual
    iCOFINSTipoManual = objTribDocItemOriginal.iCOFINSTipoManual
    iCOFINSTipoCalculoManual = objTribDocItemOriginal.iCOFINSTipoCalculoManual
    iCOFINSBaseManual = objTribDocItemOriginal.iCOFINSBaseManual
    iCOFINSAliquotaManual = objTribDocItemOriginal.iCOFINSAliquotaManual
    iCOFINSAliquotaValorManual = objTribDocItemOriginal.iCOFINSAliquotaValorManual
    iCOFINSQtdeManual = objTribDocItemOriginal.iCOFINSQtdeManual
    iCOFINSValorManual = objTribDocItemOriginal.iCOFINSValorManual
    iCOFINSSTTipoCalculoManual = objTribDocItemOriginal.iCOFINSSTTipoCalculoManual
    iCOFINSSTBaseManual = objTribDocItemOriginal.iCOFINSSTBaseManual
    iCOFINSSTAliquotaManual = objTribDocItemOriginal.iCOFINSSTAliquotaManual
    iCOFINSSTAliquotaValorManual = objTribDocItemOriginal.iCOFINSSTAliquotaValorManual
    iCOFINSSTQtdeManual = objTribDocItemOriginal.iCOFINSSTQtdeManual
    iCOFINSSTValorManual = objTribDocItemOriginal.iCOFINSSTValorManual
    iCSTManual = objTribDocItemOriginal.iCSTManual
    iISSQNManual = objTribDocItemOriginal.iISSQNManual
    iISSBaseManual = objTribDocItemOriginal.iISSBaseManual
    iISSAliquotaManual = objTribDocItemOriginal.iISSAliquotaManual
    iISSCidadeIBGEManual = objTribDocItemOriginal.iISSCidadeIBGEManual
    iISSValorManual = objTribDocItemOriginal.iISSValorManual
    iIIBaseManual = objTribDocItemOriginal.iIIBaseManual
    iIIDespAduaneiraManual = objTribDocItemOriginal.iIIDespAduaneiraManual
    iIIIOFManual = objTribDocItemOriginal.iIIIOFManual
    iIIValorManual = objTribDocItemOriginal.iIIValorManual
    iICMSBaseModalidadeManual = objTribDocItemOriginal.iICMSBaseModalidadeManual
    iICMSSubstBaseModalidadeManual = objTribDocItemOriginal.iICMSSubstBaseModalidadeManual
    iIPIEnquadramentoClasseManual = objTribDocItemOriginal.iIPIEnquadramentoClasseManual
    iIPIEnquadramentoCodigoManual = objTribDocItemOriginal.iIPIEnquadramentoCodigoManual
    iIPICNPJProdutorManual = objTribDocItemOriginal.iIPICNPJProdutorManual
    iIPISeloCodigoManual = objTribDocItemOriginal.iIPISeloCodigoManual
    iIPISeloQtdeManual = objTribDocItemOriginal.iIPISeloQtdeManual
    iIPITipoCalculoManual = objTribDocItemOriginal.iIPITipoCalculoManual
    iIPIUnidadePadraoQtdeManual = objTribDocItemOriginal.iIPIUnidadePadraoQtdeManual
    iIPIUnidadePadraoValorManual = objTribDocItemOriginal.iIPIUnidadePadraoValorManual
    iValorFreteItemManual = objTribDocItemOriginal.iValorFreteItemManual
    iValorSeguroItemManual = objTribDocItemOriginal.iValorSeguroItemManual
    iValorOutrasDespesasItemManual = objTribDocItemOriginal.iValorOutrasDespesasItemManual
    iValorDescontoItemManual = objTribDocItemOriginal.iValorDescontoItemManual
    iOrigemMercadoriaManual = objTribDocItemOriginal.iOrigemMercadoriaManual
    iFCIManual = objTribDocItemOriginal.iFCIManual

    dICMSSTCobrAntBase = objTribDocItemOriginal.dICMSSTCobrAntBase
    dICMSSTCobrAntValor = objTribDocItemOriginal.dICMSSTCobrAntValor
    iICMSSTCobrAntBaseManual = objTribDocItemOriginal.iICMSSTCobrAntBaseManual
    iICMSSTCobrAntValorManual = objTribDocItemOriginal.iICMSSTCobrAntValorManual
    
    dICMSpercBaseOperacaoPropria = objTribDocItemOriginal.dICMSpercBaseOperacaoPropria
    iICMSpercBaseOperacaoPropriaManual = objTribDocItemOriginal.iICMSpercBaseOperacaoPropriaManual
    sICMSUFDevidoST = objTribDocItemOriginal.sICMSUFDevidoST
    iICMSUFDevidoSTManual = objTribDocItemOriginal.iICMSUFDevidoSTManual
    dICMSvBCSTRet = objTribDocItemOriginal.dICMSvBCSTRet
    iICMSvBCSTRetManual = objTribDocItemOriginal.iICMSvBCSTRetManual
    dICMSvICMSSTRet = objTribDocItemOriginal.dICMSvICMSSTRet
    iICMSvICMSSTRetManual = objTribDocItemOriginal.iICMSvICMSSTRetManual
    dICMSvBCSTDest = objTribDocItemOriginal.dICMSvBCSTDest
    iICMSvBCSTDestManual = objTribDocItemOriginal.iICMSvBCSTDestManual
    dICMSvICMSSTDest = objTribDocItemOriginal.dICMSvICMSSTDest
    iICMSvICMSSTDestManual = objTribDocItemOriginal.iICMSvICMSSTDestManual
    dICMSpCredSN = objTribDocItemOriginal.dICMSpCredSN
    iICMSpCredSNManual = objTribDocItemOriginal.iICMSpCredSNManual
    dICMSvCredSN = objTribDocItemOriginal.dICMSvCredSN
    iICMSvCredSNManual = objTribDocItemOriginal.iICMSvCredSNManual
    dICMSValorIsento = objTribDocItemOriginal.dICMSValorIsento
    iICMSValorIsentoManual = objTribDocItemOriginal.iICMSValorIsentoManual
    iICMSMotivo = objTribDocItemOriginal.iICMSMotivo
    iICMSMotivoManual = objTribDocItemOriginal.iICMSMotivoManual
    iICMSSimplesTipo = objTribDocItemOriginal.iICMSSimplesTipo
    iICMSSimplesTipoManual = objTribDocItemOriginal.iICMSSimplesTipoManual
    iISSTipo = objTribDocItemOriginal.iISSTipo
    iISSTipoManual = objTribDocItemOriginal.iISSTipoManual
    iRegimeTributario = objTribDocItemOriginal.iRegimeTributario
    iRegimeTributarioManual = VAR_PREENCH_MANUAL '??? Jones 05/04/12
    sISSCST = objTribDocItemOriginal.sISSCST
    sCSOSN = objTribDocItemOriginal.sCSOSN
    sNatBCCred = objTribDocItemOriginal.sNatBCCred
    iNatBCCredManual = objTribDocItemOriginal.iNatBCCredManual
    dTotTrib = objTribDocItemOriginal.dTotTrib
    iTotTribManual = objTribDocItemOriginal.iTotTribManual
    iTotTribTipo = objTribDocItemOriginal.iTotTribTipo
    iTotTribTipoManual = objTribDocItemOriginal.iTotTribTipoManual
    
    dTotTribAliqFed = objTribDocItemOriginal.dTotTribAliqFed
    iTotTribAliqFedManual = objTribDocItemOriginal.iTotTribAliqFedManual
    dTotTribAliqEst = objTribDocItemOriginal.dTotTribAliqEst
    iTotTribAliqEstManual = objTribDocItemOriginal.iTotTribAliqEstManual
    dTotTribAliqMunic = objTribDocItemOriginal.dTotTribAliqMunic
    iTotTribAliqMunicManual = objTribDocItemOriginal.iTotTribAliqMunicManual
    
    'nfe 3.10
    dICMS51ValorOp = objTribDocItemOriginal.dICMS51ValorOp
    dICMSPercDifer = objTribDocItemOriginal.dICMSPercDifer
    dICMSValorDif = objTribDocItemOriginal.dICMSValorDif
    iICMS51ValorOpManual = objTribDocItemOriginal.iICMS51ValorOpManual
    iICMSPercDiferManual = objTribDocItemOriginal.iICMSPercDiferManual
    iICMSValorDifManual = objTribDocItemOriginal.iICMSValorDifManual
    dISSValorDeducao = objTribDocItemOriginal.dISSValorDeducao
    dISSValorOutrasRet = objTribDocItemOriginal.dISSValorOutrasRet
    dISSValorDescIncond = objTribDocItemOriginal.dISSValorDescIncond
    dISSValorDescCond = objTribDocItemOriginal.dISSValorDescCond
    dISSValorRet = objTribDocItemOriginal.dISSValorRet
    iISSIndExigibilidade = objTribDocItemOriginal.iISSIndExigibilidade
    sISSCodServ = objTribDocItemOriginal.sISSCodServ
    sISSMunicIncidImp = objTribDocItemOriginal.sISSMunicIncidImp
    iISSCodPais = objTribDocItemOriginal.iISSCodPais
    sISSNumProcesso = objTribDocItemOriginal.sISSNumProcesso
    iISSIndIncentivo = objTribDocItemOriginal.iISSIndIncentivo
    iISSValorDeducaoManual = objTribDocItemOriginal.iISSValorDeducaoManual
    iISSValorOutrasRetManual = objTribDocItemOriginal.iISSValorOutrasRetManual
    iISSValorDescIncondManual = objTribDocItemOriginal.iISSValorDescIncondManual
    iISSValorDescCondManual = objTribDocItemOriginal.iISSValorDescCondManual
    iISSValorRetManual = objTribDocItemOriginal.iISSValorRetManual
    iISSIndExigibilidadeManual = objTribDocItemOriginal.iISSIndExigibilidadeManual
    iISSCodServManual = objTribDocItemOriginal.iISSCodServManual
    iISSMunicIncidImpManual = objTribDocItemOriginal.iISSMunicIncidImpManual
    iISSCodPaisManual = objTribDocItemOriginal.iISSCodPaisManual
    iISSNumProcessoManual = objTribDocItemOriginal.iISSNumProcessoManual
    iISSIndIncentivoManual = objTribDocItemOriginal.iISSIndIncentivoManual
    
    'NT 2015.003 (ICMS Interestadual)
    sCEST = objTribDocItemOriginal.sCEST
    dICMSInterestBCUFDest = objTribDocItemOriginal.dICMSInterestBCUFDest
    iICMSInterestBCUFDestManual = objTribDocItemOriginal.iICMSInterestBCUFDestManual
    dICMSInterestPercFCPUFDest = objTribDocItemOriginal.dICMSInterestPercFCPUFDest
    iICMSInterestPercFCPUFDestManual = objTribDocItemOriginal.iICMSInterestPercFCPUFDestManual
    dICMSInterestAliqUFDest = objTribDocItemOriginal.dICMSInterestAliqUFDest
    iICMSInterestAliqUFDestManual = objTribDocItemOriginal.iICMSInterestAliqUFDestManual
    dICMSInterestAliq = objTribDocItemOriginal.dICMSInterestAliq
    iICMSInterestAliqManual = objTribDocItemOriginal.iICMSInterestAliqManual
    dICMSInterestPercPartilha = objTribDocItemOriginal.dICMSInterestPercPartilha
    iICMSInterestPercPartilhaManual = objTribDocItemOriginal.iICMSInterestPercPartilhaManual
    dICMSInterestVlrFCPUFDest = objTribDocItemOriginal.dICMSInterestVlrFCPUFDest
    iICMSInterestVlrFCPUFDestManual = objTribDocItemOriginal.iICMSInterestVlrFCPUFDestManual
    dICMSInterestVlrUFDest = objTribDocItemOriginal.dICMSInterestVlrUFDest
    iICMSInterestVlrUFDestManual = objTribDocItemOriginal.iICMSInterestVlrUFDestManual
    dICMSInterestVlrUFRemet = objTribDocItemOriginal.dICMSInterestVlrUFRemet
    iICMSInterestVlrUFRemetManual = objTribDocItemOriginal.iICMSInterestVlrUFRemetManual
    
    'UA. Tributos Devolvidos (para o item da NF-e)
    dpDevol = objTribDocItemOriginal.dpDevol
    ipDevolManual = objTribDocItemOriginal.ipDevolManual
    dIPIVlrDevolvido = objTribDocItemOriginal.dIPIVlrDevolvido
    iIPIVlrDevolvidoManual = objTribDocItemOriginal.iIPIVlrDevolvidoManual

    'nfe 4.00
    dICMSInterestBCFCPUFDest = objTribDocItemOriginal.dICMSInterestBCFCPUFDest
    iICMSInterestBCFCPUFDestManual = objTribDocItemOriginal.iICMSInterestBCFCPUFDestManual
    dICMSvBCFCPST = objTribDocItemOriginal.dICMSvBCFCPST
    iICMSvBCFCPSTManual = objTribDocItemOriginal.iICMSvBCFCPSTManual
    dICMSpFCPST = objTribDocItemOriginal.dICMSpFCPST
    iICMSpFCPSTManual = objTribDocItemOriginal.iICMSpFCPSTManual
    dICMSvFCPST = objTribDocItemOriginal.dICMSvFCPST
    iICMSvFCPSTManual = objTribDocItemOriginal.iICMSvFCPSTManual
    dICMSvBCFCPSTRet = objTribDocItemOriginal.dICMSvBCFCPSTRet
    iICMSvBCFCPSTRetManual = objTribDocItemOriginal.iICMSvBCFCPSTRetManual
    dICMSpFCPSTRet = objTribDocItemOriginal.dICMSpFCPSTRet
    iICMSpFCPSTRetManual = objTribDocItemOriginal.iICMSpFCPSTRetManual
    dICMSvFCPSTRet = objTribDocItemOriginal.dICMSvFCPSTRet
    iICMSvFCPSTRetManual = objTribDocItemOriginal.iICMSvFCPSTRetManual
    dICMSSTCobrAntAliquota = objTribDocItemOriginal.dICMSSTCobrAntAliquota
    iICMSSTCobrAntAliquotaManual = objTribDocItemOriginal.iICMSSTCobrAntAliquotaManual
    dICMSvBCFCP = objTribDocItemOriginal.dICMSvBCFCP
    iICMSvBCFCPManual = objTribDocItemOriginal.iICMSvBCFCPManual
    dICMSpFCP = objTribDocItemOriginal.dICMSpFCP
    iICMSpFCPManual = objTribDocItemOriginal.iICMSpFCPManual
    dICMSvFCP = objTribDocItemOriginal.dICMSvFCP
    iICMSvFCPManual = objTribDocItemOriginal.iICMSvFCPManual
    scBenef = objTribDocItemOriginal.scBenef
    sindEscala = objTribDocItemOriginal.sindEscala
    sCNPJFab = objTribDocItemOriginal.sCNPJFab
    
    'EC 52/2017
    iICMSSTBaseDupla = objTribDocItemOriginal.iICMSSTBaseDupla
    iICMSSTBaseDuplaManual = objTribDocItemOriginal.iICMSSTBaseDuplaManual
    
    'EC 87/2015 http://www.fazenda.mg.gov.br/empresas/legislacao_tributaria/orientacao/orientacao_002_2016.pdf
    iICMSInterestBaseDupla = objTribDocItemOriginal.iICMSInterestBaseDupla
    iICMSInterestBaseDuplaManual = objTribDocItemOriginal.iICMSInterestBaseDuplaManual
    
    icBenefManual = objTribDocItemOriginal.icBenefManual
    
    Exit Sub

Erro_Copia2:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202901)

    End Select

    Exit Sub
    
End Sub

Public Function ObterVar_Padrao(ByVal sIdentificador As String, ByVal objVariavelCalculo As ClassVariavelCalculo, Optional ByVal objClassP As Object = Nothing) As Long

Dim lErro As Long, objProduto As New ClassProduto
Dim objClass As Object

On Error GoTo Erro_ObterVar_Padrao

    If objClassP Is Nothing Then
        Set objClass = Me
    Else
        Set objClass = objClassP
    End If

    Select Case sIdentificador

        Case "NAT_OPERACAO"
            objVariavelCalculo.vValor = objClass.sNaturezaOp
            objVariavelCalculo.iStatus = objClass.iNaturezaOpManual
                    
        Case "TIPO_TRIB"
            objVariavelCalculo.vValor = objClass.iTipoTributacao
            objVariavelCalculo.iStatus = objClass.iTipoTributacaoManual

        Case "PRODUTO_CODIGO"
            objVariavelCalculo.vValor = objClass.sProduto
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_OBJ"
            
            objProduto.sCodigo = objClass.sProduto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 202856
            
            Set objVariavelCalculo.vValor = objProduto
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_OUTROS_VALORES"
            objVariavelCalculo.vValor = objClass.dValorFreteItem + objClass.dValorSeguroItem + objClass.dValorOutrasDespesasItem - objClass.dValorDescontoItem
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        'para ICMS
        Case "ICMS_TIPO"
            objVariavelCalculo.vValor = objClass.iICMSTipo
            objVariavelCalculo.iStatus = objClass.iICMSTipoManual
        
        Case "ICMS_BASE"
            objVariavelCalculo.vValor = objClass.dICMSBase
            objVariavelCalculo.iStatus = objClass.iICMSBaseManual
        
        Case "ICMS_RED_BASE"
            objVariavelCalculo.vValor = objClass.dICMSPercRedBase
            objVariavelCalculo.iStatus = objClass.iICMSPercRedBaseManual
        
        Case "ICMS_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dICMSAliquota
            objVariavelCalculo.iStatus = objClass.iICMSAliquotaManual
        
        Case "ICMS_VALOR"
            objVariavelCalculo.vValor = objClass.dICMSValor
            objVariavelCalculo.iStatus = objClass.iICMSValorManual
        
        Case "ICMS_BASE_SUBST"
            objVariavelCalculo.vValor = objClass.dICMSSubstBase
            objVariavelCalculo.iStatus = objClass.iICMSSubstBaseManual
            
        Case "ICMS_BASE_ST_COBR_ANT"
            objVariavelCalculo.vValor = objClass.dICMSSTCobrAntBase
            objVariavelCalculo.iStatus = objClass.iICMSSTCobrAntBaseManual
            
        Case "ICMS_RED_BASE_SUBST"
            objVariavelCalculo.vValor = objClass.dICMSSubstPercRedBase
            objVariavelCalculo.iStatus = objClass.iICMSSubstPercRedBaseManual
            
        Case "ICMS_ALIQUOTA_SUBST"
            objVariavelCalculo.vValor = objClass.dICMSSubstAliquota
            objVariavelCalculo.iStatus = objClass.iICMSSubstAliquotaManual
        
        Case "ICMS_VALOR_SUBST"
            objVariavelCalculo.vValor = objClass.dICMSSubstValor
            objVariavelCalculo.iStatus = objClass.iICMSSubstValorManual
        
        Case "ICMS_VALOR_ST_COBR_ANT"
            objVariavelCalculo.vValor = objClass.dICMSSTCobrAntValor
            objVariavelCalculo.iStatus = objClass.iICMSSTCobrAntValorManual
        
        Case "ICMS_MARG_SUBST"
            objVariavelCalculo.vValor = objClass.dICMSSubstPercMVA
            objVariavelCalculo.iStatus = objClass.iICMSSubstPercMVAManual
        
        Case "ICMS_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dICMSCredito
            objVariavelCalculo.iStatus = objClass.iICMSCreditoManual
        
        'para IPI
        Case "IPI_TIPO"
            objVariavelCalculo.vValor = objClass.iIPITipo
            objVariavelCalculo.iStatus = objClass.iIPITipoManual
        
        Case "IPI_BASE"
            objVariavelCalculo.vValor = objClass.dIPIBaseCalculo
            objVariavelCalculo.iStatus = objClass.iIPIBaseManual
        
        Case "IPI_RED_BASE"
            objVariavelCalculo.vValor = objClass.dIPIPercRedBase
            objVariavelCalculo.iStatus = objClass.iIPIPercRedBaseManual
        
        Case "IPI_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dIPIAliquota
            objVariavelCalculo.iStatus = objClass.iIPIAliquotaManual
        
        Case "IPI_VALOR"
            objVariavelCalculo.vValor = objClass.dIPIValor
            objVariavelCalculo.iStatus = objClass.iIPIValorManual

        Case "IPI_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dIPICredito
            objVariavelCalculo.iStatus = objClass.iIPICreditoManual
        
        Case "PIS_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dPISCredito
            objVariavelCalculo.iStatus = objClass.iPISCreditoManual
        
        Case "COFINS_VALOR_CREDITO"
            objVariavelCalculo.vValor = objClass.dCOFINSCredito
            objVariavelCalculo.iStatus = objClass.iCOFINSCreditoManual
        
        Case "DESP_IMP_ICMS_BASE"
            objVariavelCalculo.vValor = objClass.dDespImpICMSBase
            objVariavelCalculo.iStatus = objClass.iDespImpICMSBaseManual
        
        Case "DESP_IMP_ICMS_VALOR"
            objVariavelCalculo.vValor = objClass.dDespImpICMSValor
            objVariavelCalculo.iStatus = objClass.iDespImpICMSValorManual
        
        Case "DESP_IMP_ICMS_CREDITO"
            objVariavelCalculo.vValor = objClass.dDespImpICMSCredito
            objVariavelCalculo.iStatus = objClass.iDespImpICMSCreditoManual
        
        Case "ICMS_ALIQUOTA_ADICAO_DI"
            objVariavelCalculo.vValor = objClass.dICMSAliquotaAdicaoDI
            objVariavelCalculo.iStatus = objClass.iICMSAliquotaAdicaoDIManual
        
        Case "ICMS_PERCREDBASE_ADICAO_DI"
            objVariavelCalculo.vValor = objClass.dICMSPercRedBaseAdicaoDI
            objVariavelCalculo.iStatus = objClass.iICMSPercRedBaseAdicaoDIManual
        
        Case "ICMS_ST_RED_BASE"
            objVariavelCalculo.vValor = objClass.dICMSSubstPercRedBase
            objVariavelCalculo.iStatus = objClass.iICMSSubstPercRedBaseManual
        
        Case "ICMS_ST_PERC_MVA"
            objVariavelCalculo.vValor = objClass.dICMSSubstPercMVA
            objVariavelCalculo.iStatus = objClass.iICMSSubstPercMVAManual
 
        Case "PIS_TIPO"
            objVariavelCalculo.vValor = objClass.iPISTipo
            objVariavelCalculo.iStatus = objClass.iPISTipoManual
        
        Case "PIS_TIPO_CALCULO"
            objVariavelCalculo.vValor = objClass.iPISTipoCalculo
            objVariavelCalculo.iStatus = objClass.iPISTipoCalculoManual
        
        Case "PIS_BASE"
            objVariavelCalculo.vValor = objClass.dPISBase
            objVariavelCalculo.iStatus = objClass.iPISBaseManual

        Case "PIS_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dPISAliquota
            objVariavelCalculo.iStatus = objClass.iPISAliquotaManual

        Case "PIS_ALIQUOTA_ADICAO_DI"
            objVariavelCalculo.vValor = objClass.dPISAliquotaAdicaoDI
            objVariavelCalculo.iStatus = objClass.iPISAliquotaAdicaoDIManual
        
        Case "PIS_ALIQUOTA_VALOR"
            objVariavelCalculo.vValor = objClass.dPISAliquotaValor
            objVariavelCalculo.iStatus = objClass.iPISAliquotaValorManual

        Case "PIS_QTDE"
            objVariavelCalculo.vValor = objClass.dPISQtde
            objVariavelCalculo.iStatus = objClass.iPISQtdeManual

        Case "PIS_VALOR"
            objVariavelCalculo.vValor = objClass.dPISValor
            objVariavelCalculo.iStatus = objClass.iPISValorManual
            
        Case "PIS_ST_TIPO_CALCULO"
            objVariavelCalculo.vValor = objClass.iPISSTTipoCalculo
            objVariavelCalculo.iStatus = objClass.iPISSTTipoCalculoManual
        
        Case "PIS_ST_BASE"
            objVariavelCalculo.vValor = objClass.dPISSTBase
            objVariavelCalculo.iStatus = objClass.iPISSTBaseManual

        Case "PIS_ST_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dPISSTAliquota
            objVariavelCalculo.iStatus = objClass.iPISSTAliquotaManual

        Case "PIS_ST_ALIQUOTA_VALOR"
            objVariavelCalculo.vValor = objClass.dPISSTAliquotaValor
            objVariavelCalculo.iStatus = objClass.iPISSTAliquotaValorManual

        Case "PIS_ST_QTDE"
            objVariavelCalculo.vValor = objClass.dPISSTQtde
            objVariavelCalculo.iStatus = objClass.iPISSTQtdeManual

        Case "PIS_ST_VALOR"
            objVariavelCalculo.vValor = objClass.dPISSTValor
            objVariavelCalculo.iStatus = objClass.iPISSTValorManual
            
        Case "COFINS_TIPO"
            objVariavelCalculo.vValor = objClass.iCOFINSTipo
            objVariavelCalculo.iStatus = objClass.iCOFINSTipoManual
        
        Case "COFINS_TIPO_CALCULO"
            objVariavelCalculo.vValor = objClass.iCOFINSTipoCalculo
            objVariavelCalculo.iStatus = objClass.iCOFINSTipoCalculoManual
        
        Case "COFINS_BASE"
            objVariavelCalculo.vValor = objClass.dCOFINSBase
            objVariavelCalculo.iStatus = objClass.iCOFINSBaseManual

        Case "COFINS_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dCOFINSAliquota
            objVariavelCalculo.iStatus = objClass.iCOFINSAliquotaManual
            
        Case "COFINS_ALIQUOTA_ADICAO_DI"
            objVariavelCalculo.vValor = objClass.dCOFINSAliquotaAdicaoDI
            objVariavelCalculo.iStatus = objClass.iCOFINSAliquotaAdicaoDIManual
        
        Case "COFINS_ALIQUOTA_VALOR"
            objVariavelCalculo.vValor = objClass.dCOFINSAliquotaValor
            objVariavelCalculo.iStatus = objClass.iCOFINSAliquotaValorManual
            
        Case "COFINS_QTDE"
            objVariavelCalculo.vValor = objClass.dCOFINSQtde
            objVariavelCalculo.iStatus = objClass.iCOFINSQtdeManual

        Case "COFINS_VALOR"
            objVariavelCalculo.vValor = objClass.dCOFINSValor
            objVariavelCalculo.iStatus = objClass.iCOFINSValorManual
            
        Case "COFINS_ST_TIPO_CALCULO"
            objVariavelCalculo.vValor = objClass.iCOFINSSTTipoCalculo
            objVariavelCalculo.iStatus = objClass.iCOFINSSTTipoCalculoManual
        
        Case "COFINS_ST_BASE"
            objVariavelCalculo.vValor = objClass.dCOFINSSTBase
            objVariavelCalculo.iStatus = objClass.iCOFINSSTBaseManual

        Case "COFINS_ST_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dCOFINSSTAliquota
            objVariavelCalculo.iStatus = objClass.iCOFINSSTAliquotaManual

        Case "COFINS_ST_ALIQUOTA_VALOR"
            objVariavelCalculo.vValor = objClass.dCOFINSSTAliquotaValor
            objVariavelCalculo.iStatus = objClass.iCOFINSSTAliquotaValorManual

        Case "COFINS_ST_QTDE"
            objVariavelCalculo.vValor = objClass.dCOFINSSTQtde
            objVariavelCalculo.iStatus = objClass.iCOFINSSTQtdeManual

        Case "COFINS_ST_VALOR"
            objVariavelCalculo.vValor = objClass.dCOFINSSTValor
            objVariavelCalculo.iStatus = objClass.iCOFINSSTValorManual

        Case "ICMS_CST"
            objVariavelCalculo.vValor = objClass.sCST
            objVariavelCalculo.iStatus = objClass.iCSTManual

        Case "ISSQN"
            objVariavelCalculo.vValor = objClass.sISSQN
            objVariavelCalculo.iStatus = objClass.iISSQNManual

        Case "ISS_BASE"
            objVariavelCalculo.vValor = objClass.dISSBase
            objVariavelCalculo.iStatus = objClass.iISSBaseManual
            
        Case "ISS_ALIQUOTA"
            objVariavelCalculo.vValor = objClass.dISSAliquota
            objVariavelCalculo.iStatus = objClass.iISSAliquotaManual
        
        Case "ISS_VALOR"
            objVariavelCalculo.vValor = objClass.dISSValor
            objVariavelCalculo.iStatus = objClass.iISSValorManual
            
        Case "ISS_CIDADE_IBGE"
            objVariavelCalculo.vValor = objClass.sISSCidadeIBGE
            objVariavelCalculo.iStatus = objClass.iISSCidadeIBGEManual
        
        Case "II_BASE"
            objVariavelCalculo.vValor = objClass.dIIBase
            objVariavelCalculo.iStatus = objClass.iIIBaseManual
            
        Case "II_DEPESAS_ADUANEIRA"
            objVariavelCalculo.vValor = objClass.dIIDespAduaneira
            objVariavelCalculo.iStatus = objClass.iIIDespAduaneiraManual
        
        Case "II_IOF"
            objVariavelCalculo.vValor = objClass.dIIIOF
            objVariavelCalculo.iStatus = objClass.iIIIOFManual
        
        Case "II_VALOR"
            objVariavelCalculo.vValor = objClass.dIIValor
            objVariavelCalculo.iStatus = objClass.iIIValorManual
        
        Case "ICMS_BASE_MODALIDADE"
            objVariavelCalculo.vValor = objClass.iICMSBaseModalidade
            objVariavelCalculo.iStatus = objClass.iICMSBaseModalidadeManual
            
        Case "ICMS_ST_BASE_MODALIDADE"
            objVariavelCalculo.vValor = objClass.iICMSSubstBaseModalidade
            objVariavelCalculo.iStatus = objClass.iICMSSubstBaseModalidadeManual
            
         Case "IPI_ENQUADRAMENTO_CLASSE"
            objVariavelCalculo.vValor = objClass.sIPIEnquadramentoClasse
            objVariavelCalculo.iStatus = objClass.iIPIEnquadramentoClasseManual
            
         Case "IPI_ENQUADRAMENTO_CODIGO"
            objVariavelCalculo.vValor = objClass.sIPIEnquadramentoCodigo
            objVariavelCalculo.iStatus = objClass.iIPIEnquadramentoCodigoManual
            
         Case "IPI_CNPJ_PRODUTOR"
            objVariavelCalculo.vValor = objClass.sIPICNPJProdutor
            objVariavelCalculo.iStatus = objClass.iIPICNPJProdutorManual
            
         Case "IPI_SELO_CODIGO"
            objVariavelCalculo.vValor = objClass.sIPISeloCodigo
            objVariavelCalculo.iStatus = objClass.iIPISeloCodigoManual
            
        Case "IPI_SELO_QTDE"
            objVariavelCalculo.vValor = objClass.lIPISeloQtde
            objVariavelCalculo.iStatus = objClass.iIPISeloQtdeManual
            
        Case "IPI_TIPO_CALCULO"
            objVariavelCalculo.vValor = objClass.iIPITipoCalculo
            objVariavelCalculo.iStatus = objClass.iIPITipoCalculoManual
            
        Case "IPI_UM_PADRAO_QTDE"
            objVariavelCalculo.vValor = objClass.dIPIUnidadePadraoQtde
            objVariavelCalculo.iStatus = objClass.iIPIUnidadePadraoQtdeManual
            
        Case "IPI_UM_PADRAO_VALOR"
            objVariavelCalculo.vValor = objClass.dIPIUnidadePadraoValor
            objVariavelCalculo.iStatus = objClass.iIPIUnidadePadraoValorManual
            
        Case "IPI_ALIQUOTA_ADICAO_DI"
            objVariavelCalculo.vValor = objClass.dIPIAliquotaAdicaoDI
        
        Case "IPI_VALORUNIT_ADICAO_DI"
            objVariavelCalculo.vValor = objClass.dIPIValorUnitAdicaoDI
        
        Case "ORIGEM_MERCADORIA"
            objVariavelCalculo.vValor = objClass.iOrigemMercadoria
            objVariavelCalculo.iStatus = objClass.iOrigemMercadoriaManual
        
        Case "FCI"
            objVariavelCalculo.vValor = objClass.sFCI
            objVariavelCalculo.iStatus = objClass.iFCIManual
        
        Case "VALOR_FRETE_ITEM"
            objVariavelCalculo.vValor = objClass.dValorFreteItem
            objVariavelCalculo.iStatus = objClass.iValorFreteItemManual
        
        Case "VALOR_SEGURO_ITEM"
            objVariavelCalculo.vValor = objClass.dValorSeguroItem
            objVariavelCalculo.iStatus = objClass.iValorSeguroItemManual
        
        Case "VALOR_DESCONTO_ITEM"
            objVariavelCalculo.vValor = objClass.dValorDescontoItem
            objVariavelCalculo.iStatus = objClass.iValorDescontoItemManual
        
        Case "VALOR_DESCONTO_GRID"
            objVariavelCalculo.vValor = objClass.dDescontoGrid
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case "VALOR_DESPESAS_ITEM"
            objVariavelCalculo.vValor = objClass.dValorOutrasDespesasItem
            objVariavelCalculo.iStatus = objClass.iValorOutrasDespesasItemManual
        
        Case "PRODUTO_QTDE_UM_TRIB"
            objVariavelCalculo.vValor = objClass.dQtdTrib
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
        'Novos 13/12/2010
        Case "ICMS_PERC_BASE_OPER_PROP"
            objVariavelCalculo.vValor = objClass.dICMSpercBaseOperacaoPropria
            objVariavelCalculo.iStatus = objClass.iICMSpercBaseOperacaoPropriaManual
        
        Case "ICMS_ST_UF_DEVIDO"
            objVariavelCalculo.vValor = objClass.sICMSUFDevidoST
            objVariavelCalculo.iStatus = objClass.iICMSUFDevidoSTManual
        
        Case "ICMS_ST_BASE_RETIDO_UF"
            objVariavelCalculo.vValor = objClass.dICMSvBCSTRet
            objVariavelCalculo.iStatus = objClass.iICMSvBCSTRetManual
        
        Case "ICMS_ST_VALOR_RETIDO_UF"
            objVariavelCalculo.vValor = objClass.dICMSvICMSSTRet
            objVariavelCalculo.iStatus = objClass.iICMSvICMSSTRetManual
        
        Case "ICMS_ST_BASE_UF_DESTINO"
            objVariavelCalculo.vValor = objClass.dICMSvBCSTDest
            objVariavelCalculo.iStatus = objClass.iICMSvBCSTDestManual
        
        Case "ICMS_ST_VALOR_UF_DESTINO"
            objVariavelCalculo.vValor = objClass.dICMSvICMSSTDest
            objVariavelCalculo.iStatus = objClass.iICMSvICMSSTDestManual
        
        Case "ICMS_ALIQUOTA_CREDITO_SIMPLES"
            objVariavelCalculo.vValor = objClass.dICMSpCredSN
            objVariavelCalculo.iStatus = objClass.iICMSpCredSNManual
        
        Case "ICMS_VALOR_CREDITO_SIMPLES"
            objVariavelCalculo.vValor = objClass.dICMSvCredSN
            objVariavelCalculo.iStatus = objClass.iICMSvCredSNManual
        
        Case "ICMS_VALOR_ISENTO"
            objVariavelCalculo.vValor = objClass.dICMSValorIsento
            objVariavelCalculo.iStatus = objClass.iICMSValorIsentoManual
        
        Case "ICMS_MOTIVO_ISENTO"
            objVariavelCalculo.vValor = objClass.iICMSMotivo
            objVariavelCalculo.iStatus = objClass.iICMSMotivoManual
        
        Case "ICMS_TIPO_SIMPLES"
            objVariavelCalculo.vValor = objClass.iICMSSimplesTipo
            objVariavelCalculo.iStatus = objClass.iICMSSimplesTipoManual
        
        Case "ISS_TIPO"
            objVariavelCalculo.vValor = objClass.iISSTipo
            objVariavelCalculo.iStatus = objClass.iISSTipoManual
        
        Case "REGIME_TRIBUTARIO"
            objVariavelCalculo.vValor = objClass.iRegimeTributario
'            If objClassP Is Nothing Then
                objVariavelCalculo.iStatus = objClass.iRegimeTributarioManual
'            Else
'                objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'            End If
        
        Case "ISS_CST"
            objVariavelCalculo.vValor = objClass.sISSCST
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case "ICMS_CSOSN"
            objVariavelCalculo.vValor = objClass.sCSOSN
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
        
        Case "NAT_BC_CRED"
            objVariavelCalculo.vValor = objClass.sNatBCCred
            objVariavelCalculo.iStatus = objClass.iNatBCCredManual
        
        Case "TOT_TRIB"
            objVariavelCalculo.vValor = objClass.dTotTrib
            objVariavelCalculo.iStatus = objClass.iTotTribManual
        
        Case "TOT_TRIB_FED"
            objVariavelCalculo.vValor = objClass.dTotTribAliqFed
            objVariavelCalculo.iStatus = objClass.iTotTribAliqFedManual
        
        Case "TOT_TRIB_EST"
            objVariavelCalculo.vValor = objClass.dTotTribAliqEst
            objVariavelCalculo.iStatus = objClass.iTotTribAliqEstManual
        
        Case "TOT_TRIB_MUNIC"
            objVariavelCalculo.vValor = objClass.dTotTribAliqMunic
            objVariavelCalculo.iStatus = objClass.iTotTribAliqMunicManual
        
        Case "TOT_TRIB_TIPO"
            objVariavelCalculo.vValor = objClass.iTotTribTipo
            objVariavelCalculo.iStatus = objClass.iTotTribTipoManual
        
        'nfe 3.10
        Case "ICMS_51_VALOR_OP"
            objVariavelCalculo.vValor = objClass.dICMS51ValorOp
            objVariavelCalculo.iStatus = objClass.iICMS51ValorOpManual
        
        Case "ICMS_PERC_DIFER"
            objVariavelCalculo.vValor = objClass.dICMSPercDifer
            objVariavelCalculo.iStatus = objClass.iICMSPercDiferManual
        
        Case "ICMS_VALOR_DIF"
            objVariavelCalculo.vValor = objClass.dICMSValorDif
            objVariavelCalculo.iStatus = objClass.iICMSValorDifManual
        
        Case "ISS_VALOR_DEDUCAO"
            objVariavelCalculo.vValor = objClass.dISSValorDeducao
            objVariavelCalculo.iStatus = objClass.iISSValorDeducaoManual
        
        Case "ISS_VALOR_OUTRASRET"
            objVariavelCalculo.vValor = objClass.dISSValorOutrasRet
            objVariavelCalculo.iStatus = objClass.iISSValorOutrasRetManual

        Case "ISS_VALOR_DESC_INCOND"
            objVariavelCalculo.vValor = objClass.dISSValorDescIncond
            objVariavelCalculo.iStatus = objClass.iISSValorDescIncondManual
        
        Case "ISS_VALOR_DESC_COND"
            objVariavelCalculo.vValor = objClass.dISSValorDescCond
            objVariavelCalculo.iStatus = objClass.iISSValorDescCondManual
        
        Case "ISS_VALOR_RET"
            objVariavelCalculo.vValor = objClass.dISSValorRet
            objVariavelCalculo.iStatus = objClass.iISSValorRetManual
        
        Case "ISS_IND_EXIGIBILIDADE"
            objVariavelCalculo.vValor = objClass.iISSIndExigibilidade
            objVariavelCalculo.iStatus = objClass.iISSIndExigibilidadeManual
        
        Case "ISS_COD_SERV"
            objVariavelCalculo.vValor = objClass.sISSCodServ
            objVariavelCalculo.iStatus = objClass.iISSCodServManual
        
        Case "ISS_MUNIC_INCID_IMP"
            objVariavelCalculo.vValor = objClass.sISSMunicIncidImp
            objVariavelCalculo.iStatus = objClass.iISSMunicIncidImpManual
        
        Case "ISS_COD_PAIS"
            objVariavelCalculo.vValor = objClass.iISSCodPais
            objVariavelCalculo.iStatus = objClass.iISSCodPaisManual
        
        Case "ISS_NUM_PROCESSO"
            objVariavelCalculo.vValor = objClass.sISSNumProcesso
            objVariavelCalculo.iStatus = objClass.iISSNumProcessoManual
        
        Case "ISS_IND_INCENTIVO"
            objVariavelCalculo.vValor = objClass.iISSIndIncentivo
            objVariavelCalculo.iStatus = objClass.iISSIndIncentivoManual
        
        Case "PRODUTO_CEST"
            objVariavelCalculo.vValor = objClass.sCEST
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "ICMS_INTEREST_BC_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestBCUFDest
            objVariavelCalculo.iStatus = objClass.iICMSInterestBCUFDestManual
        
        Case "ICMS_INTEREST_PERC_FCP_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestPercFCPUFDest
            objVariavelCalculo.iStatus = objClass.iICMSInterestPercFCPUFDestManual
        
        Case "ICMS_INTEREST_ALIQ_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestAliqUFDest
            objVariavelCalculo.iStatus = objClass.iICMSInterestAliqUFDestManual
        
        Case "ICMS_INTEREST_ALIQ"
            objVariavelCalculo.vValor = objClass.dICMSInterestAliq
            objVariavelCalculo.iStatus = objClass.iICMSInterestAliqManual
        
        Case "ICMS_INTEREST_PERC_PARTILHA"
            objVariavelCalculo.vValor = objClass.dICMSInterestPercPartilha
            objVariavelCalculo.iStatus = objClass.iICMSInterestPercPartilhaManual

        Case "ICMS_INTEREST_VLR_FCP_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestVlrFCPUFDest
            objVariavelCalculo.iStatus = objClass.iICMSInterestVlrFCPUFDestManual

        Case "ICMS_INTEREST_VLR_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestVlrUFDest
            objVariavelCalculo.iStatus = objClass.iICMSInterestVlrUFDestManual

        Case "ICMS_INTEREST_VLR_UF_REMET"
            objVariavelCalculo.vValor = objClass.dICMSInterestVlrUFRemet
            objVariavelCalculo.iStatus = objClass.iICMSInterestVlrUFRemetManual
        
        'UA. Tributos Devolvidos (para o item da NF-e)
        Case "PERCENTUAL_DEVOLVIDO"
            objVariavelCalculo.vValor = objClass.dpDevol
            objVariavelCalculo.iStatus = objClass.ipDevolManual
        
        Case "IPI_VLR_DEVOLVIDO"
            objVariavelCalculo.vValor = objClass.dIPIVlrDevolvido
            objVariavelCalculo.iStatus = objClass.iIPIVlrDevolvidoManual
        
        'nfe 4.00
        Case "ICMS_INTEREST_BC_FCP_UF_DEST"
            objVariavelCalculo.vValor = objClass.dICMSInterestBCFCPUFDest
            objVariavelCalculo.iStatus = objClass.iICMSInterestBCFCPUFDestManual
        
        Case "ICMS_BC_FCP"
            objVariavelCalculo.vValor = objClass.dICMSvBCFCP
            objVariavelCalculo.iStatus = objClass.iICMSvBCFCPManual
        
        Case "ICMS_ALIQ_FCP"
            objVariavelCalculo.vValor = objClass.dICMSpFCP
            objVariavelCalculo.iStatus = objClass.iICMSpFCPManual
        
        Case "ICMS_VLR_FCP"
            objVariavelCalculo.vValor = objClass.dICMSvFCP
            objVariavelCalculo.iStatus = objClass.iICMSvFCPManual
        
        Case "ICMS_BC_FCP_ST"
            objVariavelCalculo.vValor = objClass.dICMSvBCFCPST
            objVariavelCalculo.iStatus = objClass.iICMSvBCFCPSTManual
        
        Case "ICMS_ALIQ_FCP_ST"
            objVariavelCalculo.vValor = objClass.dICMSpFCPST
            objVariavelCalculo.iStatus = objClass.iICMSpFCPSTManual
        
        Case "ICMS_VLR_FCP_ST"
            objVariavelCalculo.vValor = objClass.dICMSvFCPST
            objVariavelCalculo.iStatus = objClass.iICMSvFCPSTManual
        
        Case "ICMS_BC_FCP_ST_RET"
            objVariavelCalculo.vValor = objClass.dICMSvBCFCPSTRet
            objVariavelCalculo.iStatus = objClass.iICMSvBCFCPSTRetManual
        
        Case "ICMS_ALIQ_FCP_ST_RET"
            objVariavelCalculo.vValor = objClass.dICMSpFCPSTRet
            objVariavelCalculo.iStatus = objClass.iICMSpFCPSTRetManual
        
        Case "ICMS_VLR_FCP_ST_RET"
            objVariavelCalculo.vValor = objClass.dICMSvFCPSTRet
            objVariavelCalculo.iStatus = objClass.iICMSvFCPSTRetManual
        
        Case "ICMS_ALIQUOTA_ST_COBR_ANT"
            objVariavelCalculo.vValor = objClass.dICMSSTCobrAntAliquota
            objVariavelCalculo.iStatus = objClass.iICMSSTCobrAntAliquotaManual
        
        Case "ICMS_ST_BASE_DUPLA"
            objVariavelCalculo.vValor = objClass.iICMSSTBaseDupla
            objVariavelCalculo.iStatus = objClass.iICMSSTBaseDuplaManual
        
        Case "ICMS_INTEREST_BASE_DUPLA"
            objVariavelCalculo.vValor = objClass.iICMSInterestBaseDupla
            objVariavelCalculo.iStatus = objClass.iICMSInterestBaseDuplaManual
        
        Case "ICMS_CBENEF"
            objVariavelCalculo.vValor = objClass.scBenef
            objVariavelCalculo.iStatus = objClass.icBenefManual
        
        Case Else
            gError 202857
        
    End Select

    ObterVar_Padrao = SUCESSO

    Exit Function

Erro_ObterVar_Padrao:

    ObterVar_Padrao = gErr

    Select Case gErr
    
        Case 202856, 202857
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202858)

    End Select

    Exit Function

End Function

Public Function GuardarVar_Padrao(ByVal objVariavelCalculo As ClassVariavelCalculo, Optional ByVal objClassP As Object = Nothing) As Long

Dim lErro As Long, objProduto As New ClassProduto
Dim objClass As Object

On Error GoTo Erro_GuardarVar_Padrao

    If objClassP Is Nothing Then
        Set objClass = Me
    Else
        Set objClass = objClassP
    End If
    
    Select Case objVariavelCalculo.sIdentificador

        Case "NAT_OPERACAO"
            objClass.sNaturezaOp = objVariavelCalculo.vValor
            objClass.iNaturezaOpManual = objVariavelCalculo.iStatus
        
        Case "TIPO_TRIB"
            objClass.iTipoTributacao = objVariavelCalculo.vValor
            objClass.iTipoTributacaoManual = objVariavelCalculo.iStatus

        Case "PRODUTO_CODIGO"
        
        'para ICMS
        Case "ICMS_TIPO"
            objClass.iICMSTipo = objVariavelCalculo.vValor
            objClass.iICMSTipoManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE"
            objClass.dICMSBase = objVariavelCalculo.vValor
            objClass.iICMSBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_RED_BASE"
            objClass.dICMSPercRedBase = objVariavelCalculo.vValor
            objClass.iICMSPercRedBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQUOTA"
            objClass.dICMSAliquota = objVariavelCalculo.vValor
            objClass.iICMSAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR"
            objClass.dICMSValor = objVariavelCalculo.vValor
            objClass.iICMSValorManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE_SUBST"
            objClass.dICMSSubstBase = objVariavelCalculo.vValor
            objClass.iICMSSubstBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE_ST_COBR_ANT"
            objClass.dICMSSTCobrAntBase = objVariavelCalculo.vValor
            objClass.iICMSSTCobrAntBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_RED_BASE_SUBST"
            objClass.dICMSSubstPercRedBase = objVariavelCalculo.vValor
        
        Case "ICMS_ALIQUOTA_SUBST"
            objClass.dICMSSubstAliquota = objVariavelCalculo.vValor
            objClass.iICMSSubstAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_SUBST"
            objClass.dICMSSubstValor = objVariavelCalculo.vValor
            objClass.iICMSSubstValorManual = objVariavelCalculo.iStatus

        Case "ICMS_VALOR_ST_COBR_ANT"
            objClass.dICMSSTCobrAntValor = objVariavelCalculo.vValor
            objClass.iICMSSTCobrAntValorManual = objVariavelCalculo.iStatus

        Case "ICMS_MARG_SUBST"
            objClass.dICMSSubstPercMVA = objVariavelCalculo.vValor
            objClass.iICMSSubstPercMVAManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_CREDITO"
            objClass.dICMSCredito = objVariavelCalculo.vValor
            objClass.iICMSCreditoManual = objVariavelCalculo.iStatus
                
        'para IPI
        Case "IPI_TIPO"
            objClass.iIPITipo = objVariavelCalculo.vValor
            objClass.iIPITipoManual = objVariavelCalculo.iStatus
        
        Case "IPI_BASE"
            objClass.dIPIBaseCalculo = objVariavelCalculo.vValor
            objClass.iIPIBaseManual = objVariavelCalculo.iStatus
            
        Case "IPI_RED_BASE"
            objClass.dIPIPercRedBase = objVariavelCalculo.vValor
            objClass.iIPIPercRedBaseManual = objVariavelCalculo.iStatus
            
        Case "IPI_ALIQUOTA"
            objClass.dIPIAliquota = objVariavelCalculo.vValor
            objClass.iIPIAliquotaManual = objVariavelCalculo.iStatus
        
        Case "IPI_VALOR"
            objClass.dIPIValor = objVariavelCalculo.vValor
            objClass.iIPIValorManual = objVariavelCalculo.iStatus

        Case "IPI_VALOR_CREDITO"
            objClass.dIPICredito = objVariavelCalculo.vValor
            objClass.iIPICreditoManual = objVariavelCalculo.iStatus
        
        Case "PIS_VALOR_CREDITO"
            objClass.dPISCredito = objVariavelCalculo.vValor
            objClass.iPISCreditoManual = objVariavelCalculo.iStatus
        
        Case "COFINS_VALOR_CREDITO"
            objClass.dCOFINSCredito = objVariavelCalculo.vValor
            objClass.iCOFINSCreditoManual = objVariavelCalculo.iStatus
        
        Case "DESP_IMP_ICMS_BASE"
            objClass.dDespImpICMSBase = objVariavelCalculo.vValor
            objClass.iDespImpICMSBaseManual = objVariavelCalculo.iStatus
        
        Case "DESP_IMP_ICMS_VALOR"
            objClass.dDespImpICMSValor = objVariavelCalculo.vValor
            objClass.iDespImpICMSValorManual = objVariavelCalculo.iStatus
        
        Case "DESP_IMP_ICMS_CREDITO"
            objClass.dDespImpICMSCredito = objVariavelCalculo.vValor
            objClass.iDespImpICMSCreditoManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQUOTA_ADICAO_DI"
        Case "ICMS_PERCREDBASE_ADICAO_DI"

        Case "ICMS_ST_RED_BASE"
            objClass.dICMSSubstPercRedBase = objVariavelCalculo.vValor
            objClass.iICMSSubstPercRedBaseManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ST_PERC_MVA"
            objClass.dICMSSubstPercMVA = objVariavelCalculo.vValor
            objClass.iICMSSubstPercMVAManual = objVariavelCalculo.iStatus
 
        Case "PIS_TIPO"
            objClass.iPISTipo = objVariavelCalculo.vValor
            objClass.iPISTipoManual = objVariavelCalculo.iStatus
        
        Case "PIS_TIPO_CALCULO"
            objClass.iPISTipoCalculo = objVariavelCalculo.vValor
            objClass.iPISTipoCalculoManual = objVariavelCalculo.iStatus
        
        Case "PIS_BASE"
            objClass.dPISBase = objVariavelCalculo.vValor
            objClass.iPISBaseManual = objVariavelCalculo.iStatus

        Case "PIS_ALIQUOTA"
            objClass.dPISAliquota = objVariavelCalculo.vValor
            objClass.iPISAliquotaManual = objVariavelCalculo.iStatus

        Case "PIS_ALIQUOTA_VALOR"
            objClass.dPISAliquotaValor = objVariavelCalculo.vValor
            objClass.iPISAliquotaValorManual = objVariavelCalculo.iStatus

        Case "PIS_QTDE"
            objClass.dPISQtde = objVariavelCalculo.vValor
            objClass.iPISQtdeManual = objVariavelCalculo.iStatus

        Case "PIS_VALOR"
            objClass.dPISValor = objVariavelCalculo.vValor
            objClass.iPISValorManual = objVariavelCalculo.iStatus
            
        Case "PIS_ST_TIPO_CALCULO"
            objClass.iPISSTTipoCalculo = objVariavelCalculo.vValor
            objClass.iPISSTTipoCalculoManual = objVariavelCalculo.iStatus
        
        Case "PIS_ST_BASE"
            objClass.dPISSTBase = objVariavelCalculo.vValor
            objClass.iPISSTBaseManual = objVariavelCalculo.iStatus

        Case "PIS_ST_ALIQUOTA"
            objClass.dPISSTAliquota = objVariavelCalculo.vValor
            objClass.iPISSTAliquotaManual = objVariavelCalculo.iStatus

        Case "PIS_ST_ALIQUOTA_VALOR"
            objClass.dPISSTAliquotaValor = objVariavelCalculo.vValor
            objClass.iPISSTAliquotaValorManual = objVariavelCalculo.iStatus

        Case "PIS_ST_QTDE"
            objClass.dPISSTQtde = objVariavelCalculo.vValor
            objClass.iPISSTQtdeManual = objVariavelCalculo.iStatus

        Case "PIS_ST_VALOR"
            objClass.dPISSTValor = objVariavelCalculo.vValor
            objClass.iPISSTValorManual = objVariavelCalculo.iStatus
            
        Case "COFINS_TIPO"
            objClass.iCOFINSTipo = objVariavelCalculo.vValor
            objClass.iCOFINSTipoManual = objVariavelCalculo.iStatus
        
        Case "COFINS_TIPO_CALCULO"
            objClass.iCOFINSTipoCalculo = objVariavelCalculo.vValor
            objClass.iCOFINSTipoCalculoManual = objVariavelCalculo.iStatus
        
        Case "COFINS_BASE"
            objClass.dCOFINSBase = objVariavelCalculo.vValor
            objClass.iCOFINSBaseManual = objVariavelCalculo.iStatus

        Case "COFINS_ALIQUOTA"
            objClass.dCOFINSAliquota = objVariavelCalculo.vValor
            objClass.iCOFINSAliquotaManual = objVariavelCalculo.iStatus

        Case "COFINS_ALIQUOTA_VALOR"
            objClass.dCOFINSAliquotaValor = objVariavelCalculo.vValor
            objClass.iCOFINSAliquotaValorManual = objVariavelCalculo.iStatus

        Case "COFINS_QTDE"
            objClass.dCOFINSQtde = objVariavelCalculo.vValor
            objClass.iCOFINSQtdeManual = objVariavelCalculo.iStatus

        Case "COFINS_VALOR"
            objClass.dCOFINSValor = objVariavelCalculo.vValor
            objClass.iCOFINSValorManual = objVariavelCalculo.iStatus
            
        Case "COFINS_ST_TIPO_CALCULO"
            objClass.iCOFINSSTTipoCalculo = objVariavelCalculo.vValor
            objClass.iCOFINSSTTipoCalculoManual = objVariavelCalculo.iStatus
        
        Case "COFINS_ST_BASE"
            objClass.dCOFINSSTBase = objVariavelCalculo.vValor
            objClass.iCOFINSSTBaseManual = objVariavelCalculo.iStatus

        Case "COFINS_ST_ALIQUOTA"
            objClass.dCOFINSSTAliquota = objVariavelCalculo.vValor
            objClass.iCOFINSSTAliquotaManual = objVariavelCalculo.iStatus

        Case "COFINS_ST_ALIQUOTA_VALOR"
            objClass.dCOFINSSTAliquotaValor = objVariavelCalculo.vValor
            objClass.iCOFINSSTAliquotaValorManual = objVariavelCalculo.iStatus

        Case "COFINS_ST_QTDE"
            objClass.dCOFINSSTQtde = objVariavelCalculo.vValor
            objClass.iCOFINSSTQtdeManual = objVariavelCalculo.iStatus

        Case "COFINS_ST_VALOR"
            objClass.dCOFINSSTValor = objVariavelCalculo.vValor
            objClass.iCOFINSSTValorManual = objVariavelCalculo.iStatus

        Case "ICMS_CST"
            objClass.sCST = objVariavelCalculo.vValor
            objClass.iCSTManual = objVariavelCalculo.iStatus

        Case "ISSQN"
            objClass.sISSQN = objVariavelCalculo.vValor
            objClass.iISSQNManual = objVariavelCalculo.iStatus

        Case "ISS_BASE"
            objClass.dISSBase = objVariavelCalculo.vValor
            objClass.iISSBaseManual = objVariavelCalculo.iStatus

        Case "ISS_ALIQUOTA"
            objClass.dISSAliquota = objVariavelCalculo.vValor
            objClass.iISSAliquotaManual = objVariavelCalculo.iStatus
        
        Case "ISS_VALOR"
            objClass.dISSValor = objVariavelCalculo.vValor
            objClass.iISSValorManual = objVariavelCalculo.iStatus
            
        Case "ISS_CIDADE_IBGE"
            objClass.sISSCidadeIBGE = objVariavelCalculo.vValor
            objClass.iISSCidadeIBGEManual = objVariavelCalculo.iStatus
        
        Case "II_BASE"
            objClass.dIIBase = objVariavelCalculo.vValor
            objClass.iIIBaseManual = objVariavelCalculo.iStatus
            
        Case "II_DEPESAS_ADUANEIRA"
            objClass.dIIDespAduaneira = objVariavelCalculo.vValor
            objClass.iIIDespAduaneiraManual = objVariavelCalculo.iStatus
        
        Case "II_IOF"
            objClass.dIIIOF = objVariavelCalculo.vValor
            objClass.iIIIOFManual = objVariavelCalculo.iStatus
        
        Case "II_VALOR"
            objClass.dIIValor = objVariavelCalculo.vValor
            objClass.iIIValorManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BASE_MODALIDADE"
            objClass.iICMSBaseModalidade = objVariavelCalculo.vValor
            objClass.iICMSBaseModalidadeManual = objVariavelCalculo.iStatus
            
        Case "ICMS_ST_BASE_MODALIDADE"
            objClass.iICMSSubstBaseModalidade = objVariavelCalculo.vValor
            objClass.iICMSSubstBaseModalidadeManual = objVariavelCalculo.iStatus
            
         Case "IPI_ENQUADRAMENTO_CLASSE"
            objClass.sIPIEnquadramentoClasse = objVariavelCalculo.vValor
            objClass.iIPIEnquadramentoClasseManual = objVariavelCalculo.iStatus
            
         Case "IPI_ENQUADRAMENTO_CODIGO"
            objClass.sIPIEnquadramentoCodigo = objVariavelCalculo.vValor
            objClass.iIPIEnquadramentoCodigoManual = objVariavelCalculo.iStatus
            
         Case "IPI_CNPJ_PRODUTOR"
            objClass.sIPICNPJProdutor = objVariavelCalculo.vValor
            objClass.iIPICNPJProdutorManual = objVariavelCalculo.iStatus
            
         Case "IPI_SELO_CODIGO"
            objClass.sIPISeloCodigo = objVariavelCalculo.vValor
            objClass.iIPISeloCodigoManual = objVariavelCalculo.iStatus
            
        Case "IPI_SELO_QTDE"
            objClass.lIPISeloQtde = objVariavelCalculo.vValor
            objClass.iIPISeloQtdeManual = objVariavelCalculo.iStatus
            
        Case "IPI_TIPO_CALCULO"
            objClass.iIPITipoCalculo = objVariavelCalculo.vValor
            objClass.iIPITipoCalculoManual = objVariavelCalculo.iStatus
            
        Case "IPI_UM_PADRAO_QTDE"
            objClass.dIPIUnidadePadraoQtde = objVariavelCalculo.vValor
            objClass.iIPIUnidadePadraoQtdeManual = objVariavelCalculo.iStatus
            
        Case "IPI_UM_PADRAO_VALOR"
            objClass.dIPIUnidadePadraoValor = objVariavelCalculo.vValor
            objClass.iIPIUnidadePadraoValorManual = objVariavelCalculo.iStatus
            
        Case "IPI_ALIQUOTA_ADICAO_DI", "IPI_VALORUNIT_ADICAO_DI"
        
        Case "ORIGEM_MERCADORIA"
            objClass.iOrigemMercadoria = objVariavelCalculo.vValor
            objClass.iOrigemMercadoriaManual = objVariavelCalculo.iStatus
        
        Case "FCI"
            objClass.sFCI = objVariavelCalculo.vValor
            objClass.iFCIManual = objVariavelCalculo.iStatus
        
        Case "VALOR_FRETE_ITEM"
            objClass.dValorFreteItem = objVariavelCalculo.vValor
            objClass.iValorFreteItemManual = objVariavelCalculo.iStatus
        
        Case "VALOR_SEGURO_ITEM"
            objClass.dValorSeguroItem = objVariavelCalculo.vValor
            objClass.iValorSeguroItemManual = objVariavelCalculo.iStatus
        
        Case "VALOR_DESCONTO_ITEM"
            objClass.dValorDescontoItem = objVariavelCalculo.vValor
            objClass.iValorDescontoItemManual = objVariavelCalculo.iStatus
        
        Case "VALOR_DESPESAS_ITEM"
            objClass.dValorOutrasDespesasItem = objVariavelCalculo.vValor
            objClass.iValorOutrasDespesasItemManual = objVariavelCalculo.iStatus
            
        'Novos 13/12/2010
        Case "ICMS_PERC_BASE_OPER_PROP"
            objClass.dICMSpercBaseOperacaoPropria = objVariavelCalculo.vValor
            objClass.iICMSpercBaseOperacaoPropriaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ST_UF_DEVIDO"
            objClass.sICMSUFDevidoST = objVariavelCalculo.vValor
            objClass.iICMSUFDevidoSTManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ST_BASE_RETIDO_UF"
            objClass.dICMSvBCSTRet = objVariavelCalculo.vValor
            objClass.iICMSvBCSTRetManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ST_VALOR_RETIDO_UF"
            objClass.dICMSvICMSSTRet = objVariavelCalculo.vValor
            objClass.iICMSvICMSSTRetManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ST_BASE_UF_DESTINO"
            objClass.dICMSvBCSTDest = objVariavelCalculo.vValor
            objClass.iICMSvBCSTDestManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ST_VALOR_UF_DESTINO"
            objClass.dICMSvICMSSTDest = objVariavelCalculo.vValor
            objClass.iICMSvICMSSTDestManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQUOTA_CREDITO_SIMPLES"
            objClass.dICMSpCredSN = objVariavelCalculo.vValor
            objClass.iICMSpCredSNManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_CREDITO_SIMPLES"
            objClass.dICMSvCredSN = objVariavelCalculo.vValor
            objClass.iICMSvCredSNManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_ISENTO"
            objClass.dICMSValorIsento = objVariavelCalculo.vValor
            objClass.iICMSValorIsentoManual = objVariavelCalculo.iStatus
        
        Case "ICMS_MOTIVO_ISENTO"
            objClass.iICMSMotivo = objVariavelCalculo.vValor
            objClass.iICMSMotivoManual = objVariavelCalculo.iStatus
        
        Case "ICMS_TIPO_SIMPLES"
            objClass.iICMSSimplesTipo = objVariavelCalculo.vValor
            objClass.iICMSSimplesTipoManual = objVariavelCalculo.iStatus
        
        Case "ISS_TIPO"
            objClass.iISSTipo = objVariavelCalculo.vValor
            objClass.iISSTipoManual = objVariavelCalculo.iStatus
        
        Case "REGIME_TRIBUTARIO"
            objClass.iRegimeTributario = objVariavelCalculo.vValor
            If objClassP Is Nothing Then objClass.iRegimeTributarioManual = objVariavelCalculo.iStatus
        
        Case "ISS_CST"
            objClass.sISSCST = objVariavelCalculo.vValor
        
        Case "ICMS_CSOSN"
            objClass.sCSOSN = objVariavelCalculo.vValor
            
        Case "NAT_BC_CRED"
            objClass.sNatBCCred = objVariavelCalculo.vValor
            objClass.iNatBCCredManual = objVariavelCalculo.iStatus
            
        Case "TOT_TRIB"
            objClass.dTotTrib = objVariavelCalculo.vValor
            objClass.iTotTribManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_FED"
            objClass.dTotTribAliqFed = objVariavelCalculo.vValor
            objClass.iTotTribAliqFedManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_EST"
            objClass.dTotTribAliqEst = objVariavelCalculo.vValor
            objClass.iTotTribAliqEstManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_MUNIC"
            objClass.dTotTribAliqMunic = objVariavelCalculo.vValor
            objClass.iTotTribAliqMunicManual = objVariavelCalculo.iStatus
        
        Case "TOT_TRIB_TIPO"
            objClass.iTotTribTipo = objVariavelCalculo.vValor
            objClass.iTotTribTipoManual = objVariavelCalculo.iStatus
        
        'nfe 3.10
        Case "ICMS_51_VALOR_OP"
            objClass.dICMS51ValorOp = objVariavelCalculo.vValor
            objClass.iICMS51ValorOpManual = objVariavelCalculo.iStatus
        
        Case "ICMS_PERC_DIFER"
            objClass.dICMSPercDifer = objVariavelCalculo.vValor
            objClass.iICMSPercDiferManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_DIF"
            objClass.dICMSValorDif = objVariavelCalculo.vValor
            objClass.iICMSValorDifManual = objVariavelCalculo.iStatus
        
        Case "ISS_VALOR_DEDUCAO"
            objClass.dISSValorDeducao = objVariavelCalculo.vValor
            objClass.iISSValorDeducaoManual = objVariavelCalculo.iStatus
        
        Case "ISS_VALOR_OUTRASRET"
            objClass.dISSValorOutrasRet = objVariavelCalculo.vValor
            objClass.iISSValorOutrasRetManual = objVariavelCalculo.iStatus

        Case "ISS_VALOR_DESC_INCOND"
            objClass.dISSValorDescIncond = objVariavelCalculo.vValor
            objClass.iISSValorDescIncondManual = objVariavelCalculo.iStatus
        
        Case "ISS_VALOR_DESC_COND"
            objClass.dISSValorDescCond = objVariavelCalculo.vValor
            objClass.iISSValorDescCondManual = objVariavelCalculo.iStatus
        
        Case "ISS_VALOR_RET"
            objClass.dISSValorRet = objVariavelCalculo.vValor
            objClass.iISSValorRetManual = objVariavelCalculo.iStatus
        
        Case "ISS_IND_EXIGIBILIDADE"
            objClass.iISSIndExigibilidade = objVariavelCalculo.vValor
            objClass.iISSIndExigibilidadeManual = objVariavelCalculo.iStatus
        
        Case "ISS_COD_SERV"
            objClass.sISSCodServ = objVariavelCalculo.vValor
            objClass.iISSCodServManual = objVariavelCalculo.iStatus
        
        Case "ISS_MUNIC_INCID_IMP"
            objClass.sISSMunicIncidImp = objVariavelCalculo.vValor
            objClass.iISSMunicIncidImpManual = objVariavelCalculo.iStatus
        
        Case "ISS_COD_PAIS"
            objClass.iISSCodPais = objVariavelCalculo.vValor
            objClass.iISSCodPaisManual = objVariavelCalculo.iStatus
        
        Case "ISS_NUM_PROCESSO"
            objClass.sISSNumProcesso = objVariavelCalculo.vValor
            objClass.iISSNumProcessoManual = objVariavelCalculo.iStatus
        
        Case "ISS_IND_INCENTIVO"
            objClass.iISSIndIncentivo = objVariavelCalculo.vValor
            objClass.iISSIndIncentivoManual = objVariavelCalculo.iStatus
        
        Case "PRODUTO_CEST"
            objClass.sCEST = objVariavelCalculo.vValor

        Case "ICMS_INTEREST_BC_UF_DEST"
            objClass.dICMSInterestBCUFDest = objVariavelCalculo.vValor
            objClass.iICMSInterestBCUFDestManual = objVariavelCalculo.iStatus
        
        Case "ICMS_INTEREST_PERC_FCP_UF_DEST"
            objClass.dICMSInterestPercFCPUFDest = objVariavelCalculo.vValor
            objClass.iICMSInterestPercFCPUFDestManual = objVariavelCalculo.iStatus
        
        Case "ICMS_INTEREST_ALIQ_UF_DEST"
            objClass.dICMSInterestAliqUFDest = objVariavelCalculo.vValor
            objClass.iICMSInterestAliqUFDestManual = objVariavelCalculo.iStatus
        
        Case "ICMS_INTEREST_ALIQ"
            objClass.dICMSInterestAliq = objVariavelCalculo.vValor
            objClass.iICMSInterestAliqManual = objVariavelCalculo.iStatus
        
        Case "ICMS_INTEREST_PERC_PARTILHA"
            objClass.dICMSInterestPercPartilha = objVariavelCalculo.vValor
            objClass.iICMSInterestPercPartilhaManual = objVariavelCalculo.iStatus

        Case "ICMS_INTEREST_VLR_FCP_UF_DEST"
            objClass.dICMSInterestVlrFCPUFDest = objVariavelCalculo.vValor
            objClass.iICMSInterestVlrFCPUFDestManual = objVariavelCalculo.iStatus

        Case "ICMS_INTEREST_VLR_UF_DEST"
            objClass.dICMSInterestVlrUFDest = objVariavelCalculo.vValor
            objClass.iICMSInterestVlrUFDestManual = objVariavelCalculo.iStatus

        Case "ICMS_INTEREST_VLR_UF_REMET"
            objClass.dICMSInterestVlrUFRemet = objVariavelCalculo.vValor
            objClass.iICMSInterestVlrUFRemetManual = objVariavelCalculo.iStatus
                
        'UA. Tributos Devolvidos (para o item da NF-e)
        Case "PERCENTUAL_DEVOLVIDO"
            objClass.dpDevol = objVariavelCalculo.vValor
            objClass.ipDevolManual = objVariavelCalculo.iStatus
        
        Case "IPI_VLR_DEVOLVIDO"
            objClass.dIPIVlrDevolvido = objVariavelCalculo.vValor
            objClass.iIPIVlrDevolvidoManual = objVariavelCalculo.iStatus
            
        'nfe 4.00
        Case "ICMS_INTEREST_BC_FCP_UF_DEST"
            objClass.dICMSInterestBCFCPUFDest = objVariavelCalculo.vValor
            objClass.iICMSInterestBCFCPUFDestManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BC_FCP"
            objClass.dICMSvBCFCP = objVariavelCalculo.vValor
            objClass.iICMSvBCFCPManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQ_FCP"
            objClass.dICMSpFCP = objVariavelCalculo.vValor
            objClass.iICMSpFCPManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VLR_FCP"
            objClass.dICMSvFCP = objVariavelCalculo.vValor
            objClass.iICMSvFCPManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BC_FCP_ST"
            objClass.dICMSvBCFCPST = objVariavelCalculo.vValor
            objClass.iICMSvBCFCPSTManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQ_FCP_ST"
            objClass.dICMSpFCPST = objVariavelCalculo.vValor
            objClass.iICMSpFCPSTManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VLR_FCP_ST"
            objClass.dICMSvFCPST = objVariavelCalculo.vValor
            objClass.iICMSvFCPSTManual = objVariavelCalculo.iStatus
        
        Case "ICMS_BC_FCP_ST_RET"
            objClass.dICMSvBCFCPSTRet = objVariavelCalculo.vValor
            objClass.iICMSvBCFCPSTRetManual = objVariavelCalculo.iStatus
        
        Case "ICMS_ALIQ_FCP_ST_RET"
            objClass.dICMSpFCPSTRet = objVariavelCalculo.vValor
            objClass.iICMSpFCPSTRetManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VLR_FCP_ST_RET"
            objClass.dICMSvFCPSTRet = objVariavelCalculo.vValor
            objClass.iICMSvFCPSTRetManual = objVariavelCalculo.iStatus
        
        Case "ICMS_VALOR_ST_COBR_ANT"
            objClass.dICMSSTCobrAntAliquota = objVariavelCalculo.vValor
            objClass.iICMSSTCobrAntAliquotaManual = objVariavelCalculo.iStatus
            
        Case "ICMS_ST_BASE_DUPLA"
            objClass.iICMSSTBaseDupla = objVariavelCalculo.vValor
            objClass.iICMSSTBaseDuplaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_INTEREST_BASE_DUPLA"
            objClass.iICMSInterestBaseDupla = objVariavelCalculo.vValor
            objClass.iICMSInterestBaseDuplaManual = objVariavelCalculo.iStatus
        
        Case "ICMS_CBENEF"
            objClass.scBenef = objVariavelCalculo.vValor
            objClass.icBenefManual = objVariavelCalculo.iStatus
        
        Case Else
            gError 202854
        
    End Select

    GuardarVar_Padrao = SUCESSO

    Exit Function

Erro_GuardarVar_Padrao:

    GuardarVar_Padrao = gErr

    Select Case gErr
    
        Case 202854
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202855)

    End Select

    Exit Function

End Function

Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseCalculoManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseCalculoManual
End Property

Private Function Altera_VAR_PREENCH(ByVal iStatus As Integer)
    iNaturezaOpManual = iStatus
    iTipoTributacaoManual = iStatus
    iIPITipoManual = iStatus
    iIPICodProdutoManual = iStatus
    iIPIBaseCalculoManual = iStatus
    iIPIPercRedBaseManual = iStatus
    iIPIAliquotaManual = iStatus
    iIPIValorManual = iStatus
    iIPICreditoManual = iStatus
    iICMSTipoManual = iStatus
    iICMSBaseManual = iStatus
    iICMSPercRedBaseManual = iStatus
    iICMSAliquotaManual = iStatus
    iICMSValorManual = iStatus
    iICMSCreditoManual = iStatus
    iICMSSubstBaseManual = iStatus
    iICMSSubstAliquotaManual = iStatus
    iICMSSubstValorManual = iStatus
    iPISCreditoManual = iStatus
    iCOFINSCreditoManual = iStatus
    iICMSAliquotaAdicaoDIManual = iStatus
    iICMSPercRedBaseAdicaoDIManual = iStatus
    iPISAliquotaAdicaoDIManual = iStatus
    iCOFINSAliquotaAdicaoDIManual = iStatus
    iDespImpICMSBaseManual = iStatus
    iDespImpICMSValorManual = iStatus
    iDespImpICMSCreditoManual = iStatus
    iICMSSubstPercRedBaseManual = iStatus
    iICMSSubstPercMVAManual = iStatus
    iPISTipoManual = iStatus
    iPISTipoCalculoManual = iStatus
    iPISBaseManual = iStatus
    iPISAliquotaManual = iStatus
    iPISAliquotaValorManual = iStatus
    iPISQtdeManual = iStatus
    iPISValorManual = iStatus
    iPISSTTipoCalculoManual = iStatus
    iPISSTBaseManual = iStatus
    iPISSTAliquotaManual = iStatus
    iPISSTAliquotaValorManual = iStatus
    iPISSTQtdeManual = iStatus
    iPISSTValorManual = iStatus
    iCOFINSTipoManual = iStatus
    iCOFINSTipoCalculoManual = iStatus
    iCOFINSBaseManual = iStatus
    iCOFINSAliquotaManual = iStatus
    iCOFINSAliquotaValorManual = iStatus
    iCOFINSQtdeManual = iStatus
    iCOFINSValorManual = iStatus
    iCOFINSSTTipoCalculoManual = iStatus
    iCOFINSSTBaseManual = iStatus
    iCOFINSSTAliquotaManual = iStatus
    iCOFINSSTAliquotaValorManual = iStatus
    iCOFINSSTQtdeManual = iStatus
    iCOFINSSTValorManual = iStatus
    iCSTManual = iStatus
    iISSQNManual = iStatus
    iISSBaseManual = iStatus
    iISSAliquotaManual = iStatus
    iISSCidadeIBGEManual = iStatus
    iISSValorManual = iStatus
    iIIBaseManual = iStatus
    iIIDespAduaneiraManual = iStatus
    iIIIOFManual = iStatus
    iIIValorManual = iStatus
    iICMSBaseModalidadeManual = iStatus
    iICMSSubstBaseModalidadeManual = iStatus
    iIPIEnquadramentoClasseManual = iStatus
    iIPIEnquadramentoCodigoManual = iStatus
    iIPICNPJProdutorManual = iStatus
    iIPISeloCodigoManual = iStatus
    iIPISeloQtdeManual = iStatus
    iIPITipoCalculoManual = iStatus
    iIPIUnidadePadraoQtdeManual = iStatus
    iIPIUnidadePadraoValorManual = iStatus
    iValorFreteItemManual = iStatus
    iValorSeguroItemManual = iStatus
    iValorOutrasDespesasItemManual = iStatus
    iValorDescontoItemManual = iStatus
    iOrigemMercadoriaManual = iStatus
    iFCIManual = iStatus
    iICMSSTCobrAntBaseManual = iStatus
    iICMSSTCobrAntValorManual = iStatus
    
    iICMSpercBaseOperacaoPropriaManual = iStatus
    iICMSUFDevidoSTManual = iStatus
    iICMSvBCSTRetManual = iStatus
    iICMSvICMSSTRetManual = iStatus
    iICMSvBCSTDestManual = iStatus
    iICMSvICMSSTDestManual = iStatus
    iICMSpCredSNManual = iStatus
    iICMSvCredSNManual = iStatus
    iICMSValorIsentoManual = iStatus
    iICMSMotivoManual = iStatus
    iICMSSimplesTipoManual = iStatus
    iISSTipoManual = iStatus
    iNatBCCredManual = iStatus
    iTotTribManual = iStatus
    iTotTribTipoManual = iStatus

    'nfe 3.10
    iICMS51ValorOpManual = iStatus
    iICMSPercDiferManual = iStatus
    iICMSValorDifManual = iStatus
    iISSValorDeducaoManual = iStatus
    iISSValorOutrasRetManual = iStatus
    iISSValorDescIncondManual = iStatus
    iISSValorDescCondManual = iStatus
    iISSValorRetManual = iStatus
    iISSIndExigibilidadeManual = iStatus
    iISSCodServManual = iStatus
    iISSMunicIncidImpManual = iStatus
    iISSCodPaisManual = iStatus
    iISSNumProcessoManual = iStatus
    iISSIndIncentivoManual = iStatus

    'IBPTax 0.0.9
    iTotTribAliqFedManual = iStatus
    iTotTribAliqEstManual = iStatus
    iTotTribAliqMunicManual = iStatus
    
    'NT 2015.003 (ICMS Interestadual)
    iICMSInterestBCUFDestManual = iStatus
    iICMSInterestPercFCPUFDestManual = iStatus
    iICMSInterestAliqUFDestManual = iStatus
    iICMSInterestAliqManual = iStatus
    iICMSInterestPercPartilhaManual = iStatus
    iICMSInterestVlrFCPUFDestManual = iStatus
    iICMSInterestVlrUFDestManual = iStatus
    iICMSInterestVlrUFRemetManual = iStatus

    'nfe 4.00
    iICMSInterestBCFCPUFDestManual = iStatus
    iICMSvBCFCPSTManual = iStatus
    iICMSpFCPSTManual = iStatus
    iICMSvFCPSTManual = iStatus
    iICMSvBCFCPSTRetManual = iStatus
    iICMSpFCPSTRetManual = iStatus
    iICMSvFCPSTRetManual = iStatus
    iICMSSTCobrAntAliquotaManual = iStatus
    iICMSvBCFCPManual = iStatus
    iICMSpFCPManual = iStatus
    iICMSvFCPManual = iStatus
End Function

Public Sub Coloca_Manual()
    Call Altera_VAR_PREENCH(VAR_PREENCH_MANUAL)
End Sub

Public Sub Coloca_Auto()
    Call Altera_VAR_PREENCH(VAR_PREENCH_AUTOMATICO)
End Sub

