VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTribItemNFAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'utilizada em telas de entrada de NFs

Implements ClassTributoDocItem

Public objItemNFOrig As ClassItemNF

Private mvarobjTributacaoDocItem As ClassTributacaoDocItemAux
Private mvarobjDoc As ClassTributoDoc 'link p/Doc

'Private mvariFilialEmpresa As Integer
Private mvariItem As Integer
Private mvarsProduto As String
Private mvardQuantidade As Double
Private mvardValorDesconto As Double
Private mvardValorDescGlobal As Double
Private mvardValorDespGlobal As Double
Private mvardPrecoTotal As Double
Private mvardPrecoUnitario As Double
Private mvardPrecoUnitarioMoeda As Double
Private mvarsUnidadeMed As String
Private mvarsProdutoDescricao As String

Private mvarsNaturezaOp As String
Private mvariNaturezaOpManual As Integer
Private mvariTipoTributacao As Integer
Private mvariTipoTributacaoManual As Integer
Private mvariIPITipo As Integer
Private mvariIPITipoManual As Integer
Private mvardIPIBaseCalculo As Double
Private mvariIPIBaseManual As Integer
Private mvardIPIPercRedBase As Double
Private mvariIPIPercRedBaseManual As Integer
Private mvardIPIAliquota As Double
Private mvariIPIAliquotaManual As Integer
Private mvardIPIValor As Double
Private mvariIPIValorManual As Integer
Private mvardIPICredito As Double
Private mvariIPICreditoManual As Integer
Private mvariICMSTipo As Integer
Private mvariICMSTipoManual As Integer
Private mvardICMSBase As Double
Private mvariICMSBaseManual As Integer
Private mvardICMSPercRedBase As Double
Private mvariICMSPercRedBaseManual As Integer
Private mvardICMSAliquota As Double
Private mvariICMSAliquotaManual As Integer
Private mvardICMSValor As Double
Private mvariICMSValorManual As Integer
Private mvardICMSSubstBase As Double
Private mvariICMSSubstBaseManual As Integer
'Private mvardICMSSubstPercRedBase As Double
'Private mvariICMSSubstBaseRedManual As Integer
Private mvardICMSSubstAliquota As Double
Private mvariICMSSubstAliquotaManual As Integer
Private mvardICMSSubstValor As Double
Private mvariICMSSubstValorManual As Integer
'Private mvardICMSSubstPercMVA As Double
Private mvardICMSCredito As Double
Private mvariICMSCreditoManual As Integer
Private mvardPISCredito As Double
Private mvariPISCreditoManual As Integer
Private mvardCOFINSCredito As Double
Private mvariCOFINSCreditoManual As Integer

'para notas de importacao
Public dDespImpValorRateado As Double 'valor rateado de despesas de importacao
Public dDespImpICMSBase As Double 'valor da base do icms deste item obtido pelo rateio de despesas de importacao
Public dDespImpICMSValor As Double 'valor da base do icms deste item obtido pelo rateio de despesas de importacao
Public dDespImpICMSCredito As Double 'valor do credito de icms deste item obtido pelo rateio de despesas de importacao
Public dICMSAliquotaAdicaoDI As Double
Public dICMSPercRedBaseAdicaoDI As Double
Public dPISAliquotaAdicaoDI As Double
Public dIPIAliquotaAdicaoDI As Double
Public dIPIValorUnitAdicaoDI As Double
Public dCOFINSAliquotaAdicaoDI As Double
Private mvarsIPICodProduto As String
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
Private mvarobjTributacao As New ClassTributacaoDocItem

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
Private mvariRegimeTributarioManual As Integer
Private mvarsISSCST As String
Private mvarsCSOSN As String
Private mvarsNatBCCred As String
Private mvariNatBCCredManual As Integer
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
'EC 87/2015
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

Public Property Let dCOFINSCredito(ByVal vData As Double)
    mvardCOFINSCredito = vData
End Property

Public Property Get dCOFINSCredito() As Double
    dCOFINSCredito = mvardCOFINSCredito
End Property

Public Property Let dPISCredito(ByVal vData As Double)
    mvardPISCredito = vData
End Property

Public Property Get dPISCredito() As Double
    dPISCredito = mvardPISCredito
End Property

Public Property Let dIPICredito(ByVal vData As Double)
    mvardIPICredito = vData
End Property

Public Property Get dIPICredito() As Double
    dIPICredito = mvardIPICredito
End Property

Public Property Let dICMSCredito(ByVal vData As Double)
    mvardICMSCredito = vData
End Property

Public Property Get dICMSCredito() As Double
    dICMSCredito = mvardICMSCredito
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let sProdutoDescricao(ByVal vData As String)
    mvarsProdutoDescricao = vData
End Property

Public Property Get sProdutoDescricao() As String
    sProdutoDescricao = mvarsProdutoDescricao
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorDescGlobal(ByVal vData As Double)
    mvardValorDescGlobal = vData
End Property

Public Property Get dValorDescGlobal() As Double
    dValorDescGlobal = mvardValorDescGlobal
End Property

Public Property Let dValorDespGlobal(ByVal vData As Double)
    mvardValorDespGlobal = vData
End Property

Public Property Get dValorDespGlobal() As Double
    dValorDespGlobal = mvardValorDespGlobal
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

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

'Public Property Let iFilialEmpresa(ByVal vData As Integer)
'    mvariFilialEmpresa = vData
'End Property
'
'Public Property Get iFilialEmpresa() As Integer
'    iFilialEmpresa = mvariFilialEmpresa
'End Property
'
Private Function ClassTributoDocItem_Ativar(objDoc As ClassTributoDoc) As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_Ativar

    Set mvarobjDoc = objDoc
    
    Set mvarobjTributacaoDocItem = New ClassTributacaoDocItemAux

    If mvarobjTributacaoDocItem.Ativar(Me) <> SUCESSO Then
        Set mvarobjTributacaoDocItem = Nothing
        Error 27823
    End If

    ClassTributoDocItem_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_Ativar:

    ClassTributoDocItem_Ativar = Err

    Select Case Err

        Case 27823

        Case Else
            lErro = Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error, 153703)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocItem_Alterado() As Long
    ClassTributoDocItem_Alterado = mvarobjTributacaoDocItem.Alterado
End Function

Private Function ClassTributoDocItem_Calculado() As Long
    ClassTributoDocItem_Calculado = mvarobjTributacaoDocItem.Calculado
End Function

Private Function ClassTributoDocItem_iJaCalculado() As Integer
    ClassTributoDocItem_iJaCalculado = mvarobjTributacaoDocItem.iJaCalculado
End Function

Private Function ClassTributoDocItem_Desativar() As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_Desativar

    If Not (mvarobjTributacaoDocItem Is Nothing) Then
        lErro = mvarobjTributacaoDocItem.Desativar()
        Set mvarobjTributacaoDocItem = Nothing
    End If

    ClassTributoDocItem_Desativar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_Desativar:

    ClassTributoDocItem_Desativar = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error, 153704)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocItem_GuardarVarTela(sIdentificador As String, vValor As Variant) As Long
Dim lErro As Long, objVariavelCalculo As New ClassVariavelCalculo
On Error GoTo Erro_ClassTributoDocItem_GuardarVarTela

    objVariavelCalculo.sIdentificador = sIdentificador
    objVariavelCalculo.vValor = vValor
    objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
    
    ClassTributoDocItem_GuardarVarTela = ClassTributoDocItem_GuardarVar(objVariavelCalculo)

    Exit Function

Erro_ClassTributoDocItem_GuardarVarTela:

    ClassTributoDocItem_GuardarVarTela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error, 153705)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDocItem_objDoc() As ClassTributoDoc
    Set ClassTributoDocItem_objDoc = mvarobjDoc
End Property

Private Property Get ClassTributoDocItem_objTributacaoDocItem() As ClassTributacaoDocItemAux
    Set ClassTributoDocItem_objTributacaoDocItem = mvarobjTributacaoDocItem
End Property

Private Function ClassTributoDocItem_ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long, objProduto As New ClassProduto
Dim objVarTipoDocInfo As New ClassVariavelCalculo

On Error GoTo Erro_ClassTributoDocItem_ObterVar

    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    objVariavelCalculo.sIdentificador = sIdentificador
    
    Select Case sIdentificador

        Case "OBJ_ITEM_NF_ORIG"
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
            Set objVariavelCalculo.vValor = objItemNFOrig

'        Case "NAT_OPERACAO"
'            objVariavelCalculo.vValor = mvarsNaturezaOp
'            objVariavelCalculo.iStatus = mvariNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            objVariavelCalculo.vValor = mvariTipoTributacao
'            objVariavelCalculo.iStatus = mvariTipoTributacaoManual
'
'        Case "PRODUTO_CODIGO"
'            objVariavelCalculo.vValor = mvarsProduto
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
        Case "PRODUTO_VALOR"
            lErro = mvarobjDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            If mvardQuantidade = 0 And objVarTipoDocInfo.vValor = DOCINFO_NFIEIMPCC Then
                objVariavelCalculo.vValor = mvardPrecoUnitario
            Else
                objVariavelCalculo.vValor = mvardPrecoTotal
            End If
            
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_VALOR_UNITARIO"
            objVariavelCalculo.vValor = mvardPrecoUnitario
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_VALOR_UNITARIO_MOEDA"
            objVariavelCalculo.vValor = mvardPrecoUnitarioMoeda
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_DESC_GLOBAL" 'valor do desconto global rateado para o item
            objVariavelCalculo.vValor = mvardValorDescGlobal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_DESP_GLOBAL" 'valor de (frete+seguro+outras desp) rateado para o item
            objVariavelCalculo.vValor = mvardValorDespGlobal
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_VALOR_BRUTO"
        
            lErro = mvarobjDoc.ObterVar("TIPO_DOC_CODIGO", objVarTipoDocInfo)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            If mvardQuantidade = 0 And objVarTipoDocInfo.vValor = DOCINFO_NFIEIMPCC Then
                objVariavelCalculo.vValor = Arredonda_Moeda(mvardPrecoUnitario, 2)
            Else
                objVariavelCalculo.vValor = Arredonda_Moeda(mvardPrecoTotal + mvardValorDesconto, 2)
            End If
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "PRODUTO_QTDE"
            objVariavelCalculo.vValor = mvardQuantidade
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "PRODUTO_OBJ"
'
'            objProduto.sCodigo = mvarsProduto
'            lErro = CF("Produto_Le", objProduto)
'            If lErro <> SUCESSO Then Error 27825
'
'            Set objVariavelCalculo.vValor = objProduto
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        'para ICMS
'        Case "ICMS_TIPO"
'            objVariavelCalculo.vValor = mvariICMSTipo
'            objVariavelCalculo.iStatus = mvariICMSTipoManual
'
'        Case "ICMS_BASE"
'            objVariavelCalculo.vValor = mvardICMSBase
'            objVariavelCalculo.iStatus = mvariICMSBaseManual
'
'        Case "ICMS_RED_BASE"
'            objVariavelCalculo.vValor = mvardICMSPercRedBase
'            objVariavelCalculo.iStatus = mvariICMSPercRedBaseManual
'
'        Case "ICMS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardICMSAliquota
'            objVariavelCalculo.iStatus = mvariICMSAliquotaManual
'
'        Case "ICMS_VALOR"
'            objVariavelCalculo.vValor = mvardICMSValor
'            objVariavelCalculo.iStatus = mvariICMSValorManual
'
'        Case "ICMS_BASE_SUBST"
'            objVariavelCalculo.vValor = mvardICMSSubstBase
'            objVariavelCalculo.iStatus = mvariICMSSubstBaseManual
'
'        Case "ICMS_RED_BASE_SUBST"
'            objVariavelCalculo.vValor = mvardICMSSubstPercRedBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ICMS_ALIQUOTA_SUBST"
'            objVariavelCalculo.vValor = mvardICMSSubstAliquota
'            objVariavelCalculo.iStatus = mvariICMSSubstAliquotaManual
'
'        Case "ICMS_VALOR_SUBST"
'            objVariavelCalculo.vValor = mvardICMSSubstValor
'            objVariavelCalculo.iStatus = mvariICMSSubstValorManual
'
'        Case "ICMS_MARG_SUBST"
'            objVariavelCalculo.vValor = mvardICMSSubstPercMVA
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ICMS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardICMSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'para IPI
'        Case "IPI_TIPO"
'            objVariavelCalculo.vValor = mvariIPITipo
'            objVariavelCalculo.iStatus = mvariIPITipoManual
'
'        Case "IPI_BASE"
'            objVariavelCalculo.vValor = mvardIPIBaseCalculo
'            objVariavelCalculo.iStatus = mvariIPIBaseManual
'
'        Case "IPI_RED_BASE"
'            objVariavelCalculo.vValor = mvardIPIPercRedBase
'            objVariavelCalculo.iStatus = mvariIPIPercRedBaseManual
'
'        Case "IPI_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardIPIAliquota
'            objVariavelCalculo.iStatus = mvariIPIAliquotaManual
'
'        Case "IPI_VALOR"
'            objVariavelCalculo.vValor = mvardIPIValor
'            objVariavelCalculo.iStatus = mvariIPIValorManual
'
'        Case "IPI_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardIPICredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "PIS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardPISCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "COFINS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvardCOFINSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "DESP_IMP_ICMS_BASE"
'            objVariavelCalculo.vValor = dDespImpICMSBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "DESP_IMP_ICMS_VALOR"
'            objVariavelCalculo.vValor = dDespImpICMSValor
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "DESP_IMP_ICMS_CREDITO"
'            objVariavelCalculo.vValor = dDespImpICMSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ICMS_ALIQUOTA_ADICAO_DI"
'            objVariavelCalculo.vValor = dICMSAliquotaAdicaoDI
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
        Case "DESP_IMP_VALOR_RATEADO"
            objVariavelCalculo.vValor = dDespImpValorRateado
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'XXXX
'        'INÍCIO NOVOS CAMPOS
'        Case "ICMS_ST_RED_BASE"
'            objVariavelCalculo.vValor = mvardICMSSubstPercRedBase
'            objVariavelCalculo.iStatus = mvariICMSSubstPercRedBaseManual
'
'        Case "ICMS_ST_PERC_MVA"
'            objVariavelCalculo.vValor = mvardICMSSubstPercMVA
'            objVariavelCalculo.iStatus = mvariICMSSubstPercMVAManual
'
'        Case "PIS_TIPO"
'            objVariavelCalculo.vValor = mvariPISTipo
'            objVariavelCalculo.iStatus = mvariPISTipoManual
'
'        Case "PIS_TIPO_CALCULO"
'            objVariavelCalculo.vValor = mvariPISTipoCalculo
'            objVariavelCalculo.iStatus = mvariPISTipoCalculoManual
'
'        Case "PIS_BASE"
'            objVariavelCalculo.vValor = mvardPISBase
'            objVariavelCalculo.iStatus = mvariPISBaseManual
'
'        Case "PIS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardPISAliquota
'            objVariavelCalculo.iStatus = mvariPISAliquotaManual
'
'        Case "PIS_ALIQUOTA_VALOR"
'            objVariavelCalculo.vValor = mvardPISAliquotaValor
'            objVariavelCalculo.iStatus = mvariPISAliquotaValorManual
'
'        Case "PIS_QTDE"
'            objVariavelCalculo.vValor = mvardPISQtde
'            objVariavelCalculo.iStatus = mvariPISQtdeManual
'
'        Case "PIS_VALOR"
'            objVariavelCalculo.vValor = mvardPISValor
'            objVariavelCalculo.iStatus = mvariPISValorManual
'
'        Case "PIS_ST_TIPO_CALCULO"
'            objVariavelCalculo.vValor = mvariPISSTTipoCalculo
'            objVariavelCalculo.iStatus = mvariPISSTTipoCalculoManual
'
'        Case "PIS_ST_BASE"
'            objVariavelCalculo.vValor = mvardPISSTBase
'            objVariavelCalculo.iStatus = mvariPISSTBaseManual
'
'        Case "PIS_ST_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardPISSTAliquota
'            objVariavelCalculo.iStatus = mvariPISSTAliquotaManual
'
'        Case "PIS_ST_ALIQUOTA_VALOR"
'            objVariavelCalculo.vValor = mvardPISSTAliquotaValor
'            objVariavelCalculo.iStatus = mvariPISSTAliquotaValorManual
'
'        Case "PIS_ST_QTDE"
'            objVariavelCalculo.vValor = mvardPISSTQtde
'            objVariavelCalculo.iStatus = mvariPISSTQtdeManual
'
'        Case "PIS_ST_VALOR"
'            objVariavelCalculo.vValor = mvardPISSTValor
'            objVariavelCalculo.iStatus = mvariPISSTValorManual
'
'        Case "COFINS_TIPO"
'            objVariavelCalculo.vValor = mvariCOFINSTipo
'            objVariavelCalculo.iStatus = mvariCOFINSTipoManual
'
'        Case "COFINS_TIPO_CALCULO"
'            objVariavelCalculo.vValor = mvariCOFINSTipoCalculo
'            objVariavelCalculo.iStatus = mvariCOFINSTipoCalculoManual
'
'        Case "COFINS_BASE"
'            objVariavelCalculo.vValor = mvardCOFINSBase
'            objVariavelCalculo.iStatus = mvariCOFINSBaseManual
'
'        Case "COFINS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardCOFINSAliquota
'            objVariavelCalculo.iStatus = mvariCOFINSAliquotaManual
'
'        Case "COFINS_ALIQUOTA_VALOR"
'            objVariavelCalculo.vValor = mvardCOFINSAliquotaValor
'            objVariavelCalculo.iStatus = mvariCOFINSAliquotaValorManual
'
'        Case "COFINS_QTDE"
'            objVariavelCalculo.vValor = mvardCOFINSQtde
'            objVariavelCalculo.iStatus = mvariCOFINSQtdeManual
'
'        Case "COFINS_VALOR"
'            objVariavelCalculo.vValor = mvardCOFINSValor
'            objVariavelCalculo.iStatus = mvariCOFINSValorManual
'
'        Case "COFINS_ST_TIPO_CALCULO"
'            objVariavelCalculo.vValor = mvariCOFINSSTTipoCalculo
'            objVariavelCalculo.iStatus = mvariCOFINSSTTipoCalculoManual
'
'        Case "COFINS_ST_BASE"
'            objVariavelCalculo.vValor = mvardCOFINSSTBase
'            objVariavelCalculo.iStatus = mvariCOFINSSTBaseManual
'
'        Case "COFINS_ST_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardCOFINSSTAliquota
'            objVariavelCalculo.iStatus = mvariCOFINSSTAliquotaManual
'
'        Case "COFINS_ST_ALIQUOTA_VALOR"
'            objVariavelCalculo.vValor = mvardCOFINSSTAliquotaValor
'            objVariavelCalculo.iStatus = mvariCOFINSSTAliquotaValorManual
'
'        Case "COFINS_ST_QTDE"
'            objVariavelCalculo.vValor = mvardCOFINSSTQtde
'            objVariavelCalculo.iStatus = mvariCOFINSSTQtdeManual
'
'        Case "COFINS_ST_VALOR"
'            objVariavelCalculo.vValor = mvardCOFINSSTValor
'            objVariavelCalculo.iStatus = mvariCOFINSSTValorManual
'
'        Case "ICMS_CST"
'            objVariavelCalculo.vValor = mvarsCST
'            objVariavelCalculo.iStatus = mvariCSTManual
'
'        Case "ISSQN"
'            objVariavelCalculo.vValor = mvarsISSQN
'            objVariavelCalculo.iStatus = mvariISSQNManual
'
'        Case "ISS_BASE"
'            objVariavelCalculo.vValor = mvardISSBase
'            objVariavelCalculo.iStatus = mvariISSBaseManual
'
'        Case "ISS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvardISSAliquota
'            objVariavelCalculo.iStatus = mvariISSAliquotaManual
'
'        Case "ISS_VALOR"
'            objVariavelCalculo.vValor = mvardISSValor
'            objVariavelCalculo.iStatus = mvariISSValorManual
'
'        Case "ISS_CIDADE_IBGE"
'            objVariavelCalculo.vValor = mvarsISSCidadeIBGE
'            objVariavelCalculo.iStatus = mvariISSCidadeIBGEManual
'
'        Case "II_BASE"
'            objVariavelCalculo.vValor = mvardIIBase
'            objVariavelCalculo.iStatus = mvariIIBaseManual
'
'        Case "II_DEPESAS_ADUANEIRA"
'            objVariavelCalculo.vValor = mvardIIDespAduaneira
'            objVariavelCalculo.iStatus = mvariIIDespAduaneiraManual
'
'        Case "II_IOF"
'            objVariavelCalculo.vValor = mvardIIIOF
'            objVariavelCalculo.iStatus = mvariIIIOFManual
'
'        Case "II_VALOR"
'            objVariavelCalculo.vValor = mvardIIValor
'            objVariavelCalculo.iStatus = mvariIIValorManual
'
'        Case "ICMS_BASE_MODALIDADE"
'            objVariavelCalculo.vValor = mvariICMSBaseModalidade
'            objVariavelCalculo.iStatus = mvariICMSBaseModalidadeManual
'
'        Case "ICMS_ST_BASE_MODALIDADE"
'            objVariavelCalculo.vValor = mvariICMSSubstBaseModalidade
'            objVariavelCalculo.iStatus = mvariICMSSubstBaseModalidadeManual
'
'         Case "IPI_ENQUADRAMENTO_CLASSE"
'            objVariavelCalculo.vValor = mvarsIPIEnquadramentoClasse
'            objVariavelCalculo.iStatus = mvariIPIEnquadramentoClasseManual
'
'         Case "IPI_ENQUADRAMENTO_CODIGO"
'            objVariavelCalculo.vValor = mvarsIPIEnquadramentoCodigo
'            objVariavelCalculo.iStatus = mvariIPIEnquadramentoCodigoManual
'
'         Case "IPI_CNPJ_PRODUTOR"
'            objVariavelCalculo.vValor = mvarsIPICNPJProdutor
'            objVariavelCalculo.iStatus = mvariIPICNPJProdutorManual
'
'         Case "IPI_SELO_CODIGO"
'            objVariavelCalculo.vValor = mvarsIPISeloCodigo
'            objVariavelCalculo.iStatus = mvariIPISeloCodigoManual
'
'        Case "IPI_SELO_QTDE"
'            objVariavelCalculo.vValor = mvarlIPISeloQtde
'            objVariavelCalculo.iStatus = mvariIPISeloQtdeManual
'
'        Case "IPI_TIPO_CALCULO"
'            objVariavelCalculo.vValor = mvariIPITipoCalculo
'            objVariavelCalculo.iStatus = mvariIPITipoCalculoManual
'
'        Case "IPI_UM_PADRAO_QTDE"
'            objVariavelCalculo.vValor = mvardIPIUnidadePadraoQtde
'            objVariavelCalculo.iStatus = mvariIPIUnidadePadraoQtdeManual
'
'        Case "IPI_UM_PADRAO_VALOR"
'            objVariavelCalculo.vValor = mvardIPIUnidadePadraoValor
'            objVariavelCalculo.iStatus = mvariIPIUnidadePadraoValorManual
'
'        Case "ORIGEM_MERCADORIA"
'            objVariavelCalculo.vValor = mvariOrigemMercadoria
'            objVariavelCalculo.iStatus = mvariOrigemMercadoriaManual
'
'        Case "VALOR_FRETE_ITEM"
'            objVariavelCalculo.vValor = mvardValorFreteItem
'            objVariavelCalculo.iStatus = mvariValorFreteItemManual
'
'        Case "VALOR_SEGURO_ITEM"
'            objVariavelCalculo.vValor = mvardValorSeguroItem
'            objVariavelCalculo.iStatus = mvariValorSeguroItemManual
'
'        Case "VALOR_DESCONTO_ITEM"
'            objVariavelCalculo.vValor = mvardValorDescontoItem
'            objVariavelCalculo.iStatus = mvariValorDescontoItemManual
'
'        Case "VALOR_DESPESAS_ITEM"
'            objVariavelCalculo.vValor = mvardValorOutrasDespesasItem
'            objVariavelCalculo.iStatus = mvariValorOutrasDespesasItemManual
'        'FIM NOVOS CAMPOS
        
        Case Else

            lErro = objTributacao.ObterVar_Padrao(sIdentificador, objVariavelCalculo, Me)
            If lErro <> SUCESSO Then
            
                lErro = mvarobjTributacaoDocItem.ObterVar(sIdentificador, objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27824
                
            End If

    End Select

    ClassTributoDocItem_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_ObterVar:

    ClassTributoDocItem_ObterVar = Err

    Select Case Err

        Case 27824, 27825
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error, 153706)

    End Select

    Exit Function

End Function

Private Function ClassTributoDocItem_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDocItem_GuardarVar

    Select Case objVariavelCalculo.sIdentificador

'        Case "NAT_OPERACAO"
'            mvarsNaturezaOp = objVariavelCalculo.vValor
'            mvariNaturezaOpManual = objVariavelCalculo.iStatus
'
'        Case "TIPO_TRIB"
'            mvariTipoTributacao = objVariavelCalculo.vValor
'            mvariTipoTributacaoManual = objVariavelCalculo.iStatus
'
'        Case "PRODUTO_CODIGO"
        Case "PRODUTO_VALOR"
        Case "PRODUTO_VALOR_UNITARIO"
        Case "PRODUTO_VALOR_UNITARIO_MOEDA"
        Case "PRODUTO_VALOR_BRUTO"
        Case "PRODUTO_QTDE"

        Case "PRODUTO_DESC_GLOBAL" 'valor do desconto global rateado para o item
            mvardValorDescGlobal = objVariavelCalculo.vValor

        Case "PRODUTO_DESP_GLOBAL" 'valor do (frete+seguro+outras desp) rateado para o item
            mvardValorDespGlobal = objVariavelCalculo.vValor
'
'        'para ICMS
'        Case "ICMS_TIPO"
'            mvariICMSTipo = objVariavelCalculo.vValor
'            mvariICMSTipoManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_BASE"
'            mvardICMSBase = objVariavelCalculo.vValor
'            mvariICMSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_RED_BASE"
'            mvardICMSPercRedBase = objVariavelCalculo.vValor
'            mvariICMSPercRedBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_ALIQUOTA"
'            mvardICMSAliquota = objVariavelCalculo.vValor
'            mvariICMSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR"
'            mvardICMSValor = objVariavelCalculo.vValor
'            mvariICMSValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_BASE_SUBST"
'            mvardICMSSubstBase = objVariavelCalculo.vValor
'            mvariICMSSubstBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_RED_BASE_SUBST"
'            mvardICMSSubstPercRedBase = objVariavelCalculo.vValor
'
'        Case "ICMS_ALIQUOTA_SUBST"
'            mvardICMSSubstAliquota = objVariavelCalculo.vValor
'            mvariICMSSubstAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR_SUBST"
'            mvardICMSSubstValor = objVariavelCalculo.vValor
'            mvariICMSSubstValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_MARG_SUBST"
'            mvardICMSSubstPercMVA = objVariavelCalculo.vValor
'
'        Case "ICMS_VALOR_CREDITO"
'            mvardICMSCredito = objVariavelCalculo.vValor
'
'        'para IPI
'        Case "IPI_TIPO"
'            mvariIPITipo = objVariavelCalculo.vValor
'            mvariIPITipoManual = objVariavelCalculo.iStatus
'
'        Case "IPI_BASE"
'            mvardIPIBaseCalculo = objVariavelCalculo.vValor
'            mvariIPIBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_RED_BASE"
'            mvardIPIPercRedBase = objVariavelCalculo.vValor
'            mvariIPIPercRedBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_ALIQUOTA"
'            mvardIPIAliquota = objVariavelCalculo.vValor
'            mvariIPIAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR"
'            mvardIPIValor = objVariavelCalculo.vValor
'            mvariIPIValorManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR_CREDITO"
'            mvardIPICredito = objVariavelCalculo.vValor
'
'        Case "PIS_VALOR_CREDITO"
'            mvardPISCredito = objVariavelCalculo.vValor
'
'        Case "COFINS_VALOR_CREDITO"
'            mvardCOFINSCredito = objVariavelCalculo.vValor
'
'        Case "DESP_IMP_ICMS_BASE"
'            dDespImpICMSBase = objVariavelCalculo.vValor
'
'        Case "DESP_IMP_ICMS_VALOR"
'            dDespImpICMSValor = objVariavelCalculo.vValor
'
'        Case "DESP_IMP_ICMS_CREDITO"
'            dDespImpICMSCredito = objVariavelCalculo.vValor
'
'        Case "ICMS_ALIQUOTA_ADICAO_DI"
'        Case "DESP_IMP_VALOR_RATEADO"
'
'        'XXXX
'        'INÍCIO NOVOS CAMPOS
'        Case "ICMS_ST_RED_BASE"
'            mvardICMSSubstPercRedBase = objVariavelCalculo.vValor
'            mvariICMSSubstPercRedBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_ST_PERC_MVA"
'            mvardICMSSubstPercMVA = objVariavelCalculo.vValor
'            mvariICMSSubstPercMVAManual = objVariavelCalculo.iStatus
'
'        Case "PIS_TIPO"
'            mvariPISTipo = objVariavelCalculo.vValor
'            mvariPISTipoManual = objVariavelCalculo.iStatus
'
'        Case "PIS_TIPO_CALCULO"
'            mvariPISTipoCalculo = objVariavelCalculo.vValor
'            mvariPISTipoCalculoManual = objVariavelCalculo.iStatus
'
'        Case "PIS_BASE"
'            mvardPISBase = objVariavelCalculo.vValor
'            mvariPISBaseManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ALIQUOTA"
'            mvardPISAliquota = objVariavelCalculo.vValor
'            mvariPISAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ALIQUOTA_VALOR"
'            mvardPISAliquotaValor = objVariavelCalculo.vValor
'            mvariPISAliquotaValorManual = objVariavelCalculo.iStatus
'
'        Case "PIS_QTDE"
'            mvardPISQtde = objVariavelCalculo.vValor
'            mvariPISQtdeManual = objVariavelCalculo.iStatus
'
'        Case "PIS_VALOR"
'            mvardPISValor = objVariavelCalculo.vValor
'            mvariPISValorManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ST_TIPO_CALCULO"
'            mvariPISSTTipoCalculo = objVariavelCalculo.vValor
'            mvariPISSTTipoCalculoManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ST_BASE"
'            mvardPISSTBase = objVariavelCalculo.vValor
'            mvariPISSTBaseManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ST_ALIQUOTA"
'            mvardPISSTAliquota = objVariavelCalculo.vValor
'            mvariPISSTAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ST_ALIQUOTA_VALOR"
'            mvardPISSTAliquotaValor = objVariavelCalculo.vValor
'            mvariPISSTAliquotaValorManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ST_QTDE"
'            mvardPISSTQtde = objVariavelCalculo.vValor
'            mvariPISSTQtdeManual = objVariavelCalculo.iStatus
'
'        Case "PIS_ST_VALOR"
'            mvardPISSTValor = objVariavelCalculo.vValor
'            mvariPISSTValorManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_TIPO"
'            mvariCOFINSTipo = objVariavelCalculo.vValor
'            mvariCOFINSTipoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_TIPO_CALCULO"
'            mvariCOFINSTipoCalculo = objVariavelCalculo.vValor
'            mvariCOFINSTipoCalculoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_BASE"
'            mvardCOFINSBase = objVariavelCalculo.vValor
'            mvariCOFINSBaseManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ALIQUOTA"
'            mvardCOFINSAliquota = objVariavelCalculo.vValor
'            mvariCOFINSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ALIQUOTA_VALOR"
'            mvardCOFINSAliquotaValor = objVariavelCalculo.vValor
'            mvariCOFINSAliquotaValorManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_QTDE"
'            mvardCOFINSQtde = objVariavelCalculo.vValor
'            mvariCOFINSQtdeManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_VALOR"
'            mvardCOFINSValor = objVariavelCalculo.vValor
'            mvariCOFINSValorManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ST_TIPO_CALCULO"
'            mvariCOFINSSTTipoCalculo = objVariavelCalculo.vValor
'            mvariCOFINSSTTipoCalculoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ST_BASE"
'            mvardCOFINSSTBase = objVariavelCalculo.vValor
'            mvariCOFINSSTBaseManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ST_ALIQUOTA"
'            mvardCOFINSSTAliquota = objVariavelCalculo.vValor
'            mvariCOFINSSTAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ST_ALIQUOTA_VALOR"
'            mvardCOFINSSTAliquotaValor = objVariavelCalculo.vValor
'            mvariCOFINSSTAliquotaValorManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ST_QTDE"
'            mvardCOFINSSTQtde = objVariavelCalculo.vValor
'            mvariCOFINSSTQtdeManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_ST_VALOR"
'            mvardCOFINSSTValor = objVariavelCalculo.vValor
'            mvariCOFINSSTValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_CST"
'            mvarsCST = objVariavelCalculo.vValor
'            mvariCSTManual = objVariavelCalculo.iStatus
'
'        Case "ISSQN"
'            mvarsISSQN = objVariavelCalculo.vValor
'            mvariISSQNManual = objVariavelCalculo.iStatus
'
'        Case "ISS_BASE"
'            mvardISSBase = objVariavelCalculo.vValor
'            mvariISSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ISS_ALIQUOTA"
'            mvardISSAliquota = objVariavelCalculo.vValor
'            mvariISSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ISS_VALOR"
'            mvardISSValor = objVariavelCalculo.vValor
'            mvariISSValorManual = objVariavelCalculo.iStatus
'
'        Case "ISS_CIDADE_IBGE"
'            mvarsISSCidadeIBGE = objVariavelCalculo.vValor
'            mvariISSCidadeIBGEManual = objVariavelCalculo.iStatus
'
'        Case "II_BASE"
'            mvardIIBase = objVariavelCalculo.vValor
'            mvariIIBaseManual = objVariavelCalculo.iStatus
'
'        Case "II_DEPESAS_ADUANEIRA"
'            mvardIIDespAduaneira = objVariavelCalculo.vValor
'            mvariIIDespAduaneiraManual = objVariavelCalculo.iStatus
'
'        Case "II_IOF"
'            mvardIIIOF = objVariavelCalculo.vValor
'            mvariIIIOFManual = objVariavelCalculo.iStatus
'
'        Case "II_VALOR"
'            mvardIIValor = objVariavelCalculo.vValor
'            mvariIIValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_BASE_MODALIDADE"
'            mvariICMSBaseModalidade = objVariavelCalculo.vValor
'            mvariICMSBaseModalidadeManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_ST_BASE_MODALIDADE"
'            mvariICMSSubstBaseModalidade = objVariavelCalculo.vValor
'            mvariICMSSubstBaseModalidadeManual = objVariavelCalculo.iStatus
'
'         Case "IPI_ENQUADRAMENTO_CLASSE"
'            mvarsIPIEnquadramentoClasse = objVariavelCalculo.vValor
'            mvariIPIEnquadramentoClasseManual = objVariavelCalculo.iStatus
'
'         Case "IPI_ENQUADRAMENTO_CODIGO"
'            mvarsIPIEnquadramentoCodigo = objVariavelCalculo.vValor
'            mvariIPIEnquadramentoCodigoManual = objVariavelCalculo.iStatus
'
'         Case "IPI_CNPJ_PRODUTOR"
'            mvarsIPICNPJProdutor = objVariavelCalculo.vValor
'            mvariIPICNPJProdutorManual = objVariavelCalculo.iStatus
'
'         Case "IPI_SELO_CODIGO"
'            mvarsIPISeloCodigo = objVariavelCalculo.vValor
'            mvariIPISeloCodigoManual = objVariavelCalculo.iStatus
'
'        Case "IPI_SELO_QTDE"
'            mvarlIPISeloQtde = objVariavelCalculo.vValor
'            mvariIPISeloQtdeManual = objVariavelCalculo.iStatus
'
'        Case "IPI_TIPO_CALCULO"
'            mvariIPITipoCalculo = objVariavelCalculo.vValor
'            mvariIPITipoCalculoManual = objVariavelCalculo.iStatus
'
'        Case "IPI_UM_PADRAO_QTDE"
'            mvardIPIUnidadePadraoQtde = objVariavelCalculo.vValor
'            mvariIPIUnidadePadraoQtdeManual = objVariavelCalculo.iStatus
'
'        Case "IPI_UM_PADRAO_VALOR"
'            mvardIPIUnidadePadraoValor = objVariavelCalculo.vValor
'            mvariIPIUnidadePadraoValorManual = objVariavelCalculo.iStatus
'
'        Case "ORIGEM_MERCADORIA"
'            mvariOrigemMercadoria = objVariavelCalculo.vValor
'            mvariOrigemMercadoriaManual = objVariavelCalculo.iStatus
'
'        Case "VALOR_FRETE_ITEM"
'            mvardValorFreteItem = objVariavelCalculo.vValor
'            mvariValorFreteItemManual = objVariavelCalculo.iStatus
'
'        Case "VALOR_SEGURO_ITEM"
'            mvardValorSeguroItem = objVariavelCalculo.vValor
'            mvariValorSeguroItemManual = objVariavelCalculo.iStatus
'
'        Case "VALOR_DESCONTO_ITEM"
'            mvardValorDescontoItem = objVariavelCalculo.vValor
'            mvariValorDescontoItemManual = objVariavelCalculo.iStatus
'
'        Case "VALOR_DESPESAS_ITEM"
'            mvardValorOutrasDespesasItem = objVariavelCalculo.vValor
'            mvariValorOutrasDespesasItemManual = objVariavelCalculo.iStatus
'        'FIM NOVOS CAMPOS
        
        Case Else
            'delegar p/tributacaoDocItem
            '"ICMS_MARG_SUBST"
            
            lErro = objTributacao.GuardarVar_Padrao(objVariavelCalculo, Me)
            If lErro <> SUCESSO Then
            
                lErro = mvarobjTributacaoDocItem.GuardarVar(objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27826
                
            End If

    End Select

    ClassTributoDocItem_GuardarVar = SUCESSO

    Exit Function

Erro_ClassTributoDocItem_GuardarVar:

    ClassTributoDocItem_GuardarVar = Err

    Select Case Err

        Case 27826
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, ERRO_FORNECIDO_PELO_VB, Err, Error, 153707)

    End Select

    Exit Function

End Function

'Public Property Let iFilialEmpresa(ByVal vData As Integer)
'    mvariFilialEmpresa = vData
'End Property
'
'Public Property Get iFilialEmpresa() As Integer
'    iFilialEmpresa = mvariFilialEmpresa
'End Property
'
'Public Property Let iItem(ByVal vData As Integer)
'    mvariItem = vData
'End Property
'
'Public Property Get iItem() As Integer
'    iItem = mvariItem
'End Property
'
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

Public Property Let dICMSSubstValor(ByVal vData As Double)
    mvardICMSSubstValor = vData
End Property

Public Property Get dICMSSubstValor() As Double
    dICMSSubstValor = mvardICMSSubstValor
End Property

Public Property Let dICMSSubstAliquota(ByVal vData As Double)
    mvardICMSSubstAliquota = vData
End Property

Public Property Get dICMSSubstAliquota() As Double
    dICMSSubstAliquota = mvardICMSSubstAliquota
End Property

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let iICMSAliquotaManual(ByVal vData As Integer)
    mvariICMSAliquotaManual = vData
End Property

Public Property Get iICMSAliquotaManual() As Integer
    iICMSAliquotaManual = mvariICMSAliquotaManual
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
    mvardICMSAliquota = vData
End Property

Public Property Get dICMSAliquota() As Double
    dICMSAliquota = mvardICMSAliquota
End Property

Public Property Let iICMSPercRedBaseManual(ByVal vData As Integer)
    mvariICMSPercRedBaseManual = vData
End Property

Public Property Get iICMSPercRedBaseManual() As Integer
    iICMSPercRedBaseManual = mvariICMSPercRedBaseManual
End Property

Public Property Let dICMSPercRedBase(ByVal vData As Double)
    mvardICMSPercRedBase = vData
End Property

Public Property Get dICMSPercRedBase() As Double
    dICMSPercRedBase = mvardICMSPercRedBase
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let iICMSTipoManual(ByVal vData As Integer)
    mvariICMSTipoManual = vData
End Property

Public Property Get iICMSTipoManual() As Integer
    iICMSTipoManual = mvariICMSTipoManual
End Property

Public Property Let iICMSTipo(ByVal vData As Integer)
    mvariICMSTipo = vData
End Property

Public Property Get iICMSTipo() As Integer
    iICMSTipo = mvariICMSTipo
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let dIPIBaseCalculo(ByVal vData As Double)
    mvardIPIBaseCalculo = vData
End Property

Public Property Get dIPIBaseCalculo() As Double
    dIPIBaseCalculo = mvardIPIBaseCalculo
End Property

Public Property Let iIPIAliquotaManual(ByVal vData As Integer)
    mvariIPIAliquotaManual = vData
End Property

Public Property Get iIPIAliquotaManual() As Integer
    iIPIAliquotaManual = mvariIPIAliquotaManual
End Property

Public Property Let dIPIAliquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIAliquota() As Double
    dIPIAliquota = mvardIPIAliquota
End Property

Public Property Let iIPITipoManual(ByVal vData As Integer)
    mvariIPITipoManual = vData
End Property

Public Property Get iIPITipoManual() As Integer
    iIPITipoManual = mvariIPITipoManual
End Property

Public Property Let iIPITipo(ByVal vData As Integer)
    mvariIPITipo = vData
End Property

Public Property Get iIPITipo() As Integer
    iIPITipo = mvariIPITipo
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let dIPIPercRedBase(ByVal vData As Double)
    mvardIPIPercRedBase = vData
End Property

Public Property Get dIPIPercRedBase() As Double
    dIPIPercRedBase = mvardIPIPercRedBase
End Property

Public Property Let iIPIPercRedBaseManual(ByVal vData As Integer)
    mvariIPIPercRedBaseManual = vData
End Property

Public Property Get iIPIPercRedBaseManual() As Integer
    iIPIPercRedBaseManual = mvariIPIPercRedBaseManual
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    mvariNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = mvariNaturezaOpManual
End Property

Public Property Let iIPIBaseManual(ByVal vData As Integer)
    mvariIPIBaseManual = vData
End Property

Public Property Get iIPIBaseManual() As Integer
    iIPIBaseManual = mvariIPIBaseManual
End Property

'Public Property Let dICMSSubstPercRedBase(ByVal vData As Double)
'    mvardICMSSubstPercRedBase = vData
'End Property
'
'Public Property Get dICMSSubstPercRedBase() As Double
'    dICMSSubstPercRedBase = mvardICMSSubstPercRedBase
'End Property
'
'Public Property Let dICMSSubstPercMVA(ByVal vData As Double)
'    mvardICMSSubstPercMVA = vData
'End Property
'
'Public Property Get dICMSSubstPercMVA() As Double
'    dICMSSubstPercMVA = mvardICMSSubstPercMVA
'End Property

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacao = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDocItem
    Set objTributacao = mvarobjTributacao
End Property

Public Property Let iIPICreditoManual(ByVal vData As Integer)
    mvariIPICreditoManual = vData
End Property

Public Property Get iIPICreditoManual() As Integer
    iIPICreditoManual = mvariIPICreditoManual
End Property

Public Property Let iICMSCreditoManual(ByVal vData As Integer)
    mvariICMSCreditoManual = vData
End Property

Public Property Get iICMSCreditoManual() As Integer
    iICMSCreditoManual = mvariICMSCreditoManual
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

Public Property Let sIPICodProduto(ByVal vData As String)
    mvarsIPICodProduto = vData
End Property

Public Property Get sIPICodProduto() As String
    sIPICodProduto = mvarsIPICodProduto
End Property

Public Property Let dDescontoGrid(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dDescontoGrid() As Double
    dDescontoGrid = mvardValorDesconto
End Property
