VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTribItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntNF As Long
Private mvariItem As Integer
Private mvarlNumIntDoc As Long
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
Private mvardICMSSubstPercRedBase As Double
Private mvardICMSSubstAliquota As Double
Private mvardICMSSubstValor As Double
Private mvardICMSSubstPercMVA As Double

Private mvariNaturezaOpManual As Integer
Private mvariTipoTributacaoManual As Integer
Private mvariIPITipoManual As Integer
Private mvariIPIBaseManual As Integer
Private mvariIPIPercRedBaseManual As Integer
Private mvariIPIAliquotaManual As Integer
Private mvariIPIValorManual As Integer
Private mvariICMSTipoManual As Integer
Private mvariICMSBaseManual As Integer
Private mvariICMSPercRedBaseManual As Integer
Private mvariICMSAliquotaManual As Integer
Private mvariICMSValorManual As Integer
Private mvariICMSSubstBaseManual As Integer
Private mvariICMSSubstAliquotaManual As Integer
Private mvariICMSSubstValorManual As Integer

Private mvardPISCredito As Double
Private mvardCOFINSCredito As Double

'para notas de importacao
Public dICMSAliquotaAdicaoDI As Double
Public dDespImpICMSBase As Double 'valor da base do icms deste item obtido pelo rateio de despesas de importacao
Public dDespImpICMSValor As Double 'valor da base do icms deste item obtido pelo rateio de despesas de importacao
Public dDespImpICMSCredito As Double 'valor do credito de icms deste item obtido pelo rateio de despesas de importacao

'Private mvardICMSSubstPercRedBase As Double
'Private mvardICMSSubstPercMVA As Double
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
Private mvariExTIPI As Integer
Private mvarsGenero As String
Private mvariProdutoEspecifico As Integer
Private mvarsEAN As String
Private mvarsEANTrib As String
Private mvardQtdTrib As Double
Private mvarsUMTrib As String
Private mvardValorUnitTrib As Double
Private mvariICMSAliquotaAdicaoDIManual As Integer
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

Public Property Let iICMSAliquotaAdicaoDIManual(ByVal vData As Integer)
    mvariICMSAliquotaAdicaoDIManual = vData
End Property

Public Property Get iICMSAliquotaAdicaoDIManual() As Integer
    iICMSAliquotaAdicaoDIManual = mvariICMSAliquotaAdicaoDIManual
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

Public Property Get iTipo() As Integer
    iTipo = 0 'ITEM_TIPO_NORMAL
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

Public Property Let iICMSSubstBaseManual(ByVal vData As Integer)
    mvariICMSSubstBaseManual = vData
End Property

Public Property Get iICMSSubstBaseManual() As Integer
    iICMSSubstBaseManual = mvariICMSSubstBaseManual
End Property

Public Property Let iICMSValorManual(ByVal vData As Integer)
    mvariICMSValorManual = vData
End Property

Public Property Get iICMSValorManual() As Integer
    iICMSValorManual = mvariICMSValorManual
End Property

Public Property Let iICMSAliquotaManual(ByVal vData As Integer)
    mvariICMSAliquotaManual = vData
End Property

Public Property Get iICMSAliquotaManual() As Integer
    iICMSAliquotaManual = mvariICMSAliquotaManual
End Property

Public Property Let iICMSPercRedBaseManual(ByVal vData As Integer)
    mvariICMSPercRedBaseManual = vData
End Property

Public Property Get iICMSPercRedBaseManual() As Integer
    iICMSPercRedBaseManual = mvariICMSPercRedBaseManual
End Property

Public Property Let iICMSBaseManual(ByVal vData As Integer)
    mvariICMSBaseManual = vData
End Property

Public Property Get iICMSBaseManual() As Integer
    iICMSBaseManual = mvariICMSBaseManual
End Property

Public Property Let iICMSTipoManual(ByVal vData As Integer)
    mvariICMSTipoManual = vData
End Property

Public Property Get iICMSTipoManual() As Integer
    iICMSTipoManual = mvariICMSTipoManual
End Property

Public Property Let iIPIValorManual(ByVal vData As Integer)
    mvariIPIValorManual = vData
End Property

Public Property Get iIPIValorManual() As Integer
    iIPIValorManual = mvariIPIValorManual
End Property

Public Property Let iIPIAliquotaManual(ByVal vData As Integer)
    mvariIPIAliquotaManual = vData
End Property

Public Property Get iIPIAliquotaManual() As Integer
    iIPIAliquotaManual = mvariIPIAliquotaManual
End Property

Public Property Let iIPITipoManual(ByVal vData As Integer)
    mvariIPITipoManual = vData
End Property

Public Property Get iIPITipoManual() As Integer
    iIPITipoManual = mvariIPITipoManual
End Property

Public Property Let iTipoTributacaoManual(ByVal vData As Integer)
    mvariTipoTributacaoManual = vData
End Property

Public Property Get iTipoTributacaoManual() As Integer
    iTipoTributacaoManual = mvariTipoTributacaoManual
End Property

Public Property Let iIPIPercRedBaseManual(ByVal vData As Integer)
    mvariIPIPercRedBaseManual = vData
End Property

Public Property Get iIPIPercRedBaseManual() As Integer
    iIPIPercRedBaseManual = mvariIPIPercRedBaseManual
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

Public Property Let dICMSSubstBase(ByVal vData As Double)
    mvardICMSSubstBase = vData
End Property

Public Property Get dICMSSubstBase() As Double
    dICMSSubstBase = mvardICMSSubstBase
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

''Public Property Let iICMSCredita(ByVal vData As Integer)
''    mvariICMSCredita = vData
''End Property
''
''Public Property Get iICMSCredita() As Integer
''    iICMSCredita = mvariICMSCredita
''End Property
''

Public Property Let dICMSValor(ByVal vData As Double)
    mvardICMSValor = vData
End Property

Public Property Get dICMSValor() As Double
    dICMSValor = mvardICMSValor
End Property

Public Property Let dICMSAliquota(ByVal vData As Double)
    mvardICMSAliquota = vData
End Property

Public Property Get dICMSAliquota() As Double
    dICMSAliquota = mvardICMSAliquota
End Property

Public Property Let dICMSPercRedBase(ByVal vData As Double)
    mvardICMSPercRedBase = vData
End Property

Public Property Get dICMSPercRedBase() As Double
    dICMSPercRedBase = mvardICMSPercRedBase
End Property

Public Property Let dICMSBase(ByVal vData As Double)
    mvardICMSBase = vData
End Property

Public Property Get dICMSBase() As Double
    dICMSBase = mvardICMSBase
End Property

Public Property Let iICMSTipo(ByVal vData As Integer)
    mvariICMSTipo = vData
End Property

Public Property Get iICMSTipo() As Integer
    iICMSTipo = mvariICMSTipo
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

Public Property Let dIPIValor(ByVal vData As Double)
    mvardIPIValor = vData
End Property

Public Property Get dIPIValor() As Double
    dIPIValor = mvardIPIValor
End Property

Public Property Let dIPIAliquota(ByVal vData As Double)
    mvardIPIAliquota = vData
End Property

Public Property Get dIPIAliquota() As Double
    dIPIAliquota = mvardIPIAliquota
End Property

Public Property Let dIPIPercRedBase(ByVal vData As Double)
    mvardIPIPercRedBase = vData
End Property

Public Property Get dIPIPercRedBase() As Double
    dIPIPercRedBase = mvardIPIPercRedBase
End Property

Public Property Let dIPIBaseCalculo(ByVal vData As Double)
    mvardIPIBaseCalculo = vData
End Property

Public Property Get dIPIBaseCalculo() As Double
    dIPIBaseCalculo = mvardIPIBaseCalculo
End Property

Public Property Let sIPICodProduto(ByVal vData As String)
    mvarsIPICodProduto = vData
End Property

Public Property Get sIPICodProduto() As String
    sIPICodProduto = mvarsIPICodProduto
End Property

Public Property Let iIPITipo(ByVal vData As Integer)
    mvariIPITipo = vData
End Property

Public Property Get iIPITipo() As Integer
    iIPITipo = mvariIPITipo
End Property

Public Property Let iTipoTributacao(ByVal vData As Integer)
    mvariTipoTributacao = vData
End Property

Public Property Get iTipoTributacao() As Integer
    iTipoTributacao = mvariTipoTributacao
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
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

Public Property Let lNumIntNF(ByVal vData As Long)
    mvarlNumIntNF = vData
End Property

Public Property Get lNumIntNF() As Long
    lNumIntNF = mvarlNumIntNF
End Property
