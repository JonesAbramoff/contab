VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivRegESLinha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocRegES As Long
Private mvarsNaturezaOp As String
Private mvardValorTotal As Double
Private mvardValorICMSBase As Double
Private mvardValorICMS As Double
Private mvardValorICMS51 As Double
Private mvardValorICMSIsentoNaoTrib As Double
Private mvardValorICMSOutras As Double
Private mvardValorICMSSubstBase As Double
Private mvardValorICMSSubstRet As Double
Private mvardAliquotaICMS As Double
Private mvardValorDespAcess As Double
Private mvardValorIPI As Double
Private mvardAliquotaIPI As Double
Private mvardValorIPIBase As Double
Private mvardValorIPIIsentoNaoTrib As Double
Private mvardValorIPIOutras As Double
Private mvarsClassifContabil As String
Private mvardValorContabil As Double
Private mvarsObservacaoLivFisc As String
Private mvariCodigoICMS As Integer
Private mvariCodigoIPI As Integer

Public Property Let iCodigoIPI(ByVal vData As Integer)
    mvariCodigoIPI = vData
End Property

Public Property Get iCodigoIPI() As Integer
    iCodigoIPI = mvariCodigoIPI
End Property

Public Property Let iCodigoICMS(ByVal vData As Integer)
    mvariCodigoICMS = vData
End Property

Public Property Get iCodigoICMS() As Integer
    iCodigoICMS = mvariCodigoICMS
End Property

Public Property Let sObservacaoLivFisc(ByVal vData As String)
Attribute sObservacaoLivFisc.VB_Description = "observacao p/livro fiscal"
    mvarsObservacaoLivFisc = vData
End Property

Public Property Get sObservacaoLivFisc() As String
    sObservacaoLivFisc = mvarsObservacaoLivFisc
End Property

Public Property Let dValorContabil(ByVal vData As Double)
Attribute dValorContabil.VB_Description = "Valor Contabil"
    mvardValorContabil = vData
End Property

Public Property Get dValorContabil() As Double
    dValorContabil = mvardValorContabil
End Property

Public Property Let sClassifContabil(ByVal vData As String)
Attribute sClassifContabil.VB_Description = "Se for um Produto não estocavel --> Conta Aplicacao ou Se for um Produto estocavel -->  Conta Estoque Produto"
    mvarsClassifContabil = vData
End Property

Public Property Get sClassifContabil() As String
    sClassifContabil = mvarsClassifContabil
End Property

Public Property Let dValorIPIOutras(ByVal vData As Double)
Attribute dValorIPIOutras.VB_Description = "Valor do IPI Outras"
    mvardValorIPIOutras = vData
End Property

Public Property Get dValorIPIOutras() As Double
    dValorIPIOutras = mvardValorIPIOutras
End Property

Public Property Let dValorIPIIsentoNaoTrib(ByVal vData As Double)
Attribute dValorIPIIsentoNaoTrib.VB_Description = "Valor do IPI Isento ou não tributado"
    mvardValorIPIIsentoNaoTrib = vData
End Property

Public Property Get dValorIPIIsentoNaoTrib() As Double
    dValorIPIIsentoNaoTrib = mvardValorIPIIsentoNaoTrib
End Property

Public Property Let dValorIPIBase(ByVal vData As Double)
Attribute dValorIPIBase.VB_Description = "Valor Base do IPI"
    mvardValorIPIBase = vData
End Property

Public Property Get dValorIPIBase() As Double
    dValorIPIBase = mvardValorIPIBase
End Property

Public Property Let dAliquotaIPI(ByVal vData As Double)
Attribute dAliquotaIPI.VB_Description = "Aliquota do IPI"
    mvardAliquotaIPI = vData
End Property

Public Property Get dAliquotaIPI() As Double
    dAliquotaIPI = mvardAliquotaIPI
End Property

Public Property Let dValorIPI(ByVal vData As Double)
Attribute dValorIPI.VB_Description = "Valor do IPI"
    mvardValorIPI = vData
End Property

Public Property Get dValorIPI() As Double
    dValorIPI = mvardValorIPI
End Property

Public Property Let dValorDespAcess(ByVal vData As Double)
Attribute dValorDespAcess.VB_Description = "soma de frete + seguro + outras com 2 decimais"
    mvardValorDespAcess = vData
End Property

Public Property Get dValorDespAcess() As Double
    dValorDespAcess = mvardValorDespAcess
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
Attribute dAliquotaICMS.VB_Description = "Aliquota ICMS"
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dValorICMSSubstRet(ByVal vData As Double)
Attribute dValorICMSSubstRet.VB_Description = "Valor ICMS Substituição Retido"
    mvardValorICMSSubstRet = vData
End Property

Public Property Get dValorICMSSubstRet() As Double
    dValorICMSSubstRet = mvardValorICMSSubstRet
End Property

Public Property Let dValorICMSSubstBase(ByVal vData As Double)
Attribute dValorICMSSubstBase.VB_Description = "Valor Base ICMS Substituição"
    mvardValorICMSSubstBase = vData
End Property

Public Property Get dValorICMSSubstBase() As Double
    dValorICMSSubstBase = mvardValorICMSSubstBase
End Property

Public Property Let dValorICMSOutras(ByVal vData As Double)
Attribute dValorICMSOutras.VB_Description = "valor de outras despesas"
    mvardValorICMSOutras = vData
End Property

Public Property Get dValorICMSOutras() As Double
    dValorICMSOutras = mvardValorICMSOutras
End Property

Public Property Let dValorICMSIsentoNaoTrib(ByVal vData As Double)
Attribute dValorICMSIsentoNaoTrib.VB_Description = "Valor ICMS Isento Não Tributado"
    mvardValorICMSIsentoNaoTrib = vData
End Property

Public Property Get dValorICMSIsentoNaoTrib() As Double
    dValorICMSIsentoNaoTrib = mvardValorICMSIsentoNaoTrib
End Property

Public Property Let dValorICMS(ByVal vData As Double)
Attribute dValorICMS.VB_Description = "Valor ICMS"
    mvardValorICMS = vData
End Property

Public Property Get dValorICMS() As Double
    dValorICMS = mvardValorICMS
End Property

Public Property Let dValorICMS51(ByVal vData As Double)
    mvardValorICMS51 = vData
End Property

Public Property Get dValorICMS51() As Double
    dValorICMS51 = mvardValorICMS51
End Property

Public Property Let dValorICMSBase(ByVal vData As Double)
Attribute dValorICMSBase.VB_Description = "Valor ICMS Base"
    mvardValorICMSBase = vData
End Property

Public Property Get dValorICMSBase() As Double
    dValorICMSBase = mvardValorICMSBase
End Property

Public Property Let dValorTotal(ByVal vData As Double)
Attribute dValorTotal.VB_Description = "Valor da Nota (não é o valor líquido a pagar que é igual a esse menos o ValorIRRF)"
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
Attribute sNaturezaOp.VB_Description = "Natureza da Operacao"
    mvarsNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = mvarsNaturezaOp
End Property

Public Property Let lNumIntDocRegES(ByVal vData As Long)
Attribute lNumIntDocRegES.VB_Description = "Número Interno do Reg de ES associado a esta linha"
    mvarlNumIntDocRegES = vData
End Property

Public Property Get lNumIntDocRegES() As Long
    lNumIntDocRegES = mvarlNumIntDocRegES
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Número Interno"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
