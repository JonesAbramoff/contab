VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemAdicaoDIItemNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lNumIntItemAdicaoDI As Long
Public lNumIntItemNF As Long

Public iAdicao As Integer
Public iItemAdicao As Integer

Public iItemNF As Integer

Public sProduto As String
Public sDescricao As String
Public sUM As String
Public dQuantidade As Double

'valores em R$
Public dValorAduaneiro As Double
Public dValorII As Double
Public dPrecoUnitario As Double 'valor final...para constar na nf

Public dIPIAliquotaAdicaoDI As Double
Public dICMSAliquotaAdicaoDI As Double
Public dPISAliquotaAdicaoDI As Double
Public dCOFINSAliquotaAdicaoDI As Double
Public dDespImpValorRateado As Double 'valor rateado de despesas de importacao

Public dICMSPercRedBaseAdicaoDI As Double
Public dIPIUnidadePadraoValor As Double

Public Property Get dValorTotal() As Double
'valor final...para constar na nf
    dValorTotal = Arredonda_Moeda(dValorAduaneiro + dValorII)
End Property
