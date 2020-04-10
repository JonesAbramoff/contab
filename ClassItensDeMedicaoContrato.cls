VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItensMedCtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlMedicao As Long
Private mvarlNumIntItensContrato As Long
Private mvardQuantidade As Double
Private mvardCusto As Double
Private mvardVlrCobrar As Double
Private mvariStatus As Integer
Private mvarobjItensDeContrato As ClassItensDeContrato
Private mvardtDataCobranca As Date
Private mvardtDataRefIni As Date
Private mvardtDataRefFim As Date
Private mvariItem As Integer
Private mvarsContrato As String
Private mvarlNumIntItemNFiscal As Long

Public Property Let lNumIntItemNFiscal(ByVal vData As Long)
    mvarlNumIntItemNFiscal = vData
End Property

Public Property Get lNumIntItemNFiscal() As Long
    lNumIntItemNFiscal = mvarlNumIntItemNFiscal
End Property

Public Property Let sContrato(ByVal vData As String)
    mvarsContrato = vData
End Property

Public Property Get sContrato() As String
     sContrato = mvarsContrato
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
     iItem = mvariItem
End Property

Public Property Let dtDataRefIni(ByVal vData As Date)
    mvardtDataRefIni = vData
End Property

Public Property Get dtDataRefIni() As Date
     dtDataRefIni = mvardtDataRefIni
End Property

Public Property Let dtDataRefFim(ByVal vData As Date)
    mvardtDataRefFim = vData
End Property

Public Property Get dtDataRefFim() As Date
     dtDataRefFim = mvardtDataRefFim
End Property

Public Property Let dtDataCobranca(ByVal vData As Date)
    mvardtDataCobranca = vData
End Property

Public Property Get dtDataCobranca() As Date
    dtDataCobranca = mvardtDataCobranca
End Property

Public Property Set objItensDeContrato(ByVal vData As ClassItensDeContrato)
    Set mvarobjItensDeContrato = vData
End Property

Public Property Get objItensDeContrato() As ClassItensDeContrato
    Set objItensDeContrato = mvarobjItensDeContrato
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
     iStatus = mvariStatus
End Property

Public Property Let dVlrCobrar(ByVal vData As Double)
    mvardVlrCobrar = vData
End Property

Public Property Get dVlrCobrar() As Double
     dVlrCobrar = mvardVlrCobrar
End Property

Public Property Let dCusto(ByVal vData As Double)
    mvardCusto = vData
End Property

Public Property Get dCusto() As Double
     dCusto = mvardCusto
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
     dQuantidade = mvardQuantidade
End Property

Public Property Let lNumIntItensContrato(ByVal vData As Long)
    mvarlNumIntItensContrato = vData
End Property

Public Property Get lNumIntItensContrato() As Long
     lNumIntItensContrato = mvarlNumIntItensContrato
End Property

Public Property Let lMedicao(ByVal vData As Long)
    mvarlMedicao = vData
End Property

Public Property Get lMedicao() As Long
     lMedicao = mvarlMedicao
End Property



