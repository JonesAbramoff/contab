VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemCupomFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntCupom As Long
Private mvariStatus As Integer
Private mvarsProduto As String
Private mvarsUnidadeMed As String
Private mvardQuantidade As Double
Private mvardPrecoUnitario As Double
Private mvardPercDesc As Double
Private mvardValorDesconto As Double
Private mvardAliquotaICMS As Double
'Private mvariCancel As Integer
Private mvariItem As Integer
Private mvariFilialEmpresa As Integer
Private mvarsSituacaoTrib As String
Private mvariCodCaixa As Integer
Private mvarlNumIntMovEst As Long
Private mvarsProdutoNomeRed As String
Private mvarobjTributacaoDocItem As New ClassTributacaoDocItem

Public sProdutoDescricao As String

Public Property Set objTributacaoDocItem(ByVal vData As ClassTributacaoDocItem)
    Set mvarobjTributacaoDocItem = vData
End Property

Public Property Get objTributacaoDocItem() As ClassTributacaoDocItem
    Set objTributacaoDocItem = mvarobjTributacaoDocItem
End Property

Public Property Let sProdutoNomeRed(ByVal vData As String)
    mvarsProdutoNomeRed = vData
End Property

Public Property Get sProdutoNomeRed() As String
    sProdutoNomeRed = mvarsProdutoNomeRed
End Property
Public Property Let lNumIntMovEst(ByVal vData As Long)
    mvarlNumIntMovEst = vData
End Property

Public Property Get lNumIntMovEst() As Long
    lNumIntMovEst = mvarlNumIntMovEst
End Property

Public Property Let iCodCaixa(ByVal vData As Integer)
    mvariCodCaixa = vData
End Property

Public Property Get iCodCaixa() As Integer
    iCodCaixa = mvariCodCaixa
End Property

Public Property Let sSituacaoTrib(ByVal vData As String)
    mvarsSituacaoTrib = vData
End Property

Public Property Get sSituacaoTrib() As String
    sSituacaoTrib = mvarsSituacaoTrib
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

'Public Property Let icancel(ByVal vData As Integer)
'    mvariCancel = vData
'End Property
'
'Public Property Get icancel() As Integer
'    icancel = mvariCancel
'End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
    dAliquotaICMS = mvardAliquotaICMS
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dPercDesc(ByVal vData As Double)
    mvardPercDesc = vData
End Property

Public Property Get dPercDesc() As Double
    dPercDesc = mvardPercDesc
End Property

Public Property Let dPrecoUnitario(ByVal vData As Double)
    mvardPrecoUnitario = vData
End Property

Public Property Get dPrecoUnitario() As Double
    dPrecoUnitario = mvardPrecoUnitario
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let lNumIntCupom(ByVal vData As Long)
    mvarlNumIntCupom = vData
End Property

Public Property Get lNumIntCupom() As Long
    lNumIntCupom = mvarlNumIntCupom
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property



