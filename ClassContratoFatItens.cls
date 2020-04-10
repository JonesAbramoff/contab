VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContratoFatItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardCusto As Double
Private mvardVlrCobrar As Double
Private mvarsCcl As String
Private mvarsDescProd As String
Private mvarsProduto As String
Private mvardValor As Double
Private mvardQuantidade As Double
Private mvariItem As Integer
Private mvarlMedicao As Long
Private mvarlNumIntItensContrato As Long
Private mvarsUM As String
Private mvardtDataProxCobranca As Date
Private mvarsNaturezaOp As String
Private mvariTipoTributacao As Integer
Private mvardtDataRefIni As Date
Private mvardtDataRefFim As Date

Public Property Let dtDataRefFim(ByVal vData As Date)
    mvardtDataRefFim = vData
End Property

Public Property Get dtDataRefFim() As Date
    dtDataRefFim = mvardtDataRefFim
End Property

Public Property Let dtDataRefIni(ByVal vData As Date)
    mvardtDataRefIni = vData
End Property

Public Property Get dtDataRefIni() As Date
    dtDataRefIni = mvardtDataRefIni
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

Public Property Let dtDataProxCobranca(ByVal vData As Date)
    mvardtDataProxCobranca = vData
End Property

Public Property Get dtDataProxCobranca() As Date
    dtDataProxCobranca = mvardtDataProxCobranca
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
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

Public Property Let iItem(ByVal vData As Integer)
    mvariItem = vData
End Property

Public Property Get iItem() As Integer
    iItem = mvariItem
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDescProd(ByVal vData As String)
    mvarsDescProd = vData
End Property

Public Property Get sDescProd() As String
    sDescProd = mvarsDescProd
End Property

Public Property Let sCcl(ByVal vData As String)
    mvarsCcl = vData
End Property

Public Property Get sCcl() As String
    sCcl = mvarsCcl
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



