VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemPedCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object 'Inserido por Wagner 14/07/2006

Private mvardTaxa As Double
Private mvariMoeda As Integer
Private mvarlNumIntDoc As Long
Private mvardtDataLimite As Date
Private mvarsProduto As String
Private mvarsDescProduto As String
Private mvardQuantidade As Double
Private mvardQuantRecebida As Double
Private mvardQuantRecebimento As Double
Private mvarsUM As String
Private mvardPrecoUnitario As Double
Private mvardValorDesconto As Double
Private mvariTipoOrigem As Integer
Private mvarlNumIntOrigem As Long
Private mvardPercentMaisReceb As Double
Private mvardPercentMenosReceb As Double
Private mvariRebebForaFaixa As Integer
Private mvariStatus As Integer
Private mvarlObservacao As Long
Private mvarsObservacao As String
Private mvardValorIPI As Double
Private mvardAliquotaIPI As Double
Private mvardAliquotaICMS As Double
Private mvarcolLocalizacao As New Collection
Private mvarcolItemReqCompras As New Collection
Private mvarsDescCompleta As String
Private mvarlCodigoPC As Long
Private mvarcolDataEntrega As New Collection
Private mvardtDeliveryDate As Date
Private mvariTempoTransito As Integer

Public Property Let iTempoTransito(ByVal vData As Integer)
    mvariTempoTransito = vData
End Property

Public Property Get iTempoTransito() As Integer
    iTempoTransito = mvariTempoTransito
End Property

Public Property Let dtDeliveryDate(ByVal vData As Date)
    mvardtDeliveryDate = vData
End Property

Public Property Get dtDeliveryDate() As Date
    dtDeliveryDate = mvardtDeliveryDate
End Property

Public Property Set colDataEntrega(ByVal vData As Object)
    Set mvarcolDataEntrega = vData
End Property

Public Property Get colDataEntrega() As Collection
    Set colDataEntrega = mvarcolDataEntrega
End Property

Public Property Let lCodigoPC(ByVal vData As Long)
    mvarlCodigoPC = vData
End Property


Public Property Get lCodigoPC() As Long
    lCodigoPC = mvarlCodigoPC
End Property

Public Property Let sUM(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUM = 5
    mvarsUM = vData
End Property


Public Property Get sUM() As String
'Syntax: Debug.Print X.sUM
    sUM = mvarsUM
End Property

Public Property Set colItemReqCompras(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colItemReqCompras = Form1
    Set mvarcolItemReqCompras = vData
End Property


Public Property Get colItemReqCompras() As Collection

'Syntax: Debug.Print X.colItemReqCompras
    Set colItemReqCompras = mvarcolItemReqCompras
End Property



Public Property Set colLocalizacao(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colLocalizacao = Form1
    Set mvarcolLocalizacao = vData
End Property


Public Property Get colLocalizacao() As Collection

'Syntax: Debug.Print X.colLocalizacao
    Set colLocalizacao = mvarcolLocalizacao
End Property



Public Property Let dAliquotaICMS(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAliquotaICMS = 5
    mvardAliquotaICMS = vData
End Property


Public Property Get dAliquotaICMS() As Double

'Syntax: Debug.Print X.dAliquotaICMS
    dAliquotaICMS = mvardAliquotaICMS
End Property



Public Property Let dAliquotaIPI(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAliquotaIPI = 5
    mvardAliquotaIPI = vData
End Property


Public Property Get dAliquotaIPI() As Double

'Syntax: Debug.Print X.dAliquotaIPI
    dAliquotaIPI = mvardAliquotaIPI
End Property



Public Property Let dValorIPI(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorIPI = 5
    mvardValorIPI = vData
End Property


Public Property Get dValorIPI() As Double

'Syntax: Debug.Print X.dValorIPI
    dValorIPI = mvardValorIPI
End Property



Public Property Let iRebebForaFaixa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iRebebForaFaixa = 5
    mvariRebebForaFaixa = vData
End Property


Public Property Get iRebebForaFaixa() As Integer

'Syntax: Debug.Print X.iRebebForaFaixa
    iRebebForaFaixa = mvariRebebForaFaixa
End Property



Public Property Let dPercentMenosReceb(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentMenosReceb = 5
    mvardPercentMenosReceb = vData
End Property


Public Property Get dPercentMenosReceb() As Double

'Syntax: Debug.Print X.dPercentMenosReceb
    dPercentMenosReceb = mvardPercentMenosReceb
End Property



Public Property Let dPercentMaisReceb(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentMaisReceb = 5
    mvardPercentMaisReceb = vData
End Property


Public Property Get dPercentMaisReceb() As Double

'Syntax: Debug.Print X.dPercentMaisReceb
    dPercentMaisReceb = mvardPercentMaisReceb
End Property



Public Property Let lNumIntOrigem(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntOrigem = 5
    mvarlNumIntOrigem = vData
End Property


Public Property Get lNumIntOrigem() As Long

'Syntax: Debug.Print X.lNumIntOrigem
    lNumIntOrigem = mvarlNumIntOrigem
End Property



Public Property Let iTipoOrigem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoOrigem = 5
    mvariTipoOrigem = vData
End Property


Public Property Get iTipoOrigem() As Integer

'Syntax: Debug.Print X.iTipoOrigem
    iTipoOrigem = mvariTipoOrigem
End Property



Public Property Let dValorDesconto(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorDesconto = 5
    mvardValorDesconto = vData
End Property


Public Property Get dValorDesconto() As Double

'Syntax: Debug.Print X.dValorDesconto
    dValorDesconto = mvardValorDesconto
End Property



Public Property Let dPrecoUnitario(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoUnitario = 5
    mvardPrecoUnitario = vData
End Property


Public Property Get dPrecoUnitario() As Double

'Syntax: Debug.Print X.dPrecoUnitario
    dPrecoUnitario = mvardPrecoUnitario
End Property



Public Property Let dQuantRecebimento(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantRecebimento = 5
    mvardQuantRecebimento = vData
End Property


Public Property Get dQuantRecebimento() As Double

'Syntax: Debug.Print X.dQuantRecebimento
    dQuantRecebimento = mvardQuantRecebimento
End Property



Public Property Let dtDataLimite(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataLimite = 5
    mvardtDataLimite = vData
End Property


Public Property Get dtDataLimite() As Date

'Syntax: Debug.Print X.dtDataLimite
    dtDataLimite = mvardtDataLimite
End Property


Public Property Let sObservacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property


Public Property Get sObservacao() As String

'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property



Public Property Let lObservacao(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lObservacao = 5
    mvarlObservacao = vData
End Property


Public Property Get lObservacao() As Long

'Syntax: Debug.Print X.lObservacao
    lObservacao = mvarlObservacao
End Property

Public Property Let dQuantRecebida(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantRecebida = 5
    mvardQuantRecebida = vData
End Property


Public Property Get dQuantRecebida() As Double

'Syntax: Debug.Print X.dQuantRecebida
    dQuantRecebida = mvardQuantRecebida
End Property



Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double

'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let iStatus(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iStatus = 5
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer

'Syntax: Debug.Print X.iStatus
    iStatus = mvariStatus
End Property



Public Property Let sDescProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescProduto = 5
    mvarsDescProduto = vData
End Property


Public Property Get sDescProduto() As String

'Syntax: Debug.Print X.sDescProduto
    sDescProduto = mvarsDescProduto
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String

'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long

'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dTaxa(ByVal vData As Double)
    mvardTaxa = vData
End Property

Public Property Get dTaxa() As Double
    dTaxa = mvardTaxa
End Property

Private Sub Class_Initialize()
    mvardtDeliveryDate = DATA_NULA
End Sub
