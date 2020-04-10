VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOrcamentoLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntOrcamento As Long
Private mvariStatus As Integer
Private mvarsProduto As String
Private mvarsUnidadeMed As String
Private mvardQuantidade As Double
Private mvardPrecoUnitario As Double
Private mvardPercDesc As Double
Private mvardValorDesconto As Double
Private mvardAliquotaICMS As Double
Private mvariItem As Integer

Public Property Let iItem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iItem = 5
    mvariItem = vData
End Property


Public Property Get iItem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iItem
    iItem = mvariItem
End Property

Public Property Let dAliquotaICMS(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAliquotaICMS = 5
    mvardAliquotaICMS = vData
End Property

Public Property Get dAliquotaICMS() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dAliquotaICMS
    dAliquotaICMS = mvardAliquotaICMS
End Property



Public Property Let dValorDesconto(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorDesconto = 5
    mvardValorDesconto = vData
End Property


Public Property Get dValorDesconto() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorDesconto
    dValorDesconto = mvardValorDesconto
End Property



Public Property Let dPercDesc(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercDesc = 5
    mvardPercDesc = vData
End Property


Public Property Get dPercDesc() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercDesc
    dPercDesc = mvardPercDesc
End Property



Public Property Let dPrecoUnitario(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoUnitario = 5
    mvardPrecoUnitario = vData
End Property


Public Property Get dPrecoUnitario() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoUnitario
    dPrecoUnitario = mvardPrecoUnitario
End Property



Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let sUnidadeMed(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUnidadeMed = 5
    mvarsUnidadeMed = vData
End Property


Public Property Get sUnidadeMed() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUnidadeMed
    sUnidadeMed = mvarsUnidadeMed
End Property



Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property


Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property



Public Property Let iStatus(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iStatus = 5
    mvariStatus = vData
End Property


Public Property Get iStatus() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iStatus
    iStatus = mvariStatus
End Property

Public Property Let lNumIntOrcamento(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntOrcamento = 5
    mvarlNumIntOrcamento = vData
End Property


Public Property Get lNumIntOrcamento() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntOrcamento
    lNumIntOrcamento = mvarlNumIntOrcamento
End Property



Public Property Let lNumIntDoc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDoc = 5
    mvarlNumIntDoc = vData
End Property


Public Property Get lNumIntDoc() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDoc
    lNumIntDoc = mvarlNumIntDoc
End Property





