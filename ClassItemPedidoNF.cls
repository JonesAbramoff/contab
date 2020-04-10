VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemPedidoNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ClassTributacaoItemPV"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvarlCodPedido As Long 'local copy
Private mvarsProduto As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvardPrecoUnitario As Double 'local copy
Private mvardPrecoTotal As Double 'local copy
Private mvardValorDesconto As Double 'local copy
Private mvariClasseUM As Integer 'local copy
Private mvardQuantFaturada As Double 'local copy
Private mvarsDescricao As String 'local copy
Private mvarlNumIntDoc As Long 'local copy
Private mvarsUMItemNF As String 'local copy
Private mvarsUMItemPV As String 'local copy
Private mvarlNumNF As Long 'local copy
Private mvariItemNF As Integer 'local copy
Private mvariItemPV As Integer 'local copy
Private mvardPercDesconto As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarsSerie As String 'local copy
Private mvardtDataEmissao As Date 'local copy
Public Property Let dtDataEmissao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataEmissao = 5
    mvardtDataEmissao = vData
End Property


Public Property Get dtDataEmissao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataEmissao
    dtDataEmissao = mvardtDataEmissao
End Property



Public Property Let sSerie(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSerie = 5
    mvarsSerie = vData
End Property


Public Property Get sSerie() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSerie
    sSerie = mvarsSerie
End Property



Public Property Let dPercDesconto(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercDesconto = 5
    mvardPercDesconto = vData
End Property


Public Property Get dPercDesconto() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercDesconto
    dPercDesconto = mvardPercDesconto
End Property



Public Property Let iItemPV(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iItemPV = 5
    mvariItemPV = vData
End Property


Public Property Get iItemPV() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iItemPV
    iItemPV = mvariItemPV
End Property



Public Property Let iItemNF(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iItemNF = 5
    mvariItemNF = vData
End Property


Public Property Get iItemNF() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iItemNF
    iItemNF = mvariItemNF
End Property



Public Property Let lNumNF(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumNF = 5
    mvarlNumNF = vData
End Property


Public Property Get lNumNF() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumNF
    lNumNF = mvarlNumNF
End Property



Public Property Let sUMItemPV(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMItemPV = 5
    mvarsUMItemPV = vData
End Property


Public Property Get sUMItemPV() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMItemPV
    sUMItemPV = mvarsUMItemPV
End Property



Public Property Let sUMItemNF(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMItemNF = 5
    mvarsUMItemNF = vData
End Property


Public Property Get sUMItemNF() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMItemNF
    sUMItemNF = mvarsUMItemNF
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



Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
End Property



Public Property Let dQuantFaturada(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantFaturada = 5
    mvardQuantFaturada = vData
End Property


Public Property Get dQuantFaturada() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantFaturada
    dQuantFaturada = mvardQuantFaturada
End Property



Public Property Let iClasseUM(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iClasseUM = 5
    mvariClasseUM = vData
End Property


Public Property Get iClasseUM() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iClasseUM
    iClasseUM = mvariClasseUM
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



Public Property Let dPrecoTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoTotal = 5
    mvardPrecoTotal = vData
End Property


Public Property Get dPrecoTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoTotal
    dPrecoTotal = mvardPrecoTotal
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



Public Property Let lCodPedido(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodPedido = 5
    mvarlCodPedido = vData
End Property


Public Property Get lCodPedido() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodPedido
    lCodPedido = mvarlCodPedido
End Property



Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



