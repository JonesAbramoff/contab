VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86MovEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarsProduto As String 'local copy
Private mvariTipoNumIntDocOrigem As Integer 'local copy
Private mvardtDataMov As Date 'local copy
Private mvarsSiglaUM As String 'local copy
Private mvarsSiglaDoc As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsEntradaOuSaida As String 'local copy
Private mvarsEntradaSaidaCMP As String 'local copy
Private mvardCustoUnitario As Double 'local copy
Private mvardValorTotal As Double 'local copy
Private mvarlNumIntDocOrigem As Long 'local copy
Private mvariFilialEmpresa As Integer 'local copy
Private mvarsSerie As String 'local copy
Private mvarsNumeroDoc As String 'local copy
Private mvarsHistorico As String 'local copy
'local variable(s) to hold property value(s)
Private mvarlCodigo As Long 'local copy
Public Property Let lCodigo(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigo = 5
    mvarlCodigo = vData
End Property


Public Property Get lCodigo() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigo
    lCodigo = mvarlCodigo
End Property




Public Property Let sHistorico(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sHistorico = 5
    mvarsHistorico = vData
End Property


Public Property Get sHistorico() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sHistorico
    sHistorico = mvarsHistorico
End Property



Public Property Let sNumeroDoc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNumeroDoc = 5
    mvarsNumeroDoc = vData
End Property


Public Property Get sNumeroDoc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNumeroDoc
    sNumeroDoc = mvarsNumeroDoc
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



Public Property Let lNumIntDocOrigem(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocOrigem = 5
    mvarlNumIntDocOrigem = vData
End Property


Public Property Get lNumIntDocOrigem() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocOrigem
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property



Public Property Let dValorTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorTotal = 5
    mvardValorTotal = vData
End Property


Public Property Get dValorTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorTotal
    dValorTotal = mvardValorTotal
End Property



Public Property Let dCustoUnitario(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoUnitario = 5
    mvardCustoUnitario = vData
End Property


Public Property Get dCustoUnitario() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoUnitario
    dCustoUnitario = mvardCustoUnitario
End Property



Public Property Let sEntradaSaidaCMP(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEntradaSaidaCMP = 5
    mvarsEntradaSaidaCMP = vData
End Property


Public Property Get sEntradaSaidaCMP() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEntradaSaidaCMP
    sEntradaSaidaCMP = mvarsEntradaSaidaCMP
End Property



Public Property Let sEntradaOuSaida(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sEntradaOuSaida = 5
    mvarsEntradaOuSaida = vData
End Property


Public Property Get sEntradaOuSaida() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sEntradaOuSaida
    sEntradaOuSaida = mvarsEntradaOuSaida
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



Public Property Let sSiglaDoc(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaDoc = 5
    mvarsSiglaDoc = vData
End Property


Public Property Get sSiglaDoc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaDoc
    sSiglaDoc = mvarsSiglaDoc
End Property



Public Property Let sSiglaUM(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSiglaUM = 5
    mvarsSiglaUM = vData
End Property


Public Property Get sSiglaUM() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSiglaUM
    sSiglaUM = mvarsSiglaUM
End Property



Public Property Let dtDataMov(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataMov = 5
    mvardtDataMov = vData
End Property


Public Property Get dtDataMov() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataMov
    dtDataMov = mvardtDataMov
End Property



Public Property Let iTipoNumIntDocOrigem(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoNumIntDocOrigem = 5
    mvariTipoNumIntDocOrigem = vData
End Property


Public Property Get iTipoNumIntDocOrigem() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoNumIntDocOrigem
    iTipoNumIntDocOrigem = mvariTipoNumIntDocOrigem
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



