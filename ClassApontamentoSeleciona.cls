VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassApontamentoSeleciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariSoAbertos As Integer 'local copy
Private mvarlCTInicial As Long 'local copy
Private mvarlCTFinal As Long 'local copy
Private mvarsOPInicial As String 'local copy
Private mvarsOPFinal As String 'local copy
Private mvarsProdutoInicial As String 'local copy
Private mvarsProdutoFinal As String 'local copy
Private mvariTodos As Integer 'local copy
Private mvardtDataOPInicial As Date 'local copy
Private mvardtDataOPFinal As Date 'local copy
Private mvarcolPlanoOperacional As New Collection 'local copy

Public Property Set colPlanoOperacional(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colPlanoOperacional = Form1
    Set mvarcolPlanoOperacional = vData
End Property

Public Property Get colPlanoOperacional() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colPlanoOperacional
    Set colPlanoOperacional = mvarcolPlanoOperacional
End Property

Public Property Let dtDataOPFinal(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataOPFinal = 5
    mvardtDataOPFinal = vData
End Property

Public Property Get dtDataOPFinal() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataOPFinal
    dtDataOPFinal = mvardtDataOPFinal
End Property

Public Property Let dtDataOPInicial(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataOPInicial = 5
    mvardtDataOPInicial = vData
End Property

Public Property Get dtDataOPInicial() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataOPInicial
    dtDataOPInicial = mvardtDataOPInicial
End Property

Public Property Let iTodos(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTodos = 5
    mvariTodos = vData
End Property

Public Property Get iTodos() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTodos
    iTodos = mvariTodos
End Property

Public Property Let sProdutoFinal(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoFinal = 5
    mvarsProdutoFinal = vData
End Property

Public Property Get sProdutoFinal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoFinal
    sProdutoFinal = mvarsProdutoFinal
End Property

Public Property Let sProdutoInicial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProdutoInicial = 5
    mvarsProdutoInicial = vData
End Property

Public Property Get sProdutoInicial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProdutoInicial
    sProdutoInicial = mvarsProdutoInicial
End Property

Public Property Let sOPFinal(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOpFinal = 5
    mvarsOPFinal = vData
End Property

Public Property Get sOPFinal() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOpFinal
    sOPFinal = mvarsOPFinal
End Property

Public Property Let sOPInicial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sOpInicial = 5
    mvarsOPInicial = vData
End Property

Public Property Get sOPInicial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sOpInicial
    sOPInicial = mvarsOPInicial
End Property

Public Property Let lCTFinal(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCTFinal = 5
    mvarlCTFinal = vData
End Property

Public Property Get lCTFinal() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCTFinal
    lCTFinal = mvarlCTFinal
End Property

Public Property Let lCTInicial(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCTInicial = 5
    mvarlCTInicial = vData
End Property

Public Property Get lCTInicial() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCTInicial
    lCTInicial = mvarlCTInicial
End Property

Public Property Let iSoAbertos(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSoAbertos = 5
    mvariSoAbertos = vData
End Property

Public Property Get iSoAbertos() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSoAbertos
    iSoAbertos = mvariSoAbertos
End Property



