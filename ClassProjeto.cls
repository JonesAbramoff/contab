VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarsDescricao As String 'local copy
Private mvarsNomeReduzido As String 'local copy
Private mvarlCodigo As Long 'local copy
Private mvarlCodCliente As Long 'local copy
Private mvariCodFilial As Integer 'local copy
Private mvarsResponsavel As String 'local copy
Private mvardtDataCriacao As Date 'local copy
Private mvardtDataValidade As Date 'local copy
Private mvarsObservacao As String 'local copy
Private mvarcolProjetoItens As New Collection 'local copy

Public Property Set colProjetoItens(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.colProjetoItens = Form1
    Set mvarcolProjetoItens = vData
End Property

Public Property Get colProjetoItens() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.colProjetoItens
    Set colProjetoItens = mvarcolProjetoItens
End Property

Public Property Let sObservacao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sObservacao = 5
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sObservacao
    sObservacao = mvarsObservacao
End Property

Public Property Let dtDataValidade(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataValidade = 5
    mvardtDataValidade = vData
End Property

Public Property Get dtDataValidade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataValidade
    dtDataValidade = mvardtDataValidade
End Property

Public Property Let dtDataCriacao(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataCriacao = 5
    mvardtDataCriacao = vData
End Property

Public Property Get dtDataCriacao() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataCriacao
    dtDataCriacao = mvardtDataCriacao
End Property

Public Property Let sResponsavel(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sResponsavel = 5
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sResponsavel
    sResponsavel = mvarsResponsavel
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeReduzido = 5
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeReduzido
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let iCodFilial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodFilial = 5
    mvariCodFilial = vData
End Property

Public Property Get iCodFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodFilial
    iCodFilial = mvariCodFilial
End Property

Public Property Let lCodCliente(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodCliente = 5
    mvarlCodCliente = vData
End Property

Public Property Get lCodCliente() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodCliente
    lCodCliente = mvarlCodCliente
End Property

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



