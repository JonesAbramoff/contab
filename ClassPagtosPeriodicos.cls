VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPagtosPeriodicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvarlCodigo As Long 'local copy
Private mvarsDescricao As String 'local copy
Private mvarlFornecedor As Long 'local copy
Private mvariFilial As Integer 'local copy
Private mvariPeriodicidade As Integer 'local copy
Private mvardtInicio As Date 'local copy
Private mvardtTermino As Date 'local copy
Private mvardtProximo As Date 'local copy

Public Property Let dtProximo(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtTermino = 5
    mvardtProximo = vData
End Property


Public Property Get dtProximo() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtTermino
    dtProximo = mvardtProximo
End Property

Public Property Let dtTermino(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtTermino = 5
    mvardtTermino = vData
End Property


Public Property Get dtTermino() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtTermino
    dtTermino = mvardtTermino
End Property



Public Property Let dtInicio(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtInicio = 5
    mvardtInicio = vData
End Property


Public Property Get dtInicio() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtInicio
    dtInicio = mvardtInicio
End Property



Public Property Let iPeriodicidade(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iPeriodicidade = 5
    mvariPeriodicidade = vData
End Property


Public Property Get iPeriodicidade() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iPeriodicidade
    iPeriodicidade = mvariPeriodicidade
End Property



Public Property Let iFilial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilial = 5
    mvariFilial = vData
End Property


Public Property Get iFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilial
    iFilial = mvariFilial
End Property



Public Property Let lFornecedor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornecedor = 5
    mvarlFornecedor = vData
End Property


Public Property Get lFornecedor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornecedor
    lFornecedor = mvarlFornecedor
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



