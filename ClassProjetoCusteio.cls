VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProjetoCusteio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDoc As Long 'local copy
Private mvarlNumIntDocProj As Long 'local copy
Private mvardCustoTotalProjeto As Double 'local copy
Private mvardPrecoTotalProjeto As Double 'local copy
Private mvardtDataCusteio As Date 'local copy
Private mvariConcluido As Integer 'local copy
Private mvarcolProjetoCusteioItens As New Collection

Public Property Set colProjetoCusteioItens(ByVal vData As Collection)
    Set mvarcolProjetoCusteioItens = vData
End Property

Public Property Get colProjetoCusteioItens() As Collection
    Set colProjetoCusteioItens = mvarcolProjetoCusteioItens
End Property

Public Property Let dtDataCusteio(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataCusteio = 5
    mvardtDataCusteio = vData
End Property

Public Property Get dtDataCusteio() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataCusteio
    dtDataCusteio = mvardtDataCusteio
End Property

Public Property Let dPrecoTotalProjeto(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPrecoTotalProjeto = 5
    mvardPrecoTotalProjeto = vData
End Property

Public Property Get dPrecoTotalProjeto() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPrecoTotalProjeto
    dPrecoTotalProjeto = mvardPrecoTotalProjeto
End Property

Public Property Let dCustoTotalProjeto(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dCustoTotalProjeto = 5
    mvardCustoTotalProjeto = vData
End Property

Public Property Get dCustoTotalProjeto() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dCustoTotalProjeto
    dCustoTotalProjeto = mvardCustoTotalProjeto
End Property

Public Property Let lNumIntDocProj(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocProj = 5
    mvarlNumIntDocProj = vData
End Property

Public Property Get lNumIntDocProj() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocProj
    lNumIntDocProj = mvarlNumIntDocProj
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

Public Property Let iConcluido(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iConcluido = 5
    mvariConcluido = vData
End Property

Public Property Get iConcluido() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iConcluido
    iConcluido = mvariConcluido
End Property


