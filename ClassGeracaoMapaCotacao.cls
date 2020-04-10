VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGeracaoMapaCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlCodigoDe As Long 'local copy
Private mvarlCodigoAte As Long 'local copy
Private mvardtDataDe As Date 'local copy
Private mvardtDataAte As Date 'local copy
Private mvarsCategoria As String 'local copy
Private mvarColItensCategoria As New Collection
Private mvarColItensMapaCotacao As New Collection
Private mvarObjMapaCotacao As New ClassMapaCotacao

Public Property Set ColItensCategoria(ByVal vData As Object)
    Set mvarColItensCategoria = vData
End Property

Public Property Get ColItensCategoria() As Collection
    Set ColItensCategoria = mvarColItensCategoria
End Property

Public Property Let sCategoria(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCategoria = 5
    mvarsCategoria = vData
End Property


Public Property Get sCategoria() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCategoria
    sCategoria = mvarsCategoria
End Property



Public Property Let dtDataAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataAte = 5
    mvardtDataAte = vData
End Property


Public Property Get dtDataAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataAte
    dtDataAte = mvardtDataAte
End Property



Public Property Let dtDataDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataDe = 5
    mvardtDataDe = vData
End Property


Public Property Get dtDataDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataDe
    dtDataDe = mvardtDataDe
End Property



Public Property Let lCodigoAte(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigoAte = 5
    mvarlCodigoAte = vData
End Property


Public Property Get lCodigoAte() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigoAte
    lCodigoAte = mvarlCodigoAte
End Property



Public Property Let lCodigoDe(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodigoDe = 5
    mvarlCodigoDe = vData
End Property


Public Property Get lCodigoDe() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodigoDe
    lCodigoDe = mvarlCodigoDe
End Property

Public Property Set ColItensMapaCotacao(ByVal vData As Object)
    Set mvarColItensMapaCotacao = vData
End Property

Public Property Get ColItensMapaCotacao() As Collection
    Set ColItensMapaCotacao = mvarColItensMapaCotacao
End Property

Public Property Set ObjMapaCotacao(ByVal vData As Object)
    Set mvarObjMapaCotacao = vData
End Property

Public Property Get ObjMapaCotacao() As ClassMapaCotacao
    Set ObjMapaCotacao = mvarObjMapaCotacao
End Property
