VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAliquotaICMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mvarsSigla As String
Private mvardAliquota As Double
Private mvariISS As Integer
Private mvariFilialEmpresa As Integer
Private mvardValorTotalizadoLoja As Double

Public Property Let dValorTotalizadoLoja(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValorTotalizadoLoja = 5
    mvardValorTotalizadoLoja = vData
End Property


Public Property Get dValorTotalizadoLoja() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValorTotalizadoLoja
    dValorTotalizadoLoja = mvardValorTotalizadoLoja
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



Public Property Let iISS(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iISS = 5
    mvariISS = vData
End Property


Public Property Get iISS() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iISS
    iISS = mvariISS
End Property



Public Property Let dAliquota(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dAliquota = 5
    mvardAliquota = vData
End Property


Public Property Get dAliquota() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dAliquota
    dAliquota = mvardAliquota
End Property



Public Property Let sSigla(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSigla = 5
    mvarsSigla = vData
End Property


Public Property Get sSigla() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSigla
    sSigla = mvarsSigla
End Property



