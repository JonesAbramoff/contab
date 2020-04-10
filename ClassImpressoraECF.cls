VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpressoraECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilialEmpresa As Integer 'local copy
Private mvariCodigo As Integer 'local copy
Private mvariCodModelo As Integer 'local copy
Private mvarsNumSerie As String 'local copy
Public Property Let sNumSerie(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNumSerie = 5
    mvarsNumSerie = vData
End Property


Public Property Get sNumSerie() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNumSerie
    sNumSerie = mvarsNumSerie
End Property



Public Property Let iCodModelo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodModelo = 5
    mvariCodModelo = vData
End Property


Public Property Get iCodModelo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodModelo
    iCodModelo = mvariCodModelo
End Property



Public Property Let iCodigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodigo = 5
    mvariCodigo = vData
End Property


Public Property Get iCodigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodigo
    iCodigo = mvariCodigo
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



