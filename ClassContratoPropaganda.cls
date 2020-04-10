VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassContratoPropag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlCliente As Long 'local copy
Private mvardtPeriodoDe As Date 'local copy
Private mvardtPeriodoAte As Date 'local copy
Private mvardPercentual As Double 'local copy
Private mvariFilialEmpresa As Integer 'local copy
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



Public Property Let dPercentual(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPercentual = 5
    mvardPercentual = vData
End Property


Public Property Get dPercentual() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPercentual
    dPercentual = mvardPercentual
End Property



Public Property Let dtPeriodoAte(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtPeriodoAte = 5
    mvardtPeriodoAte = vData
End Property


Public Property Get dtPeriodoAte() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtPeriodoAte
    dtPeriodoAte = mvardtPeriodoAte
End Property



Public Property Let dtPeriodoDe(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtPeriodoDe = 5
    mvardtPeriodoDe = vData
End Property


Public Property Get dtPeriodoDe() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtPeriodoDe
    dtPeriodoDe = mvardtPeriodoDe
End Property



Public Property Let lCliente(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCliente = 5
    mvarlCliente = vData
End Property


Public Property Get lCliente() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCliente
    lCliente = mvarlCliente
End Property



