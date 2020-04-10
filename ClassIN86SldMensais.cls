VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIN86SldMensais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvardtData As Date 'local copy
Private mvarsConta As String 'local copy
Private mvardSaldoInicial As Double 'local copy
Private mvardTotalDebitos As Double 'local copy
Private mvardTotalCreditos As Double 'local copy
Private mvardSaldoFinal As Double 'local copy
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



Public Property Let dSaldoFinal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dSaldoFinal = 5
    mvardSaldoFinal = vData
End Property


Public Property Get dSaldoFinal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dSaldoFinal
    dSaldoFinal = mvardSaldoFinal
End Property



Public Property Let dTotalCreditos(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTotalCreditos = 5
    mvardTotalCreditos = vData
End Property


Public Property Get dTotalCreditos() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTotalCreditos
    dTotalCreditos = mvardTotalCreditos
End Property



Public Property Let dTotalDebitos(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTotalDebitos = 5
    mvardTotalDebitos = vData
End Property


Public Property Get dTotalDebitos() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTotalDebitos
    dTotalDebitos = mvardTotalDebitos
End Property



Public Property Let dSaldoInicial(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dSaldoInicial = 5
    mvardSaldoInicial = vData
End Property


Public Property Get dSaldoInicial() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dSaldoInicial
    dSaldoInicial = mvardSaldoInicial
End Property



Public Property Let sConta(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sConta = 5
    mvarsConta = vData
End Property


Public Property Get sConta() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sConta
    sConta = mvarsConta
End Property



Public Property Let dtData(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtData = 5
    mvardtData = vData
End Property


Public Property Get dtData() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtData
    dtData = mvardtData
End Property



