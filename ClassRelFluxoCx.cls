VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelFluxoCx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntRel As Long 'local copy
Private mvardtData As Date 'local copy
Private mvardRecDespesas As Double 'local copy
Private mvardTotalRec As Double 'local copy
Private mvardDespBanc As Double 'local copy
Private mvardPagDesp As Double 'local copy
Private mvardTotalPag As Double 'local copy
Private mvardSaldoInicial As Double 'local copy
Private mvardSaldoFinal As Double 'local copy
Private mvardSaldoAplicacoes As Double 'local copy
Private mvardTotalAplic As Double 'local copy
Private mvardSaldoFinalTotal As Double 'local copy

Public Property Let dSaldoFinalTotal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dSaldoFinalTotal = 5
    mvardSaldoFinalTotal = vData
End Property


Public Property Get dSaldoFinalTotal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dSaldoFinalTotal
    dSaldoFinalTotal = mvardSaldoFinalTotal
End Property



Public Property Let dTotalAplic(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTotalAplic = 5
    mvardTotalAplic = vData
End Property


Public Property Get dTotalAplic() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTotalAplic
    dTotalAplic = mvardTotalAplic
End Property



Public Property Let dSaldoAplicacoes(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dSaldoAplicacoes = 5
    mvardSaldoAplicacoes = vData
End Property


Public Property Get dSaldoAplicacoes() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dSaldoAplicacoes
    dSaldoAplicacoes = mvardSaldoAplicacoes
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



Public Property Let dTotalPag(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTotalPag = 5
    mvardTotalPag = vData
End Property


Public Property Get dTotalPag() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTotalPag
    dTotalPag = mvardTotalPag
End Property



Public Property Let dPagDesp(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dPagDesp = 5
    mvardPagDesp = vData
End Property


Public Property Get dPagDesp() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dPagDesp
    dPagDesp = mvardPagDesp
End Property



Public Property Let dDespBanc(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dDespBanc = 5
    mvardDespBanc = vData
End Property


Public Property Get dDespBanc() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dDespBanc
    dDespBanc = mvardDespBanc
End Property



Public Property Let dTotalRec(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dTotalRec = 5
    mvardTotalRec = vData
End Property


Public Property Get dTotalRec() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dTotalRec
    dTotalRec = mvardTotalRec
End Property



Public Property Let dRecDespesas(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dRecDespesas = 5
    mvardRecDespesas = vData
End Property


Public Property Get dRecDespesas() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dRecDespesas
    dRecDespesas = mvardRecDespesas
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



Public Property Let lNumIntRel(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntRel = 5
    mvarlNumIntRel = vData
End Property


Public Property Get lNumIntRel() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntRel
    lNumIntRel = mvarlNumIntRel
End Property



