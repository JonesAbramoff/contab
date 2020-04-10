VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoMoeda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariMoeda As Integer 'local copy
Private mvardtData As Date 'local copy
Private mvardValor As Double 'local copy
Public Property Let dValor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dValor = 5
    mvardValor = vData
End Property


Public Property Get dValor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dValor
    dValor = mvardValor
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



Public Property Let iMoeda(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iMoeda = 5
    mvariMoeda = vData
End Property


Public Property Get iMoeda() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iMoeda
    iMoeda = mvariMoeda
End Property



