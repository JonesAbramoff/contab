VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCondPagtoLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarobjCondPagto As New ClassCondicaoPagto
Private mvariIntervalosVariaveis As Integer
Public Property Let iIntervalosVariaveis(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iIntervalosVariaveis = 5
    mvariIntervalosVariaveis = vData
End Property


Public Property Get iIntervalosVariaveis() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iIntervalosVariaveis
    iIntervalosVariaveis = mvariIntervalosVariaveis
End Property



Public Property Set objCondPagto(ByVal vData As ClassCondicaoPagto)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objCondPagto = Form1
    Set mvarobjCondPagto = vData
End Property


Public Property Get objCondPagto() As ClassCondicaoPagto
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objCondPagto
    Set objCondPagto = mvarobjCondPagto
End Property



