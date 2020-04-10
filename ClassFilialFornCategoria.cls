VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFilialFornCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariFilial As Integer 'local copy
Private mvarlFornecedor As Long 'local copy
Private mvarsCategoria As String 'local copy
Private mvarsItem As String 'local copy
Public Property Let sItem(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sItem = 5
    mvarsItem = vData
End Property


Public Property Get sItem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sItem
    sItem = mvarsItem
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



