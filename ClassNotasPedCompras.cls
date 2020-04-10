VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassNotasPedCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntPedido As Long 'local copy
Private mvariSequencial As Integer 'local copy
Private mvarsNotaPC As String 'local copy
Public Property Let sNotaPC(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNotaPC = 5
    mvarsNotaPC = vData
End Property


Public Property Get sNotaPC() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNotaPC
    sNotaPC = mvarsNotaPC
End Property



Public Property Let iSequencial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iSequencial = 5
    mvariSequencial = vData
End Property


Public Property Get iSequencial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iSequencial
    iSequencial = mvariSequencial
End Property



Public Property Let lNumIntPedido(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntPedido = 5
    mvarlNumIntPedido = vData
End Property


Public Property Get lNumIntPedido() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntPedido
    lNumIntPedido = mvarlNumIntPedido
End Property



