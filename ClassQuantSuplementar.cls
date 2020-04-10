VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassQuantSuplementar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariTipoDestino As Integer 'local copy
Private mvarlFornCliDestino As Long 'local copy
Private mvariFilialDestino As Integer 'local copy
Private mvardQuantidade As Double 'local copy
Private mvardQuantRequisitada As Double 'local copy
Public Property Let dQuantRequisitada(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantRequisitada = 5
    mvardQuantRequisitada = vData
End Property


Public Property Get dQuantRequisitada() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantRequisitada
    dQuantRequisitada = mvardQuantRequisitada
End Property



Public Property Let dQuantidade(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dQuantidade = 5
    mvardQuantidade = vData
End Property


Public Property Get dQuantidade() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dQuantidade
    dQuantidade = mvardQuantidade
End Property



Public Property Let iFilialDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialDestino = 5
    mvariFilialDestino = vData
End Property


Public Property Get iFilialDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialDestino
    iFilialDestino = mvariFilialDestino
End Property



Public Property Let lFornCliDestino(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lFornCliDestino = 5
    mvarlFornCliDestino = vData
End Property


Public Property Get lFornCliDestino() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lFornCliDestino
    lFornCliDestino = mvarlFornCliDestino
End Property



Public Property Let iTipoDestino(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoDestino = 5
    mvariTipoDestino = vData
End Property


Public Property Get iTipoDestino() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoDestino
    iTipoDestino = mvariTipoDestino
End Property



