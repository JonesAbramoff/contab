VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMaquinasInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlNumIntDocMaq As Long 'local copy
Private mvarsProduto As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsUMProduto As String 'local copy
Private mvarsUMTempo As String 'local copy
Public Property Let sUMTempo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMTempo = 5
    mvarsUMTempo = vData
End Property

Public Property Get sUMTempo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMTempo
    sUMTempo = mvarsUMTempo
End Property

Public Property Let sUMProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sUMProduto = 5
    mvarsUMProduto = vData
End Property

Public Property Get sUMProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sUMProduto
    sUMProduto = mvarsUMProduto
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

Public Property Let sProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProduto = 5
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProduto
    sProduto = mvarsProduto
End Property

Public Property Let lNumIntDocMaq(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lNumIntDocMaq = 5
    mvarlNumIntDocMaq = vData
End Property

Public Property Get lNumIntDocMaq() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lNumIntDocMaq
    lNumIntDocMaq = mvarlNumIntDocMaq
End Property



