VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassItemClassifABC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsCodProduto As String 'local copy
Private mvarsDescProduto As String 'local copy
Private mvardDemanda As Double 'local copy
Private mvarsClasseABC As String 'local copy
Private mvarlClassifABC As Long 'local copy
Public Property Let lClassifABC(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lClassifABC = 5
    mvarlClassifABC = vData
End Property


Public Property Get lClassifABC() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lClassifABC
    lClassifABC = mvarlClassifABC
End Property



Public Property Let sClasseABC(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClasseABC = 5
    mvarsClasseABC = vData
End Property


Public Property Get sClasseABC() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClasseABC
    sClasseABC = mvarsClasseABC
End Property



Public Property Let dDemanda(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dDemanda = 5
    mvardDemanda = vData
End Property


Public Property Get dDemanda() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dDemanda
    dDemanda = mvardDemanda
End Property



Public Property Let sDescProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescProduto = 5
    mvarsDescProduto = vData
End Property


Public Property Get sDescProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescProduto
    sDescProduto = mvarsDescProduto
End Property



Public Property Let sCodProduto(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodProduto = 5
    mvarsCodProduto = vData
End Property


Public Property Get sCodProduto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodProduto
    sCodProduto = mvarsCodProduto
End Property



