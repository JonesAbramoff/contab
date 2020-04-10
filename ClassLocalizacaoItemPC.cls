VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassLocalizacaoItemPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariAlmoxarifado As Integer 'local copy
Private mvarsCcl As String 'local copy
Private mvardQuantidade As Double 'local copy
Private mvarsContaContabil As String 'local copy
Public Property Let sContaContabil(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sContaContabil = 5
    mvarsContaContabil = vData
End Property


Public Property Get sContaContabil() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sContaContabil
    sContaContabil = mvarsContaContabil
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


Public Property Let sCcl(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCcl = 5
    mvarsCcl = vData
End Property


Public Property Get sCcl() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCcl
    sCcl = mvarsCcl
End Property



Public Property Let iAlmoxarifado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iAlmoxarifado = 5
    mvariAlmoxarifado = vData
End Property


Public Property Get iAlmoxarifado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iAlmoxarifado
    iAlmoxarifado = mvariAlmoxarifado
End Property

