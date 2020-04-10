VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassPonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardX As Double 'local copy
Private mvardY As Double 'local copy
Public Property Let dY(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dY = 5
    mvardY = vData
End Property


Public Property Get dY() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dY
    dY = mvardY
End Property



Public Property Let dX(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dX = 5
    mvardX = vData
End Property


Public Property Get dX() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dX
    dX = mvardX
End Property



