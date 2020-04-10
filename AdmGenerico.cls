VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarvVariavel As Variant 'local copy

Public Property Let vVariavel(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.vVariavel = 5
    mvarvVariavel = vData
End Property


Public Property Set vVariavel(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.vVariavel = Form1
    Set mvarvVariavel = vData
End Property


Public Property Get vVariavel() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.vVariavel
    If IsObject(mvarvVariavel) Then
        Set vVariavel = mvarvVariavel
    Else
        vVariavel = mvarvVariavel
    End If
End Property



