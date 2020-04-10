VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassFilialEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarlCodEmpresa As Long 'local copy
Private mvariCodFilial As Integer 'local copy
Private mvarsNomeFilial As String 'local copy
Private mvarsNomeEmpresa As String 'local copy
Private mvarsNomeRedEmpresa As String 'local copy
Public Property Let sNomeRedEmpresa(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeRedEmpresa = 5
    mvarsNomeRedEmpresa = vData
End Property


Public Property Get sNomeRedEmpresa() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeRedEmpresa
    sNomeRedEmpresa = mvarsNomeRedEmpresa
End Property



Public Property Let sNomeEmpresa(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeEmpresa = 5
    mvarsNomeEmpresa = vData
End Property


Public Property Get sNomeEmpresa() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeEmpresa
    sNomeEmpresa = mvarsNomeEmpresa
End Property



Public Property Let sNomeFilial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeFilial = 5
    mvarsNomeFilial = vData
End Property


Public Property Get sNomeFilial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeFilial
    sNomeFilial = mvarsNomeFilial
End Property



Public Property Let iCodFilial(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodFilial = 5
    mvariCodFilial = vData
End Property


Public Property Get iCodFilial() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodFilial
    iCodFilial = mvariCodFilial
End Property



Public Property Let lCodEmpresa(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lCodEmpresa = 5
    mvarlCodEmpresa = vData
End Property


Public Property Get lCodEmpresa() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lCodEmpresa
    lCodEmpresa = mvarlCodEmpresa
End Property



