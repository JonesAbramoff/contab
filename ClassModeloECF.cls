VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassModeloECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvariCodigo As Integer 'local copy
Private mvarsNome As String 'local copy
Private mvarsFabricante As String 'local copy

Public Property Let sFabricante(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sFabricante = 5
    mvarsFabricante = vData
End Property

Public Property Get sFabricante() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sFabricante
    sFabricante = mvarsFabricante
End Property

Public Property Let sNome(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNome = 5
    mvarsNome = vData
End Property

Public Property Get sNome() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNome
    sNome = mvarsNome
End Property

Public Property Let iCodigo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iCodigo = 5
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iCodigo
    iCodigo = mvariCodigo
End Property

