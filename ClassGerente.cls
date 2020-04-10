VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGerente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariFilialEmpresa As Integer
Private mvariCodigo As Integer
Private mvarsCodUsuario As String
Private mvarsNomeReduzido As String

Public Property Let sNomeReduzido(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeReduzido = 5
    mvarsNomeReduzido = vData
End Property


Public Property Get sNomeReduzido() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeReduzido
    sNomeReduzido = mvarsNomeReduzido
End Property
Public Property Let sCodUsuario(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodUsuario = 5
    mvarsCodUsuario = vData
End Property


Public Property Get sCodUsuario() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodUsuario
    sCodUsuario = mvarsCodUsuario
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



Public Property Let iFilialEmpresa(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iFilialEmpresa = 5
    mvariFilialEmpresa = vData
End Property


Public Property Get iFilialEmpresa() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iFilialEmpresa
    iFilialEmpresa = mvariFilialEmpresa
End Property



