VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodigo As String
Private mvariFilialEmpresa As Integer
Private mvariRede As Integer
Private mvarsNome_Rede As String

Public Property Let sNome_Rede(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigo = 5
    mvarsNome_Rede = vData
End Property

Public Property Get sNome_Rede() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigo
    sNome_Rede = mvarsNome_Rede
End Property

Public Property Let iRede(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iRede = 5
    mvariRede = vData
End Property

Public Property Get iRede() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iRede
    iRede = mvariRede
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

Public Property Let sCodigo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodigo = 5
    mvarsCodigo = vData
End Property

Public Property Get sCodigo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodigo
    sCodigo = mvarsCodigo
End Property



