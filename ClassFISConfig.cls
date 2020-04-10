VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFISConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodigo As String
Private mvariFilialEmpresa As Integer
Private mvarsDescricao As String
Private mvariTipo As Integer
Private mvarsConteudo As String

Public Property Let sConteudo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sConteudo = 5
    mvarsConteudo = vData
End Property


Public Property Get sConteudo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sConteudo
    sConteudo = mvarsConteudo
End Property



Public Property Let iTipo(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipo = 5
    mvariTipo = vData
End Property


Public Property Get iTipo() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipo
    iTipo = mvariTipo
End Property



Public Property Let sDescricao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sDescricao = 5
    mvarsDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sDescricao
    sDescricao = mvarsDescricao
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


