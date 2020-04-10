VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTelasOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsNomeTela As String 'local copy
Private mvarsTituloTela As String 'local copy
Private mvarsNomeOpcao As String 'local copy
'local variable(s) to hold property value(s)
Private mvarobjTela As Object
Private mvarobjComboOpcoes As Object

Public Property Set objComboOpcoes(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objTela = Form1
    Set mvarobjComboOpcoes = vData
End Property

Public Property Get objComboOpcoes() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objTela
    Set objComboOpcoes = mvarobjComboOpcoes
End Property

Public Property Set objTela(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.objTela = Form1
    Set mvarobjTela = vData
End Property


Public Property Get objTela() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.objTela
    Set objTela = mvarobjTela
End Property



Public Property Let sNomeOpcao(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeOpcao = 5
    mvarsNomeOpcao = vData
End Property


Public Property Get sNomeOpcao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeOpcao
    sNomeOpcao = mvarsNomeOpcao
End Property



Public Property Let sTituloTela(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sTituloTela = 5
    mvarsTituloTela = vData
End Property


Public Property Get sTituloTela() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sTituloTela
    sTituloTela = mvarsTituloTela
End Property



Public Property Let sNomeTela(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeTela = 5
    mvarsNomeTela = vData
End Property


Public Property Get sNomeTela() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeTela
    sNomeTela = mvarsNomeTela
End Property



