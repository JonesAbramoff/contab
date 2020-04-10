VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassGrupoUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsCodGrupo As String 'local copy
Private mvarsDescricao As String 'local copy
Private mvardtDataValidade As Date 'local copy
Private mvariLogAtividade As Integer 'local copy
Public Property Let iLogAtividade(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLogAtividade = 5
    mvariLogAtividade = vData
End Property


Public Property Get iLogAtividade() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLogAtividade
    iLogAtividade = mvariLogAtividade
End Property



Public Property Let dtDataValidade(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.dtDataValidade = 5
    mvardtDataValidade = vData
End Property


Public Property Get dtDataValidade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.dtDataValidade
    dtDataValidade = mvardtDataValidade
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



Public Property Let sCodGrupo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodGrupo = 5
    mvarsCodGrupo = vData
End Property


Public Property Get sCodGrupo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodGrupo
    sCodGrupo = mvarsCodGrupo
End Property



