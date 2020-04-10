VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDicGrupoRelatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsCodGrupo As String 'local copy
Private mvarsCodRel As String 'local copy
Private mvariTipoDeAcesso As Integer 'local copy
Private mvarsNomeTskCustomizado As String 'local copy
Public Property Let sNomeTskCustomizado(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeTskCustomizado = 5
    mvarsNomeTskCustomizado = vData
End Property


Public Property Get sNomeTskCustomizado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeTskCustomizado
    sNomeTskCustomizado = mvarsNomeTskCustomizado
End Property



Public Property Let iTipoDeAcesso(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoDeAcesso = 5
    mvariTipoDeAcesso = vData
End Property


Public Property Get iTipoDeAcesso() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoDeAcesso
    iTipoDeAcesso = mvariTipoDeAcesso
End Property



Public Property Let sCodRel(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodRel = 5
    mvarsCodRel = vData
End Property


Public Property Get sCodRel() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodRel
    sCodRel = mvarsCodRel
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



