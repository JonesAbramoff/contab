VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDicGrupoTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Classe Grupo x Tela."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsProjeto As String 'local copy
Private mvarsClasse As String 'local copy
Private mvarsCodGrupo As String 'local copy
Private mvariTipoDeAcesso As Integer 'local copy
Private mvarsNomeTela As String 'local copy
Public Property Let sNomeTela(ByVal vData As String)
Attribute sNomeTela.VB_Description = "Nome da Tela."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNomeTela = 5
    mvarsNomeTela = vData
End Property


Public Property Get sNomeTela() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNomeTela
    sNomeTela = mvarsNomeTela
End Property



Public Property Let iTipoDeAcesso(ByVal vData As Integer)
Attribute iTipoDeAcesso.VB_Description = "Tipo de Acesso: COM_ACESSO ou SEM_ACESSO"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iTipoDeAcesso = 5
    mvariTipoDeAcesso = vData
End Property


Public Property Get iTipoDeAcesso() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iTipoDeAcesso
    iTipoDeAcesso = mvariTipoDeAcesso
End Property



Public Property Let sCodGrupo(ByVal vData As String)
Attribute sCodGrupo.VB_Description = "Código do Grupo."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sCodGrupo = 5
    mvarsCodGrupo = vData
End Property


Public Property Get sCodGrupo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sCodGrupo
    sCodGrupo = mvarsCodGrupo
End Property



Public Property Let sClasse(ByVal vData As String)
Attribute sClasse.VB_Description = "Nome da classe da rotina customizada para este grupo."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClasse = 5
    mvarsClasse = vData
End Property


Public Property Get sClasse() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClasse
    sClasse = mvarsClasse
End Property



Public Property Let sProjeto(ByVal vData As String)
Attribute sProjeto.VB_Description = "Nome do projeto da rotina customizada para este grupo."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProjeto = 5
    mvarsProjeto = vData
End Property


Public Property Get sProjeto() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProjeto
    sProjeto = mvarsProjeto
End Property



