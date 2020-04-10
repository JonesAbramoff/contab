VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDicRotina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Classe Rotina."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarsSigla As String 'local copy
Private mvarsNome As String 'local copy
Private mvarsClasse_Original As String 'local copy
Private mvarsProjeto_Original As String 'local copy
Private mvarsProjeto_Customizado As String 'local copy
Private mvarsClasse_Customizada As String 'local copy
'local variable(s) to hold property value(s)
Private mvariLogAtividade As Integer 'local copy
Public Property Let iLogAtividade(ByVal vData As Integer)
Attribute iLogAtividade.VB_Description = "Log de Atividade: ATIVIDADE ou INATIVIDADE"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.iLogAtividade = 5
    mvariLogAtividade = vData
End Property


Public Property Get iLogAtividade() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.iLogAtividade
    iLogAtividade = mvariLogAtividade
End Property



Public Property Let sClasse_Customizada(ByVal vData As String)
Attribute sClasse_Customizada.VB_Description = "Nome da classe da rotina customizada."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClasse_Customizada = 5
    mvarsClasse_Customizada = vData
End Property


Public Property Get sClasse_Customizada() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClasse_Customizada
    sClasse_Customizada = mvarsClasse_Customizada
End Property



Public Property Let sProjeto_Customizado(ByVal vData As String)
Attribute sProjeto_Customizado.VB_Description = "Nome do projeto da rotina customizada."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProjeto_Customizado = 5
    mvarsProjeto_Customizado = vData
End Property


Public Property Get sProjeto_Customizado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProjeto_Customizado
    sProjeto_Customizado = mvarsProjeto_Customizado
End Property



Public Property Let sProjeto_Original(ByVal vData As String)
Attribute sProjeto_Original.VB_Description = "Nome do projeto da rotina original."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sProjeto_Original = 5
    mvarsProjeto_Original = vData
End Property


Public Property Get sProjeto_Original() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sProjeto_Original
    sProjeto_Original = mvarsProjeto_Original
End Property



Public Property Let sClasse_Original(ByVal vData As String)
Attribute sClasse_Original.VB_Description = "Nome da classe da rotina original."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sClasse_Original = 5
    mvarsClasse_Original = vData
End Property


Public Property Get sClasse_Original() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sClasse_Original
    sClasse_Original = mvarsClasse_Original
End Property



Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição da Rotina"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sNome = 5
    mvarsNome = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sNome
    sDescricao = mvarsNome
End Property



Public Property Let sSigla(ByVal vData As String)
Attribute sSigla.VB_Description = "Sigla de Rotina"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sSigla = 5
    mvarsSigla = vData
End Property


Public Property Get sSigla() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sSigla
    sSigla = mvarsSigla
End Property



