VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassDicGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Grupo de Usuários"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCodGrupo As String 'local copy
Private mvarDescricao As String 'local copy
Private mvarDataValidade As Date 'local copy
Private mvarLogAtividade As Integer 'local copy
Public Property Let iLogAtividade(ByVal vData As Integer)
Attribute iLogAtividade.VB_Description = "1-> Faz log da atividade do Grupo    0->Não faz log"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LogAtividade = 5
    mvarLogAtividade = vData
End Property


Public Property Get iLogAtividade() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LogAtividade
    iLogAtividade = mvarLogAtividade
End Property



Public Property Let dtDataValidade(ByVal vData As Date)
Attribute dtDataValidade.VB_Description = "Data até a qual o Grupo vale."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DataValidade = 5
    mvarDataValidade = vData
End Property


Public Property Get dtDataValidade() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DataValidade
    dtDataValidade = mvarDataValidade
End Property



Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do Grupo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Descricao = 5
    mvarDescricao = vData
End Property


Public Property Get sDescricao() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Descricao
    sDescricao = mvarDescricao
End Property



Public Property Let sCodGrupo(ByVal vData As String)
Attribute sCodGrupo.VB_Description = "Código do Grupo"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodGrupo = 5
    mvarCodGrupo = vData
End Property


Public Property Get sCodGrupo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodGrupo
    sCodGrupo = mvarCodGrupo
End Property



