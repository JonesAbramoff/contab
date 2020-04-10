VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdmGrupoBrowseCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsCodGrupo As String
Private mvarsNomeTela As String
Private mvarsNomeArq As String
Private mvarsNome As String

Public Property Let sNome(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNome() As String
    sNome = mvarsNome
End Property

Public Property Let sNomeArq(ByVal vData As String)
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Property Let sCodGrupo(ByVal vData As String)
    mvarsCodGrupo = vData
End Property

Public Property Get sCodGrupo() As String
    sCodGrupo = mvarsCodGrupo
End Property



Private Sub Class_Initialize()

    mvarsCodGrupo = String(STRING_GRUPO, 0)
    mvarsNomeTela = String(STRING_NOME_TELA, 0)
    mvarsNomeArq = String(STRING_NOME_TABELA, 0)
    mvarsNome = String(STRING_NOME_CAMPO, 0)

End Sub
