VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmBrowseArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsNomeTela As String
Private mvarsNomeArq As String
Private mvarsSelecaoSQL As String
Private mvarsProjeto As String
Private mvarsClasse As String
Private mvarsClasseBrowser As String
Private mvarsTrataParametros As String
Private mvarsTituloBrowser As String
Private mvarsRotinaBotaoSeleciona As String
Private mvarsRotinaBotaoEdita As String
Private mvarsRotinaBotaoConsulta As String
Private mvariBotaoSeleciona As Integer
Private mvariBotaoEdita As Integer
Private mvariBotaoConsulta As Integer
Private mvarsProjetoObjeto As String
Private mvarsClasseObjeto As String
Private mvarsNomeTelaConsulta As String
Private mvarsNomeTelaEdita As String
Private mvariBancoDados As Integer

Public Property Let iBancoDados(ByVal vData As Integer)
    mvariBancoDados = vData
End Property

Public Property Get iBancoDados() As Integer
    iBancoDados = mvariBancoDados
End Property

Public Property Let sNomeTelaEdita(ByVal vData As String)
    mvarsNomeTelaEdita = vData
End Property

Public Property Get sNomeTelaEdita() As String
    sNomeTelaEdita = mvarsNomeTelaEdita
End Property

Public Property Let sNomeTelaConsulta(ByVal vData As String)
    mvarsNomeTelaConsulta = vData
End Property

Public Property Get sNomeTelaConsulta() As String
    sNomeTelaConsulta = mvarsNomeTelaConsulta
End Property

Public Property Let sClasseObjeto(ByVal vData As String)
    mvarsClasseObjeto = vData
End Property

Public Property Get sClasseObjeto() As String
    sClasseObjeto = mvarsClasseObjeto
End Property

Public Property Let sProjetoObjeto(ByVal vData As String)
    mvarsProjetoObjeto = vData
End Property

Public Property Get sProjetoObjeto() As String
    sProjetoObjeto = mvarsProjetoObjeto
End Property

Public Property Let iBotaoConsulta(ByVal vData As Integer)
    mvariBotaoConsulta = vData
End Property

Public Property Get iBotaoConsulta() As Integer
    iBotaoConsulta = mvariBotaoConsulta
End Property

Public Property Let iBotaoEdita(ByVal vData As Integer)
    mvariBotaoEdita = vData
End Property

Public Property Get iBotaoEdita() As Integer
    iBotaoEdita = mvariBotaoEdita
End Property

Public Property Let iBotaoSeleciona(ByVal vData As Integer)
    mvariBotaoSeleciona = vData
End Property

Public Property Get iBotaoSeleciona() As Integer
    iBotaoSeleciona = mvariBotaoSeleciona
End Property

Public Property Let sRotinaBotaoConsulta(ByVal vData As String)
    mvarsRotinaBotaoConsulta = vData
End Property

Public Property Get sRotinaBotaoConsulta() As String
    sRotinaBotaoConsulta = mvarsRotinaBotaoConsulta
End Property

Public Property Let sRotinaBotaoEdita(ByVal vData As String)
    mvarsRotinaBotaoEdita = vData
End Property

Public Property Get sRotinaBotaoEdita() As String
    sRotinaBotaoEdita = mvarsRotinaBotaoEdita
End Property

Public Property Let sRotinaBotaoSeleciona(ByVal vData As String)
    mvarsRotinaBotaoSeleciona = vData
End Property

Public Property Get sRotinaBotaoSeleciona() As String
    sRotinaBotaoSeleciona = mvarsRotinaBotaoSeleciona
End Property

Public Property Let sTituloBrowser(ByVal vData As String)
    mvarsTituloBrowser = vData
End Property

Public Property Get sTituloBrowser() As String
    sTituloBrowser = mvarsTituloBrowser
End Property

Public Property Let sTrataParametros(ByVal vData As String)
    mvarsTrataParametros = vData
End Property

Public Property Get sTrataParametros() As String
    sTrataParametros = mvarsTrataParametros
End Property

Public Property Let sClasseBrowser(ByVal vData As String)
    mvarsClasseBrowser = vData
End Property

Public Property Get sClasseBrowser() As String
    sClasseBrowser = mvarsClasseBrowser
End Property

Public Property Let sClasse(ByVal vData As String)
Attribute sClasse.VB_Description = "Classe onde se encontra a rotina de exibição dos dados"
    mvarsClasse = vData
End Property

Public Property Get sClasse() As String
    sClasse = mvarsClasse
End Property

Public Property Let sProjeto(ByVal vData As String)
Attribute sProjeto.VB_Description = "Projeto onde se encontra a rotina de exibição dos dados"
    mvarsProjeto = vData
End Property

Public Property Get sProjeto() As String
    sProjeto = mvarsProjeto
End Property

Public Property Let sSelecaoSQL(ByVal vData As String)
Attribute sSelecaoSQL.VB_Description = "Expressão de Seleção  SQL associada ao browse. Ex: Conta>? AND Lote =?"
    mvarsSelecaoSQL = vData
End Property

Public Property Get sSelecaoSQL() As String
    sSelecaoSQL = mvarsSelecaoSQL
End Property

Public Property Let sNomeArq(ByVal vData As String)
Attribute sNomeArq.VB_Description = "Nome da Tabela"
    mvarsNomeArq = vData
End Property

Public Property Get sNomeArq() As String
    sNomeArq = mvarsNomeArq
End Property

Public Property Let sNomeTela(ByVal vData As String)
Attribute sNomeTela.VB_Description = "Nome da tela de browse"
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property



