VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassTiposDeAplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvarsContaAplicacao As String
Private mvarsContaReceita As String
Private mvarsHistorico As String
Private mvariInativo As Integer

Public Property Let iInativo(ByVal vData As Integer)
Attribute iInativo.VB_Description = "Status do tipo de aplicação. 0 = ativo 1 = inativo"
    mvariInativo = vData
End Property

Public Property Get iInativo() As Integer
    iInativo = mvariInativo
End Property

Public Property Let sHistorico(ByVal vData As String)
Attribute sHistorico.VB_Description = "Historico a ser usado no movimento da conta corrente."
    mvarsHistorico = vData
End Property

Public Property Get sHistorico() As String
    sHistorico = mvarsHistorico
End Property

Public Property Let sContaReceita(ByVal vData As String)
Attribute sContaReceita.VB_Description = "Conta Contabil onde será lançado a receita proveniente da  aplicação."
    mvarsContaReceita = vData
End Property

Public Property Get sContaReceita() As String
    sContaReceita = mvarsContaReceita
End Property

Public Property Let sContaAplicacao(ByVal vData As String)
Attribute sContaAplicacao.VB_Description = "Conta Contabil onde vai ser lançada a aplicação"
    mvarsContaAplicacao = vData
End Property

Public Property Get sContaAplicacao() As String
    sContaAplicacao = mvarsContaAplicacao
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descricao do Tipo de Aplicacao."
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Identificador de um tipo de aplicação. Número com no máximo 3 algarismos"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property



