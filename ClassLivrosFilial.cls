VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivrosFilial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodLivro As Integer
Private mvariFilialEmpresa As Integer
Private mvariImprime As Integer
Private mvariNumeroProxLivro As Integer
Private mvariNumeroProxFolha As Integer
Private mvariPeriodicidade As Integer
Private mvardtDataInicial As Date
Private mvardtDataFinal As Date
Private mvardtImpressoEm As Date
Private mvarsDescricao As String

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do Livro Fiscal"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dtImpressoEm(ByVal vData As Date)
Attribute dtImpressoEm.VB_Description = "Data da Impressão Definitiva"
    mvardtImpressoEm = vData
End Property

Public Property Get dtImpressoEm() As Date
    dtImpressoEm = mvardtImpressoEm
End Property

Public Property Let dtDataFinal(ByVal vData As Date)
Attribute dtDataFinal.VB_Description = "Data do Final do livro Aberto"
    mvardtDataFinal = vData
End Property

Public Property Get dtDataFinal() As Date
    dtDataFinal = mvardtDataFinal
End Property

Public Property Let dtDataInicial(ByVal vData As Date)
Attribute dtDataInicial.VB_Description = "Data do Inicio do livro Aberto"
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let iPeriodicidade(ByVal vData As Integer)
Attribute iPeriodicidade.VB_Description = "mensal, quinzenal,..."
    mvariPeriodicidade = vData
End Property

Public Property Get iPeriodicidade() As Integer
    iPeriodicidade = mvariPeriodicidade
End Property

Public Property Let iNumeroProxFolha(ByVal vData As Integer)
Attribute iNumeroProxFolha.VB_Description = "Número da Proxima Folha"
    mvariNumeroProxFolha = vData
End Property

Public Property Get iNumeroProxFolha() As Integer
    iNumeroProxFolha = mvariNumeroProxFolha
End Property

Public Property Let iNumeroProxLivro(ByVal vData As Integer)
Attribute iNumeroProxLivro.VB_Description = "Número do Próximo Livro"
    mvariNumeroProxLivro = vData
End Property

Public Property Get iNumeroProxLivro() As Integer
    iNumeroProxLivro = mvariNumeroProxLivro
End Property

Public Property Let iImprime(ByVal vData As Integer)
Attribute iImprime.VB_Description = "0:Nao, 1:Sim"
    mvariImprime = vData
End Property

Public Property Get iImprime() As Integer
    iImprime = mvariImprime
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let iCodLivro(ByVal vData As Integer)
Attribute iCodLivro.VB_Description = "Código do Livro"
    mvariCodLivro = vData
End Property

Public Property Get iCodLivro() As Integer
    iCodLivro = mvariCodLivro
End Property
