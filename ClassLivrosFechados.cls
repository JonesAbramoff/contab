VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivrosFechados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariCodLivro As Integer
Private mvariFilialEmpresa As Integer
Private mvariNumeroLivro As Integer
Private mvardtDataInicial As Date
Private mvardtDataFinal As Date
Private mvardtDataImpressao As Date
Private mvariFolhaInicial As Integer
Private mvariFolhaFinal As Integer

Public Property Let iFolhaFinal(ByVal vData As Integer)
Attribute iFolhaFinal.VB_Description = "Folha onde o Livro terminou (um livro pode ter no máximo 500 folhas)"
    mvariFolhaFinal = vData
End Property

Public Property Get iFolhaFinal() As Integer
    iFolhaFinal = mvariFolhaFinal
End Property

Public Property Let iFolhaInicial(ByVal vData As Integer)
Attribute iFolhaInicial.VB_Description = "Folha onde o livro comecou"
    mvariFolhaInicial = vData
End Property

Public Property Get iFolhaInicial() As Integer
    iFolhaInicial = mvariFolhaInicial
End Property

Public Property Let dtDataImpressao(ByVal vData As Date)
Attribute dtDataImpressao.VB_Description = "data em que o livro foi impresso definitivamente (ou enviado, se for arquivo)"
    mvardtDataImpressao = vData
End Property

Public Property Get dtDataImpressao() As Date
    dtDataImpressao = mvardtDataImpressao
End Property

Public Property Let dtDataFinal(ByVal vData As Date)
Attribute dtDataFinal.VB_Description = "Data do Final do Livro"
    mvardtDataFinal = vData
End Property

Public Property Get dtDataFinal() As Date
    dtDataFinal = mvardtDataFinal
End Property

Public Property Let dtDataInicial(ByVal vData As Date)
Attribute dtDataInicial.VB_Description = "Data do Inicio do Livro"
    mvardtDataInicial = vData
End Property

Public Property Get dtDataInicial() As Date
    dtDataInicial = mvardtDataInicial
End Property

Public Property Let iNumeroLivro(ByVal vData As Integer)
Attribute iNumeroLivro.VB_Description = "Número do Livro"
    mvariNumeroLivro = vData
End Property

Public Property Get iNumeroLivro() As Integer
    iNumeroLivro = mvariNumeroLivro
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
Attribute iFilialEmpresa.VB_Description = "Filial Empresa do Livro"
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

Public Property Let lNumIntDoc(ByVal vData As Long)
Attribute lNumIntDoc.VB_Description = "Número Interno do Livro  Fechado"
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
