VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassLivrosFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvariCodigo As Integer
Private mvarsDescricao As String
Private mvariCodTributo As Integer
Private mvariPeriodicidade As Integer
Private mvariApuracao As Integer

Public Property Let iApuracao(ByVal vData As Integer)
Attribute iApuracao.VB_Description = "0: nao é o livro de apuracao, 1:é o livro de apuracao"
    mvariApuracao = vData
End Property

Public Property Get iApuracao() As Integer
    iApuracao = mvariApuracao
End Property

Public Property Let iPeriodicidade(ByVal vData As Integer)
Attribute iPeriodicidade.VB_Description = "mensal, quinzenal,...default p/configuracao"
    mvariPeriodicidade = vData
End Property

Public Property Get iPeriodicidade() As Integer
    iPeriodicidade = mvariPeriodicidade
End Property

Public Property Let iCodTributo(ByVal vData As Integer)
Attribute iCodTributo.VB_Description = "Código do Tributo do Livro (Tabela de Tributos)"
    mvariCodTributo = vData
End Property

Public Property Get iCodTributo() As Integer
    iCodTributo = mvariCodTributo
End Property

Public Property Let sDescricao(ByVal vData As String)
Attribute sDescricao.VB_Description = "Descrição do Livro Fiscal"
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let iCodigo(ByVal vData As Integer)
Attribute iCodigo.VB_Description = "Código do Livro Fiscal"
    mvariCodigo = vData
End Property

Public Property Get iCodigo() As Integer
    iCodigo = mvariCodigo
End Property
