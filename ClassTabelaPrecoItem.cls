VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabelaPrecoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Private mvariCodTabela As Integer
Private mvarsCodProduto As String
Private mvardPreco As Double
Private mvariFilialEmpresa As Integer
Private mvardtDataVigencia As Date
Private mvariTabelaDefault As Integer
Private mvarsDescricaoTabela As String
Private mvarsobjInfoUsu As String
Private mvarsObservacao As String
Private mvarobjTela As Object
Private mvardPrecoAtual As Double
Private mvarsTextoGrade As String
Private mvardPercDesconto As Double
Private mvardPrecoComDesconto As Double

Public dComissao As Double

Public Property Let dPrecoComDesconto(ByVal vData As Double)
    mvardPrecoComDesconto = vData
End Property

Public Property Get dPrecoComDesconto() As Double
    dPrecoComDesconto = mvardPrecoComDesconto
End Property

Public Property Let dPercDesconto(ByVal vData As Double)
    mvardPercDesconto = vData
End Property

Public Property Get dPercDesconto() As Double
    dPercDesconto = mvardPercDesconto
End Property

Public Property Let sTextoGrade(ByVal vData As String)
    mvarsTextoGrade = vData
End Property

Public Property Get sTextoGrade() As String
    sTextoGrade = mvarsTextoGrade
End Property

Public Property Let dPrecoAtual(ByVal vData As Double)
    mvardPrecoAtual = vData
End Property

Public Property Get dPrecoAtual() As Double
    dPrecoAtual = mvardPrecoAtual
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sobjInfoUsu(ByVal vData As String)
    mvarsobjInfoUsu = vData
End Property

Public Property Get sobjInfoUsu() As String
    sobjInfoUsu = mvarsobjInfoUsu
End Property

Public Property Let sDescricaoTabela(ByVal vData As String)
    mvarsDescricaoTabela = vData
End Property

Public Property Get sDescricaoTabela() As String
    sDescricaoTabela = mvarsDescricaoTabela
End Property

Public Property Let iTabelaDefault(ByVal vData As Integer)
Attribute iTabelaDefault.VB_Description = "Determina se Tabela de Preço é tabela padrão."
    mvariTabelaDefault = vData
End Property

Public Property Get iTabelaDefault() As Integer
    iTabelaDefault = mvariTabelaDefault
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property

Public Property Let sCodProduto(ByVal vData As String)
    mvarsCodProduto = vData
End Property

Public Property Get sCodProduto() As String
    sCodProduto = mvarsCodProduto
End Property

Public Property Let iCodTabela(ByVal vData As Integer)
    mvariCodTabela = vData
End Property

Public Property Get iCodTabela() As Integer
    iCodTabela = mvariCodTabela
End Property

Public Property Let dtDataVigencia(ByVal vData As Date)
    mvardtDataVigencia = vData
End Property

Public Property Get dtDataVigencia() As Date
    dtDataVigencia = mvardtDataVigencia
End Property


Private Sub Class_Initialize()
    dComissao = -1
End Sub
