VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTabelaPrecoGrupoItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsProduto As String
Private mvardPrecoNovo As Double
Private mvardPrecoAtual As Double
Private mvarsUM As String
Private mvarsDescricao As String
Private mvarsTextoGrade As String

'*******************
'Auxiliar
Private mvariGradeKitVenda As Integer
Private mvariAnaliticoComGrade As Integer
Private mvarobjProduto As ClassProduto
Private mvarobjProdutoPai As ClassProduto

Public Property Let iGradeKitVenda(ByVal vData As Integer)
    mvariGradeKitVenda = vData
End Property

Public Property Get iGradeKitVenda() As Integer
    iGradeKitVenda = mvariGradeKitVenda
End Property

Public Property Let iAnaliticoComGrade(ByVal vData As Integer)
    mvariAnaliticoComGrade = vData
End Property

Public Property Get iAnaliticoComGrade() As Integer
    iAnaliticoComGrade = mvariAnaliticoComGrade
End Property

Public Property Set objProdutoPai(ByVal vData As Object)
    Set mvarobjProdutoPai = vData
End Property

Public Property Get objProdutoPai() As ClassProduto
    Set objProdutoPai = mvarobjProdutoPai
End Property

Public Property Set objProduto(ByVal vData As Object)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As ClassProduto
    Set objProduto = mvarobjProduto
End Property
'*******************

Public Property Let sTextoGrade(ByVal vData As String)
    mvarsTextoGrade = vData
End Property

Public Property Get sTextoGrade() As String
    sTextoGrade = mvarsTextoGrade
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dPrecoAtual(ByVal vData As Double)
    mvardPrecoAtual = vData
End Property

Public Property Get dPrecoAtual() As Double
    dPrecoAtual = mvardPrecoAtual
End Property

Public Property Let dPrecoNovo(ByVal vData As Double)
    mvardPrecoNovo = vData
End Property

Public Property Get dPrecoNovo() As Double
    dPrecoNovo = mvardPrecoNovo
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

