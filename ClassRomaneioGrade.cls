VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRomaneioGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String 'Nome Da tela que está chamando a tela de Romaneio
Private mvariFilialFaturamento As Integer 'Informação usada pela tela de Pedido de Venda
Private mvarobjObjetoTela As Object 'objeto da tela a ser tratado pela tela de Romaneio Grade
Private mvariTipoNFiscal As Integer 'Informação usada pelas telas de Nota Fiscal

'****** USADOS NO FUNCIONAMENTO DA TELA DE ROMANEIO *********
Private mvarcolItensRomaneioGrade As New Collection 'Guarda todos os itens filhos de grade do produto pai.
Private mvariModoFuncionamento As Integer 'Como se trata de uma tela genérica vai servir para indicar quais procedimento devem ou não ser executados.
Private mvariItemAtual As Integer 'Indica qual item no grid está secionado no momento
'************************************************************
Private mvarobjTela As Object 'objeto tela a ser tratado em Romaneio Grade
Private mvardPreco As Double
Private mvarsTextoGrade As String
Private mvarcolRomaneioGradeAux As New Collection
Private mvarobjTabelaPrecoItem As New ClassTabelaPrecoItem
Private mvarobjProduto As New ClassProduto

Public Property Set objProduto(ByVal vData As Object)
    Set mvarobjProduto = vData
End Property

Public Property Get objProduto() As Object
    Set objProduto = mvarobjProduto
End Property

Public Property Set objTabelaPrecoItem(ByVal vData As Object)
    Set mvarobjTabelaPrecoItem = vData
End Property

Public Property Get objTabelaPrecoItem() As Object
    Set objTabelaPrecoItem = mvarobjTabelaPrecoItem
End Property

Public Property Set colRomaneioGradeAux(ByVal vData As Collection)
    Set mvarcolRomaneioGradeAux = vData
End Property

Public Property Get colRomaneioGradeAux() As Collection
    Set colRomaneioGradeAux = mvarcolRomaneioGradeAux
End Property

Public Property Let sTextoGrade(ByVal vData As String)
    mvarsTextoGrade = vData
End Property

Public Property Get sTextoGrade() As String
    sTextoGrade = mvarsTextoGrade
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let iTipoNFiscal(ByVal vData As Integer)
    mvariTipoNFiscal = vData
End Property

Public Property Get iTipoNFiscal() As Integer
    iTipoNFiscal = mvariTipoNFiscal
End Property

Public Property Let iItemAtual(ByVal vData As Integer)
    mvariItemAtual = vData
End Property

Public Property Get iItemAtual() As Integer
    iItemAtual = mvariItemAtual
End Property

Public Property Let iModoFuncionamento(ByVal vData As Integer)
    mvariModoFuncionamento = vData
End Property

Public Property Get iModoFuncionamento() As Integer
    iModoFuncionamento = mvariModoFuncionamento
End Property

Public Property Set colItensRomaneioGrade(ByVal vData As Collection)
    Set mvarcolItensRomaneioGrade = vData
End Property

Public Property Get colItensRomaneioGrade() As Collection
    Set colItensRomaneioGrade = mvarcolItensRomaneioGrade
End Property

Public Property Set objObjetoTela(ByVal vData As Object)
    Set mvarobjObjetoTela = vData
End Property

Public Property Get objObjetoTela() As Object
    Set objObjetoTela = mvarobjObjetoTela
End Property

Public Property Let iFilialFaturamento(ByVal vData As Integer)
    mvariFilialFaturamento = vData
End Property

Public Property Get iFilialFaturamento() As Integer
    iFilialFaturamento = mvariFilialFaturamento
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property
