VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassExeExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarcolMnemonico As Collection
Private mvarcolGridTransacao As ClassColGridTransacao
Private mvarcolLancamentos As ClassColLancamentos
Private mvarobjTransacao As Object
Private mvarobjContexto As Object
Private mvarcolMnemonicoValor As ClassColMnemonicoValor
Private mvariLinhaAtual As Integer 'linha do grid de expressões de formação de preço que está sendo processado no momento
Private mvariLinha As Integer 'linha do grid. Auxiliar.
Private mvarvValor As Variant
Private mvarsExpressao As String
Private mvarsProduto As String
Private mvarsExpressao1 As String 'copia de mvarsExpressao usada na formação de preço quando ao apertar o botão Ins ou Del precisa mudar a expressao de L9 para L10 (ou vice versa) ou seja aumentou/diminuiu a expressao de 1 caracter
Private mvariCaracteres As Integer 'indica o numero de caracteres adicionados ou retirados da expressao até o momento. Utilizado para tratar o clicar do botão Ins ou Del na formação de preço.
Private mvariInicio_Expressao As Integer 'indica a ultima posição que foi processada
Private mvarcolRegras As Collection

Public Property Set colRegras(ByVal vData As Object)
    Set mvarcolRegras = vData
End Property

Public Property Get colRegras() As Collection
    Set colRegras = mvarcolRegras
End Property

Public Property Set objContexto(ByVal vData As Object)
    Set mvarobjContexto = vData
End Property

Public Property Get objContexto() As Object
    Set objContexto = mvarobjContexto
End Property

Public Property Let iInicio_Expressao(ByVal vData As Integer)
    mvariInicio_Expressao = vData
End Property

Public Property Get iInicio_Expressao() As Integer
    iInicio_Expressao = mvariInicio_Expressao
End Property

Public Property Let iCaracteres(ByVal vData As Integer)
    mvariCaracteres = vData
End Property

Public Property Get iCaracteres() As Integer
    iCaracteres = mvariCaracteres
End Property

Public Property Let sExpressao1(ByVal vData As String)
    mvarsExpressao1 = vData
End Property

Public Property Get sExpressao1() As String
    sExpressao1 = mvarsExpressao1
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sExpressao(ByVal vData As String)
    mvarsExpressao = vData
End Property

Public Property Get sExpressao() As String
    sExpressao = mvarsExpressao
End Property

Public Property Let vValor(ByVal vData As Variant)
    mvarvValor = vData
End Property

Public Property Get vValor() As Variant
    vValor = mvarvValor
End Property

Public Property Let iLinha(ByVal vData As Integer)
    mvariLinha = vData
End Property

Public Property Get iLinha() As Integer
    iLinha = mvariLinha
End Property

Public Property Let iLinhaAtual(ByVal vData As Integer)
    mvariLinhaAtual = vData
End Property

Public Property Get iLinhaAtual() As Integer
    iLinhaAtual = mvariLinhaAtual
End Property

Public Property Set colMnemonicoValor(ByVal vData As Object)
    Set mvarcolMnemonicoValor = vData
End Property

Public Property Get colMnemonicoValor() As ClassColMnemonicoValor
    Set colMnemonicoValor = mvarcolMnemonicoValor
End Property

Public Property Set objTransacao(ByVal vData As Object)
    Set mvarobjTransacao = vData
End Property

Public Property Get objTransacao() As Object
    Set objTransacao = mvarobjTransacao
End Property

Public Property Set colLancamentos(ByVal vData As Object)
    Set mvarcolLancamentos = vData
End Property

Public Property Get colLancamentos() As ClassColLancamentos
    Set colLancamentos = mvarcolLancamentos
End Property

Public Property Set colGridTransacao(ByVal vData As Object)
    Set mvarcolGridTransacao = vData
End Property

Public Property Get colGridTransacao() As ClassColGridTransacao
    Set colGridTransacao = mvarcolGridTransacao
End Property

Public Property Set colMnemonico(ByVal vData As Object)
    Set mvarcolMnemonico = vData
End Property

Public Property Get colMnemonico() As Collection
    Set colMnemonico = mvarcolMnemonico
End Property



