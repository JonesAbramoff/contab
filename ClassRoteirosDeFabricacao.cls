VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRoteirosDeFabricacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarsProdutoRaiz As String
Private mvarsVersao As String
Private mvarsDescricao As String
Private mvardtDataCriacao As Date
Private mvardtDataUltModificacao As Date
Private mvardQuantidade As Double
Private mvarsUM As String
Private mvarsAutor As String
Private mvariComposicao As Integer
Private mvardPercentualPerda As Double
Private mvardCustoStandard As Double
Private mvarcolOperacoes As New Collection
Private mvariVersaoFormPreco As Integer
Private mvariNumMaxMaqPorOper As Integer

Public Property Let iNumMaxMaqPorOper(ByVal vData As Integer)
    mvariNumMaxMaqPorOper = vData
End Property

Public Property Get iNumMaxMaqPorOper() As Integer
    iNumMaxMaqPorOper = mvariNumMaxMaqPorOper
End Property

Public Property Let iVersaoFormPreco(ByVal vData As Integer)
    mvariVersaoFormPreco = vData
End Property

Public Property Get iVersaoFormPreco() As Integer
    iVersaoFormPreco = mvariVersaoFormPreco
End Property

Public Property Set colOperacoes(ByVal vData As Collection)
    Set mvarcolOperacoes = vData
End Property

Public Property Get colOperacoes() As Collection
    Set colOperacoes = mvarcolOperacoes
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let sProdutoRaiz(ByVal vData As String)
    mvarsProdutoRaiz = vData
End Property

Public Property Get sProdutoRaiz() As String
    sProdutoRaiz = mvarsProdutoRaiz
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let dtDataCriacao(ByVal vData As Date)
    mvardtDataCriacao = vData
End Property

Public Property Get dtDataCriacao() As Date
    dtDataCriacao = mvardtDataCriacao
End Property

Public Property Let dtDataUltModificacao(ByVal vData As Date)
    mvardtDataUltModificacao = vData
End Property

Public Property Get dtDataUltModificacao() As Date
    dtDataUltModificacao = mvardtDataUltModificacao
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sAutor(ByVal vData As String)
    mvarsAutor = vData
End Property

Public Property Get sAutor() As String
    sAutor = mvarsAutor
End Property

Public Property Let iComposicao(ByVal vData As Integer)
    mvariComposicao = vData
End Property

Public Property Get iComposicao() As Integer
    iComposicao = mvariComposicao
End Property

Public Property Let dPercentualPerda(ByVal vData As Double)
    mvardPercentualPerda = vData
End Property

Public Property Get dPercentualPerda() As Double
    dPercentualPerda = mvardPercentualPerda
End Property

Public Property Let dCustoStandard(ByVal vData As Double)
    mvardCustoStandard = vData
End Property

Public Property Get dCustoStandard() As Double
    dCustoStandard = mvardCustoStandard
End Property

