VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProdutoKit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public objInfoUsu As Object

Private mvarsProdutoRaiz As String
Private mvarsVersao As String
Private mvariNivel As Integer
Private mvariSeq As Integer
Private mvarsProduto As String
Private mvariSeqPai As Integer
Private mvardQuantidade As Double
Private mvarsUnidadeMed As String
Private mvariComposicao As Integer
Private mvarsProdutoDesc As String
Private mvariClasseUM As Integer
Private mvariPosicaoArvore As Integer
Private mvardPercentualPerda As Double
Private mvardCustoStandard As Double

'###########################################################
'Inserido por Wagner - Versão do Kit do Componente (Caso o componente seja um Kit)
Private mvarsVersaoKitComp As String

Public Property Let sVersaoKitComp(ByVal vData As String)
    mvarsVersaoKitComp = vData
End Property

Public Property Get sVersaoKitComp() As String
    sVersaoKitComp = mvarsVersaoKitComp
End Property
'###########################################################

Public Property Let dCustoStandard(ByVal vData As Double)
    mvardCustoStandard = vData
End Property

Public Property Get dCustoStandard() As Double
    dCustoStandard = mvardCustoStandard
End Property

Public Property Let dPercentualPerda(ByVal vData As Double)
Attribute dPercentualPerda.VB_Description = "Percentual de Perda durante a produção"
    mvardPercentualPerda = vData
End Property

Public Property Get dPercentualPerda() As Double
    dPercentualPerda = mvardPercentualPerda
End Property

Public Property Let iPosicaoArvore(ByVal vData As Integer)
    mvariPosicaoArvore = vData
End Property

Public Property Get iPosicaoArvore() As Integer
    iPosicaoArvore = mvariPosicaoArvore
End Property

Public Property Let iClasseUM(ByVal vData As Integer)
    mvariClasseUM = vData
End Property

Public Property Get iClasseUM() As Integer
    iClasseUM = mvariClasseUM
End Property

Public Property Let sProdutoDesc(ByVal vData As String)
    mvarsProdutoDesc = vData
End Property

Public Property Get sProdutoDesc() As String
    sProdutoDesc = mvarsProdutoDesc
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let iComposicao(ByVal vData As Integer)
    mvariComposicao = vData
End Property

Public Property Get iComposicao() As Integer
    iComposicao = mvariComposicao
End Property

Public Property Let sUnidadeMed(ByVal vData As String)
    mvarsUnidadeMed = vData
End Property

Public Property Get sUnidadeMed() As String
    sUnidadeMed = mvarsUnidadeMed
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let iSeqPai(ByVal vData As Integer)
    mvariSeqPai = vData
End Property

Public Property Get iSeqPai() As Integer
    iSeqPai = mvariSeqPai
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property

Public Property Let sProdutoRaiz(ByVal vData As String)
    mvarsProdutoRaiz = vData
End Property

Public Property Get sProdutoRaiz() As String
    sProdutoRaiz = mvarsProdutoRaiz
End Property



