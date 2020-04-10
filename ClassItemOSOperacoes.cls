VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassItemOSOperacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocItemOS As Long
Private mvarlNumIntDocCompet As Long
Private mvarlNumIntDocCT As Long
Private mvariSeq As Integer
Private mvariSeqPai As Integer
Private mvariSeqArvore As Integer
Private mvariSeqRoteiro As Integer
Private mvariSeqRoteiroPai As Integer
Private mvariNivel As Integer
Private mvariNivelRoteiro As Integer
Private mvarsServico As String
Private mvarsVersao As String
Private mvarsObservacao As String
Private mvariIgnoraTaxaProducao As Integer
Private mvardQuantidade As Double
Private mvarsUMedida As String
Private mvariConsideraCarga As Integer
Private mvariOrigem As Integer
Private mvarlNumIntDocOperOrigem As Long
Private mvariNumMaxMaqPorOper As Integer
Private mvariNumRepeticoes As Integer
Private mvariPosicaoArvore As Integer
Private mvarcolPecas As New Collection
Private mvarcolMaquinas As New Collection
Private mvarcolMO As New Collection

Public Property Set colMO(ByVal vData As Collection)
    Set mvarcolMO = vData
End Property

Public Property Get colMO() As Collection
    Set colMO = mvarcolMO
End Property

Public Property Set colMaquinas(ByVal vData As Collection)
    Set mvarcolMaquinas = vData
End Property

Public Property Get colMaquinas() As Collection
    Set colMaquinas = mvarcolMaquinas
End Property

Public Property Set colPecas(ByVal vData As Collection)
    Set mvarcolPecas = vData
End Property

Public Property Get colPecas() As Collection
    Set colPecas = mvarcolPecas
End Property

Public Property Let iPosicaoArvore(ByVal vData As Integer)
    mvariPosicaoArvore = vData
End Property

Public Property Get iPosicaoArvore() As Integer
    iPosicaoArvore = mvariPosicaoArvore
End Property

Public Property Let iNumRepeticoes(ByVal vData As Integer)
    mvariNumRepeticoes = vData
End Property

Public Property Get iNumRepeticoes() As Integer
    iNumRepeticoes = mvariNumRepeticoes
End Property

Public Property Let iNumMaxMaqPorOper(ByVal vData As Integer)
    mvariNumMaxMaqPorOper = vData
End Property

Public Property Get iNumMaxMaqPorOper() As Integer
    iNumMaxMaqPorOper = mvariNumMaxMaqPorOper
End Property

Public Property Let lNumIntDocOPerOrigem(ByVal vData As Long)
    mvarlNumIntDocOperOrigem = vData
End Property

Public Property Get lNumIntDocOPerOrigem() As Long
    lNumIntDocOPerOrigem = mvarlNumIntDocOperOrigem
End Property

Public Property Let iOrigem(ByVal vData As Integer)
    mvariOrigem = vData
End Property

Public Property Get iOrigem() As Integer
    iOrigem = mvariOrigem
End Property

Public Property Let iConsideraCarga(ByVal vData As Integer)
    mvariConsideraCarga = vData
End Property

Public Property Get iConsideraCarga() As Integer
    iConsideraCarga = mvariConsideraCarga
End Property

Public Property Let iIgnoraTaxaProducao(ByVal vData As Integer)
    mvariIgnoraTaxaProducao = vData
End Property

Public Property Get iIgnoraTaxaProducao() As Integer
    iIgnoraTaxaProducao = mvariIgnoraTaxaProducao
End Property

Public Property Let iSeqRoteiro(ByVal vData As Integer)
    mvariSeqRoteiro = vData
End Property

Public Property Get iSeqRoteiro() As Integer
    iSeqRoteiro = mvariSeqRoteiro
End Property

Public Property Let iSeqRoteiroPai(ByVal vData As Integer)
    mvariSeqRoteiroPai = vData
End Property

Public Property Get iSeqRoteiroPai() As Integer
    iSeqRoteiroPai = mvariSeqRoteiroPai
End Property

Public Property Let sServico(ByVal vData As String)
    mvarsServico = vData
End Property

Public Property Get sServico() As String
    sServico = mvarsServico
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let iNivelRoteiro(ByVal vData As Integer)
    mvariNivelRoteiro = vData
End Property

Public Property Get iNivelRoteiro() As Integer
    iNivelRoteiro = mvariNivelRoteiro
End Property

Public Property Let iNivel(ByVal vData As Integer)
    mvariNivel = vData
End Property

Public Property Get iNivel() As Integer
    iNivel = mvariNivel
End Property

Public Property Let iSeqArvore(ByVal vData As Integer)
    mvariSeqArvore = vData
End Property

Public Property Get iSeqArvore() As Integer
    iSeqArvore = mvariSeqArvore
End Property

Public Property Let iSeqPai(ByVal vData As Integer)
    mvariSeqPai = vData
End Property

Public Property Get iSeqPai() As Integer
    iSeqPai = mvariSeqPai
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sUMedida(ByVal vData As String)
    mvarsUMedida = vData
End Property

Public Property Get sUMedida() As String
    sUMedida = mvarsUMedida
End Property

Public Property Let lNumIntDocCT(ByVal vData As Long)
    mvarlNumIntDocCT = vData
End Property

Public Property Get lNumIntDocCT() As Long
    lNumIntDocCT = mvarlNumIntDocCT
End Property

Public Property Let lNumIntDocCompet(ByVal vData As Long)
    mvarlNumIntDocCompet = vData
End Property

Public Property Get lNumIntDocCompet() As Long
    lNumIntDocCompet = mvarlNumIntDocCompet
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let lNumIntDocItemOS(ByVal vData As Long)
    mvarlNumIntDocItemOS = vData
End Property

Public Property Get lNumIntDocItemOS() As Long
    lNumIntDocItemOS = mvarlNumIntDocItemOS
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property





