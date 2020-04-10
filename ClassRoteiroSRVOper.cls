VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRoteiroSRVOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlNumIntDoc As Long
Private mvarlNumIntDocRotSRV As Long
Private mvariSeq As Integer
Private mvarlNumIntDocCompet As Long
Private mvarlNumIntDocCT As Long
Private mvarsObservacao As String
Private mvariSeqPai As Integer
Private mvariSeqArvore As Integer
Private mvariNivel As Integer
Private mvariPosicaoArvore As Integer
Private mvarcolMP As New Collection
Private mvarcolMO As New Collection
Private mvarcolMaq As New Collection

Public Property Set colMaq(ByVal vData As Collection)
    Set mvarcolMaq = vData
End Property

Public Property Get colMaq() As Collection
    Set colMaq = mvarcolMaq
End Property

Public Property Set colMO(ByVal vData As Collection)
    Set mvarcolMO = vData
End Property

Public Property Get colMO() As Collection
    Set colMO = mvarcolMO
End Property

Public Property Set colMP(ByVal vData As Collection)
    Set mvarcolMP = vData
End Property

Public Property Get colMP() As Collection
    Set colMP = mvarcolMP
End Property

Public Property Let iPosicaoArvore(ByVal vData As Integer)
    mvariPosicaoArvore = vData
End Property

Public Property Get iPosicaoArvore() As Integer
    iPosicaoArvore = mvariPosicaoArvore
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

Public Property Let lNumIntDocRotSRV(ByVal vData As Long)
    mvarlNumIntDocRotSRV = vData
End Property

Public Property Get lNumIntDocRotSRV() As Long
    lNumIntDocRotSRV = mvarlNumIntDocRotSRV
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property
