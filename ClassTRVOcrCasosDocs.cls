VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTRVOcrCasosDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDocOcrCaso As Long
Private mvarlCodigoDoc As Long
Private mvariSeq As Integer
Private mvarsDescricao As String
Private mvarsObservacao As String
Private mvariRecebido As Integer
Private mvariNecessUsu As Integer
Private mvariNecessSist As Integer

Public Property Let iNecessSist(ByVal vData As Integer)
    mvariNecessSist = vData
End Property

Public Property Get iNecessSist() As Integer
    iNecessSist = mvariNecessSist
End Property

Public Property Let iNecessUsu(ByVal vData As Integer)
    mvariNecessUsu = vData
End Property

Public Property Get iNecessUsu() As Integer
    iNecessUsu = mvariNecessUsu
End Property

Public Property Let iRecebido(ByVal vData As Integer)
    mvariRecebido = vData
End Property

Public Property Get iRecebido() As Integer
    iRecebido = mvariRecebido
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let lNumIntDocOcrCaso(ByVal vData As Long)
    mvarlNumIntDocOcrCaso = vData
End Property

Public Property Get lNumIntDocOcrCaso() As Long
    lNumIntDocOcrCaso = mvarlNumIntDocOcrCaso
End Property

Public Property Let lCodigoDoc(ByVal vData As Long)
    mvarlCodigoDoc = vData
End Property

Public Property Get lCodigoDoc() As Long
    lCodigoDoc = mvarlCodigoDoc
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

