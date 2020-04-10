VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSimulacaoESTItens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvardtData As Date
Private mvardQuantidade As Double
Private mvarlNumIntRel As Long
Private mvariSeqRel As Integer
Private mvariSeq As Integer
Private mvariTipo As Integer
Private mvariOrdem As Integer
Private mvarsProduto As String
Private mvarsUM As String
Private mvariSimulacao As Integer
Private mvariTipoOrdem As Integer

Public Property Let iTipoOrdem(ByVal vData As Integer)
    mvariTipoOrdem = vData
End Property

Public Property Get iTipoOrdem() As Integer
    iTipoOrdem = mvariTipoOrdem
End Property

Public Property Let iSimulacao(ByVal vData As Integer)
    mvariSimulacao = vData
End Property

Public Property Get iSimulacao() As Integer
    iSimulacao = mvariSimulacao
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let iOrdem(ByVal vData As Integer)
    mvariOrdem = vData
End Property

Public Property Get iOrdem() As Integer
    iOrdem = mvariOrdem
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let iSeqRel(ByVal vData As Integer)
    mvariSeqRel = vData
End Property

Public Property Get iSeqRel() As Integer
    iSeqRel = mvariSeqRel
End Property

Public Property Let lNumIntRel(ByVal vData As Long)
    mvarlNumIntRel = vData
End Property

Public Property Get lNumIntRel() As Long
    lNumIntRel = mvarlNumIntRel
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property



