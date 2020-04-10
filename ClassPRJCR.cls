VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocPRJ As Long
Private mvarlNumIntDocEtapa As Long
Private mvariTipoOrigem As Integer
Private mvarlNumIntDocOrigem As Long
Private mvardPercent As Double
Private mvardQuantidade As Double
Private mvardValor As Double
Private mvariTipoValor As Integer
Private mvarsObservacao As String
Private mvariCalcAuto As Integer
Private mvarsItem As String
Private mvarsProjeto As String
Private mvarlNumIntDocPRJInfo As Long
Private mvariUsado As Integer
Private mvardQtdUsado As Double

Public Property Let dQtdUsado(ByVal vData As Double)
    mvardQtdUsado = vData
End Property

Public Property Get dQtdUsado() As Double
    dQtdUsado = mvardQtdUsado
End Property

Public Property Let iUsado(ByVal vData As Integer)
    mvariUsado = vData
End Property

Public Property Get iUsado() As Integer
    iUsado = mvariUsado
End Property

Public Property Let lNumIntDocPRJInfo(ByVal vData As Long)
    mvarlNumIntDocPRJInfo = vData
End Property

Public Property Get lNumIntDocPRJInfo() As Long
    lNumIntDocPRJInfo = mvarlNumIntDocPRJInfo
End Property

Public Property Let sProjeto(ByVal vData As String)
    mvarsProjeto = vData
End Property

Public Property Get sProjeto() As String
    sProjeto = mvarsProjeto
End Property

Public Property Let sItem(ByVal vData As String)
    mvarsItem = vData
End Property

Public Property Get sItem() As String
    sItem = mvarsItem
End Property

Public Property Let iCalcAuto(ByVal vData As Integer)
    mvariCalcAuto = vData
End Property

Public Property Get iCalcAuto() As Integer
    iCalcAuto = mvariCalcAuto
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let dPercent(ByVal vData As Double)
    mvardPercent = vData
End Property

Public Property Get dPercent() As Double
    dPercent = mvardPercent
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocPRJ(ByVal vData As Long)
    mvarlNumIntDocPRJ = vData
End Property

Public Property Get lNumIntDocPRJ() As Long
    lNumIntDocPRJ = mvarlNumIntDocPRJ
End Property

Public Property Let lNumIntDocEtapa(ByVal vData As Long)
    mvarlNumIntDocEtapa = vData
End Property

Public Property Get lNumIntDocEtapa() As Long
    lNumIntDocEtapa = mvarlNumIntDocEtapa
End Property

Public Property Let iTipoOrigem(ByVal vData As Integer)
    mvariTipoOrigem = vData
End Property

Public Property Get iTipoOrigem() As Integer
    iTipoOrigem = mvariTipoOrigem
End Property

Public Property Let lNumIntDocOrigem(ByVal vData As Long)
    mvarlNumIntDocOrigem = vData
End Property

Public Property Get lNumIntDocOrigem() As Long
    lNumIntDocOrigem = mvarlNumIntDocOrigem
End Property

Public Property Let dValor(ByVal vData As Double)
    mvardValor = vData
End Property

Public Property Get dValor() As Double
    dValor = mvardValor
End Property

Public Property Let iTipoValor(ByVal vData As Integer)
    mvariTipoValor = vData
End Property

Public Property Get iTipoValor() As Integer
    iTipoValor = mvariTipoValor
End Property
