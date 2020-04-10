VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPRJEtapaItensProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocEtapaPRJ As Long
Private mvariSeq As Integer
Private mvarsProduto As String
Private mvarsDescricao As String
Private mvarsVersao As String
Private mvarsUM As String
Private mvardQuantidade As Double
Private mvarsObservacao As String
Private mvariTvwChecked As Integer
Private mvariIndiceTvw As Integer
Private mvarsChaveTvw As String
Private mvardPreco As Double
Private mvardCustoInfo As Double

Public Property Let dCustoInfo(ByVal vData As Double)
    mvardCustoInfo = vData
End Property

Public Property Get dCustoInfo() As Double
    dCustoInfo = mvardCustoInfo
End Property

Public Property Let dPreco(ByVal vData As Double)
    mvardPreco = vData
End Property

Public Property Get dPreco() As Double
    dPreco = mvardPreco
End Property


Public Property Let sChaveTvw(ByVal vData As String)
    mvarsChaveTvw = vData
End Property

Public Property Get sChaveTvw() As String
    sChaveTvw = mvarsChaveTvw
End Property

Public Property Let iIndiceTvw(ByVal vData As Integer)
    mvariIndiceTvw = vData
End Property

Public Property Get iIndiceTvw() As Integer
    iIndiceTvw = mvariIndiceTvw
End Property

Public Property Let iTvwChecked(ByVal vData As Integer)
    mvariTvwChecked = vData
End Property

Public Property Get iTvwChecked() As Integer
    iTvwChecked = mvariTvwChecked
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocEtapaPRJ(ByVal vData As Long)
    mvarlNumIntDocEtapaPRJ = vData
End Property

Public Property Get lNumIntDocEtapaPRJ() As Long
    lNumIntDocEtapaPRJ = mvarlNumIntDocEtapaPRJ
End Property

Public Property Let iSeq(ByVal vData As Integer)
    mvariSeq = vData
End Property

Public Property Get iSeq() As Integer
    iSeq = mvariSeq
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQuantidade(ByVal vData As Double)
    mvardQuantidade = vData
End Property

Public Property Get dQuantidade() As Double
    dQuantidade = mvardQuantidade
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property
