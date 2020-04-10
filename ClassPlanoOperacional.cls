VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPlanoOperacional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"; Yes; ""
Option Explicit

'CLASSE CRIADA AUTOMATICAMENTE PELA TELA BROWSECRIA
Private mvarlNumIntDoc As Long
Private mvarlNumIntDocPOPai As Long
Private mvarlNumIntDocPMP As Long
Private mvariSeq As Integer
Private mvariNivel As Integer
Private mvarsCodOPOrigem As String
Private mvarsProduto As String
Private mvarsVersao As String
Private mvardQuantidade As Double
Private mvardQtdTotal As Double
Private mvarsUM As String
Private mvarlNumIntDocCT As Long
Private mvardtDataInicio As Date
Private mvardtDataFim As Date
Private mvarcolAlocacaoMaquinas As New Collection
Private mvarlNumIntDocOper As Long
Private mvarobjApontamento As New ClassApontamentoProducao
Private mvariFilialEmpresa As Integer
Private mvarcolOPFilhas As New Collection
Private mvarcolRCFilhas As New Collection
Private mvariAlterado As Integer
Private mvariStatus As Integer
Private mvarsDescErro As String
Private mvardTempoGasto As Double
Private mvarobjOP As New ClassOrdemDeProducao
Private mvariTemApontamento As Integer
Private mvariNumMaxMaqPorOper As Integer
Private mvariNumRepeticoes As Integer

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

Public Property Let iTemApontamento(ByVal vData As Integer)
    mvariTemApontamento = vData
End Property

Public Property Get iTemApontamento() As Integer
    iTemApontamento = mvariTemApontamento
End Property

Public Property Set objOP(ByVal vData As ClassOrdemDeProducao)
    Set mvarobjOP = vData
End Property

Public Property Get objOP() As ClassOrdemDeProducao
    Set objOP = mvarobjOP
End Property

Public Property Let dTempoGasto(ByVal vData As Double)
    mvardTempoGasto = vData
End Property

Public Property Get dTempoGasto() As Double
    dTempoGasto = mvardTempoGasto
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Let sDescErro(ByVal vData As String)
    mvarsDescErro = vData
End Property

Public Property Get sDescErro() As String
    sDescErro = mvarsDescErro
End Property

Public Property Let iAlterado(ByVal vData As Integer)
    mvariAlterado = vData
End Property

Public Property Get iAlterado() As Integer
    iAlterado = mvariAlterado
End Property

Public Property Let dQtdTotal(ByVal vData As Double)
    mvardQtdTotal = vData
End Property

Public Property Get dQtdTotal() As Double
    dQtdTotal = mvardQtdTotal
End Property

Public Property Set colOPFilhas(ByVal vData As Collection)
    Set mvarcolOPFilhas = vData
End Property

Public Property Get colOPFilhas() As Collection
    Set colOPFilhas = mvarcolOPFilhas
End Property

Public Property Set colRCFilhas(ByVal vData As Collection)
    Set mvarcolRCFilhas = vData
End Property

Public Property Get colRCFilhas() As Collection
    Set colRCFilhas = mvarcolRCFilhas
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Get objApontamento() As ClassApontamentoProducao
    Set objApontamento = mvarobjApontamento
End Property

Public Property Set objApontamento(ByVal vData As ClassApontamentoProducao)
    Set mvarobjApontamento = vData
End Property

Public Property Let lNumIntDocPOPai(ByVal vData As Long)
    mvarlNumIntDocPOPai = vData
End Property

Public Property Get lNumIntDocPOPai() As Long
    lNumIntDocPOPai = mvarlNumIntDocPOPai
End Property

Public Property Let lNumIntDocOper(ByVal vData As Long)
    mvarlNumIntDocOper = vData
End Property

Public Property Get lNumIntDocOper() As Long
    lNumIntDocOper = mvarlNumIntDocOper
End Property

Public Property Set colAlocacaoMaquinas(ByVal vData As Collection)
    Set mvarcolAlocacaoMaquinas = vData
End Property

Public Property Get colAlocacaoMaquinas() As Collection
    Set colAlocacaoMaquinas = mvarcolAlocacaoMaquinas
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lNumIntDocPMP(ByVal vData As Long)
    mvarlNumIntDocPMP = vData
End Property

Public Property Get lNumIntDocPMP() As Long
    lNumIntDocPMP = mvarlNumIntDocPMP
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

Public Property Let sCodOPOrigem(ByVal vData As String)
    mvarsCodOPOrigem = vData
End Property

Public Property Get sCodOPOrigem() As String
    sCodOPOrigem = mvarsCodOPOrigem
End Property

Public Property Let sProduto(ByVal vData As String)
    mvarsProduto = vData
End Property

Public Property Get sProduto() As String
    sProduto = mvarsProduto
End Property

Public Property Let sVersao(ByVal vData As String)
    mvarsVersao = vData
End Property

Public Property Get sVersao() As String
    sVersao = mvarsVersao
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

Public Property Let lNumIntDocCT(ByVal vData As Long)
    mvarlNumIntDocCT = vData
End Property

Public Property Get lNumIntDocCT() As Long
    lNumIntDocCT = mvarlNumIntDocCT
End Property

Public Property Let dtDataInicio(ByVal vData As Date)
    mvardtDataInicio = vData
End Property

Public Property Get dtDataInicio() As Date
    dtDataInicio = mvardtDataInicio
End Property

Public Property Let dtDataFim(ByVal vData As Date)
    mvardtDataFim = vData
End Property

Public Property Get dtDataFim() As Date
    dtDataFim = mvardtDataFim
End Property
