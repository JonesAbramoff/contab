VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTEF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvariTipoParcelamento As Integer
Private mvariTransacao As Integer
Private mvarsRede As String
Private mvarsDescricao As String
Private mvariTipo As Integer
Private mvariTipoMov As Integer
Private mvariNumParcelas As Integer
Private mvarsNSU As String
Private mvarsFinalizacao As String
Private mvarsAdm As String
Private mvariTipoTransacao As Integer
Private mvarsOperador As String
Private mvariTipoMeioPagto As Integer
Private mvariRede As Integer
Private mvarcolViaUnica As New Collection
Private mvarsIdentificacao As String
Private mvariOperacao As Integer
Private mvariTipoCartao As Integer
Private mvariTipoFinanciamento As Integer
Private mvarsAutorizacao As String
Private mvardtDataComprovante As Date
Private mvardtHoraComprovante As Date
Private mvardtDataPreDatado As Date
Private mvarsNSUOriginal As String
Private mvarsDataHoraOriginal As String
Private mvariTamViaUnica As Integer
Private mvardValorTotal As Double
Private mvardValorOriginal As Double
Private mvardValorTroco As Double
Private mvardValorDesconto As Double
Private mvariTamReduzido As Integer
Private mvarcolReduzido As New Collection
Private mvariTamViaCliente As Integer
Private mvarcolViaCliente As New Collection
Private mvariTamViaEstab As Integer
Private mvarcolViaEstab As New Collection
Private mvarsCodEstab As String
Private mvarsNumTerminal As String
Private mvariStatusConfirmacao As Integer
Private mvariViasComprovante As Integer
Private mvarsComando As String

Public Property Let sComando(ByVal vData As String)
    mvarsComando = vData
End Property

Public Property Get sComando() As String
    sComando = mvarsComando
End Property

Public Property Let iViasComprovante(ByVal vData As Integer)
    mvariViasComprovante = vData
End Property

Public Property Get iViasComprovante() As Integer
    iViasComprovante = mvariViasComprovante
End Property

Public Property Let iStatusConfirmacao(ByVal vData As Integer)
    mvariStatusConfirmacao = vData
End Property

Public Property Get iStatusConfirmacao() As Integer
    iStatusConfirmacao = mvariStatusConfirmacao
End Property

Public Property Let sNumTerminal(ByVal vData As String)
    mvarsNumTerminal = vData
End Property

Public Property Get sNumTerminal() As String
    sNumTerminal = mvarsNumTerminal
End Property

Public Property Let sCodEstab(ByVal vData As String)
    mvarsCodEstab = vData
End Property

Public Property Get sCodEstab() As String
    sCodEstab = mvarsCodEstab
End Property

Public Property Set colViaEstab(ByVal vData As Collection)
    Set mvarcolViaEstab = vData
End Property

Public Property Get colViaEstab() As Collection
    Set colViaEstab = mvarcolViaEstab
End Property

Public Property Let iTamViaEstab(ByVal vData As Integer)
    mvariTamViaEstab = vData
End Property

Public Property Get iTamViaEstab() As Integer
    iTamViaEstab = mvariTamViaEstab
End Property

Public Property Set colViaCliente(ByVal vData As Collection)
    Set mvarcolViaCliente = vData
End Property

Public Property Get colViaCliente() As Collection
    Set colViaCliente = mvarcolViaCliente
End Property

Public Property Let iTamViaCliente(ByVal vData As Integer)
    mvariTamViaCliente = vData
End Property

Public Property Get iTamViaCliente() As Integer
    iTamViaCliente = mvariTamViaCliente
End Property

Public Property Set colReduzido(ByVal vData As Collection)
    Set mvarcolReduzido = vData
End Property

Public Property Get colReduzido() As Collection
    Set colReduzido = mvarcolReduzido
End Property

Public Property Let iTamReduzido(ByVal vData As Integer)
    mvariTamReduzido = vData
End Property

Public Property Get iTamReduzido() As Integer
    iTamReduzido = mvariTamReduzido
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorTroco(ByVal vData As Double)
    mvardValorTroco = vData
End Property

Public Property Get dValorTroco() As Double
    dValorTroco = mvardValorTroco
End Property

Public Property Let dValorOriginal(ByVal vData As Double)
    mvardValorOriginal = vData
End Property

Public Property Get dValorOriginal() As Double
    dValorOriginal = mvardValorOriginal
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let iTamViaUnica(ByVal vData As Integer)
    mvariTamViaUnica = vData
End Property

Public Property Get iTamViaUnica() As Integer
    iTamViaUnica = mvariTamViaUnica
End Property

Public Property Let sDataHoraOriginal(ByVal vData As String)
    mvarsDataHoraOriginal = vData
End Property

Public Property Get sDataHoraOriginal() As String
    sDataHoraOriginal = mvarsDataHoraOriginal
End Property

Public Property Let sNSUOriginal(ByVal vData As String)
    mvarsNSUOriginal = vData
End Property

Public Property Get sNSUOriginal() As String
    sNSUOriginal = mvarsNSUOriginal
End Property

Public Property Let dtDataPreDatado(ByVal vData As Date)
    mvardtDataPreDatado = vData
End Property

Public Property Get dtDataPreDatado() As Date
    dtDataPreDatado = mvardtDataPreDatado
End Property

Public Property Let dtHoraComprovante(ByVal vData As Date)
    mvardtHoraComprovante = vData
End Property

Public Property Get dtHoraComprovante() As Date
    dtHoraComprovante = mvardtHoraComprovante
End Property

Public Property Let dtDataComprovante(ByVal vData As Date)
    mvardtDataComprovante = vData
End Property

Public Property Get dtDataComprovante() As Date
    dtDataComprovante = mvardtDataComprovante
End Property

Public Property Let sAutorizacao(ByVal vData As String)
    mvarsAutorizacao = vData
End Property

Public Property Get sAutorizacao() As String
    sAutorizacao = mvarsAutorizacao
End Property

Public Property Let iTipoFinanciamento(ByVal vData As Integer)
    mvariTipoFinanciamento = vData
End Property

Public Property Get iTipoFinanciamento() As Integer
    iTipoFinanciamento = mvariTipoFinanciamento
End Property

Public Property Let iTipoCartao(ByVal vData As Integer)
    mvariTipoCartao = vData
End Property

Public Property Get iTipoCartao() As Integer
    iTipoCartao = mvariTipoCartao
End Property

Public Property Let iOperacao(ByVal vData As Integer)
    mvariOperacao = vData
End Property

Public Property Get iOperacao() As Integer
    iOperacao = mvariOperacao
End Property

Public Property Let sIdentificacao(ByVal vData As String)
    mvarsIdentificacao = vData
End Property

Public Property Get sIdentificacao() As String
    sIdentificacao = mvarsIdentificacao
End Property

Public Property Set colViaUnica(ByVal vData As Collection)
    Set mvarcolViaUnica = vData
End Property

Public Property Get colViaUnica() As Collection
    Set colViaUnica = mvarcolViaUnica
End Property

Public Property Let iRede(ByVal vData As Integer)
    mvariRede = vData
End Property

Public Property Get iRede() As Integer
    iRede = mvariRede
End Property

Public Property Let iTipoMeioPagto(ByVal vData As Integer)
    mvariTipoMeioPagto = vData
End Property

Public Property Get iTipoMeioPagto() As Integer
    iTipoMeioPagto = mvariTipoMeioPagto
End Property

Public Property Let sOperador(ByVal vData As String)
    mvarsOperador = vData
End Property

Public Property Get sOperador() As String
    sOperador = mvarsOperador
End Property

Public Property Let iTipoTransacao(ByVal vData As Integer)
    mvariTipoTransacao = vData
End Property

Public Property Get iTipoTransacao() As Integer
    iTipoTransacao = mvariTipoTransacao
End Property

Public Property Let sAdm(ByVal vData As String)
    mvarsAdm = vData
End Property

Public Property Get sAdm() As String
    sAdm = mvarsAdm
End Property

Public Property Let sFinalizacao(ByVal vData As String)
    mvarsFinalizacao = vData
End Property

Public Property Get sFinalizacao() As String
    sFinalizacao = mvarsFinalizacao
End Property

Public Property Let sNSU(ByVal vData As String)
    mvarsNSU = vData
End Property

Public Property Get sNSU() As String
    sNSU = mvarsNSU
End Property

Public Property Let iNumParcelas(ByVal vData As Integer)
    mvariNumParcelas = vData
End Property

Public Property Get iNumParcelas() As Integer
    iNumParcelas = mvariNumParcelas
End Property

Public Property Let iTipoMov(ByVal vData As Integer)
    mvariTipoMov = vData
End Property

Public Property Get iTipoMov() As Integer
    iTipoMov = mvariTipoMov
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sDescricao(ByVal vData As String)
    mvarsDescricao = vData
End Property

Public Property Get sDescricao() As String
    sDescricao = mvarsDescricao
End Property

Public Property Let sRede(ByVal vData As String)
    mvarsRede = vData
End Property

Public Property Get sRede() As String
    sRede = mvarsRede
End Property

Public Property Let iTransacao(ByVal vData As Integer)
    mvariTransacao = vData
End Property

Public Property Get iTransacao() As Integer
    iTransacao = mvariTransacao
End Property

Public Property Let iTipoParcelamento(ByVal vData As Integer)
    mvariTipoParcelamento = vData
End Property

Public Property Get iTipoParcelamento() As Integer
    iTipoParcelamento = mvariTipoParcelamento
End Property

Private Sub Class_Initialize()
    mvardtDataPreDatado = DATA_NULA
End Sub
