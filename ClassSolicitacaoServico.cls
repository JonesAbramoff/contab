VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassSolicitacaoServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlNumero As Long
Private mvarlCodTabPreco As Long
Private mvarlCliente As Long
Private mvardtDataEmissao As Date
Private mvarsNumReferencia As String
Private mvardtDataPedido As Date
Private mvariTipoOperacao As Integer 'O=Importação 1=Exportação 2=Mercado Interno
Private mvariDespachante As Integer 'O=Importação 1=Exportação 2=Mercado Interno
Private mvarsMaterial As String
Private mvardQuantMaterial As Double
Private mvarsUM As String
Private mvardValorMercadoria As Double
Private mvariTipoEmbalagem As Integer
Private mvariTipoContainer As Integer
Private mvarlCodProgNavio As Long
Private mvarsBooking As String
Private mvardtDataPrevInicio As Date
Private mvardtHoraPrevInicio As Date
Private mvardtDataPrevFim As Date
Private mvardtHoraPrevFim As Date
Private mvarsObservacao As String
Private mvarlEnderecoOrigem As Long
Private mvarlEnderecoDestino As Long
Private mvarcolServico As New Collection
Private mvarsClienteNomeRed As String
Private mvarsDespachanteNomeRed As String
Private mvarsTipoContDes As String
Private mvarsTipoEmbDes As String
Private mvarobjEnderecoOrigem As ClassEndereco
Private mvarobjEnderecoDestino As ClassEndereco
Private mvarsPorto As String

Public Property Let sPorto(ByVal vData As String)
    mvarsPorto = vData
End Property

Public Property Get sPorto() As String
    sPorto = mvarsPorto
End Property

Public Property Let sTipoContDes(ByVal vData As String)
    mvarsTipoContDes = vData
End Property

Public Property Get sTipoContDes() As String
    sTipoContDes = mvarsTipoContDes
End Property

Public Property Let sTipoEmbDes(ByVal vData As String)
    mvarsTipoEmbDes = vData
End Property

Public Property Get sTipoEmbDes() As String
    sTipoEmbDes = mvarsTipoEmbDes
End Property

Public Property Set objEnderecoDestino(ByVal vData As Object)
    Set mvarobjEnderecoDestino = vData
End Property

Public Property Get objEnderecoDestino() As ClassEndereco
    Set objEnderecoDestino = mvarobjEnderecoDestino
End Property

Public Property Set objEnderecoOrigem(ByVal vData As Object)
    Set mvarobjEnderecoOrigem = vData
End Property

Public Property Get objEnderecoOrigem() As ClassEndereco
    Set objEnderecoOrigem = mvarobjEnderecoOrigem
End Property

Public Property Let sDespachanteNomeRed(ByVal vData As String)
    mvarsDespachanteNomeRed = vData
End Property

Public Property Get sDespachanteNomeRed() As String
    sDespachanteNomeRed = mvarsDespachanteNomeRed
End Property


Public Property Let sClienteNomeRed(ByVal vData As String)
    mvarsClienteNomeRed = vData
End Property

Public Property Get sClienteNomeRed() As String
    sClienteNomeRed = mvarsClienteNomeRed
End Property

Public Property Set colServico(ByVal vData As Object)
    Set mvarcolServico = vData
End Property

Public Property Get colServico() As Collection
    Set colServico = mvarcolServico
End Property

Public Property Let lEnderecoDestino(ByVal vData As Long)
    mvarlEnderecoDestino = vData
End Property

Public Property Get lEnderecoDestino() As Long
    lEnderecoDestino = mvarlEnderecoDestino
End Property

Public Property Let lEnderecoOrigem(ByVal vData As Long)
    mvarlEnderecoOrigem = vData
End Property

Public Property Get lEnderecoOrigem() As Long
    lEnderecoOrigem = mvarlEnderecoOrigem
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let dtHoraPrevFim(ByVal vData As Date)
    mvardtHoraPrevFim = vData
End Property

Public Property Get dtHoraPrevFim() As Date
    dtHoraPrevFim = mvardtHoraPrevFim
End Property

Public Property Let dtDataPrevFim(ByVal vData As Date)
    mvardtDataPrevFim = vData
End Property

Public Property Get dtDataPrevFim() As Date
    dtDataPrevFim = mvardtDataPrevFim
End Property

Public Property Let dtHoraPrevInicio(ByVal vData As Date)
    mvardtHoraPrevInicio = vData
End Property

Public Property Get dtHoraPrevInicio() As Date
    dtHoraPrevInicio = mvardtHoraPrevInicio
End Property

Public Property Let dtDataPrevInicio(ByVal vData As Date)
    mvardtDataPrevInicio = vData
End Property

Public Property Get dtDataPrevInicio() As Date
    dtDataPrevInicio = mvardtDataPrevInicio
End Property

Public Property Let sBooking(ByVal vData As String)
    mvarsBooking = vData
End Property

Public Property Get sBooking() As String
    sBooking = mvarsBooking
End Property

Public Property Let lCodProgNavio(ByVal vData As Long)
    mvarlCodProgNavio = vData
End Property

Public Property Get lCodProgNavio() As Long
    lCodProgNavio = mvarlCodProgNavio
End Property

Public Property Let iTipoContainer(ByVal vData As Integer)
    mvariTipoContainer = vData
End Property

Public Property Get iTipoContainer() As Integer
    iTipoContainer = mvariTipoContainer
End Property

Public Property Let iTipoEmbalagem(ByVal vData As Integer)
    mvariTipoEmbalagem = vData
End Property

Public Property Get iTipoEmbalagem() As Integer
    iTipoEmbalagem = mvariTipoEmbalagem
End Property

Public Property Let dValorMercadoria(ByVal vData As Double)
    mvardValorMercadoria = vData
End Property

Public Property Get dValorMercadoria() As Double
    dValorMercadoria = mvardValorMercadoria
End Property

Public Property Let sUM(ByVal vData As String)
    mvarsUM = vData
End Property

Public Property Get sUM() As String
    sUM = mvarsUM
End Property

Public Property Let dQuantMaterial(ByVal vData As Double)
    mvardQuantMaterial = vData
End Property

Public Property Get dQuantMaterial() As Double
    dQuantMaterial = mvardQuantMaterial
End Property

Public Property Let sMaterial(ByVal vData As String)
    mvarsMaterial = vData
End Property

Public Property Get sMaterial() As String
    sMaterial = mvarsMaterial
End Property

Public Property Let iDespachante(ByVal vData As Integer)
    mvariDespachante = vData
End Property

Public Property Get iDespachante() As Integer
    iDespachante = mvariDespachante
End Property

Public Property Let iTipoOperacao(ByVal vData As Integer)
    mvariTipoOperacao = vData
End Property

Public Property Get iTipoOperacao() As Integer
    iTipoOperacao = mvariTipoOperacao
End Property

Public Property Let dtDataPedido(ByVal vData As Date)
    mvardtDataPedido = vData
End Property

Public Property Get dtDataPedido() As Date
    dtDataPedido = mvardtDataPedido
End Property

Public Property Let sNumReferencia(ByVal vData As String)
    mvarsNumReferencia = vData
End Property

Public Property Get sNumReferencia() As String
    sNumReferencia = mvarsNumReferencia
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let lCodTabPreco(ByVal vData As Long)
    mvarlCodTabPreco = vData
End Property

Public Property Get lCodTabPreco() As Long
    lCodTabPreco = mvarlCodTabPreco
End Property

Public Property Let lNumero(ByVal vData As Long)
    mvarlNumero = vData
End Property

Public Property Get lNumero() As Long
    lNumero = mvarlNumero
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property





