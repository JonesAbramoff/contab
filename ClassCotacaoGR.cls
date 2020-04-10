VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCotacaoGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarlNumIntDoc As Long
Private mvariFilialEmpresa As Integer
Private mvarlCodigo As Long
Private mvardtData As Date
Private mvariTipoOperacao As Integer 'O=Importação 1=Exportação 2=Mercado Interno
Private mvarsCliente As String
Private mvariEnvio As Integer 'O=Carta 1=E-mail 2=Fax 3=Telefone
Private mvarsEnvioComplemento As String
Private mvariCodVendedor As Integer
Private mvarsIndicacao As String
Private mvarsObservacao As String
Private mvarsObsDestOrigem As String
Private mvardtDataPrevInicio As Date
Private mvariTipoEmbalagem As Integer
Private mvariAjudantes As Integer
Private mvariCarga As Integer
Private mvariCargaPorConta As Integer
Private mvariDesCarga As Integer
Private mvariDesCargaPorConta As Integer
Private mvariOva As Integer
Private mvariOvaPorConta As Integer
Private mvariDesova As Integer
Private mvariDesovaPorConta As Integer
Private mvariCargaSolta As Integer
Private mvarsDescCargaSolta As String
Private mvariCondicaoPagto As Integer
Private mvariSituacao As Integer
Private mvariJustificativa As Integer
Private mvarsObsResultado As String
Private mvarcolCotacaoOrigemDestino As New Collection
Private mvarcolCotacaoContainer As New Collection
Private mvarcolCotacaoServico As New Collection
Private mvarcolContato As New Collection
Private mvardValorMerc As Double
Private mvarsAtendente As String
Private mvarsResponsavel As String

Public Property Let sAtendente(ByVal vData As String)
    mvarsAtendente = vData
End Property

Public Property Get sAtendente() As String
    sAtendente = mvarsAtendente
End Property

Public Property Let sResponsavel(ByVal vData As String)
    mvarsResponsavel = vData
End Property

Public Property Get sResponsavel() As String
    sResponsavel = mvarsResponsavel
End Property

Public Property Let dValorMerc(ByVal vData As Double)
    mvardValorMerc = vData
End Property

Public Property Get dValorMerc() As Double
    dValorMerc = mvardValorMerc
End Property

Public Property Set colContato(ByVal vData As Object)
    Set mvarcolContato = vData
End Property

Public Property Get colContato() As Collection
    Set colContato = mvarcolContato
End Property

Public Property Set colCotacaoServico(ByVal vData As Object)
    Set mvarcolCotacaoServico = vData
End Property

Public Property Get colCotacaoServico() As Collection
    Set colCotacaoServico = mvarcolCotacaoServico
End Property

Public Property Set colCotacaoContainer(ByVal vData As Object)
    Set mvarcolCotacaoContainer = vData
End Property

Public Property Get colCotacaoContainer() As Collection
    Set colCotacaoContainer = mvarcolCotacaoContainer
End Property

Public Property Set colCotacaoOrigemDestino(ByVal vData As Object)
    Set mvarcolCotacaoOrigemDestino = vData
End Property

Public Property Get colCotacaoOrigemDestino() As Collection
    Set colCotacaoOrigemDestino = mvarcolCotacaoOrigemDestino
End Property

Public Property Let sObsResultado(ByVal vData As String)
    mvarsObsResultado = vData
End Property

Public Property Get sObsResultado() As String
    sObsResultado = mvarsObsResultado
End Property

Public Property Let iJustificativa(ByVal vData As Integer)
    mvariJustificativa = vData
End Property

Public Property Get iJustificativa() As Integer
    iJustificativa = mvariJustificativa
End Property

Public Property Let iSituacao(ByVal vData As Integer)
    mvariSituacao = vData
End Property

Public Property Get iSituacao() As Integer
    iSituacao = mvariSituacao
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let sDescCargaSolta(ByVal vData As String)
    mvarsDescCargaSolta = vData
End Property

Public Property Get sDescCargaSolta() As String
    sDescCargaSolta = mvarsDescCargaSolta
End Property

Public Property Let iCargaSolta(ByVal vData As Integer)
    mvariCargaSolta = vData
End Property

Public Property Get iCargaSolta() As Integer
    iCargaSolta = mvariCargaSolta
End Property

Public Property Let iDesovaPorConta(ByVal vData As Integer)
    mvariDesovaPorConta = vData
End Property

Public Property Get iDesovaPorConta() As Integer
    iDesovaPorConta = mvariDesovaPorConta
End Property

Public Property Let iDesova(ByVal vData As Integer)
    mvariDesova = vData
End Property

Public Property Get iDesova() As Integer
    iDesova = mvariDesova
End Property

Public Property Let iOvaPorConta(ByVal vData As Integer)
    mvariOvaPorConta = vData
End Property

Public Property Get iOvaPorConta() As Integer
    iOvaPorConta = mvariOvaPorConta
End Property

Public Property Let iOva(ByVal vData As Integer)
    mvariOva = vData
End Property

Public Property Get iOva() As Integer
    iOva = mvariOva
End Property

Public Property Let iDesCargaPorConta(ByVal vData As Integer)
    mvariDesCargaPorConta = vData
End Property

Public Property Get iDesCargaPorConta() As Integer
    iDesCargaPorConta = mvariDesCargaPorConta
End Property

Public Property Let iDesCarga(ByVal vData As Integer)
    mvariDesCarga = vData
End Property

Public Property Get iDesCarga() As Integer
    iDesCarga = mvariDesCarga
End Property

Public Property Let iCargaPorConta(ByVal vData As Integer)
    mvariCargaPorConta = vData
End Property

Public Property Get iCargaPorConta() As Integer
    iCargaPorConta = mvariCargaPorConta
End Property

Public Property Let iCarga(ByVal vData As Integer)
    mvariCarga = vData
End Property

Public Property Get iCarga() As Integer
    iCarga = mvariCarga
End Property

Public Property Let iAjudantes(ByVal vData As Integer)
    mvariAjudantes = vData
End Property

Public Property Get iAjudantes() As Integer
    iAjudantes = mvariAjudantes
End Property

Public Property Let iTipoEmbalagem(ByVal vData As Integer)
    mvariTipoEmbalagem = vData
End Property

Public Property Get iTipoEmbalagem() As Integer
    iTipoEmbalagem = mvariTipoEmbalagem
End Property

Public Property Let dtDataPrevInicio(ByVal vData As Date)
    mvardtDataPrevInicio = vData
End Property

Public Property Get dtDataPrevInicio() As Date
    dtDataPrevInicio = mvardtDataPrevInicio
End Property

Public Property Let sObsDestOrigem(ByVal vData As String)
    mvarsObsDestOrigem = vData
End Property

Public Property Get sObsDestOrigem() As String
    sObsDestOrigem = mvarsObsDestOrigem
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let sIndicacao(ByVal vData As String)
    mvarsIndicacao = vData
End Property

Public Property Get sIndicacao() As String
    sIndicacao = mvarsIndicacao
End Property

Public Property Let iCodVendedor(ByVal vData As Integer)
    mvariCodVendedor = vData
End Property

Public Property Get iCodVendedor() As Integer
    iCodVendedor = mvariCodVendedor
End Property

Public Property Let sEnvioComplemento(ByVal vData As String)
    mvarsEnvioComplemento = vData
End Property

Public Property Get sEnvioComplemento() As String
    sEnvioComplemento = mvarsEnvioComplemento
End Property

Public Property Let iEnvio(ByVal vData As Integer)
    mvariEnvio = vData
End Property

Public Property Get iEnvio() As Integer
    iEnvio = mvariEnvio
End Property

Public Property Let sCliente(ByVal vData As String)
    mvarsCliente = vData
End Property

Public Property Get sCliente() As String
    sCliente = mvarsCliente
End Property

Public Property Let iTipoOperacao(ByVal vData As Integer)
    mvariTipoOperacao = vData
End Property

Public Property Get iTipoOperacao() As Integer
    iTipoOperacao = mvariTipoOperacao
End Property

Public Property Let dtData(ByVal vData As Date)
    mvardtData = vData
End Property

Public Property Get dtData() As Date
    dtData = mvardtData
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
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




