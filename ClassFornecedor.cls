VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsRazaoSocial As String
Private mvarsNomeReduzido As String
Private mvariTipo As Integer
Private mvarsObservacao As String
Private mvardDesconto As Double
Private mvarlNumeroCompras As Long
Private mvardMediaCompra As Double
Private mvardtDataPrimeiraCompra As Date
Private mvardtDataUltimaCompra As Date
Private mvarlMediaAtraso As Long
Private mvarlMaiorAtraso As Long
Private mvardSaldoTitulos As Double
Private mvardValorAcumuladoCompras As Double
Private mvardSaldoDuplicatas As Double
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsInscricaoMunicipal As String
Private mvarsObservacao2 As String
Private mvarsContaContabil As String
Private mvarsContaFornConsig As String
Private mvarlEndereco As Long
Private mvariBanco As Integer
Private mvarsAgencia As String
Private mvarsContaCorrente As String
Private mvardtDataUltDevolucao As Date
Private mvarlNumTotalDevolucoes As Long
Private mvariProxCodFilial As Integer
Private mvardValorAcumuladoDevolucoes As Double
Private mvariCondicaoPagto As Integer
Private mvarcolCategoriaItem As New Collection

'Alteracao Daniel
Private mvarsInscricaoINSS As String
Private mvariTipoCobranca As Integer
Private mvarsContaDespesa As String
Private mvariGeraCredICMS As Integer
Private mvariTipoFrete As Integer
Private mvariAtivo As Integer
Private mvarsInscricaoSuframa As String
Private mvariRegimeTributario As Integer
Private mvariIEIsento As Integer
Private mvariIENaoContrib As Integer

Private mvarsNatureza As String

'nfe 3.10
Public sIdEstrangeiro As String

Public Property Let sNatureza(ByVal vData As String)
    mvarsNatureza = vData
End Property

Public Property Get sNatureza() As String
    sNatureza = mvarsNatureza
End Property

Public Property Let iIENaoContrib(ByVal vData As Integer)
    mvariIENaoContrib = vData
End Property

Public Property Get iIENaoContrib() As Integer
    If mvariIENaoContrib = -1 Then
        iIENaoContrib = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIENaoContrib = mvariIENaoContrib
    End If
End Property

Public Property Let iIEIsento(ByVal vData As Integer)
    mvariIEIsento = vData
End Property

Public Property Get iIEIsento() As Integer
    If mvariIEIsento = -1 Then
        iIEIsento = IIf(Len(Trim(mvarsInscricaoEstadual)) = 0, MARCADO, DESMARCADO)
    Else
        iIEIsento = mvariIEIsento
    End If
End Property

Public Property Let iRegimeTributario(ByVal vData As Integer)
    mvariRegimeTributario = vData
End Property

Public Property Get iRegimeTributario() As Integer
    iRegimeTributario = mvariRegimeTributario
End Property

Public Property Let sInscricaoSuframa(ByVal vData As String)
    mvarsInscricaoSuframa = vData
End Property

Public Property Get sInscricaoSuframa() As String
    sInscricaoSuframa = mvarsInscricaoSuframa
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Set colCategoriaItem(ByVal vData As Object)
    Set mvarcolCategoriaItem = vData
End Property

Public Property Get colCategoriaItem() As Collection
    Set colCategoriaItem = mvarcolCategoriaItem
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let dValorAcumuladoDevolucoes(ByVal vData As Double)
    mvardValorAcumuladoDevolucoes = vData
End Property

Public Property Get dValorAcumuladoDevolucoes() As Double
    dValorAcumuladoDevolucoes = mvardValorAcumuladoDevolucoes
End Property

Public Property Let iProxCodFilial(ByVal vData As Integer)
    mvariProxCodFilial = vData
End Property

Public Property Get iProxCodFilial() As Integer
    iProxCodFilial = mvariProxCodFilial
End Property

Public Property Let lNumTotalDevolucoes(ByVal vData As Long)
    mvarlNumTotalDevolucoes = vData
End Property

Public Property Get lNumTotalDevolucoes() As Long
    lNumTotalDevolucoes = mvarlNumTotalDevolucoes
End Property

Public Property Let dtDataUltDevolucao(ByVal vData As Date)
    mvardtDataUltDevolucao = vData
End Property

Public Property Get dtDataUltDevolucao() As Date
    dtDataUltDevolucao = mvardtDataUltDevolucao
End Property

Public Property Let sContaCorrente(ByVal vData As String)
    mvarsContaCorrente = vData
End Property

Public Property Get sContaCorrente() As String
    sContaCorrente = mvarsContaCorrente
End Property

Public Property Let sAgencia(ByVal vData As String)
    mvarsAgencia = vData
End Property

Public Property Get sAgencia() As String
    sAgencia = mvarsAgencia
End Property

Public Property Let iBanco(ByVal vData As Integer)
    mvariBanco = vData
End Property

Public Property Get iBanco() As Integer
    iBanco = mvariBanco
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let sContaFornConsig(ByVal vData As String)
    mvarsContaFornConsig = vData
End Property

Public Property Get sContaFornConsig() As String
    sContaFornConsig = mvarsContaFornConsig
End Property

Public Property Let sObservacao2(ByVal vData As String)
    mvarsObservacao2 = vData
End Property

Public Property Get sObservacao2() As String
    sObservacao2 = mvarsObservacao2
End Property

Public Property Let sInscricaoMunicipal(ByVal vData As String)
    mvarsInscricaoMunicipal = vData
End Property

Public Property Get sInscricaoMunicipal() As String
    sInscricaoMunicipal = mvarsInscricaoMunicipal
End Property

Public Property Let sInscricaoEstadual(ByVal vData As String)
    mvarsInscricaoEstadual = vData
End Property

Public Property Get sInscricaoEstadual() As String
    sInscricaoEstadual = mvarsInscricaoEstadual
End Property

Public Property Let sCgc(ByVal vData As String)
    mvarsCGC = vData
End Property

Public Property Get sCgc() As String
    sCgc = mvarsCGC
End Property

Public Property Let dSaldoDuplicatas(ByVal vData As Double)
    mvardSaldoDuplicatas = vData
End Property

Public Property Get dSaldoDuplicatas() As Double
    dSaldoDuplicatas = mvardSaldoDuplicatas
End Property

Public Property Let dValorAcumuladoCompras(ByVal vData As Double)
    mvardValorAcumuladoCompras = vData
End Property

Public Property Get dValorAcumuladoCompras() As Double
    dValorAcumuladoCompras = mvardValorAcumuladoCompras
End Property

Public Property Let dSaldoTitulos(ByVal vData As Double)
    mvardSaldoTitulos = vData
End Property

Public Property Get dSaldoTitulos() As Double
    dSaldoTitulos = mvardSaldoTitulos
End Property

Public Property Let lMaiorAtraso(ByVal vData As Long)
    mvarlMaiorAtraso = vData
End Property

Public Property Get lMaiorAtraso() As Long
    lMaiorAtraso = mvarlMaiorAtraso
End Property

Public Property Let lMediaAtraso(ByVal vData As Long)
    mvarlMediaAtraso = vData
End Property

Public Property Get lMediaAtraso() As Long
    lMediaAtraso = mvarlMediaAtraso
End Property

Public Property Let dtDataUltimaCompra(ByVal vData As Date)
    mvardtDataUltimaCompra = vData
End Property

Public Property Get dtDataUltimaCompra() As Date
    dtDataUltimaCompra = mvardtDataUltimaCompra
End Property

Public Property Let dtDataPrimeiraCompra(ByVal vData As Date)
    mvardtDataPrimeiraCompra = vData
End Property

Public Property Get dtDataPrimeiraCompra() As Date
    dtDataPrimeiraCompra = mvardtDataPrimeiraCompra
End Property

Public Property Let dMediaCompra(ByVal vData As Double)
    mvardMediaCompra = vData
End Property

Public Property Get dMediaCompra() As Double
    dMediaCompra = mvardMediaCompra
End Property

Public Property Let lNumeroCompras(ByVal vData As Long)
    mvarlNumeroCompras = vData
End Property

Public Property Get lNumeroCompras() As Long
    lNumeroCompras = mvarlNumeroCompras
End Property

Public Property Let dDesconto(ByVal vData As Double)
    mvardDesconto = vData
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let sObservacao(ByVal vData As String)
    mvarsObservacao = vData
End Property

Public Property Get sObservacao() As String
    sObservacao = mvarsObservacao
End Property

Public Property Let iTipo(ByVal vData As Integer)
    mvariTipo = vData
End Property

Public Property Get iTipo() As Integer
    iTipo = mvariTipo
End Property

Public Property Let sNomeReduzido(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeReduzido() As String
    sNomeReduzido = mvarsNomeReduzido
End Property

Public Property Let sRazaoSocial(ByVal vData As String)
    mvarsRazaoSocial = vData
End Property

Public Property Get sRazaoSocial() As String
    sRazaoSocial = mvarsRazaoSocial
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

'Alteracao Daniel
Public Property Let sInscricaoINSS(ByVal vData As String)
    mvarsInscricaoINSS = vData
End Property

Public Property Get sInscricaoINSS() As String
    sInscricaoINSS = mvarsInscricaoINSS
End Property

Public Property Let sContaDespesa(ByVal vData As String)
    mvarsContaDespesa = vData
End Property

Public Property Get sContaDespesa() As String
    sContaDespesa = mvarsContaDespesa
End Property

Public Property Let iTipoCobranca(ByVal vData As Integer)
    mvariTipoCobranca = vData
End Property

Public Property Get iTipoCobranca() As Integer
    iTipoCobranca = mvariTipoCobranca
End Property

Public Property Let iGeraCredICMS(ByVal vData As Integer)
    mvariGeraCredICMS = vData
End Property

Public Property Get iGeraCredICMS() As Integer
    iGeraCredICMS = mvariGeraCredICMS
End Property

Public Property Let iTipoFrete(ByVal vData As Integer)
    mvariTipoFrete = vData
End Property

Public Property Get iTipoFrete() As Integer
    iTipoFrete = mvariTipoFrete
End Property

Private Sub Class_Initialize()
    mvariIEIsento = -1
    mvariIENaoContrib = -1
End Sub
