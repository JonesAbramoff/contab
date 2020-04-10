VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarlCodigo As Long
Private mvarsRazaoSocial As String
Private mvarsNomeReduzido As String
Private mvariTipo As Integer
Private mvarsObservacao As String
Private mvardLimiteCredito As Double
Private mvardDesconto As Double
Private mvariCodMensagem As Integer
Private mvariTabelaPreco As Integer
Private mvarlNumeroCompras As Long
Private mvardMediaCompra As Double
Private mvardtDataPrimeiraCompra As Date
Private mvardtDataUltimaCompra As Date
Private mvarlMediaAtraso As Long
Private mvarlMaiorAtraso As Long
Private mvardSaldoTitulos As Double
Private mvardSaldoPedidosLiberados As Double
Private mvardSaldoAtrasados As Double
Private mvardValPagtosAtraso As Double
Private mvardValorAcumuladoCompras As Double
Private mvarlNumTitulosProtestados As Long
Private mvardtDataUltimoProtesto As Date
Private mvariNumChequesDevolvidos As Integer
Private mvardtDataUltChequeDevolvido As Date
Private mvardSaldoDuplicatas As Double
Private mvarlNumPagamentos As Long
Private mvarsCGC As String
Private mvarsInscricaoEstadual As String
Private mvarsInscricaoMunicipal As String
Private mvariCodTransportadora As Integer
Private mvarsObservacao2 As String
Private mvarsContaContabil As String
Private mvariVendedor As Integer
Private mvardComissaoVendas As Double
Private mvariRegiao As Integer
Private mvariFreqVisitas As Integer
Private mvardtDataUltVisita As Date
Private mvariCodCobrador As Integer
Private mvarlEndereco As Long
Private mvarlEnderecoCobranca As Long
Private mvarlEnderecoEntrega As Long
Private mvariProxCodFilial As Integer
Private mvariCondicaoPagto As Integer
Private mvariCodPadraoCobranca As Integer
Private mvarcolCategoriaItem As New Collection
Private mvariTipoFrete As Integer
Private mvarlCodigoLoja As Long
Private mvarsInscricaoSuframa As String
Private mvarsRG As String
Private mvariTransferido As Integer
Private mvariAtivo As Integer
Private mvariFilialEmpresaLoja As Integer
Private mvariAtivoFilial As Integer
Private mvarsGuia As String
Private mvariCodTranspRedesp As Integer
Private mvariDetPagFrete As Integer
Private mvardPercFatMaiorPV As Double 'percentual a maior que pode ser faturado em relação ao PV (Cromaton)

'########################################
'Inserido por Wagner
Private mvariBloqueado As Integer
Private mvarsUsuarioCobrador As String
Private mvarobjInfoUsu As Object
Private mvariTemFaixaReceb As Integer
Private mvariRecebForaFaixa As Integer
Private mvardPercentMaisReceb As Double
Private mvardPercentMenosReceb As Double
Private mvariIgnoraRecebPadrao As Integer

Private mvarsUsuRespCallCenter As String
Private mvariRegimeTributario As Integer
Private mvariIEIsento As Integer
Private mvariIENaoContrib As Integer

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


Public Property Let sUsuRespCallCenter(ByVal vData As String)
    mvarsUsuRespCallCenter = vData
End Property

Public Property Get sUsuRespCallCenter() As String
    sUsuRespCallCenter = mvarsUsuRespCallCenter
End Property

Public Property Get iIgnoraRecebPadrao() As Integer
    iIgnoraRecebPadrao = mvariIgnoraRecebPadrao
End Property

Public Property Let iIgnoraRecebPadrao(ByVal vData As Integer)
    mvariIgnoraRecebPadrao = vData
End Property

Public Property Get iTemFaixaReceb() As Integer
    iTemFaixaReceb = mvariTemFaixaReceb
End Property

Public Property Let iTemFaixaReceb(ByVal vData As Integer)
    mvariTemFaixaReceb = vData
End Property

Public Property Get iRecebForaFaixa() As Integer
    iRecebForaFaixa = mvariRecebForaFaixa
End Property

Public Property Let iRecebForaFaixa(ByVal vData As Integer)
    mvariRecebForaFaixa = vData
End Property

Public Property Get dPercentMaisReceb() As Double
    dPercentMaisReceb = mvardPercentMaisReceb
End Property

Public Property Let dPercentMaisReceb(ByVal vData As Double)
    mvardPercentMaisReceb = vData
End Property

Public Property Get dPercentMenosReceb() As Double
    dPercentMenosReceb = mvardPercentMenosReceb
End Property

Public Property Let dPercentMenosReceb(ByVal vData As Double)
    mvardPercentMenosReceb = vData
End Property

Public Property Let dPercFatMaiorPV(ByVal vData As Double)
    mvardPercFatMaiorPV = vData
End Property

Public Property Get dPercFatMaiorPV() As Double
    dPercFatMaiorPV = mvardPercFatMaiorPV
End Property

Public Property Set objInfoUsu(ByVal vData As Object)
    Set mvarobjInfoUsu = vData
End Property

Public Property Get objInfoUsu() As Object
    Set objInfoUsu = mvarobjInfoUsu
End Property

Public Property Let sUsuarioCobrador(ByVal vData As String)
    mvarsUsuarioCobrador = vData
End Property

Public Property Get sUsuarioCobrador() As String
    sUsuarioCobrador = mvarsUsuarioCobrador
End Property

Public Property Let iBloqueado(ByVal vData As Integer)
    mvariBloqueado = vData
End Property

Public Property Get iBloqueado() As Integer
    iBloqueado = mvariBloqueado
End Property
'########################################

Public Property Let iDetPagFrete(ByVal vData As Integer)
    mvariDetPagFrete = vData
End Property

Public Property Get iDetPagFrete() As Integer
    iDetPagFrete = mvariDetPagFrete
End Property

Public Property Let iCodTranspRedesp(ByVal vData As Integer)
    mvariCodTranspRedesp = vData
End Property

Public Property Get iCodTranspRedesp() As Integer
    iCodTranspRedesp = mvariCodTranspRedesp
End Property

Public Property Let sGuia(ByVal vData As String)
    mvarsGuia = vData
End Property

Public Property Get sGuia() As String
    sGuia = mvarsGuia
End Property

Public Property Let iAtivoFilial(ByVal vData As Integer)
    mvariAtivoFilial = vData
End Property

Public Property Get iAtivoFilial() As Integer
    iAtivoFilial = mvariAtivoFilial
End Property

Public Property Let iFilialEmpresaLoja(ByVal vData As Integer)
    mvariFilialEmpresaLoja = vData
End Property

Public Property Get iFilialEmpresaLoja() As Integer
    iFilialEmpresaLoja = mvariFilialEmpresaLoja
End Property

Public Property Let iAtivo(ByVal vData As Integer)
    mvariAtivo = vData
End Property

Public Property Get iAtivo() As Integer
    iAtivo = mvariAtivo
End Property

Public Property Let iTransferido(ByVal vData As Integer)
    mvariTransferido = vData
End Property

Public Property Get iTransferido() As Integer
    iTransferido = mvariTransferido
End Property

Public Property Let sRG(ByVal vData As String)
    mvarsRG = vData
End Property

Public Property Get sRG() As String
    sRG = mvarsRG
End Property

Public Property Let sInscricaoSuframa(ByVal vData As String)
    mvarsInscricaoSuframa = vData
End Property

Public Property Get sInscricaoSuframa() As String
    sInscricaoSuframa = mvarsInscricaoSuframa
End Property

Public Property Let lCodigoLoja(ByVal vData As Long)
    mvarlCodigoLoja = vData
End Property

Public Property Get lCodigoLoja() As Long
    lCodigoLoja = mvarlCodigoLoja
End Property

Public Property Let iTipoFrete(ByVal vData As Integer)
    mvariTipoFrete = vData
End Property

Public Property Get iTipoFrete() As Integer
    iTipoFrete = mvariTipoFrete
End Property

Public Property Let iCodPadraoCobranca(ByVal vData As Integer)
    mvariCodPadraoCobranca = vData
End Property

Public Property Get iCodPadraoCobranca() As Integer
    iCodPadraoCobranca = mvariCodPadraoCobranca
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let iProxCodFilial(ByVal vData As Integer)
    mvariProxCodFilial = vData
End Property

Public Property Get iProxCodFilial() As Integer
    iProxCodFilial = mvariProxCodFilial
End Property

Public Property Let lEnderecoEntrega(ByVal vData As Long)
    mvarlEnderecoEntrega = vData
End Property

Public Property Get lEnderecoEntrega() As Long
    lEnderecoEntrega = mvarlEnderecoEntrega
End Property

Public Property Let lEnderecoCobranca(ByVal vData As Long)
    mvarlEnderecoCobranca = vData
End Property

Public Property Get lEnderecoCobranca() As Long
    lEnderecoCobranca = mvarlEnderecoCobranca
End Property

Public Property Let lEndereco(ByVal vData As Long)
    mvarlEndereco = vData
End Property

Public Property Get lEndereco() As Long
    lEndereco = mvarlEndereco
End Property

Public Property Let iCodCobrador(ByVal vData As Integer)
    mvariCodCobrador = vData
End Property

Public Property Get iCodCobrador() As Integer
    iCodCobrador = mvariCodCobrador
End Property

Public Property Let dtDataUltVisita(ByVal vData As Date)
    mvardtDataUltVisita = vData
End Property

Public Property Get dtDataUltVisita() As Date
    dtDataUltVisita = mvardtDataUltVisita
End Property

Public Property Let iFreqVisitas(ByVal vData As Integer)
    mvariFreqVisitas = vData
End Property

Public Property Get iFreqVisitas() As Integer
    iFreqVisitas = mvariFreqVisitas
End Property

Public Property Let iRegiao(ByVal vData As Integer)
    mvariRegiao = vData
End Property

Public Property Get iRegiao() As Integer
    iRegiao = mvariRegiao
End Property

Public Property Let dComissaoVendas(ByVal vData As Double)
    mvardComissaoVendas = vData
End Property

Public Property Get dComissaoVendas() As Double
    dComissaoVendas = mvardComissaoVendas
End Property

Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let sContaContabil(ByVal vData As String)
    mvarsContaContabil = vData
End Property

Public Property Get sContaContabil() As String
    sContaContabil = mvarsContaContabil
End Property

Public Property Let sObservacao2(ByVal vData As String)
    mvarsObservacao2 = vData
End Property

Public Property Get sObservacao2() As String
    sObservacao2 = mvarsObservacao2
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
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

Public Property Let lNumPagamentos(ByVal vData As Long)
    mvarlNumPagamentos = vData
End Property

Public Property Get lNumPagamentos() As Long
    lNumPagamentos = mvarlNumPagamentos
End Property

Public Property Let dSaldoDuplicatas(ByVal vData As Double)
    mvardSaldoDuplicatas = vData
End Property

Public Property Get dSaldoDuplicatas() As Double
    dSaldoDuplicatas = mvardSaldoDuplicatas
End Property

Public Property Let dtDataUltChequeDevolvido(ByVal vData As Date)
    mvardtDataUltChequeDevolvido = vData
End Property

Public Property Get dtDataUltChequeDevolvido() As Date
    dtDataUltChequeDevolvido = mvardtDataUltChequeDevolvido
End Property

Public Property Let iNumChequesDevolvidos(ByVal vData As Integer)
    mvariNumChequesDevolvidos = vData
End Property

Public Property Get iNumChequesDevolvidos() As Integer
    iNumChequesDevolvidos = mvariNumChequesDevolvidos
End Property

Public Property Let dtDataUltimoProtesto(ByVal vData As Date)
    mvardtDataUltimoProtesto = vData
End Property

Public Property Get dtDataUltimoProtesto() As Date
    dtDataUltimoProtesto = mvardtDataUltimoProtesto
End Property

Public Property Let lNumTitulosProtestados(ByVal vData As Long)
    mvarlNumTitulosProtestados = vData
End Property

Public Property Get lNumTitulosProtestados() As Long
    lNumTitulosProtestados = mvarlNumTitulosProtestados
End Property

Public Property Let dValorAcumuladoCompras(ByVal vData As Double)
    mvardValorAcumuladoCompras = vData
End Property

Public Property Get dValorAcumuladoCompras() As Double
    dValorAcumuladoCompras = mvardValorAcumuladoCompras
End Property

Public Property Let dValPagtosAtraso(ByVal vData As Double)
    mvardValPagtosAtraso = vData
End Property

Public Property Get dValPagtosAtraso() As Double
    dValPagtosAtraso = mvardValPagtosAtraso
End Property

Public Property Let dSaldoAtrasados(ByVal vData As Double)
    mvardSaldoAtrasados = vData
End Property

Public Property Get dSaldoAtrasados() As Double
    dSaldoAtrasados = mvardSaldoAtrasados
End Property

Public Property Let dSaldoPedidosLiberados(ByVal vData As Double)
    mvardSaldoPedidosLiberados = vData
End Property

Public Property Get dSaldoPedidosLiberados() As Double
    dSaldoPedidosLiberados = mvardSaldoPedidosLiberados
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

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let iCodMensagem(ByVal vData As Integer)
    mvariCodMensagem = vData
End Property

Public Property Get iCodMensagem() As Integer
    iCodMensagem = mvariCodMensagem
End Property

Public Property Let dDesconto(ByVal vData As Double)
    mvardDesconto = vData
End Property

Public Property Get dDesconto() As Double
    dDesconto = mvardDesconto
End Property

Public Property Let dLimiteCredito(ByVal vData As Double)
    mvardLimiteCredito = vData
End Property

Public Property Get dLimiteCredito() As Double
    dLimiteCredito = mvardLimiteCredito
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

Public Property Set colCategoriaItem(ByVal vData As Object)
    Set mvarcolCategoriaItem = vData
End Property

Public Property Get colCategoriaItem() As Collection
    Set colCategoriaItem = mvarcolCategoriaItem
End Property



Private Sub Class_Initialize()
    mvariIENaoContrib = -1
    mvariIEIsento = -1
End Sub
