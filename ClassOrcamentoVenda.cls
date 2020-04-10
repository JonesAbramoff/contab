VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassOrcamentoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements ClassTributoDoc 'Inserido por Leo em 03/05/02
'Polimorfismo que implementando a ClassTributoDoc
Private mvarobjTributacaoDoc As New ClassTributacaoDocAux

Private mvardtDataEmissao As Date
Private mvardtDataReferencia As Date
Private mvardPercAcrescFinanceiro As Double
Private mvardValorOutrasDespesas As Double
Private mvardValorFrete As Double
Private mvardValorDesconto As Double
Private mvardValorSeguro As Double
Private mvardValorTotal As Double
Private mvardValorProdutos As Double
Private mvariFilialEmpresa As Integer
Private mvariFilial As Integer
Private mvariCondicaoPagto As Integer
Private mvariTabelaPreco As Integer
Private mvariCobrancaAutomatica As Integer
'Private mvariNaturezaOpManual As Integer
Private mvarlCliente As Long
Private mvarlNumIntDoc As Long
'Private mvarsNaturezaOp As String
Private mvariPrazoValidade As Integer
Private mvarcolItens As New Collection
Private mvarcolParcela As New Collection
Private mvarlCodigo As Long
Private mvarsNome As String
Private mvarsNomeReduzido As String
Private mvariVendedor As Integer
Private mvariVendedor2 As Integer
Private mvarlNumIntPedVenda As Long
Private mvarlNumIntNFiscal As Long
Private mvarobjTributacaoOV As New ClassTributacaoDoc  'incluido por Leo em 30/04/02
Private mvarsUsuario As String
Private mvarlCodigoBase As Long
Private mvarlStatus As Long
Private mvarlMotivoPerda As Long
Public mvarobjTela As Object
Private mvarobjProjetoInfo As Object

'para serem usados na gravacao de um orcamento de vendas
'  na geracao de pedidos de cotacao automaticos (implementado para a Dan-Presse)
Private mvarobjCotacao As New ClassCotacao
Private mvarcolPedidoCotacao As New Collection
Private mvariStatusComercial As Integer

Private mvarcolItensSRV As New Collection
Private mvarcolGarantiaContratoSRV As New Collection
Private mvarcolProdSolicSRV As New Collection
Private mvarlNumIntSolicSRV As Long
Private mvarcolBloqueiosPV As New ColBloqueioPV
Private mvariServico As Integer
Private mvariVersao As Integer
Private mvardtDataUltAlt As Date
Private mvardHoraUltAlt As Double
Private mvariTrocarVersao As Integer
Private mvariHistorico As Integer
Private mvardtDataProxCobr As Date
Private mvardtDataPrevReceb As Date
Private mvardValorDescontoTit As Double
Private mvarobjInfoAdic As New ClassInfoAdic
Private mvardValorDescontoItens As Double
Private mvardValorItens As Double

Private mvarsEmail As String
Private mvarsContato As String
Private mvariNaoCalcICMSST As Integer

Private mvariFilialEntrega As Integer
Private mvariPrazoEntrega As Integer
Private mvariCodTransportadora As Integer
Private mvariCodTranspRedesp As Integer
Private mvarsMensagemNota As String
Private mvarsPedidoCliente As String
Private mvarsPedidoRepresentante As String
Private mvariCanalVenda As Integer
Private mvardPesoBruto As Double
Private mvardPesoLiq As Double
Private mvarsPlaca As String
Private mvarsPlacaUF As String
Private mvarlVolumeQuant As Long
Private mvarlVolumeEspecie As Long
Private mvarlVolumeMarca As Long
Private mvarsVolumeNumero As String
Private mvardVolumeTotal As Double
Private mvardtDataEnvio As Date
Private mvardtDataEntrega As Date
Private mvariDetPagFrete As Integer
Private mvariFreteRespons As Integer
Private mvariDataEnt As Integer
Private mvariMoeda As Integer
Private mvarsPrazoTexto As String
Private mvardtDataPerda As Date
Private mvardCotacao As Double
Private mvarobjIdiomaTela As New ClassIdiomaTela
Private mvariIdioma As Integer
Private mvardPercParticVend2 As Double

Public Property Let dPercParticVend2(ByVal vData As Double)
    mvardPercParticVend2 = vData
End Property

Public Property Get dPercParticVend2() As Double
    dPercParticVend2 = mvardPercParticVend2
End Property

Public Property Let iIdioma(ByVal vData As Integer)
    mvariIdioma = vData
End Property

Public Property Get iIdioma() As Integer
    iIdioma = mvariIdioma
End Property

Public Property Set objIdiomaTela(ByVal vData As ClassIdiomaTela)
    Set mvarobjIdiomaTela = vData
End Property

Public Property Get objIdiomaTela() As ClassIdiomaTela
    Set objIdiomaTela = mvarobjIdiomaTela
End Property

Public Property Let dCotacao(ByVal vData As Double)
    mvardCotacao = vData
End Property

Public Property Get dCotacao() As Double
    dCotacao = mvardCotacao
End Property

Public Property Let dtDataPerda(ByVal vData As Date)
    mvardtDataPerda = vData
End Property

Public Property Get dtDataPerda() As Date
    dtDataPerda = mvardtDataPerda
End Property

Public Property Let sPrazoTexto(ByVal vData As String)
    mvarsPrazoTexto = vData
End Property

Public Property Get sPrazoTexto() As String
    sPrazoTexto = mvarsPrazoTexto
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let iDataEnt(ByVal vData As Integer)
    mvariDataEnt = vData
End Property

Public Property Get iDataEnt() As Integer
    iDataEnt = mvariDataEnt
End Property

Public Property Let iFreteRespons(ByVal vData As Integer)
    mvariFreteRespons = vData
End Property

Public Property Get iFreteRespons() As Integer
    iFreteRespons = mvariFreteRespons
End Property

Public Property Let iDetPagFrete(ByVal vData As Integer)
    mvariDetPagFrete = vData
End Property

Public Property Get iDetPagFrete() As Integer
    iDetPagFrete = mvariDetPagFrete
End Property

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let dtDataEnvio(ByVal vData As Date)
    mvardtDataEnvio = vData
End Property

Public Property Get dtDataEnvio() As Date
    dtDataEnvio = mvardtDataEnvio
End Property

Public Property Let dVolumeTotal(ByVal vData As Double)
    mvardVolumeTotal = vData
End Property

Public Property Get dVolumeTotal() As Double
    dVolumeTotal = mvardVolumeTotal
End Property

Public Property Let sVolumeNumero(ByVal vData As String)
    mvarsVolumeNumero = vData
End Property

Public Property Get sVolumeNumero() As String
    sVolumeNumero = mvarsVolumeNumero
End Property

Public Property Let lVolumeMarca(ByVal vData As Long)
    mvarlVolumeMarca = vData
End Property

Public Property Get lVolumeMarca() As Long
    lVolumeMarca = mvarlVolumeMarca
End Property

Public Property Let lVolumeEspecie(ByVal vData As Long)
    mvarlVolumeEspecie = vData
End Property

Public Property Get lVolumeEspecie() As Long
    lVolumeEspecie = mvarlVolumeEspecie
End Property

Public Property Let lVolumeQuant(ByVal vData As Long)
    mvarlVolumeQuant = vData
End Property

Public Property Get lVolumeQuant() As Long
    lVolumeQuant = mvarlVolumeQuant
End Property

Public Property Let sPlacaUF(ByVal vData As String)
    mvarsPlacaUF = vData
End Property

Public Property Get sPlacaUF() As String
    sPlacaUF = mvarsPlacaUF
End Property

Public Property Let sPlaca(ByVal vData As String)
    mvarsPlaca = vData
End Property

Public Property Get sPlaca() As String
    sPlaca = mvarsPlaca
End Property

Public Property Let dPesoLiq(ByVal vData As Double)
    mvardPesoLiq = vData
End Property

Public Property Get dPesoLiq() As Double
    dPesoLiq = mvardPesoLiq
End Property

Public Property Let dPesoBruto(ByVal vData As Double)
    mvardPesoBruto = vData
End Property

Public Property Get dPesoBruto() As Double
    dPesoBruto = mvardPesoBruto
End Property

Public Property Let sPedidoRepresentante(ByVal vData As String)
    mvarsPedidoRepresentante = vData
End Property

Public Property Get sPedidoRepresentante() As String
    sPedidoRepresentante = mvarsPedidoRepresentante
End Property

Public Property Let iCanalVenda(ByVal vData As Integer)
    mvariCanalVenda = vData
End Property

Public Property Get iCanalVenda() As Integer
    iCanalVenda = mvariCanalVenda
End Property

Public Property Let sPedidoCliente(ByVal vData As String)
    mvarsPedidoCliente = vData
End Property

Public Property Get sPedidoCliente() As String
    sPedidoCliente = mvarsPedidoCliente
End Property

Public Property Let sMensagemNota(ByVal vData As String)
    mvarsMensagemNota = vData
End Property

Public Property Get sMensagemNota() As String
    sMensagemNota = mvarsMensagemNota
End Property

Public Property Let iCodTranspRedesp(ByVal vData As Integer)
    mvariCodTranspRedesp = vData
End Property

Public Property Get iCodTranspRedesp() As Integer
    iCodTranspRedesp = mvariCodTranspRedesp
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
End Property

Public Property Let iFilialEntrega(ByVal vData As Integer)
    mvariFilialEntrega = vData
End Property

Public Property Get iFilialEntrega() As Integer
    iFilialEntrega = mvariFilialEntrega
End Property

Public Property Let iPrazoEntrega(ByVal vData As Integer)
    mvariPrazoEntrega = vData
End Property

Public Property Get iPrazoEntrega() As Integer
    iPrazoEntrega = mvariPrazoEntrega
End Property

Public Property Let iNaoCalcICMSST(ByVal vData As Integer)
    mvariNaoCalcICMSST = vData
End Property

Public Property Get iNaoCalcICMSST() As Integer
    iNaoCalcICMSST = mvariNaoCalcICMSST
End Property

Public Property Let sContato(ByVal vData As String)
    mvarsContato = vData
End Property

Public Property Get sContato() As String
    sContato = mvarsContato
End Property

Public Property Let sEmail(ByVal vData As String)
    mvarsEmail = vData
End Property

Public Property Get sEmail() As String
    sEmail = mvarsEmail
End Property

Public Property Let dValorItens(ByVal vData As Double)
    mvardValorItens = vData
End Property

Public Property Get dValorItens() As Double
    dValorItens = mvardValorItens
End Property

Public Property Let dValorDescontoItens(ByVal vData As Double)
    mvardValorDescontoItens = vData
End Property

Public Property Get dValorDescontoItens() As Double
    dValorDescontoItens = mvardValorDescontoItens
End Property

Public Property Set objInfoAdic(ByVal vData As ClassInfoAdic)
    Set mvarobjInfoAdic = vData
End Property

Public Property Get objInfoAdic() As ClassInfoAdic
    Set objInfoAdic = mvarobjInfoAdic
End Property

Public Property Let dValorDescontoTit(ByVal vData As Double)
    mvardValorDescontoTit = vData
End Property

Public Property Get dValorDescontoTit() As Double
    dValorDescontoTit = mvardValorDescontoTit
End Property

Public Property Let dtDataPrevReceb(ByVal vDataPrevReceb As Date)
    mvardtDataPrevReceb = vDataPrevReceb
End Property

Public Property Get dtDataPrevReceb() As Date
    dtDataPrevReceb = mvardtDataPrevReceb
End Property

Public Property Let dtDataProxCobr(ByVal vDataProxCobr As Date)
    mvardtDataProxCobr = vDataProxCobr
End Property

Public Property Get dtDataProxCobr() As Date
    dtDataProxCobr = mvardtDataProxCobr
End Property

Public Property Let iHistorico(ByVal vData As Integer)
    mvariHistorico = vData
End Property

Public Property Get iHistorico() As Integer
    iHistorico = mvariHistorico
End Property

Public Property Let iTrocarVersao(ByVal vData As Integer)
    mvariTrocarVersao = vData
End Property

Public Property Get iTrocarVersao() As Integer
    iTrocarVersao = mvariTrocarVersao
End Property

Public Property Let dHoraUltAlt(ByVal vData As Double)
    mvardHoraUltAlt = vData
End Property

Public Property Get dHoraUltAlt() As Double
    dHoraUltAlt = mvardHoraUltAlt
End Property

Public Property Let dtDataUltAlt(ByVal vData As Date)
    mvardtDataUltAlt = vData
End Property

Public Property Get dtDataUltAlt() As Date
    dtDataUltAlt = mvardtDataUltAlt
End Property

Public Property Let iVersao(ByVal vData As Integer)
    mvariVersao = vData
End Property

Public Property Get iVersao() As Integer
    iVersao = mvariVersao
End Property

Public Property Let iServico(ByVal vData As Integer)
    mvariServico = vData
End Property

Public Property Get iServico() As Integer
    iServico = mvariServico
End Property

Public Property Set colBloqueiosPV(ByVal vData As Object)
    Set mvarcolBloqueiosPV = vData
End Property

Public Property Get colBloqueiosPV() As ColBloqueioPV
    Set colBloqueiosPV = mvarcolBloqueiosPV
End Property

Public Property Let lNumIntSolicSRV(ByVal vData As Long)
    mvarlNumIntSolicSRV = vData
End Property

Public Property Get lNumIntSolicSRV() As Long
    lNumIntSolicSRV = mvarlNumIntSolicSRV
End Property

Public Property Set colProdSolicSRV(ByVal vData As Collection)
    Set mvarcolProdSolicSRV = vData
End Property

Public Property Get colProdSolicSRV() As Collection
    Set colProdSolicSRV = mvarcolProdSolicSRV
End Property

Public Property Set colGarantiaContratoSRV(ByVal vData As Collection)
    Set mvarcolGarantiaContratoSRV = vData
End Property

Public Property Get colGarantiaContratoSRV() As Collection
    Set colGarantiaContratoSRV = mvarcolGarantiaContratoSRV
End Property

Public Property Set colItensSRV(ByVal vData As Collection)
    Set mvarcolItensSRV = vData
End Property

Public Property Get colItensSRV() As Collection
    Set colItensSRV = mvarcolItensSRV
End Property

Public Property Let iStatusComercial(ByVal vData As Integer)
    mvariStatusComercial = vData
End Property

Public Property Get iStatusComercial() As Integer
    iStatusComercial = mvariStatusComercial
End Property

Public Property Set colPedidoCotacao(ByVal vData As Object)
    Set mvarcolPedidoCotacao = vData
End Property

Public Property Get colPedidoCotacao() As Collection
    Set colPedidoCotacao = mvarcolPedidoCotacao
End Property

Public Property Set objCotacao(ByVal vData As Object)
    Set mvarobjCotacao = vData
End Property

Public Property Get objCotacao() As ClassCotacao
    Set objCotacao = mvarobjCotacao
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let lMotivoPerda(ByVal vData As Long)
    mvarlMotivoPerda = vData
End Property
Public Property Get lMotivoPerda() As Long
    lMotivoPerda = mvarlMotivoPerda
End Property

Public Property Let lStatus(ByVal vData As Long)
    mvarlStatus = vData
End Property

Public Property Get lStatus() As Long
    lStatus = mvarlStatus
End Property
'####################################

Public Property Let sUsuario(ByVal vData As String)
    mvarsUsuario = vData
End Property

Public Property Get sUsuario() As String
    sUsuario = mvarsUsuario
End Property

Public Property Let lCodigoBase(ByVal vData As Long)
    mvarlCodigoBase = vData
End Property

Public Property Get lCodigoBase() As Long
    lCodigoBase = mvarlCodigoBase
End Property

Public Property Let sNomeCli(ByVal vData As String)
    mvarsNome = vData
End Property

Public Property Get sNomeCli() As String
    sNomeCli = mvarsNome
End Property

Public Property Let sNomeFilialCli(ByVal vData As String)
    mvarsNomeReduzido = vData
End Property

Public Property Get sNomeFilialCli() As String
    sNomeFilialCli = mvarsNomeReduzido
End Property

Public Property Let lCodigo(ByVal vData As Long)
    mvarlCodigo = vData
End Property

Public Property Get lCodigo() As Long
    lCodigo = mvarlCodigo
End Property

Public Property Set colParcela(ByVal vData As Collection)
    Set mvarcolParcela = vData
End Property

Public Property Get colParcela() As Collection
    Set colParcela = mvarcolParcela
End Property

Public Property Set colItens(ByVal vData As Collection)
    Set mvarcolItens = vData
End Property

Public Property Get colItens() As Collection
    Set colItens = mvarcolItens
End Property

Public Property Let iPrazoValidade(ByVal vData As Integer)
    mvariPrazoValidade = vData
End Property

Public Property Get iPrazoValidade() As Integer
    iPrazoValidade = mvariPrazoValidade
End Property

Public Property Let sNaturezaOp(ByVal vData As String)
    objTributacao.sNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = objTributacao.sNaturezaOp
End Property

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    objTributacao.iNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = objTributacao.iNaturezaOpManual
End Property

Public Property Let iCobrancaAutomatica(ByVal vData As Integer)
    mvariCobrancaAutomatica = vData
End Property

Public Property Get iCobrancaAutomatica() As Integer
    iCobrancaAutomatica = mvariCobrancaAutomatica
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let iFilialEmpresa(ByVal vData As Integer)
    mvariFilialEmpresa = vData
End Property

Public Property Get iFilialEmpresa() As Integer
    iFilialEmpresa = mvariFilialEmpresa
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Let dValorSeguro(ByVal vData As Double)
    mvardValorSeguro = vData
End Property

Public Property Get dValorSeguro() As Double
    dValorSeguro = mvardValorSeguro
End Property

Public Property Let dValorDesconto(ByVal vData As Double)
    mvardValorDesconto = vData
End Property

Public Property Get dValorDesconto() As Double
    dValorDesconto = mvardValorDesconto
End Property

Public Property Let dValorFrete(ByVal vData As Double)
    mvardValorFrete = vData
End Property

Public Property Get dValorFrete() As Double
    dValorFrete = mvardValorFrete
End Property

Public Property Let dValorOutrasDespesas(ByVal vData As Double)
    mvardValorOutrasDespesas = vData
End Property

Public Property Get dValorOutrasDespesas() As Double
    dValorOutrasDespesas = mvardValorOutrasDespesas
End Property

Public Property Let dPercAcrescFinanceiro(ByVal vData As Double)
    mvardPercAcrescFinanceiro = vData
End Property

Public Property Get dPercAcrescFinanceiro() As Double
    dPercAcrescFinanceiro = mvardPercAcrescFinanceiro
End Property

Public Property Let dtDataReferencia(ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia() As Date
    dtDataReferencia = mvardtDataReferencia
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

'Por Leo em 02/04/02 *************
Public Property Let iVendedor(ByVal vData As Integer)
    mvariVendedor = vData
End Property

Public Property Get iVendedor() As Integer
    iVendedor = mvariVendedor
End Property

Public Property Let iVendedor2(ByVal vData As Integer)
    mvariVendedor2 = vData
End Property

Public Property Get iVendedor2() As Integer
    iVendedor2 = mvariVendedor2
End Property

Public Property Let lNumIntPedVenda(ByVal vData As Long)
    mvarlNumIntPedVenda = vData
End Property

Public Property Get lNumIntPedVenda() As Long
    lNumIntPedVenda = mvarlNumIntPedVenda
End Property

Public Property Let lNumIntNFiscal(ByVal vData As Long)
    mvarlNumIntNFiscal = vData
End Property

Public Property Get lNumIntNFiscal() As Long
    lNumIntNFiscal = mvarlNumIntNFiscal
End Property

Public Property Set objTributacaoOV(ByVal vData As Object)
    Set mvarobjTributacaoOV = vData
End Property

Public Property Get objTributacaoOV() As ClassTributacaoDoc
    Set objTributacaoOV = mvarobjTributacaoOV
End Property

Private Sub Class_Initialize()
    mvardtDataProxCobr = DATA_NULA
    mvardtDataPrevReceb = DATA_NULA
End Sub

Private Function ClassTributoDoc_Ativar() As Long

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemOrcamento As ClassItemOV

On Error GoTo Erro_ClassTributoDoc_Ativar

    lErro = gobjTributacao.Tributacao_Doc_Iniciar(Me)
    If lErro <> SUCESSO Then gError 103030
    
    Set mvarobjTributacaoDoc = New ClassTributacaoDocAux
    
    'ativar o ov em si
    If mvarobjTributacaoDoc.Ativar(Me) <> SUCESSO Then
        
        Set mvarobjTributacaoDoc = Nothing
        gError 103031
    
    End If
    
    'ativar os seus itens
    For Each objItemOrcamento In colItens
        
        Set objTributoDocItem = objItemOrcamento
        
        lErro = objTributoDocItem.Ativar(Me)
        If lErro <> SUCESSO Then gError 103032
    
    Next
        
    For Each objItemOrcamento In colItensSRV
        
        Set objTributoDocItem = objItemOrcamento
        
        lErro = objTributoDocItem.Ativar(Me)
        If lErro <> SUCESSO Then gError 103032
    
    Next
        
    ClassTributoDoc_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Ativar:

    ClassTributoDoc_Ativar = gErr

    Select Case gErr

        Case 103030, 103031, 103032
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152214)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_Desativar() As Long

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemOrcamento As ClassItemOV

On Error GoTo Erro_ClassTributoDoc_Desativar
    
    'desativar os seus itens
    For Each objItemOrcamento In colItens
        
        Set objTributoDocItem = objItemOrcamento
        
        lErro = objTributoDocItem.Desativar()
        If lErro <> SUCESSO Then gError 103027
    
    Next
    
    'desativar o ov em si
    If Not (mvarobjTributacaoDoc Is Nothing) Then
        
        lErro = mvarobjTributacaoDoc.Desativar()
        If lErro <> SUCESSO Then gError 103028
        
        Set mvarobjTributacaoDoc = Nothing
    
    End If
     
    lErro = gobjTributacao.Tributacao_Doc_Terminar(Me)
    If lErro <> SUCESSO Then gError 103029
    
    ClassTributoDoc_Desativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Desativar:

    ClassTributoDoc_Desativar = gErr

    Select Case gErr

        Case 103027, 103028, 103029
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152215)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_GuardarVarTela(sIdentificador As String, vValor As Variant) As Long

Dim lErro As Long
Dim objVariavelCalculo As New ClassVariavelCalculo

On Error GoTo Erro_ClassTributoDoc_GuardarVarTela

    objVariavelCalculo.sIdentificador = sIdentificador
    objVariavelCalculo.vValor = vValor
    objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
    
    ClassTributoDoc_GuardarVarTela = ClassTributoDoc_GuardarVar(objVariavelCalculo)

    Exit Function

Erro_ClassTributoDoc_GuardarVarTela:

    ClassTributoDoc_GuardarVarTela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152216)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDoc_objTributacaoDoc() As GlobaisTRB.ClassTributacaoDocAux
    Set ClassTributoDoc_objTributacaoDoc = mvarobjTributacaoDoc
End Property

Private Function ClassTributoDoc_ObterColDocCompl(colDocCompl As Collection) As Long
'
'Dim objOVTribCompl As ClassPVTribCompl
'
'    Set colDocCompl = New Collection
'
'    Set objOVTribCompl = New ClassPVTribCompl
'    Call objOVTribCompl.Iniciar(Me, mvarobjTributacaoOV.objTributacaoFrete)
'    Call colDocCompl.Add(objOVTribCompl)
'
'    Set objOVTribCompl = New ClassPVTribCompl
'    Call objOVTribCompl.Iniciar(Me, mvarobjTributacaoOV.objTributacaoSeguro)
'    Call colDocCompl.Add(objOVTribCompl)
'
'    Set objOVTribCompl = New ClassPVTribCompl
'    Call objOVTribCompl.Iniciar(Me, mvarobjTributacaoOV.objTributacaoDesconto)
'    Call colDocCompl.Add(objOVTribCompl)
'
'    Set objOVTribCompl = New ClassPVTribCompl
'    Call objOVTribCompl.Iniciar(Me, mvarobjTributacaoOV.objTributacaoOutras)
'    Call colDocCompl.Add(objOVTribCompl)
    
    ClassTributoDoc_ObterColDocCompl = SUCESSO
    
End Function

Private Function ClassTributoDoc_ObterColDocItens(colDocItens As Collection) As Long
'retorna colecao de itens que implementam interface como TaxDocItemExemplo

Dim objItemOrcamento As ClassItemOV

   Set colDocItens = New Collection
   
   For Each objItemOrcamento In mvarcolItens
        Call colDocItens.Add(objItemOrcamento)
   Next
   
   For Each objItemOrcamento In mvarcolItensSRV
        Call colDocItens.Add(objItemOrcamento)
   Next
   
   ClassTributoDoc_ObterColDocItens = SUCESSO
   
End Function

Private Function ClassTributoDoc_ObterDocAux(objTributacaoDoc As ClassTributacaoDocAux) As Long
'retorna objeto auxiliar vinculado ao Doc
    
    Set objTributacaoDoc = mvarobjTributacaoDoc
    ClassTributoDoc_ObterDocAux = SUCESSO

End Function

Private Function ClassTributoDoc_ObterNatOpPadrao(sNatOp As String) As Long

Dim lErro As Long

On Error GoTo Erro_ClassTributoDoc_ObterNatOpPadrao

    lErro = mvarobjTributacaoDoc.ObterNaturezaOpPadrao(sNatOp)
    If lErro <> SUCESSO Then gError 103026
    
    ClassTributoDoc_ObterNatOpPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterNatOpPadrao:

    ClassTributoDoc_ObterNatOpPadrao = gErr

    Select Case gErr

        Case 103026
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152217)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterTribPadrao(iTipoTrib As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_ClassTributoDoc_ObterTribPadrao

    lErro = mvarobjTributacaoDoc.ObterTribPadrao(iTipoTrib)
    If lErro <> SUCESSO Then gError 103025
    
    ClassTributoDoc_ObterTribPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterTribPadrao:

    ClassTributoDoc_ObterTribPadrao = gErr

    Select Case gErr

        Case 103025
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152218)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long

On Error GoTo Erro_ClassTributoDoc_ObterVar
    
    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    objVariavelCalculo.sIdentificador = sIdentificador
    
    Select Case sIdentificador
        
        Case "SERIE"
            objVariavelCalculo.vValor = ""
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "DATA_ENTRADA"
            objVariavelCalculo.vValor = DATA_NULA
            objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
        
        Case "DATA_EMISSAO"
            objVariavelCalculo.vValor = mvardtDataEmissao
            If mvardtDataEmissao <> DATA_NULA Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
        
        Case "FILIAL_FATURAMENTO"
            objVariavelCalculo.vValor = mvariFilialEmpresa
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
'        Case "NAT_OPERACAO"
'            objVariavelCalculo.vValor = mvarsNaturezaOp
'            objVariavelCalculo.iStatus = mvariNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.iTipoTributacao
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iTipoTributacaoManual
        
        Case "TIPO_DOC_SIGLA"
            If iServico = MARCADO Then
                objVariavelCalculo.vValor = SIGLA_PSRV_NORMAL
            Else
                objVariavelCalculo.vValor = SIGLA_PV_NORMAL
            End If
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "TIPO_DOC_CODIGO"
            If iServico = MARCADO Then
                objVariavelCalculo.vValor = CODIGO_PSRV_NORMAL
            Else
                objVariavelCalculo.vValor = CODIGO_PV_NORMAL
            End If
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "CLIENTE_CODIGO"
            objVariavelCalculo.vValor = mvarlCliente
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
        Case "FILIAL_CLIENTE_CODIGO"
            objVariavelCalculo.vValor = mvariFilial
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "FRETE_VALOR"
            objVariavelCalculo.vValor = mvardValorFrete
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "SEGURO_VALOR"
            objVariavelCalculo.vValor = mvardValorSeguro
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
        
        Case "DESP_ACESS_VALOR"
            objVariavelCalculo.vValor = mvardValorOutrasDespesas
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL

        Case "DESCONTO_VALOR"
            objVariavelCalculo.vValor = mvardValorDesconto
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
'        'para ICMS
'        Case "ICMS_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dICMSBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iICMSBaseManual
'
'        Case "ICMS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dICMSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iICMSValorManual
'
'        Case "ICMS_SUBST_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dICMSSubstBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iICMSSubstBaseManual
'
'        Case "ICMS_SUBST_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dICMSSubstValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iICMSSubstValorManual
'
'        Case "ICMS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dICMSCredito
'
'        'para IPI
'        Case "IPI_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dIPIBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iIPIBaseManual
'
'        Case "IPI_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dIPIValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iIPIValorManual
'
'        Case "IPI_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dIPICredito
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.iISSIncluso
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iISSInclusoManual
'
'        Case "ISS_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dISSBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ISS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dISSAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iISSAliquotaManual
'
'        Case "ISS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dISSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iISSValorManual
'
'        'para IR
'        Case "IR_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dIRRFBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "IR_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dIRRFAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iIRRFAliquotaManual
'
'        Case "IR_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dIRRFValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iIRRFValorManual
'
'        Case "PIS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dPISCredito
'
'        Case "COFINS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dCOFINSCredito
'
'        Case "PIS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dPISRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iPISRetidoManual
'
'        Case "ISS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dISSRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iISSRetidoManual
'
'        Case "COFINS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dCOFINSRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iCOFINSRetidoManual
'
'        Case "CSLL_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoOV.dCSLLRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoOV.iCSLLRetidoManual
            
        Case "ICMS_ST_NAO_CALCULAR"
            objVariavelCalculo.vValor = mvariNaoCalcICMSST
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            
        Case Else
'        Case "FILIAL_CLIENTE_OBJ"
'        Case "FILIAL_CLIENTE_INSC_EST"
'        Case "CLIENTE_ENDERECO_OBJ"
'        Case "CLIENTE_UF" '(EX p/exterior)
            
            lErro = objTributacao.ObterVar_Padrao(sIdentificador, objVariavelCalculo)
            If lErro <> SUCESSO Then
                'delegar p/tributacaoDoc
                lErro = mvarobjTributacaoDoc.ObterVar(sIdentificador, objVariavelCalculo)
                If lErro <> SUCESSO Then gError 103024
            End If
            
    End Select
    
    ClassTributoDoc_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterVar:

    ClassTributoDoc_ObterVar = gErr

    Select Case gErr

        Case 103024
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152219)
        
    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_GuardarVar(objVariavelCalculo As ClassVariavelCalculo) As Long
'armazena info de variavel guardando se esta foi alterada ou nao desde ...
Dim lErro As Long
On Error GoTo Erro_ClassTributoDoc_GuardarVar
        
    Select Case objVariavelCalculo.sIdentificador
        
'        Case "NAT_OPERACAO"
'            mvarsNaturezaOp = objVariavelCalculo.vValor
'            objVariavelCalculo.iStatus = mvariNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            mvarobjTributacaoOV.iTipoTributacao = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iTipoTributacaoManual = objVariavelCalculo.iStatus
        
        Case "TIPO_DOC_SIGLA"
        Case "TIPO_DOC_CODIGO"
        
        Case "CLIENTE_CODIGO"
        
        Case "FILIAL_CLIENTE_OBJ"
        
        Case "FILIAL_CLIENTE_CODIGO"
        
        Case "FILIAL_CLIENTE_INSC_EST"

        Case "CLIENTE_ENDERECO_OBJ"
        
        Case "CLIENTE_UF" '(EX p/exterior)
        
        Case "FRETE_VALOR"
            mvardValorFrete = objVariavelCalculo.vValor
        
        Case "SEGURO_VALOR"
            mvardValorSeguro = objVariavelCalculo.vValor
        
        Case "DESP_ACESS_VALOR"
            mvardValorOutrasDespesas = objVariavelCalculo.vValor

'        'para ICMS
'        Case "ICMS_BASE"
'            mvarobjTributacaoOV.dICMSBase = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iICMSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR"
'            mvarobjTributacaoOV.dICMSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iICMSValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_BASE"
'            mvarobjTributacaoOV.dICMSSubstBase = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iICMSSubstBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_VALOR"
'            mvarobjTributacaoOV.dICMSSubstValor = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iICMSSubstValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR_CREDITO"
'            mvarobjTributacaoOV.dICMSCredito = objVariavelCalculo.vValor
'
'        'para IPI
'        Case "IPI_BASE"
'            mvarobjTributacaoOV.dIPIBase = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iIPIBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR"
'            mvarobjTributacaoOV.dIPIValor = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iIPIValorManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR_CREDITO"
'            mvarobjTributacaoOV.dIPICredito = objVariavelCalculo.vValor
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            mvarobjTributacaoOV.iISSIncluso = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iISSInclusoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_BASE"
'            mvarobjTributacaoOV.dISSBase = objVariavelCalculo.vValor
'
'        Case "ISS_ALIQUOTA"
'            mvarobjTributacaoOV.dISSAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iISSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ISS_VALOR"
'            mvarobjTributacaoOV.dISSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iISSValorManual = objVariavelCalculo.iStatus
'
'        'para IR
'        Case "IR_BASE"
'            mvarobjTributacaoOV.dIRRFBase = objVariavelCalculo.vValor
'
'        Case "IR_ALIQUOTA"
'            mvarobjTributacaoOV.dIRRFAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iIRRFAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "IR_VALOR"
'            mvarobjTributacaoOV.dIRRFValor = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iIRRFValorManual = objVariavelCalculo.iStatus
'
'        Case "PIS_VALOR_CREDITO"
'             mvarobjTributacaoOV.dPISCredito = objVariavelCalculo.vValor
'
'        Case "COFINS_VALOR_CREDITO"
'            mvarobjTributacaoOV.dCOFINSCredito = objVariavelCalculo.vValor
'
'        Case "PIS_RETIDO"
'            mvarobjTributacaoOV.dPISRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iPISRetidoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_RETIDO"
'            mvarobjTributacaoOV.dISSRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iISSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_RETIDO"
'            mvarobjTributacaoOV.dCOFINSRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iCOFINSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "CSLL_RETIDO"
'            mvarobjTributacaoOV.dCSLLRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoOV.iCSLLRetidoManual = objVariavelCalculo.iStatus
        
        Case Else
            lErro = objTributacao.GuardarVar_Padrao(objVariavelCalculo)
            If lErro <> SUCESSO Then
                'delegar p/tributacaoDoc
                lErro = mvarobjTributacaoDoc.GuardarVar(objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27369
            End If
            
    End Select

    ClassTributoDoc_GuardarVar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_GuardarVar:

    ClassTributoDoc_GuardarVar = Err

    Select Case Err

        Case 27369
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152220)

    End Select

    Exit Function

End Function

Public Function IncluirItem(ByVal sProduto As String, ByVal dQuantidade As Double, ByVal dPrecoUnitario As Double, ByVal dPrecoTotal As Double, ByVal dValorDesconto As Double, ByVal dtDataEntrega As Date, ByVal sProdutoDescricao As String, ByVal sUnidadeMed As String, ByVal iClasseUM As Integer, ByVal sDescricao As String, ByVal iStatus As Integer, ByVal lNumIntDoc As Long, ByVal objTributacaoItemOV As ClassTributacaoDocItem) As ClassItemOV
'associa um novo item ao orcamento de venda

Dim objItemOrcamento As New ClassItemOV ' Alterar para ItemOV

On Error GoTo Erro_IncluirItem

    objItemOrcamento.iFilialEmpresa = mvariFilialEmpresa
    objItemOrcamento.lCodOrcamento = mvarlCodigo
    objItemOrcamento.sProduto = sProduto
    objItemOrcamento.dQuantidade = dQuantidade
    objItemOrcamento.dPrecoUnitario = dPrecoUnitario
    objItemOrcamento.dPrecoTotal = dPrecoTotal
    objItemOrcamento.dValorDesconto = dValorDesconto
    objItemOrcamento.dtDataEntrega = dtDataEntrega
    objItemOrcamento.sProdutoDescricao = sProdutoDescricao
    objItemOrcamento.sUnidadeMed = sUnidadeMed
    objItemOrcamento.iClasseUM = iClasseUM
    objItemOrcamento.sDescricao = sDescricao
    objItemOrcamento.iStatus = iStatus
    objItemOrcamento.lNumIntDoc = lNumIntDoc
    
    Set objItemOrcamento.objTributacaoItemOV = objTributacaoItemOV
    
    mvarcolItens.Add objItemOrcamento
         
    Set IncluirItem = objItemOrcamento

    Exit Function

Erro_IncluirItem:

    Set IncluirItem = Nothing

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152221)

    End Select

    Exit Function

End Function

Public Function RemoverItem(vntIndexKey As Variant) As Long
'exclui um item do OV identificando-o como na colecao de itens

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemOrcamento As ClassItemOV

On Error GoTo Erro_RemoverItem

    Set objItemOrcamento = colItens.Item(vntIndexKey)
    
    Set objTributoDocItem = objItemOrcamento
    
    lErro = objTributoDocItem.Desativar()
    If lErro <> SUCESSO Then gError 103023
        
    Call colItens.Remove(vntIndexKey)
    
    RemoverItem = SUCESSO

    Exit Function

Erro_RemoverItem:

    RemoverItem = gErr

    Select Case gErr

        Case 103023
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152222)

    End Select

    Exit Function

End Function

'Leo até aqui *******************

Public Function RemoverItemSRV(vntIndexKey As Variant) As Long
'exclui um item do OV identificando-o como na colecao de itens

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemOrcamento As ClassItemOV

On Error GoTo Erro_RemoverItemSRV

    Set objItemOrcamento = colItensSRV.Item(vntIndexKey)
    
    Set objTributoDocItem = objItemOrcamento
    
    lErro = objTributoDocItem.Desativar()
    If lErro <> SUCESSO Then gError 103023
        
    Call colItensSRV.Remove(vntIndexKey)
    
    RemoverItemSRV = SUCESSO

    Exit Function

Erro_RemoverItemSRV:

    RemoverItemSRV = gErr

    Select Case gErr

        Case 103023
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 152222)

    End Select

    Exit Function

End Function

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacaoOV = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDoc
    Set objTributacao = mvarobjTributacaoOV
End Property

Private Function ClassTributoDoc_Alterado() As Long
    ClassTributoDoc_Alterado = mvarobjTributacaoDoc.Alterado
End Function

Public Sub AnaliseVenda_Preenche(ByVal objAnaliseVenda As ClassAnaliseVendaInfo)
Dim objItemAnalise As ClassItemAnaliseVenda
Dim objItemOV As ClassItemOV

    With objAnaliseVenda
        .dtDataEmissao = mvardtDataEmissao
        .iFilial = mvariFilial
        .iFilialEmpresa = mvariFilialEmpresa
        .iTabelaPreco = mvariTabelaPreco
        .lCliente = mvarlCliente
        .iVendedor = mvariVendedor
        .iMoeda = mvariMoeda
    End With

    Set objAnaliseVenda.colItens = New Collection
    
    For Each objItemOV In mvarcolItens
    
        If objItemOV.objInfoAdicDocItem.iIncluiValorTotal <> 0 Then
    
            Set objItemAnalise = New ClassItemAnaliseVenda
        
            With objItemAnalise
            
                '.dPercDesc = objItemOV.dPercDesc
                .dPrecoTotal = objItemOV.dPrecoTotal
                '.dPrecoTotalB = objItemOV.dPrecoTotalB
                .dPrecoUnitario = objItemOV.dPrecoUnitario
                .dPrecoBase = objItemOV.dPrecoUnitario
                .dQuantidade = objItemOV.dQuantidade
                .dValorDesconto = objItemOV.dValorDesconto
                .sProduto = objItemOV.sProduto
                .sDescricao = objItemOV.sDescricao
                .sUnidadeMed = objItemOV.sUnidadeMed
            
            End With
        
            objAnaliseVenda.colItens.Add objItemAnalise
    
        End If
    
    Next
    
End Sub
