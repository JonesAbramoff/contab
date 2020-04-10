VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassECFConfig"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get gsPPAmbiente() As String
    gsPPAmbiente = SistGlobECF_sPPAmbiente
End Property

Public Property Let gsPPAmbiente(ByVal vData As String)
    SistGlobECF_sPPAmbiente = vData
End Property

Public Property Get gsPPVersao() As String
    gsPPVersao = SistGlobECF_sPPVersao
End Property

Public Property Let gsPPVersao(ByVal vData As String)
    SistGlobECF_sPPVersao = vData
End Property

Public Property Get gsPPEmailLojista() As String
    gsPPEmailLojista = SistGlobECF_sPPEmailLojista
End Property

Public Property Let gsPPEmailLojista(ByVal vData As String)
    SistGlobECF_sPPEmailLojista = vData
End Property

Public Property Get gsPPSenhaLojista() As String
    gsPPSenhaLojista = SistGlobECF_sPPSenhaLojista
End Property

Public Property Let gsPPSenhaLojista(ByVal vData As String)
    SistGlobECF_sPPSenhaLojista = vData
End Property

Public Property Get glPPCodLoja() As Long
    glPPCodLoja = SistGlobECF_lPPCodLoja
End Property

Public Property Let glPPCodLoja(ByVal vData As Long)
    SistGlobECF_lPPCodLoja = vData
End Property

Public Property Get gsIntegracaoTipo() As String
    gsIntegracaoTipo = SistGlobECF_sIntegracaoTipo
End Property

Public Property Let gsIntegracaoTipo(ByVal vData As String)
    SistGlobECF_sIntegracaoTipo = vData
End Property

Public Property Get giIntegracaoTipo() As Integer
    giIntegracaoTipo = SistGlobECF_iIntegracaoTipo
End Property

Public Property Let giIntegracaoTipo(ByVal vData As Integer)
    SistGlobECF_iIntegracaoTipo = vData
End Property

Public Property Get gsNTKToken() As String
    gsNTKToken = SistGlobECF_sNTKToken
End Property

Public Property Let gsNTKToken(ByVal vData As String)
    SistGlobECF_sNTKToken = vData
End Property

Public Property Get gsNTKMerchantId() As String
    gsNTKMerchantId = SistGlobECF_sNTKMerchantId
End Property

Public Property Let gsNTKMerchantId(ByVal vData As String)
    SistGlobECF_sNTKMerchantId = vData
End Property

Public Property Get gsNTKURLServidor() As String
    gsNTKURLServidor = SistGlobECF_sNTKURLServidor
End Property

Public Property Let gsNTKURLServidor(ByVal vData As String)
    SistGlobECF_sNTKURLServidor = vData
End Property

Public Property Get gsEndLogradouro() As String
    gsEndLogradouro = SistGlobECF_sEndLogradouro
End Property

Public Property Let gsEndLogradouro(ByVal vData As String)
    SistGlobECF_sEndLogradouro = vData
End Property

Public Property Get gsEndNumero() As String
    gsEndNumero = SistGlobECF_sEndNumero
End Property

Public Property Let gsEndNumero(ByVal vData As String)
    SistGlobECF_sEndNumero = vData
End Property

Public Property Get gsEndComplemento() As String
    gsEndComplemento = SistGlobECF_sEndComplemento
End Property

Public Property Let gsEndComplemento(ByVal vData As String)
    SistGlobECF_sEndComplemento = vData
End Property

Public Property Get gsBairro() As String
    gsBairro = SistGlobECF_sBairro
End Property

Public Property Let gsBairro(ByVal vData As String)
    SistGlobECF_sBairro = vData
End Property

Public Property Get gsNomeReduzido() As String
    gsNomeReduzido = SistGlobECF_sNomeReduzido
End Property

Public Property Let gsNomeReduzido(ByVal vData As String)
    SistGlobECF_sNomeReduzido = vData
End Property



Public Property Get giTrocaAtivo() As Integer
    giTrocaAtivo = SistGlobECF_iTrocaAtivo
End Property

Public Property Let giTrocaAtivo(ByVal vData As Integer)
    SistGlobECF_iTrocaAtivo = vData
End Property

Public Property Get giCarneAtivo() As Integer
    giCarneAtivo = SistGlobECF_iCarneAtivo
End Property

Public Property Let giCarneAtivo(ByVal vData As Integer)
    SistGlobECF_iCarneAtivo = vData
End Property

Public Property Get giOutrosAtivo() As Integer
    giOutrosAtivo = SistGlobECF_iOutrosAtivo
End Property

Public Property Let giOutrosAtivo(ByVal vData As Integer)
    SistGlobECF_iOutrosAtivo = vData
End Property

Public Property Get giTicketAtivo() As Integer
    giTicketAtivo = SistGlobECF_iTicketAtivo
End Property

Public Property Let giTicketAtivo(ByVal vData As Integer)
    SistGlobECF_iTicketAtivo = vData
End Property

Public Property Get giCartaoDebitoAtivo() As Integer
    giCartaoDebitoAtivo = SistGlobECF_iCartaoDebitoAtivo
End Property

Public Property Let giCartaoDebitoAtivo(ByVal vData As Integer)
    SistGlobECF_iCartaoDebitoAtivo = vData
End Property

Public Property Get giCartaoCreditoAtivo() As Integer
    giCartaoCreditoAtivo = SistGlobECF_iCartaoCreditoAtivo
End Property

Public Property Let giCartaoCreditoAtivo(ByVal vData As Integer)
    SistGlobECF_iCartaoCreditoAtivo = vData
End Property

Public Property Get giChequeAtivo() As Integer
    giChequeAtivo = SistGlobECF_iChequeAtivo
End Property

Public Property Let giChequeAtivo(ByVal vData As Integer)
    SistGlobECF_iChequeAtivo = vData
End Property

Public Property Get giDinheiroAtivo() As Integer
    giDinheiroAtivo = SistGlobECF_iDinheiroAtivo
End Property

Public Property Let giDinheiroAtivo(ByVal vData As Integer)
    SistGlobECF_iDinheiroAtivo = vData
End Property

Public Property Get glNumProxTransfCaixa() As Long
    glNumProxTransfCaixa = SistGlobECF_lNumProxTransfCaixa
End Property

Public Property Let glNumProxTransfCaixa(ByVal vData As Long)
    SistGlobECF_lNumProxTransfCaixa = vData
End Property

Public Property Get giRemoveOrc() As Integer
    giRemoveOrc = SistGlobECF_iRemoveOrc
End Property

Public Property Let giRemoveOrc(ByVal vData As Integer)
    SistGlobECF_iRemoveOrc = vData
End Property

Public Property Get gcolOperadores() As Collection
    Set gcolOperadores = SistGlobECF_colOperadores
End Property

Public Property Set gcolOperadores(ByVal vData As Collection)
    Set SistGlobECF_colOperadores = vData
End Property

Public Property Get gcolMarcas() As Collection
    Set gcolMarcas = SistGlobECF_colMarcas
End Property

Public Property Set gcolMarcas(ByVal vData As Collection)
    Set SistGlobECF_colMarcas = vData
End Property

Public Property Get gcolVendedores() As Collection
    Set gcolVendedores = SistGlobECF_colVendedores
End Property

Public Property Set gcolVendedores(ByVal vData As Collection)
    Set SistGlobECF_colVendedores = vData
End Property

Public Property Get gcolAliquotasTotal() As Collection
    Set gcolAliquotasTotal = SistGlobECF_colAliquotasTotal
End Property

Public Property Set gcolAliquotasTotal(ByVal vData As Collection)
    Set SistGlobECF_colAliquotasTotal = vData
End Property

'Public Property Get gcolcupons() As Collection
'    Set gcolcupons = SistGlobECF_colCupons
'End Property
'
'Public Property Set gcolcupons(ByVal vData As Collection)
'    Set SistGlobECF_colCupons = vData
'End Property

Public Property Get gcolCheque() As Collection
    Set gcolCheque = SistGlobECF_colCheque
End Property

Public Property Set gcolCheque(ByVal vData As Collection)
    Set SistGlobECF_colCheque = vData
End Property

Public Property Get gcolCartao() As Collection
    Set gcolCartao = SistGlobECF_colCartao
End Property

Public Property Set gcolCartao(ByVal vData As Collection)
    Set SistGlobECF_colCartao = vData
End Property

Public Property Get gcolOutros() As Collection
    Set gcolOutros = SistGlobECF_colOutros
End Property

Public Property Set gcolOutros(ByVal vData As Collection)
    Set SistGlobECF_colOutros = vData
End Property

Public Property Get giCodCaixa() As Integer
    giCodCaixa = SistGlobECF_iCodCaixa
End Property

Public Property Let giCodCaixa(ByVal vData As Integer)
    SistGlobECF_iCodCaixa = vData
End Property

Public Property Get giStatusCaixa() As Integer
    giStatusCaixa = SistGlobECF_iStatusCaixa
End Property

Public Property Let giStatusCaixa(ByVal vData As Integer)
    SistGlobECF_iStatusCaixa = vData
End Property

Public Property Get giStatusSessao() As Integer
    giStatusSessao = SistGlobECF_iStatusSessao
End Property

Public Property Let giStatusSessao(ByVal vData As Integer)
    SistGlobECF_iStatusSessao = vData
End Property

Public Property Get giCodOperador() As Integer
    giCodOperador = SistGlobECF_iCodOperador
End Property

Public Property Let giCodOperador(ByVal vData As Integer)
    SistGlobECF_iCodOperador = vData
End Property

Public Property Get glSeqTransacaoEncerrada() As Long
    glSeqTransacaoEncerrada = SistGlobECF_lSeqTransacaoEncerrada
End Property

Public Property Let glSeqTransacaoEncerrada(ByVal vData As Long)
    SistGlobECF_lSeqTransacaoEncerrada = vData
End Property

Public Property Get glSeqTransacaoAberta() As Long
    glSeqTransacaoAberta = SistGlobECF_lSeqTransacaoAberta
End Property

Public Property Let glSeqTransacaoAberta(ByVal vData As Long)
    SistGlobECF_lSeqTransacaoAberta = vData
End Property

Public Property Get giCodECF() As Integer
    giCodECF = SistGlobECF_iCodECF
End Property

Public Property Let giCodECF(ByVal vData As Integer)
    SistGlobECF_iCodECF = vData
End Property

'Public Property Get gdSaldoDinheiro() As Double
'    gdSaldoDinheiro = SistGlobECF_dSaldoDinheiro
'End Property
'
'Public Property Let gdSaldoDinheiro(ByVal vData As Double)
'    SistGlobECF_dSaldoDinheiro = vData
'End Property

Public Property Get gdSaldocheques() As Double
    gdSaldocheques = SistGlobECF_dSaldoCheques
End Property

Public Property Let gdSaldocheques(ByVal vData As Double)
    SistGlobECF_dSaldoCheques = vData
End Property

Public Property Get gdSaldoBoletos() As Double
    gdSaldoBoletos = SistGlobECF_dSaldoBoletos
End Property

Public Property Let gdSaldoBoletos(ByVal vData As Double)
    SistGlobECF_dSaldoBoletos = vData
End Property

Public Property Get gdSaldoOutros() As Double
    gdSaldoOutros = SistGlobECF_dSaldoOutros
End Property

Public Property Let gdSaldoOutros(ByVal vData As Double)
    SistGlobECF_dSaldoOutros = vData
End Property

Public Property Get glNumProxOrcamento() As Long
    glNumProxOrcamento = SistGlobECF_lNumProxOrcamento
End Property

Public Property Let glNumProxOrcamento(ByVal vData As Long)
    SistGlobECF_lNumProxOrcamento = vData
End Property

Public Property Get gsNomeCaixa() As String
    gsNomeCaixa = SistGlobECF_sNomeCaixa
End Property

Public Property Let gsNomeCaixa(ByVal vData As String)
    SistGlobECF_sNomeCaixa = vData
End Property

Public Property Get gsMensagemCupom() As String
    gsMensagemCupom = SistGlobECF_sMensagemCupom
End Property

Public Property Let gsMensagemCupom(ByVal vData As String)
    SistGlobECF_sMensagemCupom = vData
End Property

Public Property Get giTEF() As Integer
    giTEF = SistGlobECF_iTEF
End Property

Public Property Let giTEF(ByVal vData As Integer)
    SistGlobECF_iTEF = vData
End Property


Public Property Get giBoletoManual() As Integer
    giBoletoManual = SistGlobECF_iBoletoManual
End Property

Public Property Let giBoletoManual(ByVal vData As Integer)
    SistGlobECF_iBoletoManual = vData
End Property

Public Property Get giPos() As Integer
    giPos = SistGlobECF_iPos
End Property

Public Property Let giPos(ByVal vData As Integer)
    SistGlobECF_iPos = vData
End Property

Public Property Get gcolVendas() As Collection
    Set gcolVendas = SistGlobECF_colVendas
End Property

Public Property Set gcolVendas(ByVal vData As Collection)
    Set SistGlobECF_colVendas = vData
End Property

Public Property Get gcolOrcamentos() As Collection
    Set gcolOrcamentos = SistGlobECF_colOrcamentos
End Property

Public Property Set gcolOrcamentos(ByVal vData As Collection)
    Set SistGlobECF_colOrcamentos = vData
End Property

Public Property Get gcolMovimentosCaixa() As Collection
    Set gcolMovimentosCaixa = SistGlobECF_colMovimentosCaixa
End Property

Public Property Set gcolMovimentosCaixa(ByVal vData As Collection)
    Set SistGlobECF_colMovimentosCaixa = vData
End Property

Public Property Get gcolAdmMeioPagto() As Collection
    Set gcolAdmMeioPagto = SistGlobECF_colAdmMeioPagto
End Property

Public Property Set gcolAdmMeioPagto(ByVal vData As Collection)
    Set SistGlobECF_colAdmMeioPagto = vData
End Property

Public Property Get giImpressoraCheque() As Integer
    giImpressoraCheque = SistGlobECF_iImpressoraCheque
End Property

Public Property Let giImpressoraCheque(ByVal vData As Integer)
    SistGlobECF_iImpressoraCheque = vData
End Property

Public Property Get gdtUltimaReducao() As Date
    If SistGlobECF_dtUltimaReducao = 0 Then SistGlobECF_dtUltimaReducao = DATA_NULA
    gdtUltimaReducao = SistGlobECF_dtUltimaReducao
End Property

Public Property Let gdtUltimaReducao(ByVal vData As Date)
    SistGlobECF_dtUltimaReducao = vData
End Property

Public Property Get giTeclado() As Integer
    giTeclado = SistGlobECF_iTeclado
End Property

Public Property Let giTeclado(ByVal vData As Integer)
    SistGlobECF_iTeclado = vData
End Property

Public Property Get giImpressoraECF() As Integer
    giImpressoraECF = SistGlobECF_iImpressoraECF
End Property

Public Property Let giImpressoraECF(ByVal vData As Integer)
    SistGlobECF_iImpressoraECF = vData
End Property

Public Property Set gcolTeclados(ByVal vData As Collection)
    Set SistGlobECF_colTeclados = vData
End Property

Public Property Get gcolTeclados() As Collection
    Set gcolTeclados = SistGlobECF_colTeclados
End Property

Public Property Set gcolRedes(ByVal vData As Collection)
    Set SistGlobECF_colRedes = vData
End Property

Public Property Get gcolRedes() As Collection
    Set gcolRedes = SistGlobECF_colRedes
End Property

Public Property Set gaobjProdutosNome(ByVal vData As ClassProdNome)
    Set SistGlobECF_aobjProdutosNome = vData
End Property

Public Property Get gaobjProdutosNome() As ClassProdNome
    Set gaobjProdutosNome = SistGlobECF_aobjProdutosNome
End Property

Public Property Set gaobjProdutosCodBarras(ByVal vData As ClassProdCodBarra)
    Set SistGlobECF_aobjProdutosCodBarras = vData
End Property

Public Property Get gaobjProdutosCodBarras() As ClassProdCodBarra
    Set gaobjProdutosCodBarras = SistGlobECF_aobjProdutosCodBarras
End Property

Public Property Set gaobjProdutosReferencia(ByVal vData As ClassProdReferencia)
    Set SistGlobECF_aobjProdutosReferencia = vData
End Property

Public Property Get gaobjProdutosReferencia() As ClassProdReferencia
    Set gaobjProdutosReferencia = SistGlobECF_aobjProdutosReferencia
End Property

Public Property Set gaobjProdutosCodigo(ByVal vData As ClassProdNome)
    Set SistGlobECF_aobjProdutosCodigo = vData
End Property

Public Property Get gaobjProdutosCodigo() As ClassProdCodigo
    Set gaobjProdutosCodigo = SistGlobECF_aobjProdutosCodigo
End Property

Public Property Set gobjLojaECF(ByVal vData As ClassLojaECF)
    Set SistGlobECF_objLojaECF = vData
End Property

Public Property Get gobjLojaECF() As ClassLojaECF
    Set gobjLojaECF = SistGlobECF_objLojaECF
End Property

Public Property Get glNumProxIdentificacao() As Long
    glNumProxIdentificacao = SistGlobECF_lNumProxIdentificacao
End Property

Public Property Let glNumProxIdentificacao(ByVal vData As Long)
    SistGlobECF_lNumProxIdentificacao = vData
End Property

Public Property Get gsNumSerie() As String
    gsNumSerie = SistGlobECF_sNumSerie
End Property

Public Property Let gsNumSerie(ByVal vData As String)
    SistGlobECF_sNumSerie = vData
End Property

Public Property Get gcolTicket() As Collection
    Set gcolTicket = SistGlobECF_colTicket
End Property

Public Property Set gcolTicket(ByVal vData As Collection)
    Set SistGlobECF_colTicket = vData
End Property

Public Property Get gcolTiposMeiosPagtos() As Collection
    Set gcolTiposMeiosPagtos = SistGlobECF_colTiposMeiosPagtos
End Property

Public Property Set gcolTiposMeiosPagtos(ByVal vData As Collection)
    Set SistGlobECF_colTiposMeiosPagtos = vData
End Property

Public Property Get gsCNPJ() As String
    gsCNPJ = SistGlobECF_sCNPJ
End Property

Public Property Let gsCNPJ(ByVal vData As String)
    SistGlobECF_sCNPJ = vData
End Property

Public Property Get gsEndereco() As String
    gsEndereco = SistGlobECF_sEndereco
End Property

Public Property Let gsEndereco(ByVal vData As String)
    SistGlobECF_sEndereco = vData
End Property

Public Property Get gsCidade() As String
    gsCidade = SistGlobECF_sCidade
End Property

Public Property Let gsCidade(ByVal vData As String)
    SistGlobECF_sCidade = vData
End Property

Public Property Get gsUF() As String
    gsUF = SistGlobECF_sUF
End Property

Public Property Let gsUF(ByVal vData As String)
    SistGlobECF_sUF = vData
End Property

Public Property Get gdtDataAnterior() As Date
    gdtDataAnterior = SistGlobECF_dtDataAnterior
End Property

Public Property Let gdtDataAnterior(ByVal vData As Date)
    SistGlobECF_dtDataAnterior = vData
End Property

Public Property Get gsInscricaoMunicipal() As String
    gsInscricaoMunicipal = SistGlobECF_sInscricaoMunicipal
End Property

Public Property Let gsInscricaoMunicipal(ByVal vData As String)
    SistGlobECF_sInscricaoMunicipal = vData
End Property

Public Property Get gsInscricaoEstadual() As String
    gsInscricaoEstadual = SistGlobECF_sInscricaoEstadual
End Property

Public Property Let gsInscricaoEstadual(ByVal vData As String)
    SistGlobECF_sInscricaoEstadual = vData
End Property

Public Property Get gsEnderecoComplemento() As String
    gsEnderecoComplemento = SistGlobECF_sEnderecoComplemento
End Property

Public Property Let gsEnderecoComplemento(ByVal vData As String)
    SistGlobECF_sEnderecoComplemento = vData
End Property

Public Property Get gcolCarne() As Collection
    Set gcolCarne = SistGlobECF_colCarne
End Property

Public Property Set gcolCarne(ByVal vData As Collection)
    Set SistGlobECF_colCarne = vData
End Property

Public Property Get gcolCliente() As Collection
    Set gcolCliente = SistGlobECF_colCliente
End Property

Public Property Set gcolCliente(ByVal vData As Collection)
    Set SistGlobECF_colCliente = vData
End Property

Public Property Set gobjClienteNome(ByVal vData As ClassClienteNome)
    Set SistGlobECF_objClienteNome = vData
End Property

Public Property Get gobjClienteNome() As ClassClienteNome
    Set gobjClienteNome = SistGlobECF_objClienteNome
End Property

Public Property Set gobjClienteCPF(ByVal vData As ClassClienteCPF)
    Set SistGlobECF_objClienteCPF = vData
End Property

Public Property Get gobjClienteCPF() As ClassClienteCPF
    Set gobjClienteCPF = SistGlobECF_objClienteCPF
End Property

Public Property Get gcolTotalizadores() As Collection
    Set gcolTotalizadores = SistGlobECF_colTotalizadores
End Property

Public Property Set gcolTotalizadores(ByVal vData As Collection)
    Set SistGlobECF_colTotalizadores = vData
End Property

Public Property Get glNumProxMovto() As Long
    glNumProxMovto = SistGlobECF_lNumProxMovto
End Property

Public Property Let glNumProxMovto(ByVal vData As Long)
    SistGlobECF_lNumProxMovto = vData
End Property

Public Property Get giOrcamentoECF() As Integer
    giOrcamentoECF = SistGlobECF_iOrcamentoECF
End Property

Public Property Let giOrcamentoECF(ByVal vData As Integer)
    SistGlobECF_iOrcamentoECF = vData
End Property

Public Property Get gsCodBarrasPorta() As String
    gsCodBarrasPorta = SistGlobECF_sCodBarrasPorta
End Property

Public Property Let gsCodBarrasPorta(ByVal vData As String)
    SistGlobECF_sCodBarrasPorta = vData
End Property

Public Property Get giCodModeloECF() As Integer
    giCodModeloECF = SistGlobECF_iCodModeloECF
End Property

Public Property Let giCodModeloECF(ByVal vData As Integer)
    SistGlobECF_iCodModeloECF = vData
End Property

Public Property Get gsDirMVTEF() As String
    gsDirMVTEF = SistGlobECF_sDirMVTEF
End Property

Public Property Let gsDirMVTEF(ByVal vData As String)
    SistGlobECF_sDirMVTEF = vData
End Property

Public Property Get giCodEmpresa() As Integer
    giCodEmpresa = SistGlobECF_iCodEmpresa
End Property

Public Property Let giCodEmpresa(ByVal vData As Integer)
    SistGlobECF_iCodEmpresa = vData
End Property

Public Property Get gsBalancaPorta() As String
    gsBalancaPorta = SistGlobECF_sBalancaPorta
End Property

Public Property Let gsBalancaPorta(ByVal vData As String)
    SistGlobECF_sBalancaPorta = vData
End Property

Public Property Get giBalancaModelo() As Integer
    giBalancaModelo = SistGlobECF_iBalancaModelo
End Property

Public Property Let giBalancaModelo(ByVal vData As Integer)
    SistGlobECF_iBalancaModelo = vData
End Property

Public Property Get giTipoTEF() As Integer
    giTipoTEF = SistGlobECF_iTipoTEF
End Property

Public Property Let giTipoTEF(ByVal vData As Integer)
    SistGlobECF_iTipoTEF = vData
End Property


Public Property Get Arquivo_Tef_Temp_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Temp_Prop = ARQUIVO_SITEF_TEMP
    Else
        Arquivo_Tef_Temp_Prop = ARQUIVO_TEF_TEMP
    End If
End Property

Public Property Get Arquivo_Tef_Req_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Req_Prop = ARQUIVO_SITEF_REQ
    Else
        Arquivo_Tef_Req_Prop = ARQUIVO_TEF_REQ
    End If
End Property

Public Property Get Arquivo_Tef_Resp1_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Resp1_Prop = ARQUIVO_SITEF_RESP1
    Else
        Arquivo_Tef_Resp1_Prop = ARQUIVO_TEF_RESP1
    End If
End Property

Public Property Get Arquivo_Tef_Resp2_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Resp2_Prop = ARQUIVO_SITEF_RESP2
    Else
        Arquivo_Tef_Resp2_Prop = ARQUIVO_TEF_RESP2
    End If
End Property

Public Property Get Arquivo_Tef_Imp_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Imp_Prop = ARQUIVO_SITEF_IMP
    Else
        Arquivo_Tef_Imp_Prop = ARQUIVO_TEF_IMP
    End If
End Property

Public Property Get Arquivo_Tef_Resp2_Backup_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Resp2_Backup_Prop = ARQUIVO_SITEF_RESP2_BACKUP
    Else
        Arquivo_Tef_Resp2_Backup_Prop = ARQUIVO_TEF_RESP2_BACKUP
    End If
End Property

Public Property Get Dir_Tef_Resp2_Backup_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Dir_Tef_Resp2_Backup_Prop = DIR_SITEF_RESP2_BACKUP
    Else
        Dir_Tef_Resp2_Backup_Prop = DIR_TEF_RESP2_BACKUP
    End If
End Property

Public Property Get Arquivo_Tef_Req_Backup_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Req_Backup_Prop = ARQUIVO_SITEF_REQ_BACKUP
    Else
        Arquivo_Tef_Req_Backup_Prop = ARQUIVO_TEF_REQ_BACKUP
    End If
End Property

Public Property Get Dir_Tef_Req_Backup_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Dir_Tef_Req_Backup_Prop = DIR_SITEF_REQ_BACKUP
    Else
        Dir_Tef_Req_Backup_Prop = DIR_TEF_REQ_BACKUP
    End If
End Property

Public Property Get Dir_Tef_Req_Backup1_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Dir_Tef_Req_Backup1_Prop = DIR_SITEF_REQ_BACKUP1
    Else
        Dir_Tef_Req_Backup1_Prop = ""
    End If
End Property

Public Property Get Arquivo_Tef_Req_Backup1_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Req_Backup1_Prop = ARQUIVO_SITEF_REQ_BACKUP1
    Else
        Arquivo_Tef_Req_Backup1_Prop = ""
    End If
End Property

Public Property Get Arquivo_Tef_Req_Backup2_Prop() As String
    If giTipoTEF = TIPOTEF_SITEF Then
        Arquivo_Tef_Req_Backup2_Prop = ARQUIVO_SITEF_REQ_BACKUP2
    Else
        Arquivo_Tef_Req_Backup2_Prop = ""
    End If
End Property

Public Property Get giPreVenda() As Integer
    giPreVenda = SistGlobECF_iPreVenda
End Property

Public Property Let giPreVenda(ByVal vData As Integer)
    SistGlobECF_iPreVenda = vData
End Property

Public Property Get giUsaImpressoraFiscal() As Integer
    giUsaImpressoraFiscal = SistGlobECF_iUsaImpressoraFiscal
End Property

Public Property Let giUsaImpressoraFiscal(ByVal vData As Integer)
    SistGlobECF_iUsaImpressoraFiscal = vData
End Property

Public Property Get giDAV() As Integer
    giDAV = SistGlobECF_iDAV
End Property

Public Property Let giDAV(ByVal vData As Integer)
    SistGlobECF_iDAV = vData
End Property

Public Property Set gcolECF(ByVal vData As Collection)
    Set SistGlobECF_colECF = vData
End Property

Public Property Get gcolECF() As Collection
    Set gcolECF = SistGlobECF_colECF
End Property

Public Property Get gsTipoECF() As String
    gsTipoECF = SistGlobECF_sTipoECF
End Property

Public Property Let gsTipoECF(ByVal vData As String)
    SistGlobECF_sTipoECF = vData
End Property

Public Property Get gsModeloECF() As String
    gsModeloECF = SistGlobECF_sModeloECF
End Property

Public Property Let gsModeloECF(ByVal vData As String)
    SistGlobECF_sModeloECF = vData
End Property

Public Property Get gsMarcaECF() As String
    gsMarcaECF = SistGlobECF_sMarcaECF
End Property

Public Property Let gsMarcaECF(ByVal vData As String)
    SistGlobECF_sMarcaECF = vData
End Property

Public Property Get gsChavePublica() As String
    gsChavePublica = SistGlobECF_sChavePublica
End Property

Public Property Let gsChavePublica(ByVal vData As String)
    SistGlobECF_sChavePublica = vData
End Property

Public Property Get gsChavePrivada() As String
    gsChavePrivada = SistGlobECF_sChavePrivada
End Property

Public Property Let gsChavePrivada(ByVal vData As String)
    SistGlobECF_sChavePrivada = vData
End Property

Public Property Get gdtDataAtualizacaoDadosCCC() As Date
    gdtDataAtualizacaoDadosCCC = SistGlobECF_dtDataAtualizacaoDadosCCC
End Property

Public Property Let gdtDataAtualizacaoDadosCCC(ByVal vData As Date)
    SistGlobECF_dtDataAtualizacaoDadosCCC = vData
End Property

Public Property Get gsCNPJDesenv() As String
    gsCNPJDesenv = SistGlobECF_sCNPJDesenv
End Property

Public Property Let gsCNPJDesenv(ByVal vData As String)
    SistGlobECF_sCNPJDesenv = vData
End Property

Public Property Get gsInscricaoMunicipalDesenv() As String
    gsInscricaoMunicipalDesenv = SistGlobECF_sInscricaoMunicipalDesenv
End Property

Public Property Let gsInscricaoMunicipalDesenv(ByVal vData As String)
    SistGlobECF_sInscricaoMunicipalDesenv = vData
End Property

Public Property Get gsInscricaoEstadualDesenv() As String
    gsInscricaoEstadualDesenv = SistGlobECF_sInscricaoEstadualDesenv
End Property

Public Property Let gsInscricaoEstadualDesenv(ByVal vData As String)
    SistGlobECF_sInscricaoEstadualDesenv = vData
End Property

Public Property Get gsNomeDesenv() As String
    gsNomeDesenv = SistGlobECF_sNomeDesenv
End Property

Public Property Let gsNomeDesenv(ByVal vData As String)
    SistGlobECF_sNomeDesenv = vData
End Property

Public Property Get gsNomePAFECF() As String
    gsNomePAFECF = SistGlobECF_sNomePAFECF
End Property

Public Property Let gsNomePAFECF(ByVal vData As String)
    SistGlobECF_sNomePAFECF = vData
End Property

Public Property Get gsVersaoPAFECF() As String
    gsVersaoPAFECF = SistGlobECF_sVersaoPAFECF
End Property

Public Property Let gsVersaoPAFECF(ByVal vData As String)
    SistGlobECF_sVersaoPAFECF = vData
End Property

Public Property Get gsMD5PAFECF() As String
    gsMD5PAFECF = SistGlobECF_sMD5PAFECF
End Property

Public Property Let gsMD5PAFECF(ByVal vData As String)
    SistGlobECF_sMD5PAFECF = vData
End Property

Public Property Get gsVersaoEspecPAFECF() As String
    gsVersaoEspecPAFECF = SistGlobECF_sVersaoEspecPAFECF
End Property

Public Property Let gsVersaoEspecPAFECF(ByVal vData As String)
    SistGlobECF_sVersaoEspecPAFECF = vData
End Property

Public Property Get gsLaudo() As String
    gsLaudo = SistGlobECF_sLaudo
End Property

Public Property Let gsLaudo(ByVal vData As String)
    SistGlobECF_sLaudo = vData
End Property

Public Property Get gsEnderecoDesenv() As String
    gsEnderecoDesenv = SistGlobECF_sEnderecoDesenv
End Property

Public Property Let gsEnderecoDesenv(ByVal vData As String)
    SistGlobECF_sEnderecoDesenv = vData
End Property

Public Property Get gsContatoDesenv() As String
    gsContatoDesenv = SistGlobECF_sContatoDesenv
End Property

Public Property Let gsContatoDesenv(ByVal vData As String)
    SistGlobECF_sContatoDesenv = vData
End Property

Public Property Get gsTelefoneDesenv() As String
    gsTelefoneDesenv = SistGlobECF_sTelefoneDesenv
End Property

Public Property Let gsTelefoneDesenv(ByVal vData As String)
    SistGlobECF_sTelefoneDesenv = vData
End Property

Public Property Get gsNomePrincArqExec() As String
    gsNomePrincArqExec = SistGlobECF_sNomePrincArqExec
End Property

Public Property Let gsNomePrincArqExec(ByVal vData As String)
    SistGlobECF_sNomePrincArqExec = vData
End Property

Public Property Get gsMD5PrincArqExec() As String
    gsMD5PrincArqExec = SistGlobECF_sMD5PrincArqExec
End Property

Public Property Let gsMD5PrincArqExec(ByVal vData As String)
    SistGlobECF_sMD5PrincArqExec = vData
End Property

Public Property Get giBloqueioReqXXII() As Integer
    giBloqueioReqXXII = SistGlobECF_iBloqueioReqXXII
End Property

Public Property Let giBloqueioReqXXII(ByVal vData As Integer)
    SistGlobECF_iBloqueioReqXXII = vData
End Property

Public Property Get gsMFAdicional() As String
    gsMFAdicional = SistGlobECF_sMFAdicional
End Property

Public Property Let gsMFAdicional(ByVal vData As String)
    SistGlobECF_sMFAdicional = vData
End Property

Public Property Get giDebug() As Integer
    giDebug = SistGlobECF_iDebug
End Property

Public Property Let giDebug(ByVal vData As Integer)
    SistGlobECF_iDebug = vData
End Property

Public Property Get glConexaoPAFECF() As Long
    glConexaoPAFECF = SistGlobECF_lConexaoPAFECF
End Property

Public Property Let glConexaoPAFECF(ByVal vData As Long)
    SistGlobECF_lConexaoPAFECF = vData
End Property

Public Property Get glConexaoOrcPAF() As Long
    glConexaoOrcPAF = SistGlobECF_lConexaoOrcPAF
End Property

Public Property Let glConexaoOrcPAF(ByVal vData As Long)
    SistGlobECF_lConexaoOrcPAF = vData
End Property

Public Property Get glTransacaoPAFECF() As Long
    glTransacaoPAFECF = SistGlobECF_lTransacaoPAFECF
End Property

Public Property Let glTransacaoPAFECF(ByVal vData As Long)
    SistGlobECF_lTransacaoPAFECF = vData
End Property

Public Property Get glTransacaoOrcPAFECF() As Long
    glTransacaoOrcPAFECF = SistGlobECF_lTransacaoOrcPAFECF
End Property

Public Property Let glTransacaoOrcPAFECF(ByVal vData As Long)
    SistGlobECF_lTransacaoOrcPAFECF = vData
End Property

Public Property Get giCodModeloECFConfig() As Integer
    giCodModeloECFConfig = SistGlobECF_iCodModeloECFConfig
End Property

Public Property Let giCodModeloECFConfig(ByVal vData As Integer)
    SistGlobECF_iCodModeloECFConfig = vData
End Property


Public Property Get gsCEP() As String
    gsCEP = SistGlobECF_sCEP
End Property

Public Property Let gsCEP(ByVal vData As String)
    SistGlobECF_sCEP = vData
End Property

Public Property Get gsCNAE() As String
    gsCNAE = SistGlobECF_sCNAE
End Property

Public Property Let gsCNAE(ByVal vData As String)
    SistGlobECF_sCNAE = vData
End Property

Public Property Get gdtDataHoraFimPapel() As Date
    gdtDataHoraFimPapel = SistGlobECF_dtDataHoraFimPapel
End Property

Public Property Let gdtDataHoraFimPapel(ByVal vData As Date)
    SistGlobECF_dtDataHoraFimPapel = vData
End Property

Public Property Get giIntervaloPapelAcabando() As Integer
    giIntervaloPapelAcabando = SistGlobECF_iIntervaloPapelAcabando

End Property

Public Property Let giIntervaloPapelAcabando(ByVal vData As Integer)
    SistGlobECF_iIntervaloPapelAcabando = vData
End Property

Public Property Get gsBalancaNome() As String
    gsBalancaNome = SistGlobECF_sBalancaNome
End Property

Public Property Let gsBalancaNome(ByVal vData As String)
    SistGlobECF_sBalancaNome = vData
End Property

Public Property Get giTotalizador_Sangria() As Integer
    giTotalizador_Sangria = SistGlobECF_iTotalizador_Sangria
End Property

Public Property Let giTotalizador_Sangria(ByVal vData As Integer)
    SistGlobECF_iTotalizador_Sangria = vData
End Property

Public Property Get giTotalizador_Suprimento() As Integer
    giTotalizador_Suprimento = SistGlobECF_iTotalizador_Suprimento
End Property

Public Property Let giTotalizador_Suprimento(ByVal vData As Integer)
    SistGlobECF_iTotalizador_Suprimento = vData
End Property

Public Property Get giTotalizador_Orcamento() As Integer
    giTotalizador_Orcamento = SistGlobECF_iTotalizador_Orcamento
End Property

Public Property Let giTotalizador_Orcamento(ByVal vData As Integer)
    SistGlobECF_iTotalizador_Orcamento = vData
End Property

Public Property Get giTEFImpressoraFolhaDupla() As Integer
    giTEFImpressoraFolhaDupla = SistGlobECF_iTEFImpressoraFolhaDupla
End Property

Public Property Let giTEFImpressoraFolhaDupla(ByVal vData As Integer)
    SistGlobECF_iTEFImpressoraFolhaDupla = vData
End Property

Public Property Get gdtDataSWBasico() As Date
    gdtDataSWBasico = SistGlobECF_dtDataSWBasico
End Property

Public Property Let gdtDataSWBasico(ByVal vData As Date)
    SistGlobECF_dtDataSWBasico = vData
End Property

Public Property Get gsHoraSWBasico() As String
    gsHoraSWBasico = SistGlobECF_sHoraSWBasico
End Property

Public Property Let gsHoraSWBasico(ByVal vData As String)
    SistGlobECF_sHoraSWBasico = vData
End Property


