VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPedidoDeVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Alterado p/Versão 2 por Luiz G.F.Nogueira em 05/03/02: colComissoes passou a ser implementada como uma Collection e não como uma coleção tipada

Implements ClassTributoDoc

Public objInfoUsu As Object
Public mvarobjTela As Object

Private mvarobjTributacaoDoc As ClassTributacaoDocAux
Private mvarcolItensPedido As New ColItemPedido

Private mvariFilialEmpresa As Integer
Private mvariFilialEmpresaFaturamento As Integer
Private mvarlCodigo As Long
Private mvarlCliente As Long
Private mvariFilial As Integer
Private mvariFilialEntrega As Integer
Private mvariTipoPedido As Integer
Private mvariCodTransportadora As Integer
Private mvariCondicaoPagto As Integer
Private mvardPercAcrescFinanceiro As Double
Private mvardtDataEmissao As Date
Private mvarsMensagemNota As String
'Private mvarsNaturezaOp As String
'Private mvariNaturezaOpManual As Integer
Private mvardValorTotal As Double
Private mvardValorFrete As Double
Private mvardValorDesconto As Double
Private mvardValorSeguro As Double
Private mvardValorOutrasDespesas As Double
Private mvarsPedidoCliente As String
Private mvariCanalVenda As Integer
Private mvariTabelaPreco As Integer
Private mvariProxSeqBloqueio As Integer
Private mvarcolComissoes As New Collection
Private mvarcolBloqueiosPV As New ColBloqueioPV
Private mvariFaturaIntegral As Integer
Private mvarcolParcelas As New ColParcelaPedidoVenda
Private mvariCobrancaAutomatica As Integer
Private mvariComissaoAutomatica As Integer
Private mvardtDataReferencia As Date
Private mvarobjTributacaoPV As New ClassTributacaoDoc
Private mvarlNumIntDoc As Long
Private mvardPesoBruto As Double
Private mvardPesoLiq As Double
Private mvarsPlaca As String
Private mvarsPlacaUF As String
Private mvarlVolumeQuant As Long
Private mvarlVolumeEspecie As Long
Private mvarlVolumeMarca As Long
Private mvardValorProdutos As Double 'local copy
Private mvarsVolumeNumero As String
Private mvariFreteRespons As Integer
Private mvarlNumIntOrcVenda As Long
Private mvardtDataEntrega As Date
Private mvariCodTranspRedesp As Integer
Private mvariDetPagFrete As Integer
Private mvardVolumeTotal As Double
Private mvariMoeda As Integer
Private mvardTaxaMoeda As Double
Private mvarsPedidoRepresentante As String
Private mvarobjProjetoInfo As Object
Private mvardtDataRefFluxo As Date
Private mvarcolItensPedidoSRV As New ColItemPedido
Private mvariStatus As Integer
Private mvarlNumIntSolicSRV As Long
Private mvariProgramacao As Integer
Private mvarsOBS As String
Private mvariAndamento As Integer
Private mvariServico As Integer
Private mvardValorDescontoTit As Double
Private mvarobjInfoAdic As New ClassInfoAdic
Private mvardValorDescontoItens As Double
Private mvardValorItens As Double
Private mvariParc As Integer
Private mvarlCodigoBase As Long

Public sEmitente As String
Public sUsuarioUltAlteracao As String
Public dtDataInclusao As Date
Public dtDataAlteracao As Date
Public dHoraInclusao As Double
Public dHoraAlteracao As Double

Public Property Let lCodigoBase(ByVal vData As Long)
    mvarlCodigoBase = vData
End Property

Public Property Get lCodigoBase() As Long
    lCodigoBase = mvarlCodigoBase
End Property

Public Property Let iParc(ByVal vData As Integer)
    mvariParc = vData
End Property

Public Property Get iParc() As Integer
    iParc = mvariParc
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

Public Property Let iServico(ByVal vData As Integer)
    mvariServico = vData
End Property

Public Property Get iServico() As Integer
    iServico = mvariServico
End Property

Public Property Let iAndamento(ByVal vData As Integer)
    mvariAndamento = vData
End Property

Public Property Get iAndamento() As Integer
    iAndamento = mvariAndamento
End Property

Public Property Let sOBS(ByVal vData As String)
    mvarsOBS = vData
End Property

Public Property Get sOBS() As String
    sOBS = mvarsOBS
End Property

Public Property Let iProgramacao(ByVal vData As Integer)
    mvariProgramacao = vData
End Property

Public Property Get iProgramacao() As Integer
    iProgramacao = mvariProgramacao
End Property

Public Property Let lNumIntSolicSRV(ByVal vData As Long)
    mvarlNumIntSolicSRV = vData
End Property

Public Property Get lNumIntSolicSRV() As Long
    lNumIntSolicSRV = mvarlNumIntSolicSRV
End Property

Public Property Let iStatus(ByVal vData As Integer)
    mvariStatus = vData
End Property

Public Property Get iStatus() As Integer
    iStatus = mvariStatus
End Property

Public Property Set colItensPedidoSRV(ByVal vData As ColItemPedido)
    Set mvarcolItensPedidoSRV = vData
End Property

Public Property Get colItensPedidoSRV() As ColItemPedido
    Set colItensPedidoSRV = mvarcolItensPedidoSRV
End Property

Public Property Let dtDataRefFluxo(ByVal vData As Date)
    mvardtDataRefFluxo = vData
End Property

Public Property Get dtDataRefFluxo() As Date
    dtDataRefFluxo = mvardtDataRefFluxo
End Property

Public Property Set objProjetoInfo(ByVal vData As Object)
    Set mvarobjProjetoInfo = vData
End Property

Public Property Get objProjetoInfo() As Object
    Set objProjetoInfo = mvarobjProjetoInfo
End Property

Public Property Let sPedidoRepresentante(ByVal vData As String)
    mvarsPedidoRepresentante = vData
End Property

Public Property Get sPedidoRepresentante() As String
    sPedidoRepresentante = mvarsPedidoRepresentante
End Property

Public Property Set objTela(ByVal vData As Object)
    Set mvarobjTela = vData
End Property

Public Property Get objTela() As Object
    Set objTela = mvarobjTela
End Property

Public Property Let dVolumeTotal(ByVal vData As Double)
    mvardVolumeTotal = vData
End Property

Public Property Get dVolumeTotal() As Double
    dVolumeTotal = mvardVolumeTotal
End Property

Public Property Let iMoeda(ByVal vData As Integer)
    mvariMoeda = vData
End Property

Public Property Get iMoeda() As Integer
    iMoeda = mvariMoeda
End Property

Public Property Let dTaxaMoeda(ByVal vData As Double)
    mvardTaxaMoeda = vData
End Property

Public Property Get dTaxaMoeda() As Double
    dTaxaMoeda = mvardTaxaMoeda
End Property

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

Public Property Let dtDataEntrega(ByVal vData As Date)
    mvardtDataEntrega = vData
End Property

Public Property Get dtDataEntrega() As Date
    dtDataEntrega = mvardtDataEntrega
End Property

Public Property Let iNaturezaOpManual(ByVal vData As Integer)
    objTributacao.iNaturezaOpManual = vData
End Property

Public Property Get iNaturezaOpManual() As Integer
    iNaturezaOpManual = objTributacao.iNaturezaOpManual
End Property

Public Property Let iFreteRespons(ByVal vData As Integer)
    mvariFreteRespons = vData
End Property

Public Property Get iFreteRespons() As Integer
    iFreteRespons = mvariFreteRespons
End Property

Public Property Let dValorProdutos(ByVal vData As Double)
    mvardValorProdutos = vData
End Property

Public Property Get dValorProdutos() As Double
    dValorProdutos = mvardValorProdutos
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

Public Property Let lNumIntDoc(ByVal vData As Long)
    mvarlNumIntDoc = vData
End Property

Public Property Get lNumIntDoc() As Long
    lNumIntDoc = mvarlNumIntDoc
End Property

Public Property Set objTributacaoPV(ByVal vData As Object)
    Set mvarobjTributacaoPV = vData
End Property

Public Property Get objTributacaoPV() As ClassTributacaoDoc
    Set objTributacaoPV = mvarobjTributacaoPV
End Property

Public Property Let dtDataReferencia(ByVal vData As Date)
    mvardtDataReferencia = vData
End Property

Public Property Get dtDataReferencia() As Date
    dtDataReferencia = mvardtDataReferencia
End Property

Public Property Let iComissaoAutomatica(ByVal vData As Integer)
    mvariComissaoAutomatica = vData
End Property

Public Property Get iComissaoAutomatica() As Integer
    iComissaoAutomatica = mvariComissaoAutomatica
End Property

Public Property Let iCobrancaAutomatica(ByVal vData As Integer)
    mvariCobrancaAutomatica = vData
End Property

Public Property Get iCobrancaAutomatica() As Integer
    iCobrancaAutomatica = mvariCobrancaAutomatica
End Property

Public Property Set colParcelas(ByVal vData As Object)
    Set mvarcolParcelas = vData
End Property

Public Property Get colParcelas() As ColParcelaPedidoVenda
    Set colParcelas = mvarcolParcelas
End Property

Public Property Let iFaturaIntegral(ByVal vData As Integer)
    mvariFaturaIntegral = vData
End Property

Public Property Get iFaturaIntegral() As Integer
    iFaturaIntegral = mvariFaturaIntegral
End Property

Public Property Let dValorTotal(ByVal vData As Double)
    mvardValorTotal = vData
End Property

Public Property Get dValorTotal() As Double
    dValorTotal = mvardValorTotal
End Property

Public Property Set colBloqueiosPV(ByVal vData As Object)
    Set mvarcolBloqueiosPV = vData
End Property

Public Property Get colBloqueiosPV() As ColBloqueioPV
    Set colBloqueiosPV = mvarcolBloqueiosPV
End Property

Public Property Set colComissoes(ByVal vData As Object)
    Set mvarcolComissoes = vData
End Property

Public Property Get colComissoes() As Collection
    Set colComissoes = mvarcolComissoes
End Property

Public Property Let iProxSeqBloqueio(ByVal vData As Integer)
    mvariProxSeqBloqueio = vData
End Property

Public Property Get iProxSeqBloqueio() As Integer
    iProxSeqBloqueio = mvariProxSeqBloqueio
End Property

Public Property Let iTabelaPreco(ByVal vData As Integer)
    mvariTabelaPreco = vData
End Property

Public Property Get iTabelaPreco() As Integer
    iTabelaPreco = mvariTabelaPreco
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

Public Property Let sNaturezaOp(ByVal vData As String)
    objTributacao.sNaturezaOp = vData
End Property

Public Property Get sNaturezaOp() As String
    sNaturezaOp = objTributacao.sNaturezaOp
End Property

Public Property Let sMensagemNota(ByVal vData As String)
    mvarsMensagemNota = vData
End Property

Public Property Get sMensagemNota() As String
    sMensagemNota = mvarsMensagemNota
End Property

Public Property Let dtDataEmissao(ByVal vData As Date)
    mvardtDataEmissao = vData
End Property

Public Property Get dtDataEmissao() As Date
    dtDataEmissao = mvardtDataEmissao
End Property

Public Property Let dPercAcrescFinanceiro(ByVal vData As Double)
    mvardPercAcrescFinanceiro = vData
End Property

Public Property Get dPercAcrescFinanceiro() As Double
    dPercAcrescFinanceiro = mvardPercAcrescFinanceiro
End Property

Public Property Let iCondicaoPagto(ByVal vData As Integer)
    mvariCondicaoPagto = vData
End Property

Public Property Get iCondicaoPagto() As Integer
    iCondicaoPagto = mvariCondicaoPagto
End Property

Public Property Let iCodTransportadora(ByVal vData As Integer)
    mvariCodTransportadora = vData
End Property

Public Property Get iCodTransportadora() As Integer
    iCodTransportadora = mvariCodTransportadora
End Property

Public Property Let iTipoPedido(ByVal vData As Integer)
    mvariTipoPedido = vData
End Property

Public Property Get iTipoPedido() As Integer
    iTipoPedido = mvariTipoPedido
End Property

Public Property Let iFilialEntrega(ByVal vData As Integer)
    mvariFilialEntrega = vData
End Property

Public Property Get iFilialEntrega() As Integer
    iFilialEntrega = mvariFilialEntrega
End Property

Public Property Let iFilial(ByVal vData As Integer)
    mvariFilial = vData
End Property

Public Property Get iFilial() As Integer
    iFilial = mvariFilial
End Property

Public Property Let lCliente(ByVal vData As Long)
    mvarlCliente = vData
End Property

Public Property Get lCliente() As Long
    lCliente = mvarlCliente
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

Public Property Let iFilialEmpresaFaturamento(ByVal vData As Integer)
    mvariFilialEmpresaFaturamento = vData
End Property

Public Property Get iFilialEmpresaFaturamento() As Integer
    iFilialEmpresaFaturamento = mvariFilialEmpresaFaturamento
End Property

'' coloquei colItensPedido como readonly p/forcar a utilizacao da funcao IncluirItem
''Public Property Set colItensPedido(ByVal vData As Object)
''    Set mvarcolItensPedido = vData
''End Property

Public Property Get colItensPedido() As ColItemPedido
    Set colItensPedido = mvarcolItensPedido
End Property
'Incluida por Leo em 25/04/02
Public Property Let lNumIntOrcVenda(ByVal vData As Long)
    mvarlNumIntOrcVenda = vData
End Property
'Incluida por Leo em 25/04/02
Public Property Get lNumIntOrcVenda() As Long
    lNumIntOrcVenda = mvarlNumIntOrcVenda
End Property


Private Sub Class_Initialize()

    dtDataInclusao = DATA_NULA
    dtDataAlteracao = DATA_NULA

End Sub

Private Function ClassTributoDoc_Ativar() As Long
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemPedido As ClassItemPedido
On Error GoTo Erro_ClassTributoDoc_Ativar

    lErro = gobjTributacao.Tributacao_Doc_Iniciar(Me)
    If lErro <> SUCESSO Then Error 27371
    
    Set mvarobjTributacaoDoc = New ClassTributacaoDocAux
    
    'ativar o pedido em si
    If mvarobjTributacaoDoc.Ativar(Me) <> SUCESSO Then
        Set mvarobjTributacaoDoc = Nothing
        Error 27366
    End If
    
    'ativar os seus itens
    For Each objItemPedido In colItensPedido
        Set objTributoDocItem = objItemPedido
        lErro = objTributoDocItem.Ativar(Me)
        If lErro <> SUCESSO Then Error 27374
    Next
        
    'ativar os seus itens
    For Each objItemPedido In colItensPedidoSRV
        Set objTributoDocItem = objItemPedido
        lErro = objTributoDocItem.Ativar(Me)
        If lErro <> SUCESSO Then Error 27374
    Next
        
    ClassTributoDoc_Ativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Ativar:

    ClassTributoDoc_Ativar = Err

    Select Case Err

        Case 27366, 27371, 27374
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152492)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_Desativar() As Long
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemPedido As ClassItemPedido
On Error GoTo Erro_ClassTributoDoc_Desativar
    
    'desativar os seus itens
    For Each objItemPedido In colItensPedido
        Set objTributoDocItem = objItemPedido
        lErro = objTributoDocItem.Desativar()
        If lErro <> SUCESSO Then Error 27685
    Next
    
    'desativar os seus itens
    For Each objItemPedido In colItensPedidoSRV
        Set objTributoDocItem = objItemPedido
        lErro = objTributoDocItem.Desativar()
        If lErro <> SUCESSO Then Error 27685
    Next
    
    'desativar o pedido em si
    If Not (mvarobjTributacaoDoc Is Nothing) Then
        lErro = mvarobjTributacaoDoc.Desativar()
        If lErro <> SUCESSO Then Error 27684
        Set mvarobjTributacaoDoc = Nothing
    End If
     
    lErro = gobjTributacao.Tributacao_Doc_Terminar(Me)
    If lErro <> SUCESSO Then Error 27372
    
    ClassTributoDoc_Desativar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_Desativar:

    ClassTributoDoc_Desativar = Err

    Select Case Err

        Case 27372, 27685, 27684
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152493)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_GuardarVarTela(sIdentificador As String, vValor As Variant) As Long
Dim lErro As Long, objVariavelCalculo As New ClassVariavelCalculo
On Error GoTo Erro_ClassTributoDoc_GuardarVarTela

    objVariavelCalculo.sIdentificador = sIdentificador
    objVariavelCalculo.vValor = vValor
    objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
    
    ClassTributoDoc_GuardarVarTela = ClassTributoDoc_GuardarVar(objVariavelCalculo)

    Exit Function

Erro_ClassTributoDoc_GuardarVarTela:

    ClassTributoDoc_GuardarVarTela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152494)

    End Select

    Exit Function

End Function

Private Property Get ClassTributoDoc_objTributacaoDoc() As GlobaisTRB.ClassTributacaoDocAux
    Set ClassTributoDoc_objTributacaoDoc = mvarobjTributacaoDoc
End Property

Private Function ClassTributoDoc_ObterColDocCompl(colDocCompl As Collection) As Long
'Dim objPVTribCompl As ClassPVTribCompl

'    Set colDocCompl = New Collection
'
'    Set objPVTribCompl = New ClassPVTribCompl
'    Call objPVTribCompl.Iniciar(Me, mvarobjTributacaoPV.objTributacaoFrete)
'    Call colDocCompl.Add(objPVTribCompl)
'
'    Set objPVTribCompl = New ClassPVTribCompl
'    Call objPVTribCompl.Iniciar(Me, mvarobjTributacaoPV.objTributacaoSeguro)
'    Call colDocCompl.Add(objPVTribCompl)
'
'    Set objPVTribCompl = New ClassPVTribCompl
'    Call objPVTribCompl.Iniciar(Me, mvarobjTributacaoPV.objTributacaoDesconto)
'    Call colDocCompl.Add(objPVTribCompl)
'
'    Set objPVTribCompl = New ClassPVTribCompl
'    Call objPVTribCompl.Iniciar(Me, mvarobjTributacaoPV.objTributacaoOutras)
'    Call colDocCompl.Add(objPVTribCompl)
    
    ClassTributoDoc_ObterColDocCompl = SUCESSO
    
End Function

Private Function ClassTributoDoc_ObterColDocItens(colDocItens As Collection) As Long
'retorna colecao de itens que implementam interface como TaxDocItemExemplo
Dim objItemPedido As ClassItemPedido

   Set colDocItens = New Collection
   
   For Each objItemPedido In mvarcolItensPedido
        Call colDocItens.Add(objItemPedido)
   Next
   
   For Each objItemPedido In mvarcolItensPedidoSRV
        Call colDocItens.Add(objItemPedido)
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
    If lErro <> SUCESSO Then Error 27401
    
    ClassTributoDoc_ObterNatOpPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterNatOpPadrao:

    ClassTributoDoc_ObterNatOpPadrao = Err

    Select Case Err

        Case 27401
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152495)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterTribPadrao(iTipoTrib As Integer) As Long
Dim lErro As Long
On Error GoTo Erro_ClassTributoDoc_ObterTribPadrao

    lErro = mvarobjTributacaoDoc.ObterTribPadrao(iTipoTrib)
    If lErro <> SUCESSO Then Error 27405
    
    ClassTributoDoc_ObterTribPadrao = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterTribPadrao:

    ClassTributoDoc_ObterTribPadrao = Err

    Select Case Err

        Case 27405
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152496)

    End Select

    Exit Function

End Function

Private Function ClassTributoDoc_ObterVar(sIdentificador As String, objVariavelCalculo As ClassVariavelCalculo) As Long
'retorna o valor solicitado se este estiver "disponivel" em seu "contexto"
'objVariavelCalculo.sIdentificador tem que vir preenchida
Dim lErro As Long, lComando As Long, iFilial As Integer

On Error GoTo Erro_ClassTributoDoc_ObterVar
    
    If objVariavelCalculo Is Nothing Then
        Set objVariavelCalculo = New ClassVariavelCalculo
    End If
    
    objVariavelCalculo.sIdentificador = sIdentificador
    
    Select Case sIdentificador
        
        Case "SERIE"
            objVariavelCalculo.vValor = ""
            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO

        Case "DATA_EMISSAO"
            objVariavelCalculo.vValor = mvardtDataEmissao
            If mvardtDataEmissao <> DATA_NULA Then
                objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
            Else
                objVariavelCalculo.iStatus = VAR_PREENCH_VAZIO
            End If
        
        Case "FILIAL_FATURAMENTO"
            
            If (mvariFilialEmpresa = 51 Or (mvariFilialEmpresa = 0 And giFilialEmpresa = 51)) And left(UCase(gsNomeEmpresa), 5) = "FEMAR" And mvarlCliente <> 0 And mvariFilialEntrega <> 0 Then
            
                lComando = Comando_Abrir()
                If lComando = 0 Then Error 30000
                
                lErro = Comando_Executar(lComando, "SELECT Filial FROM FilialClienteCategorias WHERE Categoria = 'Humaita-AM' AND Item = 'Humaita-AM' AND Cliente = ? AND Filial = ?", iFilial, mvarlCliente, mvariFilialEntrega)
                If lErro <> AD_SQL_SUCESSO Then Error 30001
            
                lErro = Comando_BuscarProximo(lComando)
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then Error 30002
            
                If lErro = AD_SQL_SEM_DADOS Then
                    objVariavelCalculo.vValor = mvariFilialEmpresaFaturamento
                Else
                    objVariavelCalculo.vValor = 52 'Humaitá
                End If
                
                Call Comando_Fechar(lComando)
            
            Else
                
                objVariavelCalculo.vValor = mvariFilialEmpresaFaturamento
            
            End If
            
            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "NAT_OPERACAO"
'            objVariavelCalculo.vValor = mvarsNaturezaOp
'            objVariavelCalculo.iStatus = mvariNaturezaOpManual
'
'        Case "TIPO_TRIB"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.iTipoTributacao
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iTipoTributacaoManual
        
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
            objVariavelCalculo.vValor = mvariFilialEntrega 'a outra filial do pedido nao interessa p/efeito de tributacao
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
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSBaseManual
'
'        Case "ICMS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSValorManual
'
'        Case "ICMS_SUBST_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSSubstBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSSubstBaseManual
'
'        Case "ICMS_SUBST_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSSubstValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iICMSSubstValorManual
'
'        Case "ICMS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dICMSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'para IPI
'        Case "IPI_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIPIBase
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIPIBaseManual
'
'        Case "IPI_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIPIValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIPIValorManual
'
'        Case "IPI_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIPICredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.iISSIncluso
'            'objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSInclusoManual
'
'        Case "ISS_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "ISS_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSAliquotaManual
'
'        Case "ISS_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSValorManual
'
'        'para IR
'        Case "IR_BASE"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIRRFBase
'            objVariavelCalculo.iStatus = VAR_PREENCH_AUTOMATICO
'
'        Case "IR_ALIQUOTA"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIRRFAliquota
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIRRFAliquotaManual
'
'        Case "IR_VALOR"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dIRRFValor
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iIRRFValorManual
'
'        Case "PIS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dPISCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "COFINS_VALOR_CREDITO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dCOFINSCredito
'            objVariavelCalculo.iStatus = VAR_PREENCH_MANUAL
'
'        Case "PIS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dPISRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iPISRetidoManual
'
'        Case "ISS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dISSRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iISSRetidoManual
'
'        Case "COFINS_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dCOFINSRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iCOFINSRetidoManual
'
'        Case "CSLL_RETIDO"
'            objVariavelCalculo.vValor = mvarobjTributacaoPV.dCSLLRetido
'            objVariavelCalculo.iStatus = mvarobjTributacaoPV.iCSLLRetidoManual
            
        Case Else
'        Case "FILIAL_CLIENTE_OBJ"
'        Case "FILIAL_CLIENTE_INSC_EST"
'        Case "CLIENTE_ENDERECO_OBJ"
'        Case "CLIENTE_UF" '(EX p/exterior)
            
            lErro = objTributacao.ObterVar_Padrao(sIdentificador, objVariavelCalculo)
            If lErro <> SUCESSO Then
                'delegar p/tributacaoDoc
                lErro = mvarobjTributacaoDoc.ObterVar(sIdentificador, objVariavelCalculo)
                If lErro <> SUCESSO Then Error 27368
            End If
            
    End Select
    
    ClassTributoDoc_ObterVar = SUCESSO

    Exit Function

Erro_ClassTributoDoc_ObterVar:

    ClassTributoDoc_ObterVar = Err

    Select Case Err

        Case 27368
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152497)

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
'            mvarobjTributacaoPV.iTipoTributacao = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iTipoTributacaoManual = objVariavelCalculo.iStatus
        
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
'            mvarobjTributacaoPV.dICMSBase = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR"
'            mvarobjTributacaoPV.dICMSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_BASE"
'            mvarobjTributacaoPV.dICMSSubstBase = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSSubstBaseManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_SUBST_VALOR"
'            mvarobjTributacaoPV.dICMSSubstValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iICMSSubstValorManual = objVariavelCalculo.iStatus
'
'        Case "ICMS_VALOR_CREDITO"
'             mvarobjTributacaoPV.dICMSCredito = objVariavelCalculo.vValor
'
'        'para IPI
'        Case "IPI_BASE"
'            mvarobjTributacaoPV.dIPIBase = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIPIBaseManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR"
'            mvarobjTributacaoPV.dIPIValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIPIValorManual = objVariavelCalculo.iStatus
'
'        Case "IPI_VALOR_CREDITO"
'            mvarobjTributacaoPV.dIPICredito = objVariavelCalculo.vValor
'
'        'para ISS
'        Case "ISS_INCLUSO"
'            mvarobjTributacaoPV.iISSIncluso = objVariavelCalculo.vValor
'            'mvarobjTributacaoPV.iISSInclusoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_BASE"
'            mvarobjTributacaoPV.dISSBase = objVariavelCalculo.vValor
'
'        Case "ISS_ALIQUOTA"
'            mvarobjTributacaoPV.dISSAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iISSAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "ISS_VALOR"
'            mvarobjTributacaoPV.dISSValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iISSValorManual = objVariavelCalculo.iStatus
'
'        'para IR
'        Case "IR_BASE"
'            mvarobjTributacaoPV.dIRRFBase = objVariavelCalculo.vValor
'
'        Case "IR_ALIQUOTA"
'            mvarobjTributacaoPV.dIRRFAliquota = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIRRFAliquotaManual = objVariavelCalculo.iStatus
'
'        Case "IR_VALOR"
'            mvarobjTributacaoPV.dIRRFValor = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iIRRFValorManual = objVariavelCalculo.iStatus
'
'        Case "PIS_VALOR_CREDITO"
'            mvarobjTributacaoPV.dPISRetido = objVariavelCalculo.vValor
'
'        Case "COFINS_VALOR_CREDITO"
'            mvarobjTributacaoPV.dCOFINSRetido = objVariavelCalculo.vValor
'
'        Case "PIS_RETIDO"
'            mvarobjTributacaoPV.dPISRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iPISRetidoManual = objVariavelCalculo.iStatus
'
'        Case "ISS_RETIDO"
'            mvarobjTributacaoPV.dISSRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iISSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "COFINS_RETIDO"
'            mvarobjTributacaoPV.dCOFINSRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iCOFINSRetidoManual = objVariavelCalculo.iStatus
'
'        Case "CSLL_RETIDO"
'            mvarobjTributacaoPV.dCSLLRetido = objVariavelCalculo.vValor
'            mvarobjTributacaoPV.iCSLLRetidoManual = objVariavelCalculo.iStatus
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152498)

    End Select

    Exit Function

End Function

Public Function IncluirItem(sProduto As String, dQuantidade As Double, dPrecoUnitario As Double, dPrecoTotal As Double, dValorDesconto As Double, dtDataEntrega As Date, sProdutoDescricao As String, dValorAbatComissao As Double, dQuantCancelada As Double, dQuantReservada As Double, colReserva As colReserva, sProdutoNomeReduzido As String, sUnidadeMed As String, sLote As String, sUMEstoque As String, iClasseUM As Integer, dQuantFaturada As Double, dQuantAFaturar As Double, sDescricao As String, iStatus As Integer, iControleEstoque As Integer, dQuantOP As Double, dQuantSC As Double, dPercDesc1 As Double, iTipoDesc1 As Integer, dPercDesc2 As Double, iTipoDesc2 As Integer, dPercDesc3 As Double, iTipoDesc3 As Integer, lNumIntDoc As Long, objTributacaoItemPV As ClassTributacaoDocItem) As ClassItemPedido
'associa um novo item ao pedido de venda
Dim lErro As Long
'Dim objTributoDocItem As ClassTributoDocItem
Dim objItemPedido As ClassItemPedido
On Error GoTo Erro_IncluirItem

    Set objItemPedido = colItensPedido.Add(mvariFilialEmpresa, mvarlCodigo, 0, sProduto, dQuantidade, dPrecoUnitario, dPrecoTotal, dValorDesconto, dtDataEntrega, sProdutoDescricao, dValorAbatComissao, dQuantCancelada, dQuantReservada, colReserva, sProdutoNomeReduzido, sUnidadeMed, sLote, sUMEstoque, iClasseUM, dQuantFaturada, dQuantAFaturar, sDescricao, iStatus, iControleEstoque, dQuantOP, dQuantSC, dPercDesc1, iTipoDesc1, dPercDesc2, iTipoDesc2, dPercDesc3, iTipoDesc3, lNumIntDoc)
        
    Set objItemPedido.objTributacaoItemPV = objTributacaoItemPV
    
'    Set objTributoDocItem = objItemPedido
'    lErro = objTributoDocItem.Ativar(Me)
'    If lErro <> SUCESSO Then Error 27374
    
    Set IncluirItem = objItemPedido

    Exit Function

Erro_IncluirItem:

    Set IncluirItem = Nothing

    Select Case Err

'        Case 27374
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152499)

    End Select

    Exit Function

End Function

Public Function RemoverItem(vntIndexKey As Variant) As Long
'exclui um item do PV identificando-o como na colecao de itens
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemPedido As ClassItemPedido
On Error GoTo Erro_RemoverItem

    Set objItemPedido = colItensPedido.Item(vntIndexKey)
    
    Set objTributoDocItem = objItemPedido
    lErro = objTributoDocItem.Desativar()
    If lErro <> SUCESSO Then Error 27375
        
    Call colItensPedido.Remove(vntIndexKey)
    
    RemoverItem = SUCESSO

    Exit Function

Erro_RemoverItem:

    RemoverItem = Err

    Select Case Err

        Case 27375
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 152500)

    End Select

    Exit Function

End Function

Public Function RemoverItemSRV(vntIndexKey As Variant) As Long
'exclui um item do PV identificando-o como na colecao de itens
Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objItemPedido As ClassItemPedido

On Error GoTo Erro_RemoverItemSRV

    Set objItemPedido = colItensPedidoSRV.Item(vntIndexKey)
    
    Set objTributoDocItem = objItemPedido
    lErro = objTributoDocItem.Desativar()
    If lErro <> SUCESSO Then gError 188910
        
    Call colItensPedidoSRV.Remove(vntIndexKey)
    
    RemoverItemSRV = SUCESSO

    Exit Function

Erro_RemoverItemSRV:

    RemoverItemSRV = gErr

    Select Case gErr

        Case 188910
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 188911)

    End Select

    Exit Function

End Function


Public Function IncluirItemSRV(sProduto As String, dQuantidade As Double, dPrecoUnitario As Double, dPrecoTotal As Double, dValorDesconto As Double, dtDataEntrega As Date, sProdutoDescricao As String, dValorAbatComissao As Double, dQuantCancelada As Double, dQuantReservada As Double, colReserva As colReserva, sProdutoNomeReduzido As String, sUnidadeMed As String, sLote As String, sUMEstoque As String, iClasseUM As Integer, dQuantFaturada As Double, dQuantAFaturar As Double, sDescricao As String, iStatus As Integer, iControleEstoque As Integer, dQuantOP As Double, dQuantSC As Double, dPercDesc1 As Double, iTipoDesc1 As Integer, dPercDesc2 As Double, iTipoDesc2 As Integer, dPercDesc3 As Double, iTipoDesc3 As Integer, lNumIntDoc As Long, objTributacaoItemPV As ClassTributacaoDocItem) As ClassItemPedido
'associa um novo item ao pedido de venda
Dim lErro As Long
'Dim objTributoDocItem As ClassTributoDocItem
Dim objItemPedido As ClassItemPedido

On Error GoTo Erro_IncluirItemSRV

    Set objItemPedido = colItensPedidoSRV.Add(mvariFilialEmpresa, mvarlCodigo, 0, sProduto, dQuantidade, dPrecoUnitario, dPrecoTotal, dValorDesconto, dtDataEntrega, sProdutoDescricao, dValorAbatComissao, dQuantCancelada, dQuantReservada, colReserva, sProdutoNomeReduzido, sUnidadeMed, sLote, sUMEstoque, iClasseUM, dQuantFaturada, dQuantAFaturar, sDescricao, iStatus, iControleEstoque, dQuantOP, dQuantSC, dPercDesc1, iTipoDesc1, dPercDesc2, iTipoDesc2, dPercDesc3, iTipoDesc3, lNumIntDoc)
        
    Set objItemPedido.objTributacaoItemPV = objTributacaoItemPV
    
'    Set objTributoDocItem = objItemPedido
'    lErro = objTributoDocItem.Ativar(Me)
'    If lErro <> SUCESSO Then Error 27374
    
    Set IncluirItemSRV = objItemPedido

    Exit Function

Erro_IncluirItemSRV:

    Set IncluirItemSRV = Nothing

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 191181)

    End Select

    Exit Function

End Function

Public Property Set colItens(ByVal vData As Object)
    Dim vValor As ClassItemPedido
    Dim colItensAux As New ColItemPedido
    For Each vValor In vData
        colItensAux.AddObj vValor
    Next
    Set mvarcolItensPedido = colItensAux
End Property

Public Property Get colItens() As Collection
    Dim vValor As Variant
    Dim colItensAux As New Collection
    For Each vValor In colItensPedido
        colItensAux.Add vValor
    Next
    Set colItens = colItensAux
End Property

Public Property Set objTributacao(ByVal vData As Object)
    Set mvarobjTributacaoPV = vData
End Property

Public Property Get objTributacao() As ClassTributacaoDoc
    Set objTributacao = mvarobjTributacaoPV
End Property

Public Property Set colItensSRV(ByVal vData As Object)
    Dim vValor As ClassItemPedido
    Dim colItensAux As New ColItemPedido
    For Each vValor In vData
        colItensAux.AddObj vValor
    Next
    Set mvarcolItensPedidoSRV = colItensSRV
End Property

Public Property Get colItensSRV() As Collection
    Dim vValor As Variant
    Dim colItensAux As New Collection
    For Each vValor In colItensPedidoSRV
        colItensAux.Add vValor
    Next
    Set colItensSRV = colItensAux
End Property

Private Function ClassTributoDoc_Alterado() As Long
    ClassTributoDoc_Alterado = mvarobjTributacaoDoc.Alterado
End Function

Public Sub AnaliseVenda_Preenche(ByVal objAnaliseVenda As ClassAnaliseVendaInfo)
Dim objItemAnalise As ClassItemAnaliseVenda
Dim objItemPV As ClassItemPedido, objComissao As ClassComissaoPedVendas

    With objAnaliseVenda
        .dtDataEmissao = mvardtDataEmissao
        .iFilial = mvariFilial
        .iFilialEmpresa = mvariFilialEmpresa
        .iTabelaPreco = mvariTabelaPreco
        .lCliente = mvarlCliente
'       .iVendedor = mvariVendedor
        .iMoeda = mvariMoeda
    End With

    For Each objComissao In mvarcolComissoes
        If objComissao.iIndireta = 0 Then
            objAnaliseVenda.iVendedor = objComissao.iCodVendedor
            Exit For
        End If
    Next
    
    Set objAnaliseVenda.colItens = New Collection
    
    For Each objItemPV In mvarcolItensPedido
    
        If objItemPV.objInfoAdicDocItem.iIncluiValorTotal <> 0 Then
    
            Set objItemAnalise = New ClassItemAnaliseVenda
        
            With objItemAnalise
            
                '.dPercDesc = objItemPV.dPercDesc
                .dPrecoTotal = objItemPV.dPrecoTotal
                '.dPrecoTotalB = objItemPV.dPrecoTotalB
                .dPrecoUnitario = objItemPV.dPrecoUnitario
                .dPrecoBase = objItemPV.dPrecoUnitario
                .dQuantidade = objItemPV.dQuantidade
                .dValorDesconto = objItemPV.dValorDesconto
                .sProduto = objItemPV.sProduto
                .sDescricao = objItemPV.sDescricao
                .sUnidadeMed = objItemPV.sUnidadeMed
            
            End With
        
            objAnaliseVenda.colItens.Add objItemAnalise
    
        End If
    
    Next
    
End Sub

