VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassTribTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public bSGECancelDummy As Boolean 'auxiliar na troca de lostfocus por validate

'classe p/encapsular tratamento do tab de tributacao
'??? no futuro incluir no grid de itens colunas p/aliquotas de IPI e ICMS

Private mvariEmitenteNF As Integer
Private mvariNFEntrada As Integer
Private mvariLivroFiscal As Integer
'Private sCidadeIBGECodigo As String
Private sCidadeIBGECodigoEmp As String

Public mvarobjForm As Object 'publico p/poder ser trocado p/tela que trata varios tipos. Jones 13/01/2000

Private mvarobjTabTrib As Object

Private WithEvents mvarobjEventoNaturezaInterna As AdmEvento
Attribute mvarobjEventoNaturezaInterna.VB_VarHelpID = -1
Private WithEvents mvarobjEventoTiposDeTributacao As AdmEvento
Attribute mvarobjEventoTiposDeTributacao.VB_VarHelpID = -1
Private WithEvents mvarobjEventoCidadeIBGE As AdmEvento
Attribute mvarobjEventoCidadeIBGE.VB_VarHelpID = -1

'-------------------------------------NFe 3.10-------------------------------------
Private WithEvents mvarobjEventoISSCodServ As AdmEvento
Attribute mvarobjEventoISSCodServ.VB_VarHelpID = -1
Private WithEvents mvarobjEventoISSMunicIncidImp As AdmEvento
Attribute mvarobjEventoISSMunicIncidImp.VB_VarHelpID = -1
Private WithEvents mvarobjEventoISSCodPais As AdmEvento
Attribute mvarobjEventoISSCodPais.VB_VarHelpID = -1
'-------------------------------------NFe 3.10-------------------------------------

Private WithEvents mvarobjEventoIPICodEnq As AdmEvento
Attribute mvarobjEventoIPICodEnq.VB_VarHelpID = -1

Private gobjVersao As ClassVersaoNFe

'??? ver p/setar sSiglaTipoNFiscal (se trocar, tem que resetar a tributacao ???)

Public mvarobjNFTributacao As ClassTribNFAux 'estrutura mantida para auxiliar a manutencao de informacoes p/tributacao
Public mvarobjPV As ClassPedidoDeVenda
Public mvarobjOV As ClassOrcamentoVenda
Public mvarobjDocTribOutros As Object
Public objFilialEmpresa As New AdmFiliais
Dim gbTemSRV As Boolean
    'todos os dados relevantes p/tributacao dentro de mvarobjNFTributacao estarao sincronizados com a tela antes da atualizacao da tributacao

Dim mvarcolTiposTribICMS As New Collection
Dim mvarcolTiposTribIPI As New Collection
Dim mvarcolTiposTribPISCOFINS As New Collection
Dim mvarcolTiposTribISS As New Collection
Dim mvarcolTiposTribICMSSimples As New Collection

Dim mvariTrazendoTribTela As Integer
Dim mvariTrazendoTribItemTela As Integer

Dim mvariNatOpInternaAlterada As Integer
Dim mvariISSAliquotaAlterada As Integer
Dim mvariISSValorAlterado As Integer
Dim mvariValorIRRFAlterado As Integer
Dim mvariPISRetidoAlterado As Integer
Dim mvariPISValorAlterado As Integer
Dim mvariISSRetidoAlterado As Integer
Dim mvariCOFINSRetidoAlterado As Integer
Dim mvariCOFINSValorAlterado As Integer
Dim mvariCSLLRetidoAlterado As Integer
Dim mvariTipoTributacaoAlterado As Integer
Dim mvariAliqIRAlterada As Integer
Dim mvariINSSBaseAlterada As Integer
Dim mvariINSSDeducoesAlterada As Integer
Dim mvariINSSValorAlterado As Integer
'Dim mvariINSSRetidoAlterado As Integer
Dim mvariIRBaseAlterado As Integer

Dim mvariNatOpItemAlterado As Integer
Dim mvariTipoTributacaoItemAlterado As Integer
Dim mvariICMSBaseItemAlterado As Integer
Dim mvariICMSPercRedBaseItemAlterado As Integer
Dim mvariICMSAliquotaItemAlterado As Integer
Dim mvariICMSValorItemAlterado As Integer
Dim mvariICMSSubstBaseItemAlterado As Integer
Dim mvariICMSSubstAliquotaItemAlterado As Integer
Dim mvariICMSSubstValorItemAlterado As Integer
Dim mvariIPIBaseItemAlterado As Integer
Dim mvariIPIPercRedBaseItemAlterado As Integer
Dim mvariIPIAliquotaItemAlterado As Integer
Dim mvariIPIValorItemAlterado As Integer
Dim mvariICMSSTCobrAntBaseItemAlterado As Integer
Dim mvariICMSSTCobrAntValorItemAlterado As Integer

Dim mvariCOFINSSTAliquotaValorItemAlterado As Integer
Dim mvariCOFINSSTAliquotaItemAlterado As Integer
Dim mvariCOFINSSTValorItemAlterado As Integer
Dim mvariCOFINSSTBaseItemAlterado As Integer
Dim mvariCOFINSSTQtdeItemAlterado As Integer
Dim mvariCOFINSAliquotaValorItemAlterado As Integer
Dim mvariCOFINSAliquotaItemAlterado As Integer
Dim mvariCOFINSValorItemAlterado As Integer
Dim mvariCOFINSBaseItemAlterado As Integer
Dim mvariCOFINSQtdeItemAlterado As Integer
Dim mvariPISAliquotaValorItemAlterado As Integer
Dim mvariPISAliquotaItemAlterado As Integer
Dim mvariPISValorItemAlterado As Integer
Dim mvariPISBaseItemAlterado As Integer
Dim mvariPISQtdeItemAlterado As Integer
Dim mvariPISSTAliquotaValorItemAlterado As Integer
Dim mvariPISSTAliquotaItemAlterado As Integer
Dim mvariPISSTValorItemAlterado As Integer
Dim mvariPISSTBaseItemAlterado As Integer
Dim mvariPISSTQtdeItemAlterado As Integer
Dim mvariISSCidadeIBGEItemAlterado As Integer
Dim mvariISSAliquotaItemAlterado As Integer
Dim mvariISSValorItemAlterado As Integer
Dim mvariISSBaseItemAlterado As Integer
Dim mvariValorFreteItemAlterado As Integer
Dim mvariValorSeguroItemAlterado As Integer
Dim mvariValorDescontoItemAlterado As Integer
Dim mvariValorOutrasDespesasItemAlterado As Integer
Dim mvariIIIOFItemAlterado As Integer
Dim mvariIIBaseItemAlterado As Integer
Dim mvariIIValorItemAlterado As Integer
Dim mvariIIDespAduaneiraItemAlterado As Integer
Dim mvariICMSSubstPercRedBaseItemAlterado As Integer
Dim mvariICMSSubstPercMVAItemAlterado As Integer
Dim mvariIPIUnidadePadraoValorItemAlterado As Integer
Dim mvariIPIUnidadePadraoQtdeItemAlterado As Integer
Dim mvariFCIAlterado As Integer

Dim mvariICMSpercBaseOperacaoPropriaAlterado As Integer
Dim mvariICMSUFDevidoSTAlterado As Integer
Dim mvariICMSvBCSTRetAlterado As Integer
Dim mvariICMSvICMSSTRetAlterado As Integer
Dim mvariICMSvBCSTDestAlterado As Integer
Dim mvariICMSvICMSSTDestAlterado As Integer
Dim mvariICMSpCredSNAlterado As Integer
Dim mvariICMSvCredSNAlterado As Integer
Dim mvariICMSValorIsentoAlterado As Integer
Dim mvariICMSMotivoAlterado As Integer
Dim mvariICMSSimplesTipoAlterado As Integer
Dim mvariISSTipoAlterado As Integer

Dim mvariTotTribAlterado As Integer
Dim mvariTotTribItemAlterado As Integer
Dim mvariTotTribTipoAlterado As Integer
Dim mvariTotTribItemTipoAlterado As Integer

'IBPTax 0.0.9
Dim mvariTotTribFedAlterado As Integer
Dim mvariTotTribFedItemAlterado As Integer
Dim mvariTotTribEstAlterado As Integer
Dim mvariTotTribEstItemAlterado As Integer
Dim mvariTotTribMunicAlterado As Integer
Dim mvariTotTribMunicItemAlterado As Integer

'nfe 3.10
'Dim mvariIndConsumidorFinalAlterado As Integer
'Dim mvariIndPresencaAlterado As Integer
Dim mvariDataPrestServicoAlterado As Integer
Dim mvariICMS51ValorOpItemAlterado As Integer
Dim mvariICMSPercDiferItemAlterado As Integer
Dim mvariICMSValorDifItemAlterado As Integer
Dim mvariISSValorDeducaoItemAlterado As Integer
Dim mvariISSValorOutrasRetItemAlterado As Integer
Dim mvariISSValorDescIncondItemAlterado As Integer
Dim mvariISSValorDescCondItemAlterado As Integer
Dim mvariISSValorRetItemAlterado As Integer
'Dim mvariISSIndExigibilidadeItemAlterado As Integer
Dim mvariISSCodServItemAlterado As Integer
Dim mvariISSMunicIncidImpItemAlterado As Integer
Dim mvariISSCodPaisItemAlterado As Integer
Dim mvariISSNumProcessoItemAlterado As Integer
'Dim mvariISSIndIncentivoItemAlterado As Integer

'NT 2015.003 (ICMS Interestadual)
Dim mvariICMSInterestPercFCPUFDestItemAlterado As Integer
Dim mvariICMSInterestBCUFDestItemAlterado As Integer
Dim mvariICMSInterestAliqUFDestItemAlterado As Integer
Dim mvariICMSInterestVlrUFDestItemAlterado As Integer
Dim mvariICMSInterestVlrUFRemetItemAlterado As Integer
Dim mvariICMSInterestVlrFCPUFDestItemAlterado As Integer

'UA. Tributos Devolvidos (para o item da NF-e)
Dim mvaripDevolItemAlterado As Integer
Dim mvariIPIVlrDevolvidoItemAlterado As Integer

'nfe 4.00
Dim mvariICMSInterestBCFCPUFDestItemAlterado As Integer
Dim mvariICMSvBCFCPSTItemAlterado As Integer
Dim mvariICMSpFCPSTItemAlterado As Integer
Dim mvariICMSvFCPSTItemAlterado As Integer
Dim mvariICMSvBCFCPSTRetItemAlterado As Integer
Dim mvariICMSpFCPSTRetItemAlterado As Integer
Dim mvariICMSvFCPSTRetItemAlterado As Integer
Dim mvariICMSSTCobrAntAliquotaItemAlterado As Integer
Dim mvariICMSvBCFCPItemAlterado As Integer
Dim mvariICMSpFCPItemAlterado As Integer
Dim mvariICMSvFCPItemAlterado As Integer
Dim mvaricBenefItemAlterado As Integer
'nfe 4.00

Public iTipoTela As Integer
Public iSubTipoTela As Integer

Dim mvariRecalculandoTributacao As Integer

Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"

'Mnemônicos para formação da mensagem da NF
Private Const MN_EXISTE_PRODUTO As String = "Existe_Produto"
Private Const MN_ENTRADA As String = "Entrada"
Private Const MN_SAIDA As String = "Saida"
Private Const MN_TIPOTRIB As String = "TipoTrib"
Private Const MN_TIPONF As String = "TipoNF"
Private Const MN_FILEMPCGC As String = "FilEmpCGC"
Private Const MN_FILEMP As String = "FilEmp"
Private Const MN_EXISTE_CATPRODITEM As String = "Existe_CatProdItem"
Private Const MN_EXISTE_CATPROD As String = "Existe_CatProd"
Private Const MN_EXISTE_TIPOPROD As String = "Existe_TipoProd"
Private Const MN_CODBACENDEST As String = "CodBacenDest"
Private Const MN_UFDEST As String = "UFDest"
Private Const MN_ENDREFDEST As String = "EndRefDest"
Private Const MN_REGRA As String = "Regra"
Private Const MN_BASECREDICMS As String = "BaseCredICMS"
Private Const MN_ALIQUOTACREDICMS As String = "AliquotaCredICMS"
Private Const MN_VALORCREDICMS As String = "ValorCredICMS"
Private Const MN_MSGPV As String = "MsgPV"
Private Const MN_MSGFILCLI As String = "MsgFilCli"
Private Const MN_ENDENTREGA As String = "EndEntrega"
Private Const MN_ENDTRANSREDESP As String = "EndTranspRedesp"
Private Const MN_FILCLIFORNINSCSUF As String = "FilCliFornInscSuf"
Private Const MN_NUMPEDIDOTERC As String = "NumPedidoTerc"
Private Const MN_TIPOCLIFORN As String = "TipoCliForn"
Private Const MN_CFOP As String = "CFOP"
Private Const MN_CLIFORNNOMERED As String = "CliFornNomeRed"
Private Const MN_CLIFORN As String = "CliForn"
Private Const MN_FILCLIFORNINSCEST As String = "FilCliFornInscEst"
Private Const MN_FILCLIFORNINSCMUN As String = "FilCliFornInscMun"
Private Const MN_POSSUI_ICMSST As String = "Possui_ICMSST"
Private Const MN_POSSUI_ICMSRED As String = "Possui_ICMSRED"
Private Const MN_POSSUI_CATCLIFORITEM As String = "Possui_CatCliForItem"
Private Const MN_POSSUI_CATCLIFORN As String = "Possui_CatCliForn"
Private Const MN_INTERNA As String = "Interna"
Private Const MN_EXTERNA As String = "Externa"
Private Const MN_SERIE As String = "Serie"
Private Const MN_VLRDESCONTONF As String = "VlrDescontoNF"
Private Const MN_VLRTOTALNF As String = "VlrTotalNF"
Private Const MN_VLRFRETENF As String = "VlrFreteNF"
Private Const MN_VLRSEGURONF As String = "VlrSeguroNF"
Private Const MN_VLRDESPESASNF As String = "VlrDespesasNF"
Private Const MN_VLRPRODUTONF As String = "VlrProdutoNF"
Private Const MN_NUMPEDIDOVENDA As String = "NumPedidoVenda"
Private Const MN_TOTTRIB As String = "TotTrib"
Private Const MN_PERCTOTTRIB As String = "PercTotTrib"
Private Const MN_TOTTRIBFED As String = "TotTribFed"
Private Const MN_PERCTOTTRIBFED As String = "PercTotTribFed"
Private Const MN_TOTTRIBEST As String = "TotTribEst"
Private Const MN_PERCTOTTRIBEST As String = "PercTotTribEst"
Private Const MN_TOTTRIBMUNIC As String = "TotTribMunic"
Private Const MN_PERCTOTTRIBMUNIC As String = "PercTotTribMunic"

Private Const MN_DEVOLUCAO As String = "Devolucao"
Private Const MN_TIPONFORIG As String = "TipoNFOrig"
Private Const MN_NUMNFORIG As String = "NumNFOrig"
Private Const MN_NUMNFORIGVINC As String = "NumNFOrigVinc"
Private Const MN_CFOPNFORIG As String = "CFOPNFOrig"
Private Const MN_EMISSAONFORIG As String = "EmissaoNFOrig"
Private Const MN_CLIFORNNFORIG As String = "CliFornNFOrig"
Private Const MN_NUMNFORIGDET As String = "NumNFOrigDet"
Private Const MN_NUMNFORIGDETBENEF As String = "NumNFOrigDetBenef"

Private Const MN_VLRICMS As String = "VlrICMS"
Private Const MN_VLRICMSINTERESTFCP As String = "VlrICMSInterestFCP"
Private Const MN_VLRICMSINTERESTUFDEST As String = "VlrICMSInterestDest"
Private Const MN_VLRICMSINTERESTUFREMET As String = "VlrICMSInterestRemet"

Private Const MN_VLRICMSST As String = "VlrICMSST"
Private Const MN_BASEICMSST As String = "BaseICMSST"
Private Const MN_VLRICMSSTCOBRANT As String = "VlrICMSSTCobrAnt"
Private Const MN_BASEICMSSTCOBRANT As String = "BaseICMSSTCobrAnt"

Private Const MN_BASEICMSSTNFORIG As String = "BaseICMSSTNFOrig"
Private Const MN_VLRICMSSTNFORIG As String = "VlrICMSSTNFOrig"

Private gcolRetRegras As Collection
Private gcolRegras As Collection

Public mvariFrameAtual As Integer
Public mvariFrameAtualDet As Integer
Public mvariFrameAtualRes As Integer

Public Property Get objNFTributacao() As ClassTribNFAux
    Set objNFTributacao = mvarobjNFTributacao
End Property

Public Property Get TIPOTELA_NF() As Integer
    TIPOTELA_NF = 0
End Property

Public Property Get TIPOTELA_PV() As Integer
    TIPOTELA_PV = 1
End Property

Public Property Get TIPOTELA_OV() As Integer
    TIPOTELA_OV = 2
End Property

Public Property Get TIPOTELA_PRJ_PROP() As Integer
    TIPOTELA_PRJ_PROP = 3
End Property

Public Property Get TIPOTELA_PRJ_CTR() As Integer
    TIPOTELA_PRJ_CTR = 4
End Property

Public Property Get SUBTIPOTELA_EDICAO() As Integer
    SUBTIPOTELA_EDICAO = 0
End Property

Public Property Get SUBTIPOTELA_CONSULTA() As Integer
    SUBTIPOTELA_CONSULTA = 1
End Property

Private Sub AtualizarTabTrib()

Dim lErro As Long

On Error GoTo Erro_AtualizarTabTrib

    lErro = Tributacao_GravarTela()
    If lErro <> SUCESSO Then Error 27786

    Call gobjDocTrib2.Alterado

    lErro = mvarobjForm.ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 27787
    
    lErro = Trata_Comercio_Exterior
    'If lErro <> SUCESSO Then Error 27787
    
    Exit Sub

Erro_AtualizarTabTrib:

    Select Case Err

        Case 27786, 27787, 27788

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153717)

    End Select

    Exit Sub

End Sub

Private Sub AtualizarTabTribItem()

Dim lErro As Long, iIndice As Integer
Dim objTributacaoItem As Object
Dim objTributoDocItem As ClassTributoDocItem

On Error GoTo Erro_AtualizarTabTribItem

    'atualizar dados da colecao p/o item ou complemento corrente

    'se um item estiver selecionado
    iIndice = TabTrib_Controle("ComboItensTrib").ListIndex
    If iIndice <> -1 Then
        If gbTemSRV Then
            If iIndice + 1 > gcolDocItem(gobjDocTrib, True).Count Then
                Set objTributacaoItem = gobjDocTribItem(iIndice + 1 - gcolDocItem(gobjDocTrib, True).Count) 'gobjPedidoVenda.colItensPedido.Item(iIndice + 1 - gobjPedidoVenda.colItensPedidoSRV.Count)
                Set objTributoDocItem = gcolDocItem(gobjDocTrib)(iIndice + 1 - gcolDocItem(gobjDocTrib, True).Count)
            Else
                Set objTributacaoItem = gobjDocTribItem(iIndice + 1, True)
                Set objTributoDocItem = gcolDocItem(gobjDocTrib, True)(iIndice + 1)
            End If
        Else
            'preenche os campos da tela em funcao do item selecionado
            Set objTributacaoItem = gobjDocTribItem(iIndice + 1)
            Set objTributoDocItem = gcolDocItem(gobjDocTrib)(iIndice + 1)
        End If
    Else
        Error 27789
    End If

    lErro = TributacaoItem_GravarTela(objTributacaoItem)
    If lErro <> SUCESSO Then Error 27790

    Call objTributoDocItem.Alterado

    lErro = mvarobjForm.ValorTotal_Calcula()
    If lErro <> SUCESSO Then Error 27791

    lErro = Trata_Comercio_Exterior_Item(objTributacaoItem)
    'If lErro <> SUCESSO Then Error 27792

    lErro = TributacaoItem_TrazerTela(objTributacaoItem)
    If lErro <> SUCESSO Then Error 27792
    
    lErro = Tributacao_TrazerTela(gobjDocTrib)
    If lErro <> SUCESSO Then Error 27792

    Exit Sub

Erro_AtualizarTabTribItem:

    Select Case Err

        Case 27790, 27791, 27792

        Case 27789
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_ITEM_TRIB_SEL", Err, Error$)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153718)

    End Select
    
'    Resume Next

    Exit Sub

End Sub

Sub ComboICMSTipo_Click()

    If TabTrib_Controle("ComboICMSTipo").ListIndex = -1 Then Exit Sub

    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem

End Sub

Sub ComboIPITipo_Click()

    If TabTrib_Controle("ComboIPITipo").ListIndex = -1 Then Exit Sub

    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem

End Sub

Sub ComboItensTrib_Click()

Dim iIndice As Integer, objItem As Object

    iIndice = TabTrib_Controle("ComboItensTrib").ListIndex

    If iIndice <> -1 Then

        If gbTemSRV Then
            If iIndice + 1 > gcolDocItem(gobjDocTrib, True).Count Then
                Set objItem = gcolDocItem(gobjDocTrib).Item(iIndice + 1 - gcolDocItem(gobjDocTrib, True).Count) 'gobjPedidoVenda.colItensPedido.Item(iIndice + 1 - gobjPedidoVenda.colItensPedidoSRV.Count)
            Else
                Set objItem = gcolDocItem(gobjDocTrib, True).Item(iIndice + 1)
            End If
        Else
            'preenche os campos da tela em funcao do item selecionado
            Set objItem = gcolDocItem(gobjDocTrib).Item(iIndice + 1)
        End If

        TabTrib_Controle("LabelValorItem").Caption = Format(objItem.dPrecoTotal, "Standard")
        If iTipoTela <> TIPOTELA_PV Then
            TabTrib_Controle("LabelQtdeItem").Caption = CStr(objItem.dQuantidade)
        Else
            TabTrib_Controle("LabelQtdeItem").Caption = CStr(objItem.dQuantidade - objItem.dQuantCancelada)
        End If
        TabTrib_Controle("LabelUMItem").Caption = objItem.sUnidadeMed

        Call TributacaoItem_TrazerTela(gobjDocTribItem(iIndice + 1, gbTemSRV))

    End If

End Sub

Sub LblNatOpInterna_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    If Len(Trim(TabTrib_Controle("NatOpInterna").ClipText)) > 0 Then objNaturezaOp.sCodigo = TabTrib_Controle("NatOpInterna").Text

    If Len(Trim(TabTrib_Controle("DataEmissao").ClipText)) > 0 Then
        dtDataRef = MaskedParaDate(TabTrib_Controle("DataEmissao"))
    Else
        dtDataRef = DATA_NULA
    End If

    'Adiciona os limites de Natureza
    Select Case mvariNFEntrada
    
        Case NF_SAIDA
            colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
            colSelecao.Add NATUREZA_SAIDA_COD_FINAL
            
        Case NF_ENTRADA
        
            If dtDataRef < DATA_INICIO_CFOP4 Then
            
                If Len(Trim(TabTrib_Controle("DataEntrada").ClipText)) > 0 Then
                    dtDataRef = MaskedParaDate(TabTrib_Controle("DataEntrada"))
                End If
            
            End If
            
            colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
            colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
            
    End Select

    'Chama a Tela de browse de NaturezaOp p/naturezas de entrada
    If mvariEmitenteNF = DOCINFO_EMPRESA And gobjCRFAT.iUsaNFe = MARCADO Then
        Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, mvarobjEventoNaturezaInterna, gobjVersao.sColunaCFOP & " >= 0")
    Else
        Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, mvarobjEventoNaturezaInterna, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))
    End If
End Sub

Sub NaturezaItemLabel_Click()

Dim objNaturezaOp As New ClassNaturezaOp
Dim colSelecao As New Collection
Dim dtDataRef As Date

    'Se NaturezaOP estiver preenchida coloca no Obj
    If Len(Trim(TabTrib_Controle("NaturezaOpItem").ClipText)) > 0 Then objNaturezaOp.sCodigo = TabTrib_Controle("NaturezaOpItem").Text
    
    If iTipoTela = TIPOTELA_PRJ_CTR Or iTipoTela = TIPOTELA_PRJ_CTR Then
        If Len(Trim(TabTrib_Controle("DataCriacao").ClipText)) > 0 Then
            dtDataRef = MaskedParaDate(TabTrib_Controle("DataCriacao"))
        Else
            dtDataRef = DATA_NULA
        End If
    Else
        If Len(Trim(TabTrib_Controle("DataEmissao").ClipText)) > 0 Then
            dtDataRef = MaskedParaDate(TabTrib_Controle("DataEmissao"))
        Else
            dtDataRef = DATA_NULA
        End If
    End If

    'Adiciona os limites de Natureza
    If mvariNFEntrada = 1 Then

        If dtDataRef < DATA_INICIO_CFOP4 And iTipoTela = TIPOTELA_NF Then
        
            If Len(Trim(TabTrib_Controle("DataEntrada").ClipText)) > 0 Then
                dtDataRef = MaskedParaDate(TabTrib_Controle("DataEntrada"))
            End If
        
        End If
        colSelecao.Add NATUREZA_ENTRADA_COD_INICIAL
        colSelecao.Add NATUREZA_ENTRADA_COD_FINAL
    Else
        colSelecao.Add NATUREZA_SAIDA_COD_INICIAL
        colSelecao.Add NATUREZA_SAIDA_COD_FINAL
    End If

    'Chama a Tela de browse de NaturezaOp p/naturezas de entrada
    If mvariEmitenteNF = DOCINFO_EMPRESA And gobjCRFAT.iUsaNFe = MARCADO Then
        Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, mvarobjEventoNaturezaInterna, gobjVersao.sColunaCFOP & " >= 0")
    Else
        Call Chama_Tela("NaturezaOpLista", colSelecao, objNaturezaOp, mvarobjEventoNaturezaInterna, "{fn LENGTH(Codigo) } = " & IIf(dtDataRef < DATA_INICIO_CFOP4, "3", "4"))
    End If
    
    Exit Sub

End Sub

Sub NaturezaOpItem_Change()
    mvariNatOpItemAlterado = 1
End Sub

Sub NaturezaOpItem_Validate(Cancel As Boolean)

Dim sNatOp As String, lErro As Long, objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_NaturezaOpItem_Validate

    If mvariNatOpItemAlterado = 0 Then Exit Sub

    sNatOp = Trim(TabTrib_Controle("NaturezaOpItem").Text)

    If sNatOp <> "" Then
    
        objNaturezaOp.sCodigo = sNatOp
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then Error 56864
        If lErro <> SUCESSO Then Error 56879
        
        TabTrib_Controle("LabelDescrNatOpItem").Caption = objNaturezaOp.sDescricao
        
        'se for entrada garantir que tenha codigo < 500
        If mvariNFEntrada = 1 And objNaturezaOp.sCodigo >= NATUREZA_SAIDA_COD_INICIAL Then Error 56992

        'se for saida garantir que tenha codigo > 500
        If mvariNFEntrada = 0 And objNaturezaOp.sCodigo <= NATUREZA_SAIDA_COD_INICIAL Then Error 56993

        lErro = NatOp_VerificaTamanho(sNatOp)
        If lErro <> SUCESSO Then Error 32284
        
        If mvariEmitenteNF = DOCINFO_EMPRESA And gobjCRFAT.iUsaNFe = MARCADO Then
            lErro = CF("CFOP_Testa_Versao", sNatOp, gobjVersao.sColunaCFOP, gobjVersao.sDescricao)
            If lErro <> SUCESSO Then Error 32284
        End If
        
    Else
    
        TabTrib_Controle("LabelDescrNatOpItem").Caption = ""
    
    End If

    If mvariTrazendoTribTela = 0 Then Call AtualizarTabTribItem
        
    mvariNatOpItemAlterado = 0

    Exit Sub

Erro_NaturezaOpItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 32284
        
        Case 56864
''???            If TabTrib_Controle("NaturezaOpItem").Visible Then TabTrib_Controle("NaturezaOpItem").SetFocus
        
        Case 56879
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, sNatOp)
''???            If TabTrib_Controle("NaturezaOpItem").Visible Then TabTrib_Controle("NaturezaOpItem").SetFocus
        
        Case 56992
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ENTRADA", Err, sNatOp)
''???            If TabTrib_Controle("NaturezaOpItem").Visible Then TabTrib_Controle("NaturezaOpItem").SetFocus
        
        Case 56993
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", Err, sNatOp)
''???            If TabTrib_Controle("NaturezaOpItem").Visible Then TabTrib_Controle("NaturezaOpItem").SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153719)

    End Select

End Sub

Sub NatOpInterna_Change()
    mvariNatOpInternaAlterada = 1
    If mvariNFEntrada = NF_SAIDA Then TabTrib_Controle("NatOpInternaEspelho").Caption = TabTrib_Controle("NatOpInterna").Text
End Sub

Sub NatOpInterna_Validate(Cancel As Boolean, Optional iEntradaSaida As Integer = 0)
'iEntradaSaida: 0=nao infomado, 1=entrada, 2= saida
Dim sNatOp As String, lErro As Long, objNaturezaOp As New ClassNaturezaOp

On Error GoTo Erro_NatOpInterna_Validate

    If mvariNatOpInternaAlterada = 0 Then Exit Sub

    sNatOp = Trim(TabTrib_Controle("NatOpInterna").Text)

    If sNatOp <> "" Then
    
        objNaturezaOp.sCodigo = sNatOp
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then Error 32006
        
        'se nao encontrou a natureza
        If lErro <> SUCESSO Then Error 32007
        
        TabTrib_Controle("DescNatOpInterna").Caption = objNaturezaOp.sDescricao
        
        If mvariTrazendoTribTela = 0 Then
        
            'se for entrada garantir que tenha codigo < 500
            If mvariNFEntrada = 1 And objNaturezaOp.sCodigo >= NATUREZA_SAIDA_COD_INICIAL Then Error 32008
    
            'se for saida garantir que tenha codigo > 500
            If mvariNFEntrada = 0 And objNaturezaOp.sCodigo <= NATUREZA_SAIDA_COD_INICIAL Then Error 32009
            
            lErro = NatOp_VerificaTamanho(sNatOp)
            If lErro <> SUCESSO Then Error 32283
            
            If mvariEmitenteNF = DOCINFO_EMPRESA And gobjCRFAT.iUsaNFe = MARCADO Then
                lErro = CF("CFOP_Testa_Versao", sNatOp, gobjVersao.sColunaCFOP, gobjVersao.sDescricao)
                If lErro <> SUCESSO Then Error 32283
            End If
                
        End If
        
    Else
    
        TabTrib_Controle("DescNatOpInterna").Caption = ""
    
    End If
    
    If mvariTrazendoTribTela = 0 Then Call AtualizarTabTrib

    If mvariEmitenteNF = DOCINFO_EMPRESA And mvariNFEntrada = 1 Then
    
        TabTrib_Controle("NaturezaOp").Text = TabTrib_Controle("NatOpInterna").Text
        If mvariLivroFiscal = 0 Then mvarobjForm.iNaturezaOpAlterada = 0
        
    End If
    
    mvariNatOpInternaAlterada = 0

    Exit Sub

Erro_NatOpInterna_Validate:

    Cancel = True
    
    Select Case Err

        Case 32283
        
        Case 32006
''???            If TabTrib_Controle("NatOpInterna").Visible Then TabTrib_Controle("NatOpInterna").SetFocus
        
        Case 32007
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, sNatOp)
''???            If TabTrib_Controle("NatOpInterna").Visible Then TabTrib_Controle("NatOpInterna").SetFocus
            
        Case 32008
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ENTRADA", Err, sNatOp)
''???            If TabTrib_Controle("NatOpInterna").Visible Then TabTrib_Controle("NatOpInterna").SetFocus
        
        Case 32009
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SAIDA", Err, sNatOp)
''???            If TabTrib_Controle("NatOpInterna").Visible Then TabTrib_Controle("NatOpInterna").SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153720)

    End Select

End Sub

Sub TipoTributacao_Change()

    mvariTipoTributacaoAlterado = 1

End Sub

Sub TipoTributacao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objTipoDeTributacao As New ClassTipoDeTributacaoMovto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_TipoTributacao_Validate

    If mvariTipoTributacaoAlterado = 0 Then Exit Sub

    If Len(Trim(TabTrib_Controle("TipoTributacao").Text)) <> 0 Then

        objTipoDeTributacao.iTipo = CInt(TabTrib_Controle("TipoTributacao").Text)

        lErro = CF("TipoTributacao_Le", objTipoDeTributacao)
        If lErro <> SUCESSO And lErro <> 27259 Then Error 27793

        If lErro = 27259 Then Error 27794

        If mvariNFEntrada = 1 Then
        
            If objTipoDeTributacao.iEntrada = 0 Then Error 59378
        
        Else
            
            If objTipoDeTributacao.iEntrada = 1 Then Error 59379
        
        End If
        
        With mvarobjForm
        
            TabTrib_Controle("DescTipoTrib").Caption = objTipoDeTributacao.sDescricao
    
'            'se nao incide ISS
'            If objTipoDeTributacao.iISSIncide = 0 Then
'                TabTrib_Controle("FrameTabTribDet")(6).Enabled = False 'Frame de ISS Item
'            Else
'                If iSubTipoTela <> SUBTIPOTELA_CONSULTA Then
'                    TabTrib_Controle("FrameTabTribDet")(6).Enabled = True 'Frame de ISS Item
'                Else
'                    TabTrib_Controle("FrameTabTribDet")(6).Enabled = False 'Frame de ISS Item
'                End If
'            End If
    
            'se nao incide IR
            If objTipoDeTributacao.iIRIncide = 0 Then
                TabTrib_Controle("ValorIRRF").Enabled = False
                TabTrib_Controle("IRAliquota").Enabled = False
                TabTrib_Controle("IRBase").Enabled = False
            Else
                TabTrib_Controle("ValorIRRF").Enabled = True
                TabTrib_Controle("IRAliquota").Enabled = True
                TabTrib_Controle("IRBase").Enabled = True
            End If

            'se nao retem INSS
            If objTipoDeTributacao.iINSSIncide = 0 Then
                TabTrib_Controle("INSSBase").Enabled = False
                TabTrib_Controle("INSSDeducoes").Enabled = False
                TabTrib_Controle("INSSValor").Enabled = False
                TabTrib_Controle("INSSRetido").Enabled = False
            Else
                TabTrib_Controle("INSSBase").Enabled = True
                TabTrib_Controle("INSSDeducoes").Enabled = True
                TabTrib_Controle("INSSValor").Enabled = True
                TabTrib_Controle("INSSRetido").Enabled = True
            End If
            
            'se nao retem PIS
            If objTipoDeTributacao.iPISRetencao = 0 Then
                TabTrib_Controle("PISRetido").Enabled = False
            Else
                TabTrib_Controle("PISRetido").Enabled = True
            End If

            'se nao retem ISS
            If objTipoDeTributacao.iISSRetencao = 0 Then
                TabTrib_Controle("ISSRetido").Enabled = False
            Else
                TabTrib_Controle("ISSRetido").Enabled = True
            End If

            'se nao retem COFINS
            If objTipoDeTributacao.iCOFINSRetencao = 0 Then
                TabTrib_Controle("COFINSRetido").Enabled = False
            Else
                TabTrib_Controle("COFINSRetido").Enabled = True
            End If
        
            'se nao retem CSLL
            If objTipoDeTributacao.iCSLLRetencao = 0 Then
                TabTrib_Controle("CSLLRetido").Enabled = False
            Else
                TabTrib_Controle("CSLLRetido").Enabled = True
            End If
        
        End With

    Else
    
        TabTrib_Controle("DescTipoTrib").Caption = ""
    
    End If
    
    If mvariTrazendoTribTela = 0 Then Call AtualizarTabTrib

    mvariTipoTributacaoAlterado = 0

    Exit Sub

Erro_TipoTributacao_Validate:

    Cancel = True
    
    Select Case Err

        Case 27793
''???            If TabTrib_Controle("TipoTributacao").Visible Then TabTrib_Controle("TipoTributacao").SetFocus

        Case 27794
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TabTrib_Controle("TipoTributacao").Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("TipoDeTributacao", objTipoDeTributacao)
            Else
''???                If TabTrib_Controle("TipoTributacao").Visible Then TabTrib_Controle("TipoTributacao").SetFocus
            End If

        Case 59378
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIB_INCOMPAT_ENTRADA", Err)
''???            If TabTrib_Controle("TipoTributacao").Visible Then TabTrib_Controle("TipoTributacao").SetFocus
        
        Case 59379
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIB_INCOMPAT_SAIDA", Err)
''???            If TabTrib_Controle("TipoTributacao").Visible Then TabTrib_Controle("TipoTributacao").SetFocus
        
        Case Else
            vbMsgRes = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153721)

    End Select

    Exit Sub

End Sub

Sub TipoTributacaoItem_Change()

    mvariTipoTributacaoItemAlterado = 1

End Sub

Sub TipoTributacaoItem_Validate(Cancel As Boolean)

Dim lErro As Long, vbMsgRes As VbMsgBoxResult
Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_TipoTributacaoItem_Validate

    'se trocou o tipo de tributacao
    If mvariTipoTributacaoItemAlterado Then

        objTributacaoTipo.iTipo = StrParaInt(TabTrib_Controle("TipoTributacaoItem").Text)
        If objTributacaoTipo.iTipo <> 0 Then

            lErro = CF("TipoTributacao_Le", objTributacaoTipo)
            If lErro <> SUCESSO And lErro <> 27259 Then Error 27795

            If lErro = 27259 Then Error 19305

            If mvariNFEntrada = 1 Then
            
                If objTributacaoTipo.iEntrada = 0 Then Error 59380
            
            Else
                
                If objTributacaoTipo.iEntrada = 1 Then Error 59381
            
            End If
        
        Else
            
            TabTrib_Controle("DescTipoTribItem").Caption = ""
        
        End If

        Call AtualizarTabTribItem

        mvariTipoTributacaoItemAlterado = 0

    End If

    Exit Sub

Erro_TipoTributacaoItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 27795

        Case 19305
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_TIPOTRIBUTACAO", TabTrib_Controle("TipoTributacaoItem").Text)
            If vbMsgRes = vbYes Then
                Call Chama_Tela("TipoDeTributacao", objTributacaoTipo)
            Else
''???                If TabTrib_Controle("TipoTributacaoItem").Visible Then TabTrib_Controle("TipoTributacaoItem").SetFocus
            End If
        
        Case 59380
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIB_INCOMPAT_ENTRADA", Err)
''???            If TabTrib_Controle("TipoTributacaoItem").Visible Then TabTrib_Controle("TipoTributacaoItem").SetFocus
        
        Case 59381
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOTRIB_INCOMPAT_SAIDA", Err)
''???            If TabTrib_Controle("TipoTributacaoItem").Visible Then TabTrib_Controle("TipoTributacaoItem").SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153722)

    End Select

    Exit Sub

End Sub
Sub TribSobreItem_Click()

Dim iCount As Integer

    iCount = gcolDocTribItem.Count
    If gbTemSRV Then
        iCount = iCount + gcolDocTribItem(, True).Count
    End If
    
    'se houver itens na combo
    If iCount <> 0 Then
    
        If TabTrib_Controle("ComboItensTrib").ListCount <> iCount Then
        
            Call TributacaoItem_InicializaTab
        
        Else
        
            'mostra o frame de itens e esconde o de outros
            TabTrib_Controle("FrameItensTrib").Visible = True
            'TabTrib_Controle("FrameOutrosTrib").Visible = False
            
            'selecionar o 1o item
            TabTrib_Controle("ComboItensTrib").ListIndex = 0
            
            Call ComboItensTrib_Click
        
        End If
        
    End If

End Sub

Sub TributacaoAutomatica_Click()
    
    If TabTrib_Controle("TributacaoAutomatica").Value = 1 Then
    
        Call gobjDocTrib2.Alterado

        Call mvarobjForm.ValorTotal_Calcula
    
    End If
    
End Sub

Sub TributacaoRecalcular_Click()

    mvariRecalculandoTributacao = 1

    Call gobjDocTrib2.Alterado

    Call mvarobjForm.ValorTotal_Calcula

    mvariRecalculandoTributacao = 0

End Sub

Sub OpcaoTributacao_Click()

Dim lErro As Long, objAux As Object

On Error GoTo Erro_OpcaoTributacao_Click

    'Se frame selecionado não for o atual
    If TabTrib_Controle("OpcaoTributacao").SelectedItem.Index <> mvariFrameAtual Then

        If TabStrip_PodeTrocarTab(mvariFrameAtual, TabTrib_Controle("OpcaoTributacao"), mvarobjForm) <> SUCESSO Then Exit Sub
        
        'Esconde o frame atual, mostra o novo
        Set objAux = TabTrib_Controle("FrameTributacao")
        objAux(TabTrib_Controle("OpcaoTributacao").SelectedItem.Index).Visible = True
        objAux(mvariFrameAtual).Visible = False
        'Armazena novo valor de mvariFrameAtual
        mvariFrameAtual = TabTrib_Controle("OpcaoTributacao").SelectedItem.Index

        'se abriu o tab de detalhamento
        If TabTrib_Controle("OpcaoTributacao").SelectedItem.Index = 2 Then
            lErro = TributacaoItem_InicializaTab
            If lErro <> SUCESSO Then Error 27796
        End If

    End If

    Exit Sub

Erro_OpcaoTributacao_Click:

    Select Case Err

        Case 27796

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153723)

    End Select

    Exit Sub

End Sub

Sub ValorIRRF_Change()

    mvariValorIRRFAlterado = REGISTRO_ALTERADO

End Sub

Sub ValorIRRF_Validate(Cancel As Boolean)

Dim dValorIRRF As Double
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_ValorIRRF_Validate

    If mvariValorIRRFAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ValorIRRF").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ValorIRRF").Text)
        If lErro <> SUCESSO Then Error 27797

        dValorIRRF = CDbl(TabTrib_Controle("ValorIRRF").Text)

        'este teste é necessário pois nas nfs externas o valor total é uma masked pois é informado pelo usuario p/efeito de conferencia
        If mvariNFEntrada = NF_SAIDA Then
            If Len(Trim(TabTrib_Controle("ValorTotal").Caption)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("ValorTotal").Caption)
        Else
            If Len(Trim(TabTrib_Controle("Total").Text)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("Total").Text)
        End If

        If dValorIRRF > dValorTotal Then Error 27798

        TabTrib_Controle("ValorIRRF").Text = Format(dValorIRRF, "Standard")

    End If

    mvariValorIRRFAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True
    
    Select Case Err

        Case 26654, 27677, 27678, 27797
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case 27798
            Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", Err, dValorIRRF, dValorTotal)
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153724)

    End Select

    Exit Sub

End Sub

Private Function TributacaoNF_Terminar() As Long

Dim lErro As Long, objTributoDoc As ClassTributoDoc

On Error GoTo Erro_TributacaoNF_Terminar

    'se mvarobjNFTributacao já foi inicializado
    Set objTributoDoc = gobjDocTrib
    lErro = objTributoDoc.Desativar
    If lErro <> SUCESSO Then Error 27799
    gobjDocTrib = Nothing

    TributacaoNF_Terminar = SUCESSO

    Exit Function

Erro_TributacaoNF_Terminar:

    TributacaoNF_Terminar = Err

    Select Case Err

        Case 27799

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153725)

    End Select

End Function

Public Function TributacaoNF_Reset(Optional objAux As Object) As Long
'cria ou atualiza mvarobjNFTributacao com dados "padrao"

Dim lErro As Long, objTributoDoc As ClassTributoDoc, objItemNF As ClassItemNF
Dim iTipoNFiscal As Integer, sSiglaTipoNFiscal As String, iEmitenteNF As Integer
Dim dtDataEmissao As Date, dtDataEntrada As Date
Dim sSerie As String, objCidade As New ClassCidades

On Error GoTo Erro_TributacaoNF_Reset

    objFilialEmpresa.iCodFilial = giFilialEmpresa

    'se mvarobjNFTributacao já foi inicializado
    If Not (gobjDocTrib Is Nothing) Then
               
        'salva valores anteriores
        If iTipoTela = TIPOTELA_NF Then
            iTipoNFiscal = mvarobjNFTributacao.iTipoNFiscal
            sSiglaTipoNFiscal = mvarobjNFTributacao.sSiglaTipoNFiscal
            iEmitenteNF = mvariEmitenteNF
            dtDataEmissao = mvarobjNFTributacao.dtDataEmissao
            dtDataEntrada = mvarobjNFTributacao.dtDataEntrada
            sSerie = mvarobjNFTributacao.sSerie
        End If
        
        Set objTributoDoc = gobjDocTrib
        lErro = objTributoDoc.Desativar
        If lErro <> SUCESSO Then Error 27800
        
        If iTipoTela = TIPOTELA_NF Then
            Set mvarobjNFTributacao = New ClassTribNFAux
            
            'retorna os valores salvos acima
            mvarobjNFTributacao.iTipoNFiscal = iTipoNFiscal
            mvarobjNFTributacao.sSiglaTipoNFiscal = sSiglaTipoNFiscal
            mvariEmitenteNF = iEmitenteNF
            mvarobjNFTributacao.dtDataEmissao = dtDataEmissao
            mvarobjNFTributacao.dtDataEntrada = dtDataEntrada
            mvarobjNFTributacao.sSerie = sSerie
        ElseIf iTipoTela = TIPOTELA_PV Then
            If Not (objAux Is Nothing) Then
                Set mvarobjPV = objAux
            Else
                Set mvarobjPV = New ClassPedidoDeVenda
                mvarobjPV.dtDataEmissao = gdtDataAtual
            End If
            If gbTemSRV Then mvarobjPV.iServico = MARCADO
            Set mvarobjForm.gobjPedidoVenda = mvarobjPV
            iTipoNFiscal = 52
            dtDataEntrada = gdtDataAtual
        ElseIf iTipoTela = TIPOTELA_OV Then
            If Not (objAux Is Nothing) Then
                Set mvarobjOV = objAux
            Else
                Set mvarobjOV = New ClassOrcamentoVenda
                mvarobjOV.dtDataEmissao = gdtDataAtual
            End If
            If gbTemSRV Then mvarobjOV.iServico = MARCADO
            Set mvarobjForm.gobjOrcamentoVenda = mvarobjOV
            iTipoNFiscal = 52
            dtDataEntrada = gdtDataAtual
        Else
            If Not (objAux Is Nothing) Then
                Set mvarobjDocTribOutros = objAux
            Else
                Call mvarobjForm.Novo_GOBJ(mvarobjDocTribOutros)
                mvarobjDocTribOutros.dtData = gdtDataAtual
            End If
            Call mvarobjForm.Atualiza_GOBJ(mvarobjDocTribOutros)
            iTipoNFiscal = 52
            dtDataEntrada = gdtDataAtual
        End If
        
    Else
    
        If iTipoTela = TIPOTELA_NF Then
            Set mvarobjNFTributacao = New ClassTribNFAux
            mvarobjNFTributacao.dtDataEmissao = gdtDataAtual
            mvarobjNFTributacao.dtDataEntrada = gdtDataAtual
            mvarobjNFTributacao.sSerie = "" '???
        ElseIf iTipoTela = TIPOTELA_PV Then
            Set mvarobjPV = New ClassPedidoDeVenda
            mvarobjPV.dtDataEmissao = gdtDataAtual
            Set mvarobjForm.gobjPedidoVenda = mvarobjPV
        ElseIf iTipoTela = TIPOTELA_OV Then
            Set mvarobjOV = New ClassOrcamentoVenda
            mvarobjOV.dtDataEmissao = gdtDataAtual
            Set mvarobjForm.gobjOrcamentoVenda = mvarobjOV
        Else
            Call mvarobjForm.Novo_GOBJ(mvarobjDocTribOutros)
            mvarobjDocTribOutros.dtDataEmissao = gdtDataAtual
            Call mvarobjForm.Atualiza_GOBJ(mvarobjDocTribOutros)
        End If

    End If

    Set objTributoDoc = gobjDocTrib
    lErro = objTributoDoc.Ativar
    If lErro <> SUCESSO Then Error 27801

    If (iTipoTela = TIPOTELA_NF And Not (objAux Is Nothing)) Then
    
        'obtem as informacoes de uma nf (inclusive seus complementos)
        lErro = ExtraiNF(objAux)
        If lErro <> SUCESSO Then Error 22881
        
        'obtem as informacoes dos itens da NF
        For Each objItemNF In objAux.colItensNF
            lErro = ExtraiNFItem(objAux, objItemNF)
            If lErro <> SUCESSO Then Error 22882
        Next
        
        If gbTemSRV Then
            'obtem as informacoes dos itens da NF
            For Each objItemNF In objAux.colItensSRV
                lErro = ExtraiNFItem(objAux, objItemNF, True)
                If lErro <> SUCESSO Then Error 22882
            Next
        End If
        
    End If
    
    If mvarcolTiposTribICMS.Count = 0 Then
        lErro = CF("TiposTribICMS_Le_Todos", mvarcolTiposTribICMS)
        If lErro <> SUCESSO Then Error 22882
    End If
    
    If mvarcolTiposTribICMSSimples.Count = 0 Then
        lErro = CF("TiposTribICMSSimples_Le_Todos", mvarcolTiposTribICMSSimples)
        If lErro <> SUCESSO Then Error 22882
    End If
    
    If mvarcolTiposTribIPI.Count = 0 Then
        lErro = CF("TiposTribIPI_Le_Todos", mvarcolTiposTribIPI)
        If lErro <> SUCESSO Then Error 22882
    End If
    
    If mvarcolTiposTribPISCOFINS.Count = 0 Then
        lErro = CF("TiposTribPISCOFINS_Le_Todos", mvarcolTiposTribPISCOFINS)
        If lErro <> SUCESSO Then Error 22882
    End If
    
    If mvarcolTiposTribISS.Count = 0 Then
        lErro = CF("TiposTribISS_Le_Todos", mvarcolTiposTribISS)
        If lErro <> SUCESSO Then Error 22882
    End If
    
    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO Then Error 22882
    
    objCidade.sDescricao = objFilialEmpresa.objEndereco.sCidade

    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then Error 22882

    sCidadeIBGECodigoEmp = objCidade.sCodIBGE
    
    Set gobjDocTribTributacao.colTiposTribICMS = mvarcolTiposTribICMS
    Set gobjDocTribTributacao.colTiposTribIPI = mvarcolTiposTribIPI
    Set gobjDocTribTributacao.colTiposTribPISCOFINS = mvarcolTiposTribPISCOFINS
    Set gobjDocTribTributacao.colTiposTribICMSSimples = mvarcolTiposTribICMSSimples
    Set gobjDocTribTributacao.colTiposTribISS = mvarcolTiposTribISS

    mvariNatOpInternaAlterada = 0
    mvariTipoTributacaoAlterado = 0
    mvariISSAliquotaAlterada = 0
    mvariISSValorAlterado = 0
    mvariValorIRRFAlterado = 0
    mvariPISRetidoAlterado = 0
    mvariPISValorAlterado = 0
    mvariISSRetidoAlterado = 0
    mvariCOFINSRetidoAlterado = 0
    mvariCOFINSValorAlterado = 0
    mvariCSLLRetidoAlterado = 0
    mvariAliqIRAlterada = 0

    mvariINSSBaseAlterada = 0
    mvariINSSDeducoesAlterada = 0
    mvariINSSValorAlterado = 0
'    mvariINSSRetidoAlterado = 0
    mvariIRBaseAlterado = 0
    
    mvariTotTribAlterado = 0
    mvariTotTribFedAlterado = 0
    mvariTotTribEstAlterado = 0
    mvariTotTribMunicAlterado = 0
    mvariTotTribTipoAlterado = 0
    
    'nfe 3.10
'    mvariIndConsumidorFinalAlterado = 0
'    mvariIndPresencaAlterado = 0
    mvariDataPrestServicoAlterado = 0
    mvariICMS51ValorOpItemAlterado = 0
    mvariICMSPercDiferItemAlterado = 0
    mvariICMSValorDifItemAlterado = 0
    mvariISSValorDeducaoItemAlterado = 0
    mvariISSValorOutrasRetItemAlterado = 0
    mvariISSValorDescIncondItemAlterado = 0
    mvariISSValorDescCondItemAlterado = 0
    mvariISSValorRetItemAlterado = 0
'    mvariISSIndExigibilidadeItemAlterado = 0
    mvariISSCodServItemAlterado = 0
    mvariISSMunicIncidImpItemAlterado = 0
    mvariISSCodPaisItemAlterado = 0
    mvariISSNumProcessoItemAlterado = 0
'    mvariISSIndIncentivoItemAlterado = 0
    
    mvariNatOpItemAlterado = 0
    mvariTipoTributacaoItemAlterado = 0
    mvariICMSBaseItemAlterado = 0
    mvariICMSPercRedBaseItemAlterado = 0
    mvariICMSAliquotaItemAlterado = 0
    mvariICMSValorItemAlterado = 0
    mvariICMSSubstBaseItemAlterado = 0
    mvariICMSSubstAliquotaItemAlterado = 0
    mvariICMSSubstValorItemAlterado = 0
    mvariIPIBaseItemAlterado = 0
    mvariIPIPercRedBaseItemAlterado = 0
    mvariIPIAliquotaItemAlterado = 0
    mvariIPIValorItemAlterado = 0
    mvariICMSSTCobrAntBaseItemAlterado = 0
    mvariICMSSTCobrAntValorItemAlterado = 0
    
    mvariCOFINSSTAliquotaValorItemAlterado = 0
    mvariCOFINSSTAliquotaItemAlterado = 0
    mvariCOFINSSTValorItemAlterado = 0
    mvariCOFINSSTBaseItemAlterado = 0
    mvariCOFINSSTQtdeItemAlterado = 0
    mvariCOFINSAliquotaValorItemAlterado = 0
    mvariCOFINSAliquotaItemAlterado = 0
    mvariCOFINSValorItemAlterado = 0
    mvariCOFINSBaseItemAlterado = 0
    mvariCOFINSQtdeItemAlterado = 0
    mvariPISAliquotaValorItemAlterado = 0
    mvariPISAliquotaItemAlterado = 0
    mvariPISValorItemAlterado = 0
    mvariPISBaseItemAlterado = 0
    mvariPISQtdeItemAlterado = 0
    mvariPISSTAliquotaValorItemAlterado = 0
    mvariPISSTAliquotaItemAlterado = 0
    mvariPISSTValorItemAlterado = 0
    mvariPISSTBaseItemAlterado = 0
    mvariPISSTQtdeItemAlterado = 0
    mvariISSCidadeIBGEItemAlterado = 0
    mvariISSAliquotaItemAlterado = 0
    mvariISSValorItemAlterado = 0
    mvariISSBaseItemAlterado = 0
    mvariValorFreteItemAlterado = 0
    mvariValorSeguroItemAlterado = 0
    mvariValorDescontoItemAlterado = 0
    mvariValorOutrasDespesasItemAlterado = 0
    mvariIIIOFItemAlterado = 0
    mvariIIBaseItemAlterado = 0
    mvariIIValorItemAlterado = 0
    mvariIIDespAduaneiraItemAlterado = 0
    mvariICMSSubstPercRedBaseItemAlterado = 0
    mvariICMSSubstPercMVAItemAlterado = 0
    mvariIPIUnidadePadraoValorItemAlterado = 0
    mvariIPIUnidadePadraoQtdeItemAlterado = 0
    mvariFCIAlterado = 0
    
    mvariICMSpercBaseOperacaoPropriaAlterado = 0
    mvariICMSUFDevidoSTAlterado = 0
    mvariICMSvBCSTRetAlterado = 0
    mvariICMSvICMSSTRetAlterado = 0
    mvariICMSvBCSTDestAlterado = 0
    mvariICMSvICMSSTDestAlterado = 0
    mvariICMSpCredSNAlterado = 0
    mvariICMSvCredSNAlterado = 0
    mvariICMSValorIsentoAlterado = 0
    mvariICMSMotivoAlterado = 0
    mvariICMSSimplesTipoAlterado = 0
    mvariISSTipoAlterado = 0
    mvariTotTribItemAlterado = 0
    mvariTotTribFedItemAlterado = 0
    mvariTotTribEstItemAlterado = 0
    mvariTotTribMunicItemAlterado = 0
    mvariTotTribItemTipoAlterado = 0
    
    'NT 2015.003 (ICMS Interestadual)
    mvariICMSInterestPercFCPUFDestItemAlterado = 0
    mvariICMSInterestBCUFDestItemAlterado = 0
    mvariICMSInterestAliqUFDestItemAlterado = 0
    mvariICMSInterestVlrUFDestItemAlterado = 0
    mvariICMSInterestVlrUFRemetItemAlterado = 0
    mvariICMSInterestVlrFCPUFDestItemAlterado = 0
    
    'nfe 4.0
    mvariICMSInterestBCFCPUFDestItemAlterado = 0
    mvariICMSpFCPSTItemAlterado = 0
    mvariICMSvFCPSTItemAlterado = 0
    mvariICMSvBCFCPSTRetItemAlterado = 0
    mvariICMSpFCPSTRetItemAlterado = 0
    mvariICMSvFCPSTRetItemAlterado = 0
    mvariICMSSTCobrAntAliquotaItemAlterado = 0
    mvariICMSvBCFCPItemAlterado = 0
    mvariICMSpFCPItemAlterado = 0
    mvariICMSvFCPItemAlterado = 0
    
    TributacaoNF_Reset = SUCESSO

    Exit Function

Erro_TributacaoNF_Reset:

    TributacaoNF_Reset = Err

    Select Case Err

        Case 27800, 27801, 22881, 22882

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153726)

    End Select

    Exit Function

End Function

Private Function CarregaTiposTrib() As Long

Dim lErro As Long, sCodigo As String
Dim objTipoTribICMS As ClassTipoTribICMS
Dim objTipoTribICMSSimples As ClassTipoTribICMSSimples
Dim objTipoTribIPI As ClassTipoTribIPI

On Error GoTo Erro_CarregaTiposTrib

    Set mvarcolTiposTribICMS = New Collection

    lErro = CF("TiposTribICMS_Le_Todos", mvarcolTiposTribICMS)
    If lErro <> SUCESSO Then Error 27802

    'Preenche ComboICMSTipo
    For Each objTipoTribICMS In mvarcolTiposTribICMS
        'sCodigo = ""
        'sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribICMS.iTipo)))
        sCodigo = Format(objTipoTribICMS.iTipoTribCST, "00") & SEPARADOR & objTipoTribICMS.sDescricao
        If Not (mvarobjForm Is Nothing) Then
            TabTrib_Controle("ComboICMSTipo").AddItem (sCodigo)
            TabTrib_Controle("ComboICMSTipo").ItemData(TabTrib_Controle("ComboICMSTipo").NewIndex) = objTipoTribICMS.iTipo
        End If
    Next
    
'    Set mvarcolTiposTribICMSSimples = New Collection
'
'    lErro = CF("TiposTribICMSSimples_Le_Todos", mvarcolTiposTribICMSSimples)
'    If lErro <> SUCESSO Then Error 27802
'
'    'Preenche ComboICMSTipo
'    For Each objTipoTribICMSSimples In mvarcolTiposTribICMSSimples
'        'sCodigo = ""
'        'sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribICMS.iTipo)))
'        sCodigo = Format(objTipoTribICMSSimples.iCSOSN, "000") & SEPARADOR & objTipoTribICMSSimples.sDescricao
'        If Not (mvarobjForm Is Nothing) Then
'            TabTrib_Controle("ComboICMSSimplesTipo").AddItem (sCodigo)
'            TabTrib_Controle("ComboICMSSimplesTipo").ItemData(TabTrib_Controle("ComboICMSSimplesTipo").NewIndex) = objTipoTribICMSSimples.iTipo
'        End If
'    Next
    
    Set mvarcolTiposTribIPI = New Collection

    lErro = CF("TiposTribIPI_Le_Todos", mvarcolTiposTribIPI)
    If lErro <> SUCESSO Then Error 27803

    'Preenche ComboIPITipo
    For Each objTipoTribIPI In mvarcolTiposTribIPI
        'sCodigo = ""
        'sCodigo = Space(STRING_TIPO_ICMS_CODIGO - Len(CStr(objTipoTribIPI.iTipo)))
        If mvariNFEntrada = MARCADO Then
            sCodigo = Format(objTipoTribIPI.iCSTEntrada, "00") & SEPARADOR & objTipoTribIPI.sDescricao
        Else
            sCodigo = Format(objTipoTribIPI.iCSTSaida, "00") & SEPARADOR & objTipoTribIPI.sDescricao
        End If
        TabTrib_Controle("ComboIPITipo").AddItem (sCodigo)
        TabTrib_Controle("ComboIPITipo").ItemData(TabTrib_Controle("ComboIPITipo").NewIndex) = objTipoTribIPI.iTipo

    Next

    CarregaTiposTrib = SUCESSO

    Exit Function

Erro_CarregaTiposTrib:

    CarregaTiposTrib = Err

    Select Case Err

        Case 27802, 27803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153727)

    End Select

    Exit Function

End Function

Private Function Carrega_Tab_Tributacao(objTributacao As Object) As Long

Dim lErro As Long
Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto

On Error GoTo Erro_Carrega_Tab_Tributacao

    mvariTrazendoTribTela = 1

    'no frame de "resumo"
    If (TabTrib_Controle("DescNatOpInterna").Caption = "" Or (TabTrib_Controle("NatOpInterna").Text <> gobjDocTrib.sNaturezaOp)) Then
        
        TabTrib_Controle("NatOpInterna").Text = objTributacao.sNaturezaOp ' gobjDocTrib.sNaturezaOp
        Call NatOpInterna_Validate(bSGECancelDummy)
        
    Else
    
        If mvariEmitenteNF = DOCINFO_EMPRESA And mvariNFEntrada = 1 Then
        
            TabTrib_Controle("NaturezaOp").Text = TabTrib_Controle("NatOpInterna").Text
            If mvariLivroFiscal = 0 Then mvarobjForm.iNaturezaOpAlterada = 0
            
        End If
    
    End If
    
    objTributacaoTipo.iTipo = objTributacao.iTipoTributacao
    If objTributacaoTipo.iTipo <> 0 Then
        If (TabTrib_Controle("DescTipoTrib").Caption = "" Or (TabTrib_Controle("TipoTributacao").Text <> CStr(objTributacao.iTipoTributacao))) Then
            TabTrib_Controle("TipoTributacao").Text = CStr(objTributacao.iTipoTributacao)
            Call TipoTributacao_Validate(bSGECancelDummy)
        End If
    Else
        TabTrib_Controle("TipoTributacao").Text = ""
        TabTrib_Controle("DescTipoTrib").Caption = ""
    End If

    With mvarobjForm
        TabTrib_Controle("IPIBase").Caption = Format(objTributacao.dIPIBase, "Standard")
        TabTrib_Controle("IPIValor").Caption = Format(objTributacao.dIPIValor, "Standard")
        TabTrib_Controle("IPIValor1").Caption = Format(objTributacao.dIPIValor, "Standard")
        'If TabTrib_Controle("IPICredito").Enabled Then TabTrib_Controle("IPICredito").Caption = Format(objTributacao.dIPICredito, "Standard")
        TabTrib_Controle("IPICredito").Caption = Format(objTributacao.dIPICredito, "Standard")
        TabTrib_Controle("ISSBase").Caption = Format(objTributacao.dISSBase, "Standard")
        TabTrib_Controle("ISSBase1").Caption = Format(objTributacao.dISSBase, "Standard")
        'TabTrib_Controle("ISSAliquota").Text = CStr(objTributacao.dISSAliquota * 100)
        TabTrib_Controle("ISSValor").Caption = Format(objTributacao.dISSValor, "Standard")
        TabTrib_Controle("ISSValor1").Caption = Format(objTributacao.dISSValor, "Standard")
        TabTrib_Controle("ISSIncluso").Value = objTributacao.iISSIncluso
        TabTrib_Controle("IRBase").Text = Format(objTributacao.dIRRFBase, "Standard")
        TabTrib_Controle("IRAliquota").Text = CStr(objTributacao.dIRRFAliquota * 100)
        TabTrib_Controle("ValorIRRF").Text = Format(objTributacao.dIRRFValor, "Standard")
        TabTrib_Controle("ICMSBase").Caption = Format(objTributacao.dICMSBase, "Standard")
        TabTrib_Controle("ICMSBase1").Caption = Format(objTributacao.dICMSBase, "Standard")
        TabTrib_Controle("ICMSValor").Caption = Format(objTributacao.dICMSValor, "Standard")
        TabTrib_Controle("ICMSValor1").Caption = Format(objTributacao.dICMSValor, "Standard")
        TabTrib_Controle("ICMSSubstBase").Caption = Format(objTributacao.dICMSSubstBase, "Standard")
        TabTrib_Controle("ICMSSubstBase1").Caption = Format(objTributacao.dICMSSubstBase, "Standard")
        TabTrib_Controle("ICMSSubstValor").Caption = Format(objTributacao.dICMSSubstValor, "Standard")
        TabTrib_Controle("ICMSSubstValor1").Caption = Format(objTributacao.dICMSSubstValor, "Standard")
        If TabTrib_Controle("ICMSCredito").Enabled Then TabTrib_Controle("ICMSCredito").Caption = Format(objTributacao.dICMSCredito, "Standard")
        TabTrib_Controle("INSSBase").Text = Format(objTributacao.dINSSBase, "Standard")
        TabTrib_Controle("INSSDeducoes").Text = Format(objTributacao.dINSSDeducoes, "Standard")
        TabTrib_Controle("INSSValor").Text = Format(objTributacao.dINSSValor, "Standard")
        If TabTrib_Controle("INSSRetido").Value <> objTributacao.iINSSRetido Then
            TabTrib_Controle("INSSRetido").Value = objTributacao.iINSSRetido
        End If
        TabTrib_Controle("PISCredito").Caption = Format(objTributacao.dPISCredito, "Standard")
        TabTrib_Controle("COFINSCredito").Caption = Format(objTributacao.dCOFINSCredito, "Standard")
        
        TabTrib_Controle("PISRetido").Text = Format(objTributacao.dPISRetido, "Standard")
        TabTrib_Controle("ISSRetido").Text = Format(objTributacao.dISSRetido, "Standard")
        TabTrib_Controle("COFINSRetido").Text = Format(objTributacao.dCOFINSRetido, "Standard")
        TabTrib_Controle("CSLLRetido").Text = Format(objTributacao.dCSLLRetido, "Standard")
        TabTrib_Controle("TotTrib").Text = Format(objTributacao.dTotTrib, "Standard")
        TabTrib_Controle("TotTribFed").Text = Format(objTributacao.dTotTribFed, "Standard")
        TabTrib_Controle("TotTribEst").Text = Format(objTributacao.dTotTribEst, "Standard")
        TabTrib_Controle("TotTribMunic").Text = Format(objTributacao.dTotTribMunic, "Standard")
        Call Combo_Seleciona_ItemData(TabTrib_Controle("TotTribTipo"), objTributacao.iTotTribTipo)
        
        'nfe 3.10
        TabTrib_Controle("PISValorST").Caption = Format(objTributacao.dPISSTValor, "Standard")
        TabTrib_Controle("COFINSValorST").Caption = Format(objTributacao.dCOFINSSTValor, "Standard")
        TabTrib_Controle("IIValor").Caption = Format(objTributacao.dIIValor, "Standard")
        
        TabTrib_Controle("PISValor").Caption = Format(objTributacao.dPISValor, "Standard")
        TabTrib_Controle("COFINSValor").Caption = Format(objTributacao.dCOFINSValor, "Standard")
        TabTrib_Controle("ICMSDesonerado").Caption = Format(objTributacao.dICMSValorIsento, "Standard")
        TabTrib_Controle("ISSValorDeducao").Caption = Format(objTributacao.dISSValorDeducao, "Standard")
        TabTrib_Controle("ISSValorOutrasRet").Caption = Format(objTributacao.dISSValorOutrasRet, "Standard")
        TabTrib_Controle("ISSValorDescIncond").Caption = Format(objTributacao.dISSValorDescIncond, "Standard")
        TabTrib_Controle("ISSValorDescCond").Caption = Format(objTributacao.dISSValorDescCond, "Standard")
        Call DateParaMasked(TabTrib_Controle("DataPrestServico"), objTributacao.dtDataPrestServico)

        If TabTrib_Controle("IndConsumidorFinal").Value <> objTributacao.iIndConsumidorFinal Then
            TabTrib_Controle("IndConsumidorFinal").Value = objTributacao.iIndConsumidorFinal
        End If

        Call Combo_Seleciona_ItemData(TabTrib_Controle("IndPresenca"), objTributacao.iIndPresenca)
    
        TabTrib_Controle("ISSMunicFaturamento").PromptInclude = False
        TabTrib_Controle("ISSMunicFaturamento").Text = IIf(objTributacao.sISSMunicFaturamento = "", sCidadeIBGECodigoEmp, objTributacao.sISSMunicFaturamento)
        TabTrib_Controle("ISSMunicFaturamento").PromptInclude = True
    
        TabTrib_Controle("ICMSInterestVlrFCPUFDest").Caption = Format(objTributacao.dICMSInterestVlrFCPUFDest, "Standard")
        TabTrib_Controle("ICMSInterestVlrUFDest").Caption = Format(objTributacao.dICMSInterestVlrUFDest, "Standard")
        TabTrib_Controle("ICMSInterestVlrUFRemet").Caption = Format(objTributacao.dICMSInterestVlrUFRemet, "Standard")
        
        'NFe 4.00
        TabTrib_Controle("ICMSVlrFCP").Caption = Format(objTributacao.dICMSVlrFCP, "Standard")
        TabTrib_Controle("ICMSVlrFCPSTRet").Caption = Format(objTributacao.dICMSVlrFCPSTRet, "Standard")
        TabTrib_Controle("ICMSVlrFCPST").Caption = Format(objTributacao.dICMSVlrFCPST, "Standard")
        TabTrib_Controle("IPIVlrDevolvido").Caption = Format(objTributacao.dIPIVlrDevolvido, "Standard")
        
    End With

    'obs.: o frame de "detalhamento" vou deixar p/carregar qdo o usuario entrar nele

    mvariISSAliquotaAlterada = 0
    mvariISSValorAlterado = 0
    mvariValorIRRFAlterado = 0
    mvariPISRetidoAlterado = 0
    mvariPISValorAlterado = 0
    mvariISSRetidoAlterado = 0
    mvariCOFINSRetidoAlterado = 0
    mvariCOFINSValorAlterado = 0
    mvariCSLLRetidoAlterado = 0
    mvariTipoTributacaoAlterado = 0
    mvariAliqIRAlterada = 0

    mvariINSSBaseAlterada = 0
    mvariINSSDeducoesAlterada = 0
    mvariINSSValorAlterado = 0
'    mvariINSSRetidoAlterado = 0
    mvariIRBaseAlterado = 0

    mvariTotTribAlterado = 0
    mvariTotTribFedAlterado = 0
    mvariTotTribEstAlterado = 0
    mvariTotTribMunicAlterado = 0
    mvariTotTribTipoAlterado = 0
    
    'nfe 3.10
'    mvariIndConsumidorFinalAlterado = 0
'    mvariIndPresencaAlterado = 0
    mvariDataPrestServicoAlterado = 0
    
    mvariTrazendoTribTela = 0

    Carrega_Tab_Tributacao = SUCESSO

    Exit Function

Erro_Carrega_Tab_Tributacao:

    mvariTrazendoTribTela = 0

    Carrega_Tab_Tributacao = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153728)

    End Select

End Function

Private Function TributacaoItem_InicializaTab() As Long
'deve ser chamada na entrada do tab de detalhamento dentro do tab de tributacao

Dim lErro As Long, objItem As Object, sItem As String

On Error GoTo Erro_TributacaoItem_InicializaTab
    
    'esvaziar a combo de itens
    TabTrib_Controle("ComboItensTrib").Clear
    
    If gbTemSRV Then
    
        'preencher a combo de itens: com "codigo do produto - descricao"
        For Each objItem In gcolDocTribItem(, True)
    
            lErro = Mascara_MascararProduto(objItem.sProduto, sItem)
            If lErro <> SUCESSO Then Error 27805
            
            If iTipoTela <> TIPOTELA_NF And Len(Trim(objItem.sProdutoDescricao)) = 0 Then objItem.sProdutoDescricao = objItem.sDescricao
            If iTipoTela = TIPOTELA_PV Then objItem.sProdutoDescricao = objItem.sDescricao
    
            sItem = sItem & " - " & objItem.sProdutoDescricao
            TabTrib_Controle("ComboItensTrib").AddItem sItem
        Next
    
    End If

    'preencher a combo de itens: com "codigo do produto - descricao"
    For Each objItem In gcolDocTribItem

        lErro = Mascara_MascararProduto(objItem.sProduto, sItem)
        If lErro <> SUCESSO Then Error 27805
        
        If iTipoTela <> TIPOTELA_NF And Len(Trim(objItem.sProdutoDescricao)) = 0 Then objItem.sProdutoDescricao = objItem.sDescricao
        If iTipoTela = TIPOTELA_PV Then objItem.sProdutoDescricao = objItem.sDescricao

        sItem = sItem & " - " & objItem.sProdutoDescricao
        TabTrib_Controle("ComboItensTrib").AddItem sItem
    Next

    Call TribSobreItem_Click

    TributacaoItem_InicializaTab = SUCESSO

    Exit Function

Erro_TributacaoItem_InicializaTab:

    TributacaoItem_InicializaTab = Err

    Select Case Err

        Case 27805

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153729)

    End Select

    Exit Function

End Function

Private Function TributacaoItem_TrazerTela(objTributacaoItem As Object) As Long
'traz para a tela dados de tributacao de um item

Dim iIndice As Integer
Dim lErro As Long, objTipoTribIPI As ClassTipoTribIPI, objTipoTribICMS As ClassTipoTribICMS, objTipoTribICMSSimples As ClassTipoTribICMSSimples, objTipoTribISS As ClassTipoTribISS
Dim objTributacaoTipo As New ClassTipoDeTributacaoMovto
Dim objNaturezaOp As New ClassNaturezaOp, sNatOp As String
Dim X As New ClassFatGlobal, iTop As Integer
'Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TributacaoItem_TrazerTela

    mvariTrazendoTribItemTela = 1

    TabTrib_Controle("NaturezaOpItem").Text = objTributacaoItem.sNaturezaOp
    sNatOp = Trim(TabTrib_Controle("NaturezaOpItem").Text)
    If sNatOp <> "" Then
            
        objNaturezaOp.sCodigo = sNatOp
        lErro = CF("NaturezaOperacao_Le", objNaturezaOp)
        If lErro <> SUCESSO And lErro <> 17958 Then Error 56878
        If lErro <> SUCESSO Then Error 56880

        TabTrib_Controle("LabelDescrNatOpItem").Caption = objNaturezaOp.sDescricao
    Else
        TabTrib_Controle("LabelDescrNatOpItem").Caption = ""
    End If
    
'    objTipoDocInfo.iCodigo = mvarobjNFTributacao.iTipoNFiscal
'
'    'Lê o Tipo de Documento
'    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
'    If lErro <> SUCESSO And lErro <> 31415 Then Error 56878
    
    If gobjDocTribTributacao.iFinalidadeNFe = 4 Then 'Devolução
        TabTrib_Controle("FrameTabTribDet")(10).Enabled = True
    Else
        TabTrib_Controle("FrameTabTribDet")(10).Enabled = False
    End If

    objTributacaoTipo.iTipo = objTributacaoItem.iTipoTributacao
    If objTributacaoTipo.iTipo <> 0 Then

        lErro = CF("TipoTributacao_Le", objTributacaoTipo)
        If lErro <> SUCESSO Then Error 27806

        With mvarobjForm
        
            TabTrib_Controle("TipoTributacaoItem").Text = CStr(objTributacaoItem.iTipoTributacao)
            TabTrib_Controle("DescTipoTribItem").Caption = objTributacaoTipo.sDescricao
    
            'se nao incide IPI
            If objTributacaoTipo.iIPIIncide = 0 Then
                Call IPI_Frame_Habilita(False)
            Else
                Call IPI_Frame_Habilita(True)
            End If
    
'            'se nao incide ICMS
'            If objTributacaoTipo.iICMSIncide = 0 Then
'                Call ICMS_Frame_Habilita(False)
'            Else
'                Call ICMS_Frame_Habilita(True)
'            End If
            
'            'se nao incide ICMS
'            If objTributacaoTipo.iISSIncide = 0 Then
'                Call ISS_Frame_Habilita(False)
'            Else
'                Call ISS_Frame_Habilita(True)
'            End If
    
            'setar se credita ou nao IPI e ICMS
           TabTrib_Controle("IPICredita").Value = IIf(objTributacaoTipo.iIPICredita = 0, 0, 1)
           'TabTrib_Controle("ICMSCredita").Value = IIf(objTributacaoTipo.iICMSCredita = 0, 0, 1)
           TabTrib_Controle("ICMSCredita").Value = IIf(objTributacaoItem.dICMSCredito < DELTA_VALORMONETARIO, 0, 1)
'            If TabTrib_Controle("IPICredito").Enabled Then TabTrib_Controle("IPICredita").Value = objTributacaoTipo.iIPICredita
'            If TabTrib_Controle("ICMSCredito").Enabled Then TabTrib_Controle("ICMSCredita").Value = objTributacaoTipo.iICMSCredita

        End With
        
    Else
    
        TabTrib_Controle("TipoTributacaoItem").Text = ""
        TabTrib_Controle("DescTipoTribItem").Caption = ""
    
    End If

    With mvarobjForm
    
        'setar dados de ICMS
        Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSTipo"), objTributacaoItem.iICMSTipo)
        TabTrib_Controle("ICMSBaseItem").Text = CStr(objTributacaoItem.dICMSBase)
        TabTrib_Controle("ICMSPercRedBaseItem").Text = CStr(objTributacaoItem.dICMSPercRedBase * 100)
        TabTrib_Controle("ICMSAliquotaItem").Text = CStr(objTributacaoItem.dICMSAliquota * 100)
        TabTrib_Controle("ICMSValorItem").Text = CStr(objTributacaoItem.dICMSValor)
    
        'setar dados ICMS Substituicao
        TabTrib_Controle("ICMSSubstBaseItem").Text = CStr(objTributacaoItem.dICMSSubstBase)
        TabTrib_Controle("ICMSSubstAliquotaItem").Text = CStr(objTributacaoItem.dICMSSubstAliquota * 100)
        TabTrib_Controle("ICMSSubstValorItem").Text = CStr(objTributacaoItem.dICMSSubstValor)

        TabTrib_Controle("ICMSSTCobrAntBaseItem").Text = CStr(objTributacaoItem.dICMSSTCobrAntBase)
        TabTrib_Controle("ICMSSTCobrAntValorItem").Text = CStr(objTributacaoItem.dICMSSTCobrAntValor)

    End With

    For Each objTipoTribICMS In mvarcolTiposTribICMS
        If objTipoTribICMS.iTipo = objTributacaoItem.iICMSTipo Then
            'objTributacaoItem.sCST = CStr(objTributacaoItem.iOrigemMercadoria) & Format(objTipoTribICMS.iTipoTribCST, "00")
            Exit For
        End If
    Next
    If objTributacaoItem.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL Then
        TabTrib_Controle("ComboICMSTipo").Visible = True
        TabTrib_Controle("ComboICMSSimplesTipo").Visible = False
    Else
        For Each objTipoTribICMSSimples In mvarcolTiposTribICMSSimples
            If objTipoTribICMSSimples.iTipo = objTributacaoItem.iICMSSimplesTipo Then
                'objTributacaoItem.sCSOSN = CStr(objTributacaoItem.iOrigemMercadoria) & Format(objTipoTribICMSSimples.iCSOSN, "000")
                Exit For
            End If
        Next
        If objTipoTribICMSSimples Is Nothing Then Set objTipoTribICMSSimples = mvarcolTiposTribICMSSimples.Item(1)
        TabTrib_Controle("ComboICMSTipo").Visible = False
        TabTrib_Controle("ComboICMSSimplesTipo").Visible = True
    End If

    If objTributacaoTipo.iISSIncide = 0 Then
        Call ISS_Frame_Habilita(False)
    Else
        Call ISS_Frame_Habilita(True)
    End If
    
    With mvarobjForm
    
        'LAYOUT ICMS
        If objTributacaoItem.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL Then
            
'            If objTipoTribICMS.iPermiteAliquota = MARCADO Then
'                TabTrib_Controle("FrameICMS")(1).Visible = True
'                TabTrib_Controle("FrameICMS")(2).Visible = False
'            ElseIf objTipoTribICMS.iPermiteMotDesoneracao = MARCADO Then
'                TabTrib_Controle("FrameICMS")(1).Visible = False
'                TabTrib_Controle("FrameICMS")(2).Visible = True
'            Else
'                TabTrib_Controle("FrameICMS")(1).Visible = False
'                TabTrib_Controle("FrameICMS")(2).Visible = False
'            End If

            If objTipoTribICMS.iPermiteAliquota = MARCADO Then
                TabTrib_Controle("FrameICMS")(1).Visible = True
                TabTrib_Controle("FrameFCP")(0).Enabled = True
            Else
                TabTrib_Controle("FrameICMS")(1).Visible = False
                TabTrib_Controle("FrameFCP")(0).Enabled = False
            End If
            
            If objTipoTribICMS.iPermiteMotDesoneracao = MARCADO Then
                TabTrib_Controle("FrameICMS")(2).Visible = True
                TabTrib_Controle("FrameICMS")(1).Height = 1215
            Else
                TabTrib_Controle("FrameICMS")(2).Visible = False
                TabTrib_Controle("FrameICMS")(1).Height = 1575
            End If
                
            If objTipoTribICMS.iPermiteST = MARCADO Then
                If objTipoTribICMS.iPermiteSTRetAnt = MARCADO Then
                    TabTrib_Controle("FrameICMSST")(1).Visible = False
                    TabTrib_Controle("FrameICMSST")(2).Visible = True
                    TabTrib_Controle("FrameICMSST")(3).Visible = False
                    TabTrib_Controle("FrameFCP")(1).Enabled = False
                    TabTrib_Controle("FrameFCP")(2).Enabled = True
                ElseIf objTipoTribICMS.iPermiteSTRetUFRem = MARCADO Then
                    TabTrib_Controle("FrameICMSST")(1).Visible = False
                    TabTrib_Controle("FrameICMSST")(2).Visible = False
                    TabTrib_Controle("FrameICMSST")(3).Visible = True
                    TabTrib_Controle("FrameFCP")(1).Enabled = True
                    TabTrib_Controle("FrameFCP")(2).Enabled = False
                Else
                    TabTrib_Controle("FrameICMSST")(1).Visible = True
                    TabTrib_Controle("FrameICMSST")(2).Visible = False
                    TabTrib_Controle("FrameICMSST")(3).Visible = False
                    TabTrib_Controle("FrameFCP")(1).Enabled = True
                    TabTrib_Controle("FrameFCP")(2).Enabled = False
                End If
            Else
                TabTrib_Controle("FrameICMSST")(1).Visible = False
                TabTrib_Controle("FrameICMSST")(2).Visible = False
                TabTrib_Controle("FrameICMSST")(3).Visible = False
                TabTrib_Controle("FrameFCP")(1).Enabled = False
                TabTrib_Controle("FrameFCP")(2).Enabled = False
            
            End If
            
            For iIndice = 1 To 3
                If objTipoTribICMS.iPermiteAliquota = DESMARCADO And objTipoTribICMS.iPermiteST = MARCADO Then
                    TabTrib_Controle("FrameICMSST")(iIndice).top = TabTrib_Controle("FrameICMS")(1).top
                Else
                    TabTrib_Controle("FrameICMSST")(iIndice).top = 2025
                End If
            Next
            If objTipoTribICMS.iPermiteAliquota = DESMARCADO And objTipoTribICMS.iPermiteST = DESMARCADO Then
                    TabTrib_Controle("FrameICMS")(2).top = TabTrib_Controle("FrameICMS")(1).top
                Else
                    TabTrib_Controle("FrameICMS")(2).top = 1525
            End If
            
            If objTipoTribICMS.iPermiteReducaoBase = MARCADO Then
                TabTrib_Controle("FrameICMSRedBase").Visible = True
            Else
                TabTrib_Controle("FrameICMSRedBase").Visible = False
            End If
            
            If objTipoTribICMS.iPermitePercBaseOperProp = MARCADO Then
                TabTrib_Controle("FrameICMSBaseOperProp").Visible = True
                TabTrib_Controle("FrameICMS")(1).Height = 1575
            Else
                TabTrib_Controle("FrameICMSBaseOperProp").Visible = False
            End If
            
            If objTipoTribICMS.iPermiteUFICMSSTDevido = MARCADO Then
                TabTrib_Controle("FrameICMSSTUFDevido").Visible = True
            Else
                TabTrib_Controle("FrameICMSSTUFDevido").Visible = False
            End If
        
            If objTipoTribICMS.iTipoTribCST = 51 Then
                TabTrib_Controle("FrameICMS51").Visible = True
                TabTrib_Controle("FrameICMS")(1).Height = 1575
            Else
                TabTrib_Controle("FrameICMS51").Visible = False
            End If
        
        Else
        
            TabTrib_Controle("FrameICMS")(1).Height = 1230
            If objTipoTribICMSSimples.iPermiteCredito = MARCADO Then
                TabTrib_Controle("FrameICMSCred").Visible = True
                TabTrib_Controle("FrameICMS")(1).top = TabTrib_Controle("FrameICMSCred").top + TabTrib_Controle("FrameICMSCred").Height + 15
            Else
                TabTrib_Controle("FrameICMSCred").Visible = False
            End If
        
            TabTrib_Controle("FrameICMS")(2).Visible = False
            If objTipoTribICMSSimples.iPermiteAliquota = MARCADO Then
                TabTrib_Controle("FrameICMS")(1).Visible = True
                TabTrib_Controle("FrameFCP")(0).Enabled = True
            Else
                TabTrib_Controle("FrameICMS")(1).Visible = False
                TabTrib_Controle("FrameFCP")(0).Enabled = False
            End If
                
            If objTipoTribICMSSimples.iPermiteST = MARCADO Then
                If objTipoTribICMSSimples.iPermiteSTRetAnt = MARCADO Then
                    TabTrib_Controle("FrameICMSST")(1).Visible = False
                    TabTrib_Controle("FrameICMSST")(2).Visible = True
                    TabTrib_Controle("FrameICMSST")(3).Visible = False
                    TabTrib_Controle("FrameFCP")(1).Enabled = False
                    TabTrib_Controle("FrameFCP")(2).Enabled = True
                Else
                    TabTrib_Controle("FrameICMSST")(1).Visible = True
                    TabTrib_Controle("FrameICMSST")(2).Visible = False
                    TabTrib_Controle("FrameICMSST")(3).Visible = False
                    TabTrib_Controle("FrameFCP")(1).Enabled = True
                    TabTrib_Controle("FrameFCP")(2).Enabled = False
                End If
            Else
                TabTrib_Controle("FrameICMSST")(1).Visible = False
                TabTrib_Controle("FrameICMSST")(2).Visible = False
                TabTrib_Controle("FrameICMSST")(3).Visible = False
                TabTrib_Controle("FrameFCP")(1).Enabled = False
                TabTrib_Controle("FrameFCP")(2).Enabled = False
            End If
            
            For iIndice = 1 To 3
                If objTipoTribICMSSimples.iPermiteAliquota = DESMARCADO And objTipoTribICMSSimples.iPermiteST = MARCADO Then
                    TabTrib_Controle("FrameICMSST")(iIndice).top = TabTrib_Controle("FrameICMS")(1).top
                Else
                    TabTrib_Controle("FrameICMSST")(iIndice).top = 2025
                End If
            Next

            TabTrib_Controle("FrameICMSBaseOperProp").Visible = False
            TabTrib_Controle("FrameICMSSTUFDevido").Visible = False
            
        End If
        'LAYOUT ICMS
                    
        'setar dados de IPI
        Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboIPITipo"), objTributacaoItem.iIPITipo)
        TabTrib_Controle("IPIBaseItem").Text = CStr(objTributacaoItem.dIPIBaseCalculo)
        TabTrib_Controle("IPIPercRedBaseItem").Text = CStr(objTributacaoItem.dIPIPercRedBase * 100)
        TabTrib_Controle("IPIAliquotaItem").Text = CStr(objTributacaoItem.dIPIAliquota * 100)
        TabTrib_Controle("IPIValorItem").Text = CStr(objTributacaoItem.dIPIValor)
    
        For Each objTipoTribIPI In mvarcolTiposTribIPI
            If objTipoTribIPI.iTipo = objTributacaoItem.iIPITipo Then Exit For
        Next
        If objTipoTribIPI Is Nothing Then
            objTributacaoItem.iIPITipo = 6 'Se o tipo de IPI já não existe mais força ser Outras
            For Each objTipoTribIPI In mvarcolTiposTribIPI
                If objTipoTribIPI.iTipo = objTributacaoItem.iIPITipo Then Exit For
            Next
        End If
    
'        'se permite reducao de base habilitar este campo
'        If objTipoTribIPI.iPermiteReducaoBase Then
'            TabTrib_Controle("IPIPercRedBaseItem").Enabled = True
'        Else
'            'desabilita-lo e limpa-lo em caso contrario
'            TabTrib_Controle("IPIPercRedBaseItem").Enabled = False
'        End If
    
        'se permite aliquota habilitar este campo e valor.
        If objTipoTribIPI.iPermiteAliquota = MARCADO Then
            TabTrib_Controle("FrameIPIItem").Enabled = True
        Else
            'Desabilitar os dois campos e coloca-los com zero
            TabTrib_Controle("FrameIPIItem").Enabled = False
        End If

    End With
    
    TabTrib_Controle("ValorDescontoItem").Text = Format(objTributacaoItem.dDescontoGrid, "STANDARD")
    TabTrib_Controle("ValorFreteItem").Text = Format(objTributacaoItem.dValorFreteItem, "STANDARD")
    TabTrib_Controle("ValorSeguroItem").Text = Format(objTributacaoItem.dValorSeguroItem, "STANDARD")
    TabTrib_Controle("ValorDespesasItem").Text = Format(objTributacaoItem.dValorOutrasDespesasItem, "STANDARD")
    
    TabTrib_Controle("UMTrib").Caption = objTributacaoItem.sUMTrib
    TabTrib_Controle("QtdeTrib").Caption = Formata_Estoque(objTributacaoItem.dQtdTrib)
    TabTrib_Controle("ValorTrib").Caption = Format(objTributacaoItem.dValorUnitTrib, X.gobjFAT.sFormatoPrecoUnitario)
    Call Combo_Seleciona_ItemData(TabTrib_Controle("OrigemMercadoria"), objTributacaoItem.iOrigemMercadoria)
    TabTrib_Controle("FCI").Text = objTributacaoItem.sFCI
    Call Combo_Seleciona_ItemData(TabTrib_Controle("ICMSModalidade"), objTributacaoItem.iICMSBaseModalidade)
    Call ICMSModalidade_Click(False)
    Call Combo_Seleciona_ItemData(TabTrib_Controle("ICMSSTModalidade"), objTributacaoItem.iICMSSubstBaseModalidade)
    Call ICMSSTModalidade_Click(False)
    Call Combo_Seleciona_ItemData(TabTrib_Controle("IPITipoCalculo"), objTributacaoItem.iIPITipoCalculo)
    Call IPITipoCalculo_Click(False)
    
    Call Combo_Seleciona_ItemData(TabTrib_Controle("PISTipoCalculo"), objTributacaoItem.iPISTipoCalculo)
    Call PISTipoCalculo_Click(False)
    Call Combo_Seleciona_ItemData(TabTrib_Controle("PISSTTipoCalculo"), objTributacaoItem.iPISSTTipoCalculo)
    Call PISSTTipoCalculo_Click(False)
    Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboPISTipo"), objTributacaoItem.iPISTipo)
    Call ComboPISTipo_Click(False)
    
    Call Combo_Seleciona_ItemData(TabTrib_Controle("COFINSTipoCalculo"), objTributacaoItem.iCOFINSTipoCalculo)
    Call COFINSTipoCalculo_Click(False)
    Call Combo_Seleciona_ItemData(TabTrib_Controle("COFINSSTTipoCalculo"), objTributacaoItem.iCOFINSSTTipoCalculo)
    Call COFINSSTTipoCalculo_Click(False)
    If Len(Trim(objTributacaoItem.sISSQN)) = 0 Then
        TabTrib_Controle("ISSListaServ").ListIndex = -1
    Else
        Call CF("sCombo_Seleciona2", TabTrib_Controle("ISSListaServ"), objTributacaoItem.sISSQN)
    End If
    Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboCOFINSTipo"), objTributacaoItem.iCOFINSTipo)
    Call ComboCOFINSTipo_Click(False)

    TabTrib_Controle("ICMSSubstPercRedBaseItem").Text = CStr(objTributacaoItem.dICMSSubstPercRedBase * 100)
    TabTrib_Controle("ICMSSubstPercMVAItem").Text = CStr(objTributacaoItem.dICMSSubstPercMVA * 100)
    TabTrib_Controle("PISAliquotaItem").Text = CStr(objTributacaoItem.dPISAliquota * 100)
    TabTrib_Controle("PISSTAliquotaItem").Text = CStr(objTributacaoItem.dPISSTAliquota * 100)
    TabTrib_Controle("COFINSAliquotaItem").Text = CStr(objTributacaoItem.dCOFINSAliquota * 100)
    TabTrib_Controle("COFINSSTAliquotaItem").Text = CStr(objTributacaoItem.dCOFINSSTAliquota * 100)
    TabTrib_Controle("ISSAliquotaItem").Text = CStr(objTributacaoItem.dISSAliquota * 100)

    TabTrib_Controle("IPIUnidadePadraoValorItem").Text = Format(objTributacaoItem.dIPIUnidadePadraoValor, X.gobjFAT.sFormatoPrecoUnitario)
    TabTrib_Controle("IPIUnidadePadraoQtdeItem").Text = Formata_Estoque(objTributacaoItem.dIPIUnidadePadraoQtde)
    
    TabTrib_Controle("PISBaseItem").Text = Format(objTributacaoItem.dPISBase, "STANDARD")
    TabTrib_Controle("PISAliquotaRSItem").Text = Format(objTributacaoItem.dPISAliquotaValor, X.gobjFAT.sFormatoPrecoUnitario)
    TabTrib_Controle("PISQtdeItem").Text = Formata_Estoque(objTributacaoItem.dPISQtde)
    TabTrib_Controle("PISValorItem").Text = Format(objTributacaoItem.dPISValor, "STANDARD")
    TabTrib_Controle("PISSTBaseItem").Text = Format(objTributacaoItem.dPISSTBase, "STANDARD")
    TabTrib_Controle("PISSTAliquotaRSItem").Text = Format(objTributacaoItem.dPISSTAliquotaValor, "STANDARD")
    TabTrib_Controle("PISSTQtdeItem").Text = Formata_Estoque(objTributacaoItem.dPISSTQtde)
    TabTrib_Controle("PISSTValorItem").Text = Format(objTributacaoItem.dPISSTValor, "STANDARD")
    
    TabTrib_Controle("COFINSBaseItem").Text = Format(objTributacaoItem.dCOFINSBase, "STANDARD")
    TabTrib_Controle("COFINSAliquotaRSItem").Text = Format(objTributacaoItem.dCOFINSAliquotaValor, X.gobjFAT.sFormatoPrecoUnitario)
    TabTrib_Controle("COFINSQtdeItem").Text = Formata_Estoque(objTributacaoItem.dCOFINSQtde)
    TabTrib_Controle("COFINSValorItem").Text = Format(objTributacaoItem.dCOFINSValor, "STANDARD")
    TabTrib_Controle("COFINSSTBaseItem").Text = Format(objTributacaoItem.dCOFINSSTBase, "STANDARD")
    TabTrib_Controle("COFINSSTAliquotaRSItem").Text = Format(objTributacaoItem.dCOFINSSTAliquotaValor, "STANDARD")
    TabTrib_Controle("COFINSSTQtdeItem").Text = Formata_Estoque(objTributacaoItem.dCOFINSSTQtde)
    TabTrib_Controle("COFINSSTValorItem").Text = Format(objTributacaoItem.dCOFINSSTValor, "STANDARD")
    
    TabTrib_Controle("ISSBaseItem").Text = Format(objTributacaoItem.dISSBase, "STANDARD")
    TabTrib_Controle("ISSValorItem").Text = Format(objTributacaoItem.dISSValor, "STANDARD")
    TabTrib_Controle("IIBaseItem").Text = Format(objTributacaoItem.dIIBase, "STANDARD")
    TabTrib_Controle("IIDespAduItem").Text = Format(objTributacaoItem.dIIDespAduaneira, "STANDARD")
    TabTrib_Controle("IIIOFItem").Text = Format(objTributacaoItem.dIIIOF, "STANDARD")
    TabTrib_Controle("IIValorItem").Text = Format(objTributacaoItem.dIIValor, "STANDARD")
    
    If objTributacaoItem.iISSCodPais <> 0 Then
        Call Combo_Seleciona_ItemData(TabTrib_Controle("ISSCodPaisItem"), objTributacaoItem.iISSCodPais)
    Else
        TabTrib_Controle("ISSCodPaisItem").ListIndex = -1
    End If
        
    TabTrib_Controle("ISSCodIBGE").PromptInclude = False
    TabTrib_Controle("ISSCodIBGE").Text = objTributacaoItem.sISSCidadeIBGE
    TabTrib_Controle("ISSCodIBGE").PromptInclude = True
    Call ISSCodIBGE_Validate(bSGECancelDummy, False)
    
    TabTrib_Controle("IPIClasseEnq").Text = objTributacaoItem.sIPIEnquadramentoClasse
    TabTrib_Controle("IPICodEnq").Text = objTributacaoItem.sIPIEnquadramentoCodigo
    Call IPICodEnq_Validate(bSGECancelDummy)
    
    TabTrib_Controle("IPICodSelo").PromptInclude = False
    TabTrib_Controle("IPICodSelo").Text = objTributacaoItem.sIPISeloCodigo
    TabTrib_Controle("IPICodSelo").PromptInclude = True
    Call IPICodSelo_Validate(bSGECancelDummy, False)

    TabTrib_Controle("IPISeloQtde").PromptInclude = False
    TabTrib_Controle("IPISeloQtde").Text = CStr(objTributacaoItem.lIPISeloQtde)
    TabTrib_Controle("IPISeloQtde").PromptInclude = True
    Call IPISeloQtde_Validate(bSGECancelDummy, False)
    
    TabTrib_Controle("IPICNPJProdutor").PromptInclude = False
    TabTrib_Controle("IPICNPJProdutor").Text = objTributacaoItem.sIPICNPJProdutor
    TabTrib_Controle("IPICNPJProdutor").PromptInclude = True
    Call IPICNPJProdutor_Validate(bSGECancelDummy, False)
    
    Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboISSTipo"), objTributacaoItem.iISSTipo)
    For Each objTipoTribISS In mvarcolTiposTribISS
        If objTipoTribISS.iTipo = objTributacaoItem.iISSTipo Then Exit For
    Next
    'objTributacaoItem.sISSCST = objTipoTribISS.sTipoTribCST
    
    If objTributacaoItem.iRegimeTributario = REGIME_TRIBUTARIO_NORMAL Then
        TabTrib_Controle("ComboICMSSimplesTipo").Visible = False
        TabTrib_Controle("ComboICMSTipo").Visible = True
    Else
        TabTrib_Controle("ComboICMSSimplesTipo").Visible = True
        TabTrib_Controle("ComboICMSTipo").Visible = False
        If objTributacaoItem.iICMSSimplesTipo <> 0 Then
            Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSSimplesTipo"), objTributacaoItem.iICMSSimplesTipo)
        Else
            TabTrib_Controle("ComboICMSSimplesTipo").ListIndex = 0
        End If
    End If

    TabTrib_Controle("ICMSpercBaseOperacaoPropria").Text = CStr(objTributacaoItem.dICMSpercBaseOperacaoPropria * 100)
    
    TabTrib_Controle("ICMSUFDevidoST").Text = objTributacaoItem.sICMSUFDevidoST
    Call ICMSUFDevidoST_Validate(bSGECancelDummy, False)
    
    TabTrib_Controle("ICMSvBCSTRet").Text = Format(objTributacaoItem.dICMSvBCSTRet, "STANDARD")
    TabTrib_Controle("ICMSvICMSSTRet").Text = Format(objTributacaoItem.dICMSvICMSSTRet, "STANDARD")
    TabTrib_Controle("ICMSvBCSTDest").Text = Format(objTributacaoItem.dICMSvBCSTDest, "STANDARD")
    TabTrib_Controle("ICMSvICMSSTDest").Text = Format(objTributacaoItem.dICMSvICMSSTDest, "STANDARD")
    TabTrib_Controle("ICMSvCredSN").Text = Format(objTributacaoItem.dICMSvCredSN, "STANDARD")
    TabTrib_Controle("ICMSValorIsento").Text = Format(objTributacaoItem.dICMSValorIsento, "STANDARD")
    TabTrib_Controle("ICMSpCredSN").Text = CStr(objTributacaoItem.dICMSpCredSN * 100)

    If objTributacaoItem.iICMSMotivo <> 0 Then
        Call Combo_Seleciona_ItemData(TabTrib_Controle("ICMSMotivo"), objTributacaoItem.iICMSMotivo)
    Else
        TabTrib_Controle("ICMSMotivo").ListIndex = -1
    End If
    
    Call CF("sCombo_Seleciona2", TabTrib_Controle("NatBCCred"), objTributacaoItem.sNatBCCred)
    Call NatBCCred_Click(False)

    TabTrib_Controle("TotTribItem").Text = Format(objTributacaoItem.dTotTrib, "STANDARD")
    TabTrib_Controle("TotTribAliqFedItem").Text = CStr(objTributacaoItem.dTotTribAliqFed * 100)
    TabTrib_Controle("TotTribAliqEstItem").Text = CStr(objTributacaoItem.dTotTribAliqEst * 100)
    TabTrib_Controle("TotTribAliqMunicItem").Text = CStr(objTributacaoItem.dTotTribAliqMunic * 100)
    Call Combo_Seleciona_ItemData(TabTrib_Controle("TotTribItemTipo"), objTributacaoItem.iTotTribTipo)
    
    'nfe 3.10
    TabTrib_Controle("ICMS51ValorOpItem").Text = Format(objTributacaoItem.dICMS51ValorOp, "STANDARD")
    TabTrib_Controle("ICMSPercDiferItem").Text = CStr(objTributacaoItem.dICMSPercDifer * 100)
    TabTrib_Controle("ICMSValorDifItem").Text = Format(objTributacaoItem.dICMSValorDif, "STANDARD")
    
    TabTrib_Controle("ISSValorDeducaoItem").Text = Format(objTributacaoItem.dISSValorDeducao, "STANDARD")
    TabTrib_Controle("ISSValorOutrasRetItem").Text = Format(objTributacaoItem.dISSValorOutrasRet, "STANDARD")
    TabTrib_Controle("ISSValorDescIncondItem").Text = Format(objTributacaoItem.dISSValorDescIncond, "STANDARD")
    TabTrib_Controle("ISSValorDescCondItem").Text = Format(objTributacaoItem.dISSValorDescCond, "STANDARD")
    TabTrib_Controle("ISSValorRetItem").Text = Format(objTributacaoItem.dISSValorRet, "STANDARD")
    
    If objTributacaoItem.iISSIndExigibilidade <> 0 Then
        Call Combo_Seleciona_ItemData(TabTrib_Controle("ISSIndExigibilidadeItem"), objTributacaoItem.iISSIndExigibilidade)
    Else
        TabTrib_Controle("ISSIndExigibilidadeItem").ListIndex = -1
    End If
    
    TabTrib_Controle("ISSMunicIncidImpItem").PromptInclude = False
    TabTrib_Controle("ISSMunicIncidImpItem").Text = objTributacaoItem.sISSMunicIncidImp
    TabTrib_Controle("ISSMunicIncidImpItem").PromptInclude = True
    Call ISSMunicIncidImpItem_Validate(bSGECancelDummy, False)
    
    TabTrib_Controle("ISSCodServItem").PromptInclude = False
    TabTrib_Controle("ISSCodServItem").Text = objTributacaoItem.sISSCodServ
    TabTrib_Controle("ISSCodServItem").PromptInclude = True
    Call ISSCodServItem_Validate(bSGECancelDummy, False)
    
    If TabTrib_Controle("ISSIndIncentivoItem").Value <> objTributacaoItem.iISSIndIncentivo Then
        TabTrib_Controle("ISSIndIncentivoItem").Value = objTributacaoItem.iISSIndIncentivo
    End If
    
    TabTrib_Controle("ISSNumProcessoItem").Text = objTributacaoItem.sISSNumProcesso
    
    
    Call Combo_Seleciona_ItemData(TabTrib_Controle("ICMSInterestAliqItem"), CInt(objTributacaoItem.dICMSInterestAliq * 100))
    Call Combo_Seleciona_ItemData(TabTrib_Controle("ICMSInterestPercPartilhaItem"), CInt(objTributacaoItem.dICMSInterestPercPartilha * 100))
    
    TabTrib_Controle("ICMSInterestBCUFDestItem").Text = CStr(objTributacaoItem.dICMSInterestBCUFDest)
    TabTrib_Controle("ICMSInterestPercFCPUFDestItem").Text = CStr(objTributacaoItem.dICMSInterestPercFCPUFDest * 100)
    TabTrib_Controle("ICMSInterestAliqUFDestItem").Text = CStr(objTributacaoItem.dICMSInterestAliqUFDest * 100)
    TabTrib_Controle("ICMSInterestVlrFCPUFDestItem").Text = CStr(objTributacaoItem.dICMSInterestVlrFCPUFDest)
    TabTrib_Controle("ICMSInterestVlrUFDestItem").Text = CStr(objTributacaoItem.dICMSInterestVlrUFDest)
    TabTrib_Controle("ICMSInterestVlrUFRemetItem").Text = CStr(objTributacaoItem.dICMSInterestVlrUFRemet)
    
    'UA. Tributos Devolvidos (para o item da NF-e)
    TabTrib_Controle("pDevolItem").Text = CStr(objTributacaoItem.dpDevol * 100)
    TabTrib_Controle("IPIVlrDevolvidoItem").Text = CStr(objTributacaoItem.dIPIVlrDevolvido)
    
    'NFe 4.00
    TabTrib_Controle("cBenefItem").Text = objTributacaoItem.scBenef
    
    TabTrib_Controle("ICMSInterestBCFCPUFDestItem").Text = CStr(objTributacaoItem.dICMSInterestBCFCPUFDest)
    TabTrib_Controle("ICMSSTCobrAntAliquotaItem").Text = CStr(objTributacaoItem.dICMSSTCobrAntAliquota * 100)
    
    TabTrib_Controle("ICMSvBCFCPItem").Text = CStr(objTributacaoItem.dICMSvBCFCP)
    TabTrib_Controle("ICMSpFCPItem").Text = CStr(objTributacaoItem.dICMSpFCP * 100)
    TabTrib_Controle("ICMSvFCPItem").Text = CStr(objTributacaoItem.dICMSvFCP)
    
    TabTrib_Controle("ICMSvBCFCPSTItem").Text = CStr(objTributacaoItem.dICMSvBCFCPST)
    TabTrib_Controle("ICMSpFCPSTItem").Text = CStr(objTributacaoItem.dICMSpFCPST * 100)
    TabTrib_Controle("ICMSvFCPSTItem").Text = CStr(objTributacaoItem.dICMSvFCPST)
    
    TabTrib_Controle("ICMSvBCFCPSTRetItem").Text = CStr(objTributacaoItem.dICMSvBCFCPSTRet)
    TabTrib_Controle("ICMSpFCPSTRetItem").Text = CStr(objTributacaoItem.dICMSpFCPSTRet * 100)
    TabTrib_Controle("ICMSvFCPSTRetItem").Text = CStr(objTributacaoItem.dICMSvFCPSTRet)
    
    If objTributacaoItem.iICMSSTBaseDupla = MARCADO Then
        TabTrib_Controle("ICMSSTBaseDupla").Value = vbChecked
    Else
        TabTrib_Controle("ICMSSTBaseDupla").Value = vbUnchecked
    End If
    
    If objTributacaoItem.iICMSInterestBaseDupla = MARCADO Then
        TabTrib_Controle("ICMSInterestBaseDupla").Value = vbChecked
    Else
        TabTrib_Controle("ICMSInterestBaseDupla").Value = vbUnchecked
    End If
        
    mvariNatOpItemAlterado = 0
    mvariTipoTributacaoItemAlterado = 0
    mvariICMSBaseItemAlterado = 0
    mvariICMSPercRedBaseItemAlterado = 0
    mvariICMSAliquotaItemAlterado = 0
    mvariICMSValorItemAlterado = 0
    mvariICMSSubstBaseItemAlterado = 0
    mvariICMSSubstAliquotaItemAlterado = 0
    mvariICMSSubstValorItemAlterado = 0
    mvariIPIBaseItemAlterado = 0
    mvariIPIPercRedBaseItemAlterado = 0
    mvariIPIAliquotaItemAlterado = 0
    mvariIPIValorItemAlterado = 0
    mvariICMSSTCobrAntBaseItemAlterado = 0
    mvariICMSSTCobrAntValorItemAlterado = 0
    
    mvariCOFINSSTAliquotaValorItemAlterado = 0
    mvariCOFINSSTAliquotaItemAlterado = 0
    mvariCOFINSSTValorItemAlterado = 0
    mvariCOFINSSTBaseItemAlterado = 0
    mvariCOFINSSTQtdeItemAlterado = 0
    mvariCOFINSAliquotaValorItemAlterado = 0
    mvariCOFINSAliquotaItemAlterado = 0
    mvariCOFINSValorItemAlterado = 0
    mvariCOFINSBaseItemAlterado = 0
    mvariCOFINSQtdeItemAlterado = 0
    mvariPISAliquotaValorItemAlterado = 0
    mvariPISAliquotaItemAlterado = 0
    mvariPISValorItemAlterado = 0
    mvariPISBaseItemAlterado = 0
    mvariPISQtdeItemAlterado = 0
    mvariPISSTAliquotaValorItemAlterado = 0
    mvariPISSTAliquotaItemAlterado = 0
    mvariPISSTValorItemAlterado = 0
    mvariPISSTBaseItemAlterado = 0
    mvariPISSTQtdeItemAlterado = 0
    mvariISSCidadeIBGEItemAlterado = 0
    mvariISSAliquotaItemAlterado = 0
    mvariISSValorItemAlterado = 0
    mvariISSBaseItemAlterado = 0
    mvariValorFreteItemAlterado = 0
    mvariValorSeguroItemAlterado = 0
    mvariValorDescontoItemAlterado = 0
    mvariValorOutrasDespesasItemAlterado = 0
    mvariIIIOFItemAlterado = 0
    mvariIIBaseItemAlterado = 0
    mvariIIValorItemAlterado = 0
    mvariIIDespAduaneiraItemAlterado = 0
    mvariICMSSubstPercRedBaseItemAlterado = 0
    mvariICMSSubstPercMVAItemAlterado = 0
    mvariIPIUnidadePadraoValorItemAlterado = 0
    mvariIPIUnidadePadraoQtdeItemAlterado = 0
    mvariFCIAlterado = 0
    
    mvariICMSpercBaseOperacaoPropriaAlterado = 0
    mvariICMSUFDevidoSTAlterado = 0
    mvariICMSvBCSTRetAlterado = 0
    mvariICMSvICMSSTRetAlterado = 0
    mvariICMSvBCSTDestAlterado = 0
    mvariICMSvICMSSTDestAlterado = 0
    mvariICMSpCredSNAlterado = 0
    mvariICMSvCredSNAlterado = 0
    mvariICMSValorIsentoAlterado = 0
    mvariICMSMotivoAlterado = 0
    mvariICMSSimplesTipoAlterado = 0
    mvariISSTipoAlterado = 0
    mvariTotTribItemAlterado = 0
    mvariTotTribFedItemAlterado = 0
    mvariTotTribEstItemAlterado = 0
    mvariTotTribMunicItemAlterado = 0
    mvariTotTribItemTipoAlterado = 0
    
    'nfe 3.10
    mvariICMS51ValorOpItemAlterado = 0
    mvariICMSPercDiferItemAlterado = 0
    mvariICMSValorDifItemAlterado = 0
    mvariISSValorDeducaoItemAlterado = 0
    mvariISSValorOutrasRetItemAlterado = 0
    mvariISSValorDescIncondItemAlterado = 0
    mvariISSValorDescCondItemAlterado = 0
    mvariISSValorRetItemAlterado = 0
    mvariISSCodServItemAlterado = 0
    mvariISSMunicIncidImpItemAlterado = 0
    mvariISSCodPaisItemAlterado = 0
    mvariISSNumProcessoItemAlterado = 0
    
    'NT 2015.003 (ICMS Interestadual)
    mvariICMSInterestPercFCPUFDestItemAlterado = 0
    mvariICMSInterestBCUFDestItemAlterado = 0
    mvariICMSInterestAliqUFDestItemAlterado = 0
    mvariICMSInterestVlrUFDestItemAlterado = 0
    mvariICMSInterestVlrUFRemetItemAlterado = 0
    mvariICMSInterestVlrFCPUFDestItemAlterado = 0
    
    'UA. Tributos Devolvidos (para o item da NF-e)
    mvaripDevolItemAlterado = 0
    mvariIPIVlrDevolvidoItemAlterado = 0
    
    'nfe 4.0
    mvariICMSInterestBCFCPUFDestItemAlterado = 0
    mvariICMSpFCPSTItemAlterado = 0
    mvariICMSvFCPSTItemAlterado = 0
    mvariICMSvBCFCPSTRetItemAlterado = 0
    mvariICMSpFCPSTRetItemAlterado = 0
    mvariICMSvFCPSTRetItemAlterado = 0
    mvariICMSSTCobrAntAliquotaItemAlterado = 0
    mvariICMSvBCFCPItemAlterado = 0
    mvariICMSpFCPItemAlterado = 0
    mvariICMSvFCPItemAlterado = 0

    mvariTrazendoTribItemTela = 0

    TributacaoItem_TrazerTela = SUCESSO

    Exit Function

Erro_TributacaoItem_TrazerTela:

    mvariTrazendoTribItemTela = 0

    TributacaoItem_TrazerTela = Err

    Select Case Err

        Case 27806, 56878

        Case 56880
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_INEXISTENTE", Err, sNatOp)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153730)

    End Select

    Exit Function

End Function

Private Function Tributacao_TrazerTela(objTributacao As Object) As Long

On Error GoTo Erro_Tributacao_TrazerTela

    If iSubTipoTela = SUBTIPOTELA_CONSULTA Then
        TabTrib_Controle("ValorFrete").Caption = Format(objTributacao.dValorFrete, "STANDARD")
        TabTrib_Controle("ValorSeguro").Caption = Format(objTributacao.dValorSeguro, "STANDARD")
        TabTrib_Controle("ValorDesconto").Caption = Format(objTributacao.dValorDesconto, "STANDARD")
        TabTrib_Controle("ValorDespesas").Caption = Format(objTributacao.dValorOutrasDespesas, "STANDARD")
    Else
        TabTrib_Controle("ValorFrete").Text = Format(objTributacao.dValorFrete, "STANDARD")
        TabTrib_Controle("ValorSeguro").Text = Format(objTributacao.dValorSeguro, "STANDARD")
        TabTrib_Controle("ValorDesconto").Text = Format(objTributacao.dValorDesconto, "STANDARD")
        TabTrib_Controle("ValorDespesas").Text = Format(objTributacao.dValorOutrasDespesas, "STANDARD")
    End If
    
    Tributacao_TrazerTela = SUCESSO

    Exit Function

Erro_Tributacao_TrazerTela:

    Tributacao_TrazerTela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153730)

    End Select

    Exit Function
    
End Function

Function AtualizarTributacao(Optional ByVal objAux As Object) As Long

Dim lErro As Long

On Error GoTo Erro_AtualizarTributacao

    If iTipoTela = TIPOTELA_NF Then
        'para notas de importacao
        If Not (objAux Is Nothing) Then Set mvarobjNFTributacao.objNFiscal = objAux
    End If
        
    lErro = gobjTributacao.AtualizaImpostos(gobjDocTrib, mvariRecalculandoTributacao)
    If lErro <> SUCESSO Then Error 27807
    
'    If iTipoTela = TIPOTELA_NF Then
'        Call AtualizarMsg
'    End If

    'jogar dados do obj atualizado p/a tela
    lErro = Carrega_Tab_Tributacao(gobjDocTribTributacao)
    If lErro <> SUCESSO Then Error 27808

    If mvariLivroFiscal <> 0 And iTipoTela = TIPOTELA_NF Then
        
        If Not (mvarobjForm.gobjNFiscal Is Nothing) Then
        
            'Mover a parte de tributação
            lErro = PreencherNF(mvarobjForm.gobjNFiscal)
            If lErro <> SUCESSO Then Error 27808
        
            'Repreenche as Linhas do livro de Acordo com a Tributacao
            lErro = CF("Repreenche_Linha_LivroFiscal", mvarobjForm.gobjNFiscal, mvarobjForm.gobjLivRegES, mvarobjForm)
            If lErro <> SUCESSO Then Error 27808
    
        End If
    
    End If
    
    If iTipoTela = TIPOTELA_NF Then
        If mvarobjNFTributacao.iTipoNFiscal = DOCINFO_NFIEDV Then
            lErro = Tributacao_TrazerTela(gobjDocTrib)
            If lErro <> SUCESSO Then Error 27808
        End If
    End If
    
    AtualizarTributacao = SUCESSO

    Exit Function

Erro_AtualizarTributacao:

    AtualizarTributacao = Err

    Select Case Err

        Case 27807, 27808

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153731)

    End Select

    Exit Function

End Function

Private Function Tributacao_GravarTela() As Long
'transfere dados de tributacao da tela para mvarobjNFTributacao
'os dados que estiverem diferentes devem ser marcados como "manuais"
Dim iIndice As Integer, iTemp As Integer, dTemp As Double, dtTemp As Date
Dim lErro As Long

On Error GoTo Erro_Tributacao_GravarTela

    With mvarobjForm
    
        'setar NatOpInterna
        If gobjDocTrib.sNaturezaOp <> TabTrib_Controle("NatOpInterna").Text Then
            gobjDocTrib.sNaturezaOp = TabTrib_Controle("NatOpInterna").Text
            gobjDocTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
        End If
        
        iTemp = StrParaInt(TabTrib_Controle("TipoTributacao").Text)

        If iTemp <> gobjDocTribTributacao.iTipoTributacao Then
            gobjDocTribTributacao.iTipoTributacao = iTemp
            gobjDocTribTributacao.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        End If
    
        'setar dados de ISS
        iTemp = TabTrib_Controle("ISSIncluso").Value
        If iTemp <> gobjDocTribTributacao.iISSIncluso Then
            gobjDocTribTributacao.iISSIncluso = iTemp
        End If
    
'        If TabTrib_Controle("ISSValor").Caption <> CStr(gobjDocTribTributacao.dISSValor) Then
'            dTemp = StrParaDbl(TabTrib_Controle("ISSValor").Caption)
'            If gobjDocTribTributacao.dISSValor <> dTemp Then
'                gobjDocTribTributacao.dISSValor = dTemp
'                gobjDocTribTributacao.iISSValorManual = VAR_PREENCH_MANUAL
'            End If
'        End If
    
        'setar dados de IR
        If TabTrib_Controle("IRBase").Text <> Format(gobjDocTribTributacao.dIRRFBase, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("IRBase").Text)
            If gobjDocTribTributacao.dIRRFBase <> dTemp Then
                gobjDocTribTributacao.dIRRFBase = dTemp
                gobjDocTribTributacao.iIRRFBaseManual = VAR_PREENCH_MANUAL
            End If
        End If

        If TabTrib_Controle("IRAliquota").Text <> CStr(gobjDocTribTributacao.dIRRFAliquota * 100) Then
            dTemp = StrParaDbl(TabTrib_Controle("IRAliquota").Text) / 100
            If gobjDocTribTributacao.dIRRFAliquota <> dTemp Then
                gobjDocTribTributacao.dIRRFAliquota = dTemp
                gobjDocTribTributacao.iIRRFAliquotaManual = VAR_PREENCH_MANUAL
            End If
        End If
    
        If TabTrib_Controle("ValorIRRF").Text <> Format(gobjDocTribTributacao.dIRRFValor, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("ValorIRRF").Text)
            If gobjDocTribTributacao.dIRRFValor <> dTemp Then
                gobjDocTribTributacao.dIRRFValor = dTemp
                gobjDocTribTributacao.iIRRFValorManual = VAR_PREENCH_MANUAL
            End If
        End If

        'setar dados de INSS
        If TabTrib_Controle("INSSBase").Text <> Format(gobjDocTribTributacao.dINSSBase, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("INSSBase").Text)
            If gobjDocTribTributacao.dINSSBase <> dTemp Then
                gobjDocTribTributacao.dINSSBase = dTemp
                gobjDocTribTributacao.iINSSBaseManual = VAR_PREENCH_MANUAL
            End If
        End If
    
        If TabTrib_Controle("INSSDeducoes").Text <> Format(gobjDocTribTributacao.dINSSDeducoes, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("INSSDeducoes").Text)
            If gobjDocTribTributacao.dINSSDeducoes <> dTemp Then
                gobjDocTribTributacao.dINSSDeducoes = dTemp
                gobjDocTribTributacao.iINSSDeducoesManual = VAR_PREENCH_MANUAL
            End If
        End If
    
        If TabTrib_Controle("INSSValor").Text <> Format(gobjDocTribTributacao.dINSSValor, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("INSSValor").Text)
            If gobjDocTribTributacao.dINSSValor <> dTemp Then
                gobjDocTribTributacao.dINSSValor = dTemp
                gobjDocTribTributacao.iINSSValorManual = VAR_PREENCH_MANUAL
            End If
        End If
    
        iTemp = TabTrib_Controle("INSSRetido").Value
        If iTemp <> gobjDocTribTributacao.iINSSRetido Then
            gobjDocTribTributacao.iINSSRetido = iTemp
            gobjDocTribTributacao.iINSSRetidoManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("PISRetido").Text <> Format(gobjDocTribTributacao.dPISRetido, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("PISRetido").Text)
            If gobjDocTribTributacao.dPISRetido <> dTemp Then
                gobjDocTribTributacao.dPISRetido = dTemp
                gobjDocTribTributacao.iPISRetidoManual = VAR_PREENCH_MANUAL
            End If
        End If

'        If TabTrib_Controle("PISValor").Caption <> CStr(gobjDocTribTributacao.dPISCredito) Then
'            dTemp = StrParaDbl(TabTrib_Controle("PISValor").Caption)
'            If gobjDocTribTributacao.dPISCredito <> dTemp Then
'                gobjDocTribTributacao.dPISCredito = dTemp
'                gobjDocTribTributacao.iPISCreditoManual = VAR_PREENCH_MANUAL
'            End If
'        End If

        If TabTrib_Controle("ISSRetido").Text <> Format(gobjDocTribTributacao.dISSRetido, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSRetido").Text)
            If gobjDocTribTributacao.dISSRetido <> dTemp Then
                gobjDocTribTributacao.dISSRetido = dTemp
                gobjDocTribTributacao.iISSRetidoManual = VAR_PREENCH_MANUAL
            End If
        End If

        If TabTrib_Controle("COFINSRetido").Text <> Format(gobjDocTribTributacao.dCOFINSRetido, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSRetido").Text)
            If gobjDocTribTributacao.dCOFINSRetido <> dTemp Then
                gobjDocTribTributacao.dCOFINSRetido = dTemp
                gobjDocTribTributacao.iCOFINSRetidoManual = VAR_PREENCH_MANUAL
            End If
        End If

'        If TabTrib_Controle("COFINSValor").Caption <> CStr(gobjDocTribTributacao.dCOFINSCredito) Then
'            dTemp = StrParaDbl(TabTrib_Controle("COFINSValor").Caption)
'            If gobjDocTribTributacao.dCOFINSCredito <> dTemp Then
'                gobjDocTribTributacao.dCOFINSCredito = dTemp
'                gobjDocTribTributacao.iCOFINSCreditoManual = VAR_PREENCH_MANUAL
'            End If
'        End If

        If TabTrib_Controle("CSLLRetido").Text <> Format(gobjDocTribTributacao.dCSLLRetido, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("CSLLRetido").Text)
            If gobjDocTribTributacao.dCSLLRetido <> dTemp Then
                gobjDocTribTributacao.dCSLLRetido = dTemp
                gobjDocTribTributacao.iCSLLRetidoManual = VAR_PREENCH_MANUAL
            End If
        End If

        If TabTrib_Controle("TotTrib").Text <> Format(gobjDocTribTributacao.dTotTrib, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTrib").Text)
            If gobjDocTribTributacao.dTotTrib <> dTemp Then
                gobjDocTribTributacao.dTotTrib = dTemp
                gobjDocTribTributacao.iTotTribManual = VAR_PREENCH_MANUAL
            End If
        End If
        
        If TabTrib_Controle("TotTribFed").Text <> Format(gobjDocTribTributacao.dTotTribFed, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTribFed").Text)
            If gobjDocTribTributacao.dTotTribFed <> dTemp Then
                gobjDocTribTributacao.dTotTribFed = dTemp
                gobjDocTribTributacao.iTotTribFedManual = VAR_PREENCH_MANUAL
            End If
        End If
        
        If TabTrib_Controle("TotTribEst").Text <> Format(gobjDocTribTributacao.dTotTribEst, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTribEst").Text)
            If gobjDocTribTributacao.dTotTribEst <> dTemp Then
                gobjDocTribTributacao.dTotTribEst = dTemp
                gobjDocTribTributacao.iTotTribEstManual = VAR_PREENCH_MANUAL
            End If
        End If
        
        If TabTrib_Controle("TotTribMunic").Text <> Format(gobjDocTribTributacao.dTotTribMunic, "Standard") Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTribMunic").Text)
            If gobjDocTribTributacao.dTotTribMunic <> dTemp Then
                gobjDocTribTributacao.dTotTribMunic = dTemp
                gobjDocTribTributacao.iTotTribMunicManual = VAR_PREENCH_MANUAL
            End If
        End If

        iTemp = TabTrib_Controle("TotTribTipo").ItemData(TabTrib_Controle("TotTribTipo").ListIndex)
        If iTemp <> gobjDocTribTributacao.iTotTribTipo Then
            gobjDocTribTributacao.iTotTribTipo = iTemp
            gobjDocTribTributacao.iTotTribTipoManual = VAR_PREENCH_MANUAL
        End If
        
        'nfe 3.10
        iTemp = TabTrib_Controle("IndConsumidorFinal").Value
        If iTemp <> gobjDocTribTributacao.iIndConsumidorFinal Then
            gobjDocTribTributacao.iIndConsumidorFinal = iTemp
            gobjDocTribTributacao.iIndConsumidorFinalManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("IndPresenca").ListIndex <> -1 Then
            iTemp = TabTrib_Controle("IndPresenca").ItemData(TabTrib_Controle("IndPresenca").ListIndex)
            If iTemp <> gobjDocTribTributacao.iIndPresenca Then
                gobjDocTribTributacao.iIndPresenca = iTemp
                gobjDocTribTributacao.iIndPresencaManual = VAR_PREENCH_MANUAL
            End If
        End If
        
        dtTemp = MaskedParaDate(TabTrib_Controle("DataPrestServico"))
        If dtTemp <> gobjDocTribTributacao.dtDataPrestServico Then
            gobjDocTribTributacao.dtDataPrestServico = dtTemp
            gobjDocTribTributacao.iDataPrestServicoManual = VAR_PREENCH_MANUAL
        End If
        
'        'nfe 4.00
'        If TabTrib_Controle("IPIVlrDevolvido").Text <> Format(gobjDocTribTributacao.dIPIVlrDevolvido, "Standard") Then
'            dTemp = StrParaDbl(TabTrib_Controle("IPIVlrDevolvido").Text)
'            If gobjDocTribTributacao.dIPIVlrDevolvido <> dTemp Then
'                gobjDocTribTributacao.dIPIVlrDevolvido = dTemp
'                gobjDocTribTributacao.iIPIVlrDevolvidoManual = VAR_PREENCH_MANUAL
'            End If
'        End If
        
    End With
    
    Tributacao_GravarTela = SUCESSO

    Exit Function

Erro_Tributacao_GravarTela:

    Tributacao_GravarTela = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153732)

    End Select

    Exit Function

End Function

Private Function TributacaoItem_GravarTela(objTributacaoItem As Object) As Long
'transfere dados de tributacao de um item da tela para objTributacaoItem
'os dados que estiverem diferentes devem ser marcados como "manuais"
Dim iIndice As Integer, iTemp As Integer, dTemp As Double, sTemp As String
Dim lErro As Long

On Error GoTo Erro_TributacaoItem_GravarTela

    With mvarobjForm
    
        sTemp = Trim(TabTrib_Controle("NaturezaOpItem").Text)
        If Trim(objTributacaoItem.sNaturezaOp) <> sTemp Then
            objTributacaoItem.sNaturezaOp = sTemp
            objTributacaoItem.iNaturezaOpManual = VAR_PREENCH_MANUAL
        End If
    
        iTemp = StrParaInt(TabTrib_Controle("TipoTributacaoItem").Text)
        If iTemp <> objTributacaoItem.iTipoTributacao Then
            objTributacaoItem.iTipoTributacao = iTemp
            objTributacaoItem.iTipoTributacaoManual = VAR_PREENCH_MANUAL
        End If
    
        'setar dados de ICMS
    
        iTemp = TabTrib_Controle("ComboICMSTipo").ItemData(TabTrib_Controle("ComboICMSTipo").ListIndex)
        If iTemp <> objTributacaoItem.iICMSTipo Then
            objTributacaoItem.iICMSTipo = iTemp
            objTributacaoItem.iICMSTipoManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("ICMSBaseItem").Text <> CStr(objTributacaoItem.dICMSBase) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSBaseItem").Text)
            objTributacaoItem.dICMSBase = dTemp
            objTributacaoItem.iICMSBaseManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("ICMSPercRedBaseItem").Text <> CStr(objTributacaoItem.dICMSPercRedBase * 100) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSPercRedBaseItem").Text) / 100
            objTributacaoItem.dICMSPercRedBase = dTemp
            objTributacaoItem.iICMSPercRedBaseManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("ICMSAliquotaItem").Text <> CStr(objTributacaoItem.dICMSAliquota * 100) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSAliquotaItem").Text) / 100
            objTributacaoItem.dICMSAliquota = dTemp
            objTributacaoItem.iICMSAliquotaManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("ICMSValorItem").Text <> CStr(objTributacaoItem.dICMSValor) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSValorItem").Text)
            objTributacaoItem.dICMSValor = dTemp
            objTributacaoItem.iICMSValorManual = VAR_PREENCH_MANUAL
        End If
    
        'setar dados ICMS Substituicao
    
        If TabTrib_Controle("ICMSSubstBaseItem").Text <> CStr(objTributacaoItem.dICMSSubstBase) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSubstBaseItem").Text)
            objTributacaoItem.dICMSSubstBase = dTemp
            objTributacaoItem.iICMSSubstBaseManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("ICMSSubstAliquotaItem").Text <> CStr(objTributacaoItem.dICMSSubstAliquota * 100) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSubstAliquotaItem").Text) / 100
            objTributacaoItem.dICMSSubstAliquota = dTemp
            objTributacaoItem.iICMSSubstAliquotaManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("ICMSSubstValorItem").Text <> CStr(objTributacaoItem.dICMSSubstValor) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSubstValorItem").Text)
            objTributacaoItem.dICMSSubstValor = dTemp
            objTributacaoItem.iICMSSubstValorManual = VAR_PREENCH_MANUAL
        End If
                
        If TabTrib_Controle("ICMSSTCobrAntBaseItem").Text <> CStr(objTributacaoItem.dICMSSTCobrAntBase) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSTCobrAntBaseItem").Text)
            objTributacaoItem.dICMSSTCobrAntBase = dTemp
            objTributacaoItem.iICMSSTCobrAntBaseManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("ICMSSTCobrAntValorItem").Text <> CStr(objTributacaoItem.dICMSSTCobrAntValor) Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSTCobrAntValorItem").Text)
            objTributacaoItem.dICMSSTCobrAntValor = dTemp
            objTributacaoItem.iICMSSTCobrAntValorManual = VAR_PREENCH_MANUAL
        End If
    
        'setar dados de IPI
        iTemp = TabTrib_Controle("ComboIPITipo").ItemData(TabTrib_Controle("ComboIPITipo").ListIndex)
        If iTemp <> objTributacaoItem.iIPITipo Then
            objTributacaoItem.iIPITipo = iTemp
            objTributacaoItem.iIPITipoManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("IPIBaseItem").Text <> CStr(objTributacaoItem.dIPIBaseCalculo) Then
            dTemp = StrParaDbl(TabTrib_Controle("IPIBaseItem").Text)
            objTributacaoItem.dIPIBaseCalculo = dTemp
            objTributacaoItem.iIPIBaseManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("IPIPercRedBaseItem").Text <> CStr(objTributacaoItem.dIPIPercRedBase * 100) Then
            dTemp = StrParaDbl(TabTrib_Controle("IPIPercRedBaseItem").Text) / 100
            objTributacaoItem.dIPIPercRedBase = dTemp
            objTributacaoItem.iIPIPercRedBaseManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("IPIAliquotaItem").Text <> CStr(objTributacaoItem.dIPIAliquota * 100) Then
            dTemp = StrParaDbl(TabTrib_Controle("IPIAliquotaItem").Text) / 100
            objTributacaoItem.dIPIAliquota = dTemp
            objTributacaoItem.iIPIAliquotaManual = VAR_PREENCH_MANUAL
        End If
    
        If TabTrib_Controle("IPIValorItem").Text <> CStr(objTributacaoItem.dIPIValor) Then
            dTemp = StrParaDbl(TabTrib_Controle("IPIValorItem").Text)
            objTributacaoItem.dIPIValor = dTemp
            objTributacaoItem.iIPIValorManual = VAR_PREENCH_MANUAL
        End If
        'XXXX
        'INÍCIO - NOVOS CAMPOS
        If Abs(StrParaDbl(TabTrib_Controle("ValorFreteItem").Text) - objTributacaoItem.dValorFreteItem) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ValorFreteItem").Text)
            objTributacaoItem.dValorFreteItem = dTemp
            objTributacaoItem.iValorFreteItemManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ValorSeguroItem").Text) - objTributacaoItem.dValorSeguroItem) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ValorSeguroItem").Text)
            objTributacaoItem.dValorSeguroItem = dTemp
            objTributacaoItem.iValorSeguroItemManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ValorDespesasItem").Text) - objTributacaoItem.dValorOutrasDespesasItem) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ValorDespesasItem").Text)
            objTributacaoItem.dValorOutrasDespesasItem = dTemp
            objTributacaoItem.iValorOutrasDespesasItemManual = VAR_PREENCH_MANUAL
        End If
        
        'If Abs(StrParaDbl(TabTrib_Controle("ValorDescontoItem").Text) - objTributacaoItem.dValorDescontoItem) > DELTA_VALORMONETARIO Then
            'dTemp = StrParaDbl(TabTrib_Controle("ValorDescontoItem").Text)
            objTributacaoItem.dValorDescontoItem = 0 'dTemp
            objTributacaoItem.iValorDescontoItemManual = VAR_PREENCH_MANUAL
        'End If
        
        If Codigo_Extrai(TabTrib_Controle("OrigemMercadoria").Text) <> objTributacaoItem.iOrigemMercadoria Then
            objTributacaoItem.iOrigemMercadoria = Codigo_Extrai(TabTrib_Controle("OrigemMercadoria").Text)
            objTributacaoItem.iOrigemMercadoriaManual = VAR_PREENCH_MANUAL
        End If
        
        If Trim(TabTrib_Controle("FCI").Text) <> objTributacaoItem.sFCI Then
            objTributacaoItem.sFCI = Trim(TabTrib_Controle("FCI").Text)
            objTributacaoItem.iFCIManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("ICMSModalidade").Text) <> objTributacaoItem.iICMSBaseModalidade Then
            objTributacaoItem.iICMSBaseModalidade = Codigo_Extrai(TabTrib_Controle("ICMSModalidade").Text)
            objTributacaoItem.iICMSBaseModalidadeManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("ICMSSTModalidade").Text) <> objTributacaoItem.iICMSSubstBaseModalidade Then
            objTributacaoItem.iICMSSubstBaseModalidade = Codigo_Extrai(TabTrib_Controle("ICMSSTModalidade").Text)
            objTributacaoItem.iICMSSubstBaseModalidadeManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSSubstPercRedBaseItem").Text) - objTributacaoItem.dICMSSubstPercRedBase * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSubstPercRedBaseItem").Text) / 100
            objTributacaoItem.dICMSSubstPercRedBase = dTemp
            objTributacaoItem.iICMSSubstPercRedBaseManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSSubstPercMVAItem").Text) - objTributacaoItem.dICMSSubstPercMVA * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSubstPercMVAItem").Text) / 100
            objTributacaoItem.dICMSSubstPercMVA = dTemp
            objTributacaoItem.iICMSSubstPercMVAManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("IPITipoCalculo").Text) <> objTributacaoItem.iIPITipoCalculo Then
            objTributacaoItem.iIPITipoCalculo = Codigo_Extrai(TabTrib_Controle("IPITipoCalculo").Text)
            objTributacaoItem.iIPITipoCalculoManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("IPIClasseEnq").Text <> objTributacaoItem.sIPIEnquadramentoClasse Then
            objTributacaoItem.sIPIEnquadramentoClasse = Trim(TabTrib_Controle("IPIClasseEnq").Text)
            objTributacaoItem.iIPIEnquadramentoClasseManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("IPICodEnq").Text <> objTributacaoItem.sIPIEnquadramentoCodigo Then
            objTributacaoItem.sIPIEnquadramentoCodigo = Format(Trim(TabTrib_Controle("IPICodEnq").Text), "000")
            objTributacaoItem.iIPIEnquadramentoCodigoManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("IPICNPJProdutor").ClipText <> objTributacaoItem.sIPICNPJProdutor Then
            objTributacaoItem.sIPICNPJProdutor = Trim(TabTrib_Controle("IPICNPJProdutor").Text)
            objTributacaoItem.iIPICNPJProdutorManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("IPICodSelo").Text <> objTributacaoItem.sIPISeloCodigo Then
            objTributacaoItem.sIPISeloCodigo = Trim(TabTrib_Controle("IPICodSelo").Text)
            objTributacaoItem.iIPISeloCodigoManual = VAR_PREENCH_MANUAL
        End If

        If StrParaInt(TabTrib_Controle("IPISeloQtde").Text) <> objTributacaoItem.lIPISeloQtde Then
            objTributacaoItem.lIPISeloQtde = StrParaLong(TabTrib_Controle("IPISeloQtde").Text)
            objTributacaoItem.iIPISeloQtdeManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("IPIUnidadePadraoQtdeItem").Text) - objTributacaoItem.dIPIUnidadePadraoQtde) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IPIUnidadePadraoQtdeItem").Text)
            objTributacaoItem.dIPIUnidadePadraoQtde = dTemp
            objTributacaoItem.iIPIUnidadePadraoQtdeManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("IPIUnidadePadraoValorItem").Text) - objTributacaoItem.dIPIUnidadePadraoValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IPIUnidadePadraoValorItem").Text)
            objTributacaoItem.dIPIUnidadePadraoValor = dTemp
            objTributacaoItem.iIPIUnidadePadraoValorManual = VAR_PREENCH_MANUAL
        End If
        

        If Codigo_Extrai(TabTrib_Controle("ComboPISTipo").Text) <> objTributacaoItem.iPISTipo Then
            objTributacaoItem.iPISTipo = Codigo_Extrai(TabTrib_Controle("ComboPISTipo").Text)
            objTributacaoItem.iPISTipoManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("PISTipoCalculo").Text) <> objTributacaoItem.iPISTipoCalculo Then
            objTributacaoItem.iPISTipoCalculo = Codigo_Extrai(TabTrib_Controle("PISTipoCalculo").Text)
            objTributacaoItem.iPISTipoCalculoManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("PISSTTipoCalculo").Text) <> objTributacaoItem.iPISSTTipoCalculo Then
            objTributacaoItem.iPISSTTipoCalculo = Codigo_Extrai(TabTrib_Controle("PISSTTipoCalculo").Text)
            objTributacaoItem.iPISSTTipoCalculoManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("PISBaseItem").Text) - objTributacaoItem.dPISBase) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISBaseItem").Text)
            objTributacaoItem.dPISBase = dTemp
            objTributacaoItem.iPISBaseManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("PISAliquotaItem").Text) - objTributacaoItem.dPISAliquota * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISAliquotaItem").Text) / 100
            objTributacaoItem.dPISAliquota = dTemp
            objTributacaoItem.iPISAliquotaManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("PISQtdeItem").Text) - objTributacaoItem.dPISQtde) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISQtdeItem").Text)
            objTributacaoItem.dPISQtde = dTemp
            objTributacaoItem.iPISQtdeManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("PISValorItem").Text) - objTributacaoItem.dPISValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISValorItem").Text)
            objTributacaoItem.dPISValor = dTemp
            objTributacaoItem.iPISValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("PISSTBaseItem").Text) - objTributacaoItem.dPISSTBase) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISSTBaseItem").Text)
            objTributacaoItem.dPISSTBase = dTemp
            objTributacaoItem.iPISSTBaseManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("PISSTAliquotaItem").Text) - objTributacaoItem.dPISSTAliquota * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISSTAliquotaItem").Text) / 100
            objTributacaoItem.dPISSTAliquota = dTemp
            objTributacaoItem.iPISSTAliquotaManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("PISSTQtdeItem").Text) - objTributacaoItem.dPISSTQtde) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISSTQtdeItem").Text)
            objTributacaoItem.dPISSTQtde = dTemp
            objTributacaoItem.iPISSTQtdeManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("PISSTValorItem").Text) - objTributacaoItem.dPISSTValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISSTValorItem").Text)
            objTributacaoItem.dPISSTValor = dTemp
            objTributacaoItem.iPISSTValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("PISAliquotaRSItem").Text) - objTributacaoItem.dPISAliquotaValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISAliquotaRSItem").Text)
            objTributacaoItem.dPISAliquotaValor = dTemp
            objTributacaoItem.iPISAliquotaValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("PISSTAliquotaRSItem").Text) - objTributacaoItem.dPISSTAliquotaValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("PISSTAliquotaRSItem").Text)
            objTributacaoItem.dPISSTAliquotaValor = dTemp
            objTributacaoItem.iPISSTAliquotaValorManual = VAR_PREENCH_MANUAL
        End If

        If Codigo_Extrai(TabTrib_Controle("ComboCOFINSTipo").Text) <> objTributacaoItem.iCOFINSTipo Then
            objTributacaoItem.iCOFINSTipo = Codigo_Extrai(TabTrib_Controle("ComboCOFINSTipo").Text)
            objTributacaoItem.iCOFINSTipoManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("COFINSTipoCalculo").Text) <> objTributacaoItem.iCOFINSTipoCalculo Then
            objTributacaoItem.iCOFINSTipoCalculo = Codigo_Extrai(TabTrib_Controle("COFINSTipoCalculo").Text)
            objTributacaoItem.iCOFINSTipoCalculoManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("COFINSSTTipoCalculo").Text) <> objTributacaoItem.iCOFINSSTTipoCalculo Then
            objTributacaoItem.iCOFINSSTTipoCalculo = Codigo_Extrai(TabTrib_Controle("COFINSSTTipoCalculo").Text)
            objTributacaoItem.iCOFINSSTTipoCalculoManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("COFINSBaseItem").Text) - objTributacaoItem.dCOFINSBase) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSBaseItem").Text)
            objTributacaoItem.dCOFINSBase = dTemp
            objTributacaoItem.iCOFINSBaseManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("COFINSAliquotaItem").Text) - objTributacaoItem.dCOFINSAliquota * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSAliquotaItem").Text) / 100
            objTributacaoItem.dCOFINSAliquota = dTemp
            objTributacaoItem.iCOFINSAliquotaManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("COFINSQtdeItem").Text) - objTributacaoItem.dCOFINSQtde) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSQtdeItem").Text)
            objTributacaoItem.dCOFINSQtde = dTemp
            objTributacaoItem.iCOFINSQtdeManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("COFINSValorItem").Text) - objTributacaoItem.dCOFINSValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSValorItem").Text)
            objTributacaoItem.dCOFINSValor = dTemp
            objTributacaoItem.iCOFINSValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("COFINSSTBaseItem").Text) - objTributacaoItem.dCOFINSSTBase) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSSTBaseItem").Text)
            objTributacaoItem.dCOFINSSTBase = dTemp
            objTributacaoItem.iCOFINSSTBaseManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("COFINSSTAliquotaItem").Text) - objTributacaoItem.dCOFINSSTAliquota * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSSTAliquotaItem").Text) / 100
            objTributacaoItem.dCOFINSSTAliquota = dTemp
            objTributacaoItem.iCOFINSSTAliquotaManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("COFINSSTQtdeItem").Text) - objTributacaoItem.dCOFINSSTQtde) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSSTQtdeItem").Text)
            objTributacaoItem.dCOFINSSTQtde = dTemp
            objTributacaoItem.iCOFINSSTQtdeManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("COFINSSTValorItem").Text) - objTributacaoItem.dCOFINSSTValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSSTValorItem").Text)
            objTributacaoItem.dCOFINSSTValor = dTemp
            objTributacaoItem.iCOFINSSTValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("COFINSAliquotaRSItem").Text) - objTributacaoItem.dCOFINSAliquotaValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSAliquotaRSItem").Text)
            objTributacaoItem.dCOFINSAliquotaValor = dTemp
            objTributacaoItem.iCOFINSAliquotaValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("COFINSSTAliquotaRSItem").Text) - objTributacaoItem.dCOFINSSTAliquotaValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("COFINSSTAliquotaRSItem").Text)
            objTributacaoItem.dCOFINSSTAliquotaValor = dTemp
            objTributacaoItem.iCOFINSSTAliquotaValorManual = VAR_PREENCH_MANUAL
        End If
        
        If SCodigo_Extrai(TabTrib_Controle("ISSListaServ").Text) <> objTributacaoItem.sISSQN Then
            objTributacaoItem.sISSQN = SCodigo_Extrai(TabTrib_Controle("ISSListaServ").Text)
            objTributacaoItem.iISSQNManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ISSBaseItem").Text) - objTributacaoItem.dISSBase) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSBaseItem").Text)
            objTributacaoItem.dISSBase = dTemp
            objTributacaoItem.iISSBaseManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ISSAliquotaItem").Text) - objTributacaoItem.dISSAliquota * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSAliquotaItem").Text) / 100
            objTributacaoItem.dISSAliquota = dTemp
            objTributacaoItem.iISSAliquotaManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ISSValorItem").Text) - objTributacaoItem.dISSValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSValorItem").Text)
            objTributacaoItem.dISSValor = dTemp
            objTributacaoItem.iISSValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("IIBaseItem").Text) - objTributacaoItem.dIIBase) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IIBaseItem").Text)
            objTributacaoItem.dIIBase = dTemp
            objTributacaoItem.iIIBaseManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("IIValorItem").Text) - objTributacaoItem.dIIValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IIValorItem").Text)
            objTributacaoItem.dIIValor = dTemp
            objTributacaoItem.iIIValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("IIValorItem").Text) - objTributacaoItem.dIIValor) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IIValorItem").Text)
            objTributacaoItem.dIIValor = dTemp
            objTributacaoItem.iIIValorManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("IIIOFItem").Text) - objTributacaoItem.dIIIOF) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IIIOFItem").Text)
            objTributacaoItem.dIIIOF = dTemp
            objTributacaoItem.iIIIOFManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("IIDespAduItem").Text) - objTributacaoItem.dIIDespAduaneira) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IIDespAduItem").Text)
            objTributacaoItem.dIIDespAduaneira = dTemp
            objTributacaoItem.iIIDespAduaneiraManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("ISSCodIBGE").ClipText <> objTributacaoItem.sISSCidadeIBGE Then
            objTributacaoItem.sISSCidadeIBGE = Trim(TabTrib_Controle("ISSCodIBGE").ClipText)
            objTributacaoItem.iISSCidadeIBGEManual = VAR_PREENCH_MANUAL
        End If
        'FIM - NOVOS CAMPOS
        
        If TabTrib_Controle("ComboISSTipo").ListIndex <> -1 Then
            iTemp = TabTrib_Controle("ComboISSTipo").ItemData(TabTrib_Controle("ComboISSTipo").ListIndex)
        Else
            iTemp = 0
        End If
        If iTemp <> objTributacaoItem.iISSTipo Then
            objTributacaoItem.iISSTipo = iTemp
            objTributacaoItem.iISSTipoManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("ComboICMSSimplesTipo").ListIndex <> -1 Then
            iTemp = TabTrib_Controle("ComboICMSSimplesTipo").ItemData(TabTrib_Controle("ComboICMSSimplesTipo").ListIndex)
        Else
            iTemp = 0
        End If
        If iTemp <> objTributacaoItem.iICMSSimplesTipo Then
            objTributacaoItem.iICMSSimplesTipo = iTemp
            objTributacaoItem.iICMSSimplesTipoManual = VAR_PREENCH_MANUAL
        End If
    
        If Abs(StrParaDbl(TabTrib_Controle("ICMSpercBaseOperacaoPropria").Text) - objTributacaoItem.dICMSpercBaseOperacaoPropria * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSpercBaseOperacaoPropria").Text) / 100
            objTributacaoItem.dICMSpercBaseOperacaoPropria = dTemp
            objTributacaoItem.iICMSpercBaseOperacaoPropriaManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvBCSTRet").Text) - objTributacaoItem.dICMSvBCSTRet) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvBCSTRet").Text)
            objTributacaoItem.dICMSvBCSTRet = dTemp
            objTributacaoItem.iICMSvBCSTRetManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvICMSSTRet").Text) - objTributacaoItem.dICMSvICMSSTRet) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvICMSSTRet").Text)
            objTributacaoItem.dICMSvICMSSTRet = dTemp
            objTributacaoItem.iICMSvICMSSTRetManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvBCSTDest").Text) - objTributacaoItem.dICMSvBCSTDest) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvBCSTDest").Text)
            objTributacaoItem.dICMSvBCSTDest = dTemp
            objTributacaoItem.iICMSvBCSTDestManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvICMSSTDest").Text) - objTributacaoItem.dICMSvICMSSTDest) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvICMSSTDest").Text)
            objTributacaoItem.dICMSvICMSSTDest = dTemp
            objTributacaoItem.iICMSvICMSSTDestManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvCredSN").Text) - objTributacaoItem.dICMSvCredSN) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvCredSN").Text)
            objTributacaoItem.dICMSvCredSN = dTemp
            objTributacaoItem.iICMSvCredSNManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSValorIsento").Text) - objTributacaoItem.dICMSValorIsento) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSValorIsento").Text)
            objTributacaoItem.dICMSValorIsento = dTemp
            objTributacaoItem.iICMSValorIsentoManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ICMSpCredSN").Text) - objTributacaoItem.dICMSpCredSN * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSpCredSN").Text) / 100
            objTributacaoItem.dICMSpCredSN = dTemp
            objTributacaoItem.iICMSpCredSNManual = VAR_PREENCH_MANUAL
        End If
        
        If Codigo_Extrai(TabTrib_Controle("ICMSMotivo").Text) <> objTributacaoItem.iICMSMotivo Then
            objTributacaoItem.iICMSMotivo = Codigo_Extrai(TabTrib_Controle("ICMSMotivo").Text)
            objTributacaoItem.iICMSMotivoManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("ICMSUFDevidoST").Text <> objTributacaoItem.sICMSUFDevidoST Then
            objTributacaoItem.sICMSUFDevidoST = Trim(TabTrib_Controle("ICMSUFDevidoST").Text)
            objTributacaoItem.iICMSUFDevidoSTManual = VAR_PREENCH_MANUAL
        End If
        
        If SCodigo_Extrai(TabTrib_Controle("NatBCCred").Text) <> objTributacaoItem.sNatBCCred Then
            objTributacaoItem.sNatBCCred = SCodigo_Extrai(TabTrib_Controle("NatBCCred").Text)
            objTributacaoItem.iNatBCCredManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("TotTribItem").Text) - objTributacaoItem.dTotTrib) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTribItem").Text)
            objTributacaoItem.dTotTrib = dTemp
            objTributacaoItem.iTotTribManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("TotTribAliqFedItem").Text) - objTributacaoItem.dTotTribAliqFed * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTribAliqFedItem").Text) / 100
            objTributacaoItem.dTotTribAliqFed = dTemp
            objTributacaoItem.iTotTribAliqFedManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("TotTribAliqEstItem").Text) - objTributacaoItem.dTotTribAliqEst * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTribAliqEstItem").Text) / 100
            objTributacaoItem.dTotTribAliqEst = dTemp
            objTributacaoItem.iTotTribAliqEstManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("TotTribAliqMunicItem").Text) - objTributacaoItem.dTotTribAliqMunic * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("TotTribAliqMunicItem").Text) / 100
            objTributacaoItem.dTotTribAliqMunic = dTemp
            objTributacaoItem.iTotTribAliqMunicManual = VAR_PREENCH_MANUAL
        End If

        iTemp = TabTrib_Controle("TotTribItemTipo").ItemData(TabTrib_Controle("TotTribItemTipo").ListIndex)
        If iTemp <> objTributacaoItem.iTotTribTipo Then
            objTributacaoItem.iTotTribTipo = iTemp
            objTributacaoItem.iTotTribTipoManual = VAR_PREENCH_MANUAL
        End If
    
        'nfe 3.10
        If Abs(StrParaDbl(TabTrib_Controle("ICMS51ValorOpItem").Text) - objTributacaoItem.dICMS51ValorOp) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMS51ValorOpItem").Text)
            objTributacaoItem.dICMS51ValorOp = dTemp
            objTributacaoItem.iICMS51ValorOpManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ICMSPercDiferItem").Text) - objTributacaoItem.dICMSPercDifer * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSPercDiferItem").Text) / 100
            objTributacaoItem.dICMSPercDifer = dTemp
            objTributacaoItem.iICMSPercDiferManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSValorDifItem").Text) - objTributacaoItem.dICMSValorDif) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSValorDifItem").Text)
            objTributacaoItem.dICMSValorDif = dTemp
            objTributacaoItem.iICMSValorDifManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ISSValorDeducaoItem").Text) - objTributacaoItem.dISSValorDeducao) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSValorDeducaoItem").Text)
            objTributacaoItem.dISSValorDeducao = dTemp
            objTributacaoItem.iISSValorDeducaoManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ISSValorOutrasRetItem").Text) - objTributacaoItem.dISSValorOutrasRet) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSValorOutrasRetItem").Text)
            objTributacaoItem.dISSValorOutrasRet = dTemp
            objTributacaoItem.iISSValorOutrasRetManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ISSValorDescIncondItem").Text) - objTributacaoItem.dISSValorDescIncond) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSValorDescIncondItem").Text)
            objTributacaoItem.dISSValorDescIncond = dTemp
            objTributacaoItem.iISSValorDescIncondManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ISSValorDescCondItem").Text) - objTributacaoItem.dISSValorDescCond) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSValorDescCondItem").Text)
            objTributacaoItem.dISSValorDescCond = dTemp
            objTributacaoItem.iISSValorDescCondManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ISSValorRetItem").Text) - objTributacaoItem.dISSValorRet) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ISSValorRetItem").Text)
            objTributacaoItem.dISSValorRet = dTemp
            objTributacaoItem.iISSValorRetManual = VAR_PREENCH_MANUAL
        End If

        If Codigo_Extrai(TabTrib_Controle("ISSIndExigibilidadeItem").Text) <> objTributacaoItem.iISSIndExigibilidade Then
            objTributacaoItem.iISSIndExigibilidade = Codigo_Extrai(TabTrib_Controle("ISSIndExigibilidadeItem").Text)
            objTributacaoItem.iISSIndExigibilidadeManual = VAR_PREENCH_MANUAL
        End If

        iTemp = TabTrib_Controle("ISSIndIncentivoItem").Value
        If iTemp <> objTributacaoItem.iISSIndIncentivo Then
            objTributacaoItem.iISSIndIncentivo = iTemp
            objTributacaoItem.iISSIndIncentivoManual = VAR_PREENCH_MANUAL
        End If

        If TabTrib_Controle("ISSMunicIncidImpItem").ClipText <> objTributacaoItem.sISSMunicIncidImp Then
            objTributacaoItem.sISSMunicIncidImp = Trim(TabTrib_Controle("ISSMunicIncidImpItem").ClipText)
            objTributacaoItem.iISSMunicIncidImpManual = VAR_PREENCH_MANUAL
        End If

        If TabTrib_Controle("ISSCodServItem").ClipText <> objTributacaoItem.sISSCodServ Then
            objTributacaoItem.sISSCodServ = Trim(TabTrib_Controle("ISSCodServItem").ClipText)
            objTributacaoItem.iISSCodServManual = VAR_PREENCH_MANUAL
        End If

        sTemp = Trim(TabTrib_Controle("ISSNumProcessoItem").Text)
        If Trim(objTributacaoItem.sISSNumProcesso) <> sTemp Then
            objTributacaoItem.sISSNumProcesso = sTemp
            objTributacaoItem.iISSNumProcessoManual = VAR_PREENCH_MANUAL
        End If

        If TabTrib_Controle("ISSCodPaisItem").ListIndex <> -1 Then
            iTemp = TabTrib_Controle("ISSCodPaisItem").ItemData(TabTrib_Controle("ISSCodPaisItem").ListIndex)
        Else
            iTemp = 0
        End If
        If iTemp <> objTributacaoItem.iISSCodPais Then
            objTributacaoItem.iISSCodPais = iTemp
            objTributacaoItem.iISSCodPaisManual = VAR_PREENCH_MANUAL
        End If
               
        'NT 2015.003 (ICMS Interestadual)
        If Abs(StrParaDbl(TabTrib_Controle("ICMSInterestBCUFDestItem").Text) - objTributacaoItem.dICMSInterestBCUFDest) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSInterestBCUFDestItem").Text)
            objTributacaoItem.dICMSInterestBCUFDest = dTemp
            objTributacaoItem.iICMSInterestBCUFDestManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSInterestPercFCPUFDestItem").Text) - objTributacaoItem.dICMSInterestPercFCPUFDest * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSInterestPercFCPUFDestItem").Text) / 100
            objTributacaoItem.dICMSInterestPercFCPUFDest = dTemp
            objTributacaoItem.iICMSInterestPercFCPUFDestManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSInterestAliqUFDestItem").Text) - objTributacaoItem.dICMSInterestAliqUFDest * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSInterestAliqUFDestItem").Text) / 100
            objTributacaoItem.dICMSInterestAliqUFDest = dTemp
            objTributacaoItem.iICMSInterestAliqUFDestManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("ICMSInterestAliqItem").ListIndex <> -1 Then
            iTemp = TabTrib_Controle("ICMSInterestAliqItem").ItemData(TabTrib_Controle("ICMSInterestAliqItem").ListIndex)
        Else
            iTemp = 0
        End If
        If iTemp <> CInt(objTributacaoItem.dICMSInterestAliq * 100) Then
            dTemp = iTemp / 100
            objTributacaoItem.dICMSInterestAliq = dTemp
            objTributacaoItem.iICMSInterestAliqManual = VAR_PREENCH_MANUAL
        End If
        
        If TabTrib_Controle("ICMSInterestPercPartilhaItem").ListIndex <> -1 Then
            iTemp = TabTrib_Controle("ICMSInterestPercPartilhaItem").ItemData(TabTrib_Controle("ICMSInterestPercPartilhaItem").ListIndex)
        Else
            iTemp = 0
        End If
        If iTemp <> CInt(objTributacaoItem.dICMSInterestPercPartilha * 100) Then
            dTemp = iTemp / 100
            objTributacaoItem.dICMSInterestPercPartilha = dTemp
            objTributacaoItem.iICMSInterestPercPartilhaManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSInterestVlrFCPUFDestItem").Text) - objTributacaoItem.dICMSInterestVlrFCPUFDest) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSInterestVlrFCPUFDestItem").Text)
            objTributacaoItem.dICMSInterestVlrFCPUFDest = dTemp
            objTributacaoItem.iICMSInterestVlrFCPUFDestManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSInterestVlrUFDestItem").Text) - objTributacaoItem.dICMSInterestVlrUFDest) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSInterestVlrUFDestItem").Text)
            objTributacaoItem.dICMSInterestVlrUFDest = dTemp
            objTributacaoItem.iICMSInterestVlrUFDestManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSInterestVlrUFRemetItem").Text) - objTributacaoItem.dICMSInterestVlrUFRemet) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSInterestVlrUFRemetItem").Text)
            objTributacaoItem.dICMSInterestVlrUFRemet = dTemp
            objTributacaoItem.iICMSInterestVlrUFRemetManual = VAR_PREENCH_MANUAL
        End If
        
        'UA. Tributos Devolvidos (para o item da NF-e)
        If Abs(StrParaDbl(TabTrib_Controle("IPIVlrDevolvidoItem").Text) - objTributacaoItem.dIPIVlrDevolvido) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("IPIVlrDevolvidoItem").Text)
            objTributacaoItem.dIPIVlrDevolvido = dTemp
            objTributacaoItem.iIPIVlrDevolvidoManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("pDevolItem").Text) - objTributacaoItem.dpDevol * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("pDevolItem").Text) / 100
            objTributacaoItem.dpDevol = dTemp
            objTributacaoItem.ipDevolManual = VAR_PREENCH_MANUAL
        End If
        
        'NFe 4.00
        sTemp = Trim(TabTrib_Controle("cBenefItem").Text)
        If Trim(objTributacaoItem.scBenef) <> sTemp Then
            objTributacaoItem.scBenef = sTemp
            objTributacaoItem.icBenefManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSInterestBCFCPUFDestItem").Text) - objTributacaoItem.dICMSInterestBCFCPUFDest) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSInterestBCFCPUFDestItem").Text)
            objTributacaoItem.dICMSInterestBCFCPUFDest = dTemp
            objTributacaoItem.iICMSInterestBCFCPUFDestManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ICMSSTCobrAntAliquotaItem").Text) - objTributacaoItem.dICMSSTCobrAntAliquota * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSSTCobrAntAliquotaItem").Text) / 100
            objTributacaoItem.dICMSSTCobrAntAliquota = dTemp
            objTributacaoItem.iICMSSTCobrAntAliquotaManual = VAR_PREENCH_MANUAL
        End If
        
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvBCFCPItem").Text) - objTributacaoItem.dICMSvBCFCP) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvBCFCPItem").Text)
            objTributacaoItem.dICMSvBCFCP = dTemp
            objTributacaoItem.iICMSvBCFCPManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ICMSpFCPItem").Text) - objTributacaoItem.dICMSpFCP * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSpFCPItem").Text) / 100
            objTributacaoItem.dICMSpFCP = dTemp
            objTributacaoItem.iICMSpFCPManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvFCPItem").Text) - objTributacaoItem.dICMSvFCP) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvFCPItem").Text)
            objTributacaoItem.dICMSvFCP = dTemp
            objTributacaoItem.iICMSvFCPManual = VAR_PREENCH_MANUAL
        End If
    
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvBCFCPSTItem").Text) - objTributacaoItem.dICMSvBCFCPST) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvBCFCPSTItem").Text)
            objTributacaoItem.dICMSvBCFCPST = dTemp
            objTributacaoItem.iICMSvBCFCPSTManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ICMSpFCPSTItem").Text) - objTributacaoItem.dICMSpFCPST * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSpFCPSTItem").Text) / 100
            objTributacaoItem.dICMSpFCPST = dTemp
            objTributacaoItem.iICMSpFCPSTManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvFCPSTItem").Text) - objTributacaoItem.dICMSvFCPST) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvFCPSTItem").Text)
            objTributacaoItem.dICMSvFCPST = dTemp
            objTributacaoItem.iICMSvFCPSTManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvBCFCPSTRetItem").Text) - objTributacaoItem.dICMSvBCFCPSTRet) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvBCFCPSTRetItem").Text)
            objTributacaoItem.dICMSvBCFCPSTRet = dTemp
            objTributacaoItem.iICMSvBCFCPSTRetManual = VAR_PREENCH_MANUAL
        End If

        If Abs(StrParaDbl(TabTrib_Controle("ICMSpFCPSTRetItem").Text) - objTributacaoItem.dICMSpFCPSTRet * 100) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSpFCPSTRetItem").Text) / 100
            objTributacaoItem.dICMSpFCPSTRet = dTemp
            objTributacaoItem.iICMSpFCPSTRetManual = VAR_PREENCH_MANUAL
        End If
        
        If Abs(StrParaDbl(TabTrib_Controle("ICMSvFCPSTRetItem").Text) - objTributacaoItem.dICMSvFCPSTRet) > DELTA_VALORMONETARIO Then
            dTemp = StrParaDbl(TabTrib_Controle("ICMSvFCPSTRetItem").Text)
            objTributacaoItem.dICMSvFCPSTRet = dTemp
            objTributacaoItem.iICMSvFCPSTRetManual = VAR_PREENCH_MANUAL
        End If
        
        iTemp = IIf(TabTrib_Controle("ICMSSTBaseDupla").Value = vbChecked, MARCADO, DESMARCADO)
        If iTemp <> objTributacaoItem.iICMSSTBaseDupla Then
            objTributacaoItem.iICMSSTBaseDupla = iTemp
            objTributacaoItem.iICMSSTBaseDuplaManual = VAR_PREENCH_MANUAL
        End If
        
        iTemp = IIf(TabTrib_Controle("ICMSInterestBaseDupla").Value = vbChecked, MARCADO, DESMARCADO)
        If iTemp <> objTributacaoItem.iICMSInterestBaseDupla Then
            objTributacaoItem.iICMSInterestBaseDupla = iTemp
            objTributacaoItem.iICMSInterestBaseDuplaManual = VAR_PREENCH_MANUAL
        End If
                
    End With
    
    TributacaoItem_GravarTela = SUCESSO

    Exit Function

Erro_TributacaoItem_GravarTela:

    TributacaoItem_GravarTela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153733)

    End Select

    Exit Function

End Function

Function Inclusao_Item_Grid(iLinha As Integer, sProduto As String, Optional ByVal bSRV As Boolean = False) As Long
'trata a inclusao de uma linha de item no grid

Dim lErro As Long
Dim objTributoDocItem As ClassTributoDocItem
Dim objTributoDocItemAux As Object
Dim objProduto As New ClassProduto
Dim vCodBarras As Variant
Dim sProdutoFormatado As String, iPreenchido As Integer
Dim dFator As Double
Dim objItem As Object
Dim objCamposGenericosValores As New ClassCamposGenericosValores, sCNPJ As String

On Error GoTo Erro_Inclusao_Item_Grid

    lErro = Move_GridItem_Memoria(gobjDocTrib, iLinha, sProduto, bSRV)
    If lErro <> SUCESSO Then Error 27809

    Set objTributoDocItem = gcolDocTribItem(gobjDocTrib, bSRV)(iLinha)
    Set objTributoDocItemAux = gobjDocTribItem(iLinha, bSRV)
    lErro = objTributoDocItem.Ativar(gobjDocTrib)
    If lErro <> SUCESSO Then Error 27810
    
'    'Da tela de PV o produto não vem formatado
'    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iPreenchido)
'    If lErro <> SUCESSO Then Error 27809
    sProdutoFormatado = sProduto
    
    If Len(Trim(sProdutoFormatado)) > 0 Then
    
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 27809
        
        objTributoDocItemAux.sProduto = sProdutoFormatado
        objTributoDocItemAux.iExTIPI = objProduto.iExTIPI
        objTributoDocItemAux.sGenero = Trim(objProduto.sGenero)
        objTributoDocItemAux.iProdutoEspecifico = objProduto.iProdutoEspecifico
        objTributoDocItemAux.sUMTrib = objProduto.sSiglaUMTrib
        objTributoDocItemAux.sIPICodProduto = Trim(objProduto.sIPICodigo)
        objTributoDocItemAux.sProdutoDescricao = Trim(objProduto.sDescricao)
        objTributoDocItemAux.iOrigemMercadoria = objProduto.iOrigemMercadoria
        objTributoDocItemAux.sFCI = Trim(objProduto.sFCI)
        objTributoDocItemAux.sCEST = Trim(objProduto.sCEST)
        
        'NFe 4.00
        If objProduto.iProdEmEscalaRelev = MARCADO Then
            objTributoDocItemAux.sindEscala = "S"
        ElseIf objProduto.iProdEmEscalaRelev = DESMARCADO Then
            objTributoDocItemAux.sindEscala = "N"
        Else
            objTributoDocItemAux.sindEscala = ""
        End If
        If objProduto.lFabricante <> 0 Then
        
            objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_PROD_FABR
            objCamposGenericosValores.lCodValor = objProduto.lFabricante
            
            lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
            If lErro <> SUCESSO Then Error 27810
        
            Call Formata_String_Numero(objCamposGenericosValores.sComplemento1, sCNPJ)
        
            objTributoDocItemAux.sCNPJFab = sCNPJ
        Else
            objTributoDocItemAux.sCNPJFab = ""
        End If
        'NFe 4.00
       
        For Each vCodBarras In objProduto.colCodBarras
            objTributoDocItemAux.sEAN = vCodBarras
            objTributoDocItemAux.sEANTrib = vCodBarras
            Exit For
        Next
        
        Set objItem = gcolDocItem(gobjDocTrib, bSRV).Item(iLinha)
    
        dFator = 0
        If Len(Trim(objItem.sUnidadeMed)) <> 0 And Len(Trim(objTributoDocItemAux.sUMTrib)) <> 0 Then
        
            'Faz a conversão da UM da tela para a UM tributavel
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objItem.sUnidadeMed, objTributoDocItemAux.sUMTrib, dFator)
            If lErro <> SUCESSO Then Error 27811
        
        End If
        
        If dFator = 0 Then dFator = 1
        
        objTributoDocItemAux.dQtdTrib = objItem.dQuantidade * dFator
        objTributoDocItemAux.dValorUnitTrib = objItem.dPrecoUnitario / dFator
            
        lErro = Trata_Comercio_Exterior_Item(objTributoDocItemAux)
        'If lErro <> SUCESSO Then Error 27809
            
'        objTributoDocItemAux.sISSQN = objProduto.sISSQN
        
'        '*****************************
'        'Preenche o código do IBGE da cidade onde o serviço foi realizado
'        'com a cidade do cliente ou a cidade da empresa
'        'de acordo com a configuração do produto
'        lErro = CF("ProdutoCNAE_Le", objProduto)
'        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 27809
'
'        If lErro = SUCESSO Then
'            If objProduto.objProdutoCNAE.iLocServCliente = MARCADO Then
'                objTributoDocItemAux.sISSCidadeIBGE = sCidadeIBGECodigo
'            Else
'                objTributoDocItemAux.sISSCidadeIBGE = sCidadeIBGECodigoEmp
'            End If
'        Else
'            objTributoDocItemAux.sISSCidadeIBGE = sCidadeIBGECodigoEmp
'        End If
'        '*****************************
        
    End If

    Inclusao_Item_Grid = SUCESSO

    Exit Function

Erro_Inclusao_Item_Grid:

    Inclusao_Item_Grid = Err

    Select Case Err

        Case 27809, 27810, 27811

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153734)

    End Select

    Exit Function

End Function

Function Exclusao_Item_Grid(iLinha As Integer, Optional ByVal bSRV As Boolean = False) As Long
'trata a exclusao de uma linha de item no grid
Dim objItem As Object, objTributoDocItem As ClassTributoDocItem

        Set objItem = gcolDocTribItem(gobjDocTrib, bSRV).Item(iLinha)
        Set objTributoDocItem = objItem
        Call objTributoDocItem.Desativar
        If bSRV Then
            Call gobjDocTrib.RemoverItemSRV(iLinha)
        Else
            Call gobjDocTrib.RemoverItem(iLinha)
        End If

End Function

Function Alteracao_Item_Grid(iIndice As Integer, Optional ByVal bSRV As Boolean = False, Optional lNumIntItemNFOrig As Long = 0) As Long
'trata a alteracao de uma linha de item no grid

Dim lErro As Long, sProduto As String, iPreenchido As Integer
Dim objItem As Object
Dim objTributoDocItemAux As Object
Dim objProduto As New ClassProduto
Dim vCodBarras As Variant
Dim dFator As Double, objTributoDocItem As ClassTributoDocItem
Dim objCamposGenericosValores As New ClassCamposGenericosValores, sCNPJ As String

On Error GoTo Erro_Alteracao_Item_Grid

    Set objItem = gcolDocItem(gobjDocTrib, bSRV).Item(iIndice)
    Set objTributoDocItemAux = gobjDocTribItem(iIndice, bSRV)

    Set objTributoDocItem = objItem
    Call objTributoDocItem.Alterado
    
    With mvarobjForm
    
        If bSRV Then
        
            If Len(Trim(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_ServicoSRV_Col))) > 0 Then
        
                lErro = CF("Produto_Formata", TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_ServicoSRV_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then Error 27811
        
                objItem.sProduto = sProduto
            End If
        
            objItem.sUnidadeMed = TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_UnidadeMedSRV_Col)
            objItem.dQuantidade = dQuantidadeTela(iIndice, bSRV)
            objItem.dPrecoTotal = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_PrecoTotalSRV_Col))
            
            If iTipoTela = TIPOTELA_OV Then
                objItem.dPrecoUnitario = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_PrecoUnitarioSRV_Col))
            Else
                objItem.dPrecoUnitario = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitarioSRV_Col))
            End If
            If iTipoTela = TIPOTELA_NF Then
                If mvarobjForm.iGrid_ValorUnitarioMoeda_Col <> 0 Then
                    objItem.dPrecoUnitarioMoeda = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitarioMoedaSRV_Col))
                Else
                    objItem.dPrecoUnitarioMoeda = 0
                End If
            End If
            objItem.dValorDesconto = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_DescontoSRV_Col))
        
        Else
    
            If Len(Trim(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_Produto_Col))) > 0 Then
        
                lErro = CF("Produto_Formata", TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_Produto_Col), sProduto, iPreenchido)
                If lErro <> SUCESSO Then Error 27811
        
                objItem.sProduto = sProduto
            End If
        
            objItem.sUnidadeMed = TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_UnidadeMed_Col)
            objItem.dQuantidade = dQuantidadeTela(iIndice)
            objItem.dPrecoTotal = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_PrecoTotal_Col))
            
            If iTipoTela = TIPOTELA_OV Or iTipoTela = TIPOTELA_PRJ_PROP Or iTipoTela = TIPOTELA_PRJ_CTR Then
                objItem.dPrecoUnitario = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_PrecoUnitario_Col))
            Else
                objItem.dPrecoUnitario = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitario_Col))
            End If
            If iTipoTela = TIPOTELA_NF Then
                If mvarobjForm.iGrid_ValorUnitarioMoeda_Col <> 0 Then
                    objItem.dPrecoUnitarioMoeda = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitarioMoeda_Col))
                Else
                    objItem.dPrecoUnitarioMoeda = 0
                End If
            End If
            objItem.dValorDesconto = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_Desconto_Col))

        End If

    End With
    
    If lNumIntItemNFOrig <> 0 Then
    
        Set objTributoDocItemAux.objItemNFOrig = New ClassItemNF
        objTributoDocItemAux.objItemNFOrig.lNumIntDoc = lNumIntItemNFOrig
        
        lErro = CF("ItemNFiscal_Le", objTributoDocItemAux.objItemNFOrig)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
           
        Set objTributoDocItemAux.objItemNFOrig.objTributacao = New ClassTributacaoDocItem
        objTributoDocItemAux.objItemNFOrig.objTributacao.iTipoDoc = TIPODOC_TRIB_NF
        objTributoDocItemAux.objItemNFOrig.objTributacao.lNumIntDocItem = objTributoDocItemAux.objItemNFOrig.lNumIntDoc
        lErro = CF("TributacaoDocItem_Le", objTributoDocItemAux.objItemNFOrig.objTributacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If
    
    If Len(Trim(sProduto)) > 0 Then
    
        objProduto.sCodigo = sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 27811
        
        objTributoDocItemAux.sProduto = sProduto
        objTributoDocItemAux.iExTIPI = objProduto.iExTIPI
        objTributoDocItemAux.sGenero = objProduto.sGenero
        objTributoDocItemAux.iProdutoEspecifico = objProduto.iProdutoEspecifico
        objTributoDocItemAux.sUMTrib = objProduto.sSiglaUMTrib
        objTributoDocItemAux.sIPICodProduto = Trim(objProduto.sIPICodigo)
        objTributoDocItemAux.sCEST = Trim(objProduto.sCEST)
    
        'NFe 4.00
        If objProduto.iProdEmEscalaRelev = MARCADO Then
            objTributoDocItemAux.sindEscala = "S"
        ElseIf objProduto.iProdEmEscalaRelev = DESMARCADO Then
            objTributoDocItemAux.sindEscala = "N"
        Else
            objTributoDocItemAux.sindEscala = ""
        End If
        If objProduto.lFabricante <> 0 Then
        
            objCamposGenericosValores.lCodCampo = CAMPOSGENERICOS_PROD_FABR
            objCamposGenericosValores.lCodValor = objProduto.lFabricante
            
            lErro = CF("CamposGenericosValores_Le_CodCampo_CodValor", objCamposGenericosValores)
            If lErro <> SUCESSO Then Error 27811
        
            Call Formata_String_Numero(objCamposGenericosValores.sComplemento1, sCNPJ)
        
            objTributoDocItemAux.sCNPJFab = sCNPJ
        Else
            objTributoDocItemAux.sCNPJFab = ""
        End If
        'NFe 4.00
        
        For Each vCodBarras In objProduto.colCodBarras
            objTributoDocItemAux.sEAN = vCodBarras
            objTributoDocItemAux.sEANTrib = vCodBarras
            Exit For
        Next
        
        'Faz a conversão da UM da tela para a UM tributavel
        lErro = CF("UM_Conversao", objProduto.iClasseUM, objItem.sUnidadeMed, objTributoDocItemAux.sUMTrib, dFator)
        If lErro <> SUCESSO Then Error 27811
        
        objTributoDocItemAux.dQtdTrib = objItem.dQuantidade * dFator
        objTributoDocItemAux.dValorUnitTrib = objItem.dPrecoUnitario / dFator
        objTributoDocItemAux.dDescontoGrid = objItem.dValorDesconto
        objTributoDocItemAux.dPrecoTotal = objItem.dPrecoTotal
        objTributoDocItemAux.sProdutoDescricao = objProduto.sDescricao
    
        lErro = Trata_Comercio_Exterior_Item(objTributoDocItemAux)
        'If lErro <> SUCESSO Then Error 27811
    
    End If
        
    Alteracao_Item_Grid = SUCESSO

    Exit Function

Erro_Alteracao_Item_Grid:

    Alteracao_Item_Grid = Err

    Select Case Err

        Case 27811

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153735)

    End Select

    Exit Function

End Function

Private Function Move_GridItem_Memoria(objTrib As Object, iIndice As Integer, sProduto As String, Optional ByVal bSRV As Boolean = False) As Long
'obtem informacao de uma linha de grid p/"criacao" de um item p/tratamento da tributacao

Dim lErro As Long
Dim objItem As ClassTribItemNFAux

On Error GoTo Erro_Move_GridItem_Memoria

    If iTipoTela = TIPOTELA_NF Then

        Set objItem = New ClassTribItemNFAux
    
        If bSRV Then
    
            With mvarobjForm
            
                If Len(Trim(sProduto)) = 0 Then Error 20768
            
                objItem.sProduto = sProduto
                objItem.sUnidadeMed = TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_UnidadeMedSRV_Col)
                objItem.dQuantidade = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_QuantidadeSRV_Col))
                objItem.dPrecoTotal = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_PrecoTotalSRV_Col))
                objItem.dPrecoUnitario = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitarioSRV_Col))
                If mvarobjForm.iGrid_ValorUnitarioMoeda_Col <> 0 Then
                    objItem.dPrecoUnitarioMoeda = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitarioMoedaSRV_Col))
                Else
                    objItem.dPrecoUnitarioMoeda = 0
                End If
                objItem.sProdutoDescricao = TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_DescServicoSRV_Col)
                objItem.dValorDesconto = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_DescontoSRV_Col))
        
            End With
            
            lErro = objTrib.IncluirItemSRV(objItem)
            If lErro <> SUCESSO Then Error 27831
            
        Else
            
            With mvarobjForm
            
                If Len(Trim(sProduto)) = 0 Then Error 20768
            
                objItem.sProduto = sProduto
                objItem.sUnidadeMed = TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_UnidadeMed_Col)
                objItem.dQuantidade = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_Quantidade_Col))
                objItem.dPrecoTotal = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_PrecoTotal_Col))
                objItem.dPrecoUnitario = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitario_Col))
                If mvarobjForm.iGrid_ValorUnitarioMoeda_Col <> 0 Then
                    objItem.dPrecoUnitarioMoeda = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_ValorUnitarioMoeda_Col))
                Else
                    objItem.dPrecoUnitarioMoeda = 0
                End If
                objItem.sProdutoDescricao = TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_DescProduto_Col)
                objItem.dValorDesconto = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_Desconto_Col))
        
            End With
            
            lErro = objTrib.IncluirItem(objItem)
            If lErro <> SUCESSO Then Error 27831
            
        End If
        
    Else
        If bSRV Then
            lErro = mvarobjForm.Move_GridItem_MemoriaSRV(objTrib, iIndice, sProduto)
        Else
            lErro = mvarobjForm.Move_GridItem_Memoria(objTrib, iIndice, sProduto)
        End If
        If lErro <> SUCESSO Then Error 27831
    End If

    Move_GridItem_Memoria = SUCESSO

    Exit Function

Erro_Move_GridItem_Memoria:

    Move_GridItem_Memoria = Err

    Select Case Err

        Case 20768
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)

        Case 27831

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153736)

    End Select

    Exit Function

End Function

Function Ativar(objForm As Object, Optional iEntrada As Integer = 0, Optional iLivroFiscal As Integer = 0, Optional ByVal iTipo As Integer = 0, Optional ByVal iSubTipo As Integer = 0, Optional ByVal bTemSRV As Boolean = False) As Long

Dim lErro As Long, objControl As Object
On Error GoTo Erro_Ativar

    mvariNFEntrada = iEntrada
    mvariLivroFiscal = iLivroFiscal
    iTipoTela = iTipo
    iSubTipoTela = iSubTipo
    gbTemSRV = bTemSRV
    mvariTrazendoTribItemTela = 1
    
    Set gobjVersao = New ClassVersaoNFe
    
    gobjVersao.iCodigo = gobjCRFAT.iVersaoNFE
    
    lErro = CF("VersaoNFe_Le", gobjVersao)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then Error 27838

    Set mvarobjForm = objForm
    
    Set mvarobjTabTrib = Nothing
    For Each objControl In mvarobjForm.Controls
        If objControl.Name = "TabTrib" Then
            Set mvarobjTabTrib = objControl
            Set mvarobjTabTrib.gobjForm = mvarobjForm
            Exit For
        End If
    Next
    
    lErro = CarregaTiposTrib()
    If lErro <> SUCESSO Then Error 27838

    lErro = Carrega_Outros_Controles()
    If lErro <> SUCESSO Then Error 27838

    Set mvarobjEventoTiposDeTributacao = New AdmEvento
    Set mvarobjEventoNaturezaInterna = New AdmEvento
    Set mvarobjEventoCidadeIBGE = New AdmEvento

'-------------------------------------NFe 3.10-------------------------------------
    Set mvarobjEventoISSCodServ = New AdmEvento
    Set mvarobjEventoISSMunicIncidImp = New AdmEvento
    Set mvarobjEventoISSCodPais = New AdmEvento
'-------------------------------------NFe 3.10-------------------------------------

    Set mvarobjEventoIPICodEnq = New AdmEvento

    lErro = TributacaoNF_Reset()
    If lErro <> SUCESSO Then Error 27839
    
    If mvariNFEntrada = MARCADO Then
        TabTrib_Controle("ICMSCredita").Caption = "Credita"
        TabTrib_Controle("IPICredita").Caption = "Credita"
        TabTrib_Controle("LabelICMSCredito").Caption = "Crédito"
        TabTrib_Controle("LabelIPICredito").Caption = "Crédito"
        TabTrib_Controle("LabelPISCredito").Caption = "Crédito"
        TabTrib_Controle("LabelCOFINSCredito").Caption = "Crédito"
    Else
        TabTrib_Controle("ICMSCredita").Caption = "Debita"
        TabTrib_Controle("IPICredita").Caption = "Debita"
        TabTrib_Controle("LabelICMSCredito").Caption = "Débito"
        TabTrib_Controle("LabelIPICredito").Caption = "Débito"
        TabTrib_Controle("LabelPISCredito").Caption = "Débito"
        TabTrib_Controle("LabelCOFINSCredito").Caption = "Débito"
    End If
    TabTrib_Controle("ICMSCredito").Enabled = True
    TabTrib_Controle("IPICredito").Enabled = True
    
    If iSubTipoTela = SUBTIPOTELA_CONSULTA Then
        TabTrib_Controle("FrameTabTribDet")(1).Enabled = False
        TabTrib_Controle("FrameTabTribDet")(2).Enabled = False
        TabTrib_Controle("FrameTabTribDet")(3).Enabled = False
        TabTrib_Controle("FrameTabTribDet")(4).Enabled = False
        TabTrib_Controle("FrameTabTribDet")(5).Enabled = False
        TabTrib_Controle("FrameTabTribDet")(6).Enabled = False
        TabTrib_Controle("FrameTabTribDet")(7).Enabled = False
        TabTrib_Controle("FrameImpostoImportacao").Enabled = False
        TabTrib_Controle("FrameResumo")(1).Enabled = False
        TabTrib_Controle("FrameResumo")(2).Enabled = False
        TabTrib_Controle("FrameResumo")(3).Enabled = False
    Else
        TabTrib_Controle("FrameTabTribDet")(1).Enabled = True
        TabTrib_Controle("FrameTabTribDet")(2).Enabled = True
        TabTrib_Controle("FrameTabTribDet")(3).Enabled = True
        TabTrib_Controle("FrameTabTribDet")(4).Enabled = True
        TabTrib_Controle("FrameTabTribDet")(5).Enabled = True
        TabTrib_Controle("FrameTabTribDet")(6).Enabled = True
        TabTrib_Controle("FrameTabTribDet")(7).Enabled = True
        TabTrib_Controle("FrameImpostoImportacao").Enabled = False
        TabTrib_Controle("FrameResumo")(1).Enabled = True
        TabTrib_Controle("FrameResumo")(2).Enabled = True
        TabTrib_Controle("FrameResumo")(3).Enabled = True
    End If

    mvariFrameAtual = 1
    mvariTrazendoTribItemTela = 0

    Ativar = SUCESSO

    Exit Function

Erro_Ativar:

    Ativar = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153737)

    End Select

    Exit Function

End Function

Function Desativar() As Long

Dim lErro As Long

On Error GoTo Erro_Desativar
    
    lErro = TributacaoNF_Terminar()
    If lErro <> SUCESSO Then Error 27837

    Set mvarobjEventoTiposDeTributacao = Nothing
    Set mvarobjEventoNaturezaInterna = Nothing
    Set mvarobjEventoCidadeIBGE = Nothing
    
'-------------------------------------NFe 3.10-------------------------------------
    Set mvarobjEventoISSCodServ = Nothing
    Set mvarobjEventoISSMunicIncidImp = Nothing
    Set mvarobjEventoISSCodPais = Nothing
'-------------------------------------NFe 3.10-------------------------------------

    Set mvarobjEventoIPICodEnq = Nothing
    
    Set mvarcolTiposTribICMS = Nothing
    Set mvarcolTiposTribIPI = Nothing
    Set mvarcolTiposTribPISCOFINS = Nothing
    Set mvarcolTiposTribICMSSimples = Nothing
    Set mvarcolTiposTribISS = Nothing
    
    Set gobjVersao = Nothing

    Set mvarobjForm = Nothing
    
    Set objFilialEmpresa = Nothing
    
    Desativar = SUCESSO

    Exit Function

Erro_Desativar:

    Desativar = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153738)

    End Select

    Exit Function

End Function

Function TabClick() As Long
    TabClick = TributacaoItem_InicializaTab()
End Function

Sub FilialFornecedor_Alterada(lCodigo As Long, iCodFilial As Integer)

    mvarobjNFTributacao.lFornecedor = lCodigo
    mvarobjNFTributacao.iFilialForn = iCodFilial
    
    Call gobjDocTrib2.Alterado

    Call mvarobjForm.ValorTotal_Calcula
    
End Sub

Sub FilialCliente_Alterada(lCodigo As Long, iCodFilial As Integer)

Dim lErro As Long
Dim objTribItem As Object
Dim objFilialCliente As New ClassFilialCliente
'Dim objCidade As New ClassCidades
Dim objEndereco As New ClassEndereco
Dim objEnderecoEnt As New ClassEndereco
Dim bTrocouCli As Boolean

On Error GoTo Erro_FilialCliente_Alterada

    bTrocouCli = False

    If iTipoTela = TIPOTELA_NF Then
        If mvarobjNFTributacao.lCliente <> lCodigo Or mvarobjNFTributacao.iFilialCli <> iCodFilial Then
            bTrocouCli = True
        End If
        mvarobjNFTributacao.lCliente = lCodigo
        mvarobjNFTributacao.iFilialCli = iCodFilial
    Else
        If gobjDocTrib.lCliente <> lCodigo Or gobjDocTrib.iFilial <> iCodFilial Then
            bTrocouCli = True
        End If
        gobjDocTrib.lCliente = lCodigo
        gobjDocTrib.iFilial = iCodFilial
    End If
    
    If lCodigo <> 0 And iCodFilial <> 0 Then
        
        objFilialCliente.lCodCliente = lCodigo
        objFilialCliente.iCodFilial = iCodFilial

        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM

        objEndereco.lCodigo = objFilialCliente.lEndereco

        lErro = CF("Endereco_Le", objEndereco)
        If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM

        objEnderecoEnt.lCodigo = objFilialCliente.lEnderecoEntrega

        lErro = CF("Endereco_Le", objEnderecoEnt)
        If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM

'        objCidade.sDescricao = objEndereco.sCidade
'        lErro = CF("Cidade_Le_Nome", objCidade)
'        If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
'
'        For Each objTribItem In gcolDocTribItem
'            gobjDocTribItemTributacao(objTribItem).sISSCidadeIBGE = objCidade.sCodIBGE
'        Next
'
'        sCidadeIBGECodigo = objCidade.sCodIBGE

    End If
    
    If bTrocouCli Then
        Call Atualiza_End_Entrega(objFilialCliente, objEndereco, objEnderecoEnt)
    End If
    
    Call gobjDocTrib2.Alterado

    Call mvarobjForm.ValorTotal_Calcula
    
    Exit Sub

Erro_FilialCliente_Alterada:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153739)

    End Select

    Exit Sub
    
End Sub

Sub Atualiza_End_Entrega(ByVal objFilialCliente As ClassFilialCliente, ByVal objEndereco As ClassEndereco, ByVal objEnderecoEnt As ClassEndereco)

Dim objInfoAdi As ClassInfoAdic
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Atualiza_End_Entrega

    'Se não tiver apenas trazendo dados já cadastrados
    If Not mvarobjForm.gbCarregandoTela Then

        'Se é uma saída
        If mvariNFEntrada = NF_SAIDA Then
        
            'Instância o gobj se ele ainda estiver com nothing
            'Se não tiver o obj vai pular fora e não fazer nada
            If mvarobjForm.gobjInfoAdic Is Nothing Then Set mvarobjForm.gobjInfoAdic = New ClassInfoAdic
            
            'Se o cliente tem endereço de entrega preenche com ele
            If Len(Trim(objEnderecoEnt.sEndereco)) > 0 Then
                
                Set objInfoAdi = mvarobjForm.gobjInfoAdic
            
                'Limpa os dados atuais pois trocou de cliente
                Set objInfoAdi.objRetEnt = New ClassRetiradaEntrega
            
                Set objInfoAdi.objRetEnt.objEnderecoEnt = New ClassEndereco
                'Call objInfoAdi.objRetEnt.objEnderecoEnt.Copiar(objEnderecoEnt)
                
                objInfoAdi.objRetEnt.objEnderecoEnt.iCodigoPais = objEnderecoEnt.iCodigoPais
                objInfoAdi.objRetEnt.objEnderecoEnt.lNumero = objEnderecoEnt.lNumero
                objInfoAdi.objRetEnt.objEnderecoEnt.sBairro = objEnderecoEnt.sBairro
                objInfoAdi.objRetEnt.objEnderecoEnt.sCEP = objEnderecoEnt.sCEP
                objInfoAdi.objRetEnt.objEnderecoEnt.sCidade = objEnderecoEnt.sCidade
                objInfoAdi.objRetEnt.objEnderecoEnt.sComplemento = objEnderecoEnt.sComplemento
                objInfoAdi.objRetEnt.objEnderecoEnt.sEndereco = objEnderecoEnt.sEndereco
                objInfoAdi.objRetEnt.objEnderecoEnt.sLogradouro = objEnderecoEnt.sLogradouro
                objInfoAdi.objRetEnt.objEnderecoEnt.sSiglaEstado = objEnderecoEnt.sSiglaEstado
                objInfoAdi.objRetEnt.objEnderecoEnt.sTipoLogradouro = objEnderecoEnt.sTipoLogradouro
                
                objInfoAdi.objRetEnt.iFilialCliEnt = objFilialCliente.iCodFilial
                objInfoAdi.objRetEnt.lClienteEnt = objFilialCliente.lCodCliente
                objInfoAdi.objRetEnt.sCNPJCPFEnt = objFilialCliente.sCgc
                'objInfoAdi.objRetEnt.objEnderecoEnt.lCodigo = 0
                
                Set objInfoAdi.objRetEnt.objEnderecoRet = New ClassEndereco
                objInfoAdi.objRetEnt.objEnderecoRet.iCodigoPais = PAIS_BRASIL
            
            Else
                
                Set objInfoAdi = mvarobjForm.gobjInfoAdic
            
                Set objInfoAdi.objRetEnt = Nothing
            
            End If
            
        End If

    End If

    Exit Sub

Erro_Atualiza_End_Entrega:

    Exit Sub
    
End Sub

Private Property Get gobjInfoAdic() As ClassInfoAdic

On Error GoTo Erro_gobjInfoAdic

    Set gobjInfoAdic = mvarobjForm.gobjInfoAdic

    Exit Property

Erro_gobjInfoAdic:

    Set gobjInfoAdic = Nothing

    Exit Property
    
End Property

Function Traz_NFiscal_Tela(Optional objAux As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_NFiscal_Tela

    lErro = TributacaoNF_Reset(objAux)
    If lErro <> SUCESSO Then Error 27833

    lErro = Carrega_Tab_Tributacao(gobjDocTribTributacao)
    If lErro <> SUCESSO Then Error 27834

    If mvariFrameAtual = 2 Then
    
        lErro = TributacaoItem_InicializaTab
        If lErro <> SUCESSO Then Error 56958

    End If
    
    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    Traz_NFiscal_Tela = Err

    Select Case Err

        Case 27833, 27834, 56958
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153739)

    End Select

    Exit Function

End Function

Function Traz_NFiscalOrig_Tela(objNFiscal As ClassNFiscal) As Long

Dim lErro As Long

On Error GoTo Erro_Traz_NFiscalOrig_Tela

    lErro = TributacaoNFOrig_Reset(objNFiscal)
    If lErro <> SUCESSO Then Error 59300

    lErro = gobjTributacao.AtualizaImpostos(mvarobjNFTributacao, 1)
    If lErro <> SUCESSO Then Error 59301
    
    lErro = Carrega_Tab_Tributacao(mvarobjNFTributacao)
    If lErro <> SUCESSO Then Error 59302

    If mvariFrameAtual = 2 Then
    
        lErro = TributacaoItem_InicializaTab
        If lErro <> SUCESSO Then Error 59303

    End If
    
    If iTipoTela = TIPOTELA_NF Then
        If mvarobjNFTributacao.iTipoNFiscal = DOCINFO_NFIEDV Then
            Call Tributacao_TrazerTela(mvarobjNFTributacao)
        End If
    End If
    
    Traz_NFiscalOrig_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscalOrig_Tela:

    Traz_NFiscalOrig_Tela = Err

    Select Case Err

        Case 59300 To 59303
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153740)

    End Select

    Exit Function

End Function

Sub Limpa_Tela()
    
    mvariTrazendoTribItemTela = 1
    mvariTrazendoTribTela = 1
    
    With mvarobjForm
    
        'tab de tributacao resumo
        TabTrib_Controle("DescTipoTrib").Caption = ""
        TabTrib_Controle("DescNatOpInterna").Caption = ""
'        TabTrib_Controle("TributacaoAutomatica").Value = 1
'       TabTrib_Controle("ISSIncluso").Value = 0
        TabTrib_Controle("ISSValor").Caption = ""
        TabTrib_Controle("ISSValor1").Caption = ""
        TabTrib_Controle("IPIBase").Caption = ""
        TabTrib_Controle("IPIValor").Caption = ""
        TabTrib_Controle("IPIValor1").Caption = ""
        TabTrib_Controle("IPICredito").Caption = ""
        TabTrib_Controle("ISSBase").Caption = ""
        TabTrib_Controle("ISSBase1").Caption = ""
        TabTrib_Controle("IRBase").Text = ""
        TabTrib_Controle("IRAliquota").Text = ""
        TabTrib_Controle("ValorIRRF").Text = ""
        TabTrib_Controle("ICMSBase").Caption = ""
        TabTrib_Controle("ICMSBase1").Caption = ""
        TabTrib_Controle("ICMSValor").Caption = ""
        TabTrib_Controle("ICMSValor1").Caption = ""
        TabTrib_Controle("ICMSSubstBase").Caption = ""
        TabTrib_Controle("ICMSSubstBase1").Caption = ""
        TabTrib_Controle("ICMSSubstValor").Caption = ""
        TabTrib_Controle("ICMSSubstValor1").Caption = ""
        TabTrib_Controle("ICMSCredito").Caption = ""
        TabTrib_Controle("INSSBase").Text = ""
        TabTrib_Controle("INSSDeducoes").Text = ""
        TabTrib_Controle("INSSValor").Text = ""
        TabTrib_Controle("INSSRetido").Value = vbUnchecked
        TabTrib_Controle("PISValor").Caption = ""
        TabTrib_Controle("COFINSValor").Caption = ""
        TabTrib_Controle("TotTrib").Text = ""
        TabTrib_Controle("TotTribFed").Text = ""
        TabTrib_Controle("TotTribEst").Text = ""
        TabTrib_Controle("TotTribMunic").Text = ""
        TabTrib_Controle("TotTribTipo").ListIndex = 0
        TabTrib_Controle("PISCredito").Caption = ""
        TabTrib_Controle("COFINSCredito").Caption = ""
    
        'nfe 3.10
        TabTrib_Controle("ICMSDesonerado").Caption = ""
        TabTrib_Controle("ISSValorDeducao").Caption = ""
        TabTrib_Controle("ISSValorOutrasRet").Caption = ""
        TabTrib_Controle("ISSValorDescIncond").Caption = ""
        TabTrib_Controle("ISSValorDescCond").Caption = ""
        TabTrib_Controle("PISValorST").Caption = ""
        TabTrib_Controle("COFINSValorST").Caption = ""
        TabTrib_Controle("IIValor").Caption = ""
        
        'tab de tributacao itens
        TabTrib_Controle("ComboItensTrib").Clear
        TabTrib_Controle("LabelValorItem").Caption = ""
        TabTrib_Controle("LabelQtdeItem").Caption = ""
        TabTrib_Controle("LabelUMItem").Caption = ""
        TabTrib_Controle("LabelDescrNatOpItem").Caption = ""
        TabTrib_Controle("DescTipoTribItem").Caption = ""


        TabTrib_Controle("ISSDescIBGE").Caption = ""
        TabTrib_Controle("ISSListaServ").ListIndex = -1
        TabTrib_Controle("ComboCOFINSTipo").ListIndex = -1
        TabTrib_Controle("COFINSTipoCalculo").ListIndex = -1
        TabTrib_Controle("COFINSSTTipoCalculo").ListIndex = -1
        TabTrib_Controle("ComboPISTipo").ListIndex = -1
        TabTrib_Controle("PISTipoCalculo").ListIndex = -1
        TabTrib_Controle("PISSTTipoCalculo").ListIndex = -1
        TabTrib_Controle("UMTrib").Caption = ""
        TabTrib_Controle("QtdeTrib").Caption = ""
        TabTrib_Controle("ValorTrib").Caption = ""
        TabTrib_Controle("OrigemMercadoria").ListIndex = -1
        TabTrib_Controle("FCI").Text = ""
        TabTrib_Controle("ICMSModalidade").ListIndex = -1
        TabTrib_Controle("ICMSSTModalidade").ListIndex = -1
        TabTrib_Controle("IPITipoCalculo").ListIndex = -1

        TabTrib_Controle("ComboICMSSimplesTipo").ListIndex = -1
        TabTrib_Controle("ComboISSTipo").ListIndex = -1
        TabTrib_Controle("ICMSUFDevidoST").ListIndex = -1
        TabTrib_Controle("ICMSMotivo").ListIndex = -1

        TabTrib_Controle("NatBCCred").ListIndex = -1
        TabTrib_Controle("TotTribItemTipo").ListIndex = 0

        'nfe 3.10
        TabTrib_Controle("ISSMunicIncidDesc").Caption = ""
        TabTrib_Controle("ISSCodServDesc").Caption = ""
        TabTrib_Controle("ISSIndIncentivoItem").Value = vbUnchecked
        TabTrib_Controle("ISSIndExigibilidadeItem").ListIndex = -1
        TabTrib_Controle("ISSCodPaisItem").ListIndex = -1
        
        If Not (mvarobjTabTrib Is Nothing) Then
            Call mvarobjTabTrib.Limpa_TabTrib
        End If
        
        'NT 2015.003 (ICMS Interestadual)
        TabTrib_Controle("ICMSInterestVlrFCPUFDest").Caption = ""
        TabTrib_Controle("ICMSInterestVlrUFDest").Caption = ""
        TabTrib_Controle("ICMSInterestVlrUFRemet").Caption = ""
        TabTrib_Controle("ICMSInterestAliqItem").ListIndex = -1
        TabTrib_Controle("ICMSInterestPercPartilhaItem").ListIndex = -1
        
        TabTrib_Controle("IPICodEnqDesc").Caption = ""
        
        'NFe 4.00
        TabTrib_Controle("ICMSVlrFCP").Caption = ""
        TabTrib_Controle("ICMSVlrFCPSTRet").Caption = ""
        TabTrib_Controle("ICMSVlrFCPST").Caption = ""
        TabTrib_Controle("IPIVlrDevolvido").Caption = ""
        
'        sCidadeIBGECodigo = ""

    End With
    
    Call TributacaoNF_Reset

    mvariTrazendoTribItemTela = 0
    mvariTrazendoTribTela = 0

End Sub

Sub ValorFrete_Validate(bCancel As Boolean, dValor As Double)
    
    gobjDocTrib.dValorFrete = dValor
    
    Call gobjDocTrib2.Alterado

End Sub

Sub ValorSeguro_Validate(bCancel As Boolean, dValor As Double)
    
    gobjDocTrib.dValorSeguro = dValor
    
    Call gobjDocTrib2.Alterado

End Sub

Sub ValorDesconto_Validate(bCancel As Boolean, dValor As Double)
    
    gobjDocTrib.dValorDesconto = dValor
    
    Call gobjDocTrib2.Alterado

End Sub

Sub ValorDespesas_Validate(bCancel As Boolean, dValor As Double)
    
    gobjDocTrib.dValorOutrasDespesas = dValor
    
    Call gobjDocTrib2.Alterado

End Sub

Sub TipoNFiscal_Definir(iTipoNFiscal As Integer, sSiglaTipoNFiscal As String)

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_TipoNFiscal_Definir

    mvarobjNFTributacao.iTipoNFiscal = iTipoNFiscal
    mvarobjNFTributacao.sSiglaTipoNFiscal = sSiglaTipoNFiscal

    objTipoDocInfo.iCodigo = iTipoNFiscal

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then Error 56990

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then Error 56991

    mvariEmitenteNF = objTipoDocInfo.iEmitente
    
    If Not (mvarobjForm Is Nothing) Then
    
        Call gobjDocTrib2.Alterado

        Call mvarobjForm.ValorTotal_Calcula
    
    End If
    
    Exit Sub
     
Erro_TipoNFiscal_Definir:

    Select Case Err
          
        Case 56990
        
        Case 56991
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", Err, objTipoDocInfo.iCodigo)
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153741)
     
    End Select
     
    Exit Sub

End Sub

Sub ICMSAliquotaItem_Change()

    mvariICMSAliquotaItemAlterado = 1
    
End Sub

Sub ICMSAliquotaItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_ICMSAliquotaItem_Validate

    If mvariICMSAliquotaItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSAliquotaItem").Text)
            If lErro <> SUCESSO Then Error 56865
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSAliquotaItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56865
''???            If TabTrib_Controle("ICMSAliquotaItem").Visible Then TabTrib_Controle("ICMSAliquotaItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153742)

    End Select

    Exit Sub

End Sub

Sub ICMSBaseItem_Change()
    
    mvariICMSBaseItemAlterado = 1

End Sub

Sub ICMSBaseItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_ICMSBaseItem_Validate

    If mvariICMSBaseItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSBaseItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSBaseItem").Text)
            If lErro <> SUCESSO Then Error 56863
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSBaseItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56863
''???            If TabTrib_Controle("ICMSBaseItem").Visible Then TabTrib_Controle("ICMSBaseItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153743)

    End Select

    Exit Sub

End Sub

Sub ICMSPercRedBaseItem_Change()
    
    mvariICMSPercRedBaseItemAlterado = 1

End Sub

Sub ICMSPercRedBaseItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_ICMSPercRedBaseItem_Validate

    If mvariICMSPercRedBaseItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSPercRedBaseItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSPercRedBaseItem").Text)
            If lErro <> SUCESSO Then Error 56866
            
        End If
            
        Call AtualizarTabTribItem
        
        mvariICMSPercRedBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSPercRedBaseItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56866
''???            If TabTrib_Controle("ICMSPercRedBaseItem").Visible Then TabTrib_Controle("ICMSPercRedBaseItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153744)

    End Select

    Exit Sub

End Sub

Sub ICMSSubstAliquotaItem_Change()
    
    mvariICMSSubstAliquotaItemAlterado = 1

End Sub

Sub ICMSSubstAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstAliquotaItem_Validate

    If mvariICMSSubstAliquotaItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSSubstAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSSubstAliquotaItem").Text)
            If lErro <> SUCESSO Then Error 56867
            
        End If
                
        Call AtualizarTabTribItem
        
        mvariICMSSubstAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSubstAliquotaItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56867
''???            If TabTrib_Controle("ICMSSubstAliquotaItem").Visible Then TabTrib_Controle("ICMSSubstAliquotaItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153745)

    End Select

    Exit Sub

End Sub

Sub ICMSSubstBaseItem_Change()

    mvariICMSSubstBaseItemAlterado = 1
    
End Sub

Sub ICMSSubstBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstBaseItem_Validate

    If mvariICMSSubstBaseItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSSubstBaseItem").Text)) > 0 Then
    
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSSubstBaseItem").Text)
            If lErro <> SUCESSO Then Error 56872
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSSubstBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSubstBaseItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56872
''???            If TabTrib_Controle("ICMSSubstBaseItem").Visible Then TabTrib_Controle("ICMSSubstBaseItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153746)

    End Select

    Exit Sub

End Sub

Sub ICMSSubstValorItem_Change()

    mvariICMSSubstValorItemAlterado = 1

End Sub

Sub ICMSSubstValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstValorItem_Validate

    If mvariICMSSubstValorItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSSubstValorItem").Text)) > 0 Then
    
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSSubstValorItem").Text)
            If lErro <> SUCESSO Then Error 56873
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSSubstValorItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSubstValorItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56873
''???            If TabTrib_Controle("ICMSSubstValorItem").Visible Then TabTrib_Controle("ICMSSubstValorItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153747)

    End Select

    Exit Sub

End Sub

Sub ICMSValorItem_Change()

    mvariICMSValorItemAlterado = 1

End Sub

Sub ICMSValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSValorItem_Validate

    If mvariICMSValorItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSValorItem").Text)) > 0 Then
    
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSValorItem").Text)
            If lErro <> SUCESSO Then Error 56874
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSValorItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSValorItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56874
''???            If TabTrib_Controle("ICMSValorItem").Visible Then TabTrib_Controle("ICMSValorItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153748)

    End Select

    Exit Sub

End Sub

Sub IPIAliquotaItem_Change()

    mvariIPIAliquotaItemAlterado = 1
    
End Sub

Sub IPIAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIAliquotaItem_Validate

    If mvariIPIAliquotaItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("IPIAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("IPIAliquotaItem").Text)
            If lErro <> SUCESSO Then Error 56868
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariIPIAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_IPIAliquotaItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56868
''???            If TabTrib_Controle("IPIAliquotaItem").Visible Then TabTrib_Controle("IPIAliquotaItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub IPIBaseItem_Change()

    mvariIPIBaseItemAlterado = 1
    
End Sub

Sub IPIBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIBaseItem_Validate

    If mvariIPIBaseItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("IPIBaseItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IPIBaseItem").Text)
            If lErro <> SUCESSO Then Error 56875
        
        End If
    
        Call AtualizarTabTribItem
        
        mvariIPIBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_IPIBaseItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56875
''???            If TabTrib_Controle("IPIBaseItem").Visible Then TabTrib_Controle("IPIBaseItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153750)

    End Select

    Exit Sub

End Sub

Sub IPIPercRedBaseItem_Change()

    mvariIPIPercRedBaseItemAlterado = 1
    
End Sub

Sub IPIPercRedBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIPercRedBaseItem_Validate

    If mvariIPIPercRedBaseItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("IPIPercRedBaseItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("IPIPercRedBaseItem").Text)
            If lErro <> SUCESSO Then Error 56869
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariIPIPercRedBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_IPIPercRedBaseItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56869
''???            If TabTrib_Controle("IPIPercRedBaseItem").Visible Then TabTrib_Controle("IPIPercRedBaseItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153751)

    End Select

    Exit Sub

End Sub

Sub IPIValorItem_Change()

    mvariIPIValorItemAlterado = 1
    
End Sub

Sub IPIValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIValorItem_Validate

    If mvariIPIValorItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("IPIValorItem").Text)) > 0 Then
    
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IPIValorItem").Text)
            If lErro <> SUCESSO Then Error 56876
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariIPIValorItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_IPIValorItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56876
''???            If TabTrib_Controle("IPIValorItem").Visible Then TabTrib_Controle("IPIValorItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153752)

    End Select

    Exit Sub

End Sub

Sub IRAliquota_Change()
    
    mvariAliqIRAlterada = 1

End Sub

Sub IRAliquota_Validate(Cancel As Boolean)

Dim dIRAliquota As Double, dIRValor As Double
Dim lErro As Long

On Error GoTo Erro_IRAliquota_Validate

    If mvariAliqIRAlterada = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IRAliquota").Text)) > 0 Then
    
        lErro = Porcentagem_Critica(TabTrib_Controle("IRAliquota").Text)
        If lErro <> SUCESSO Then Error 56870
        
    End If
    
    Call AtualizarTabTrib
        
    mvariAliqIRAlterada = 0
    
    Exit Sub
    
Erro_IRAliquota_Validate:

    Cancel = True
    
    Select Case Err

        Case 56870
''???            If TabTrib_Controle("IRAliquota").Visible Then TabTrib_Controle("IRAliquota").SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153753)

    End Select

    Exit Sub

End Sub

Private Sub mvarobjEventoNaturezaInterna_evSelecao(obj1 As Object)

Dim objNaturezaOp As ClassNaturezaOp

    Set objNaturezaOp = obj1
    
    With mvarobjForm

        If mvariFrameAtual = 1 Then
    
            'Coloca a natureza Operação na tela
            TabTrib_Controle("NatOpInterna").Text = objNaturezaOp.sCodigo
            Call NatOpInterna_Validate(bSGECancelDummy)
    
        Else
    
            'Coloca a natureza Operação na tela
            TabTrib_Controle("NaturezaOpItem").Text = objNaturezaOp.sCodigo
            Call NaturezaOpItem_Validate(bSGECancelDummy)
    
        End If

        .Show

    End With

    Exit Sub

End Sub

Private Sub mvarobjEventoTiposDeTributacao_evSelecao(obj1 As Object)

Dim objTipoTrib As ClassTipoDeTributacaoMovto

    Set objTipoTrib = obj1

    If mvariFrameAtual = 1 Then

        TabTrib_Controle("TipoTributacao").Text = objTipoTrib.iTipo
        Call TipoTributacao_Validate(bSGECancelDummy)

    Else

        TabTrib_Controle("TipoTributacaoItem").Text = objTipoTrib.iTipo
        Call TipoTributacaoItem_Validate(bSGECancelDummy)

    End If

    mvarobjForm.Show

    Exit Sub

End Sub

Sub LblTipoTrib_Click()

Dim colSelecao As Collection
Dim objTipoTrib As New ClassTipoDeTributacaoMovto

    Set colSelecao = New Collection
    
    If mvariNFEntrada = 1 Then
        colSelecao.Add "1"
        colSelecao.Add "1"
    Else
        colSelecao.Add "0"
        colSelecao.Add "0"
    End If
    
    Call Chama_Tela("TiposDeTribMovtoLista", colSelecao, objTipoTrib, mvarobjEventoTiposDeTributacao)
    
End Sub

Sub LblTipoTribItem_Click()

    Call LblTipoTrib_Click
    
End Sub

Public Function PreencherNF(objNFiscal As ClassNFiscal) As Long
'copia as informacoes sobre tributacao para objNFiscal

Dim lErro As Long, objTributacaoNF As New ClassTributacaoDoc, objTribItemNFAux As ClassTribItemNFAux
Dim objTribItemNF As ClassTributacaoDocItem, iIndice As Integer

On Error GoTo Erro_PreencherNF

    With objTributacaoNF

        .dICMSBase = mvarobjNFTributacao.dICMSBase
        .dICMSCredito = mvarobjNFTributacao.dICMSCredito
        .dICMSSubstBase = mvarobjNFTributacao.dICMSSubstBase
        .dICMSSubstValor = mvarobjNFTributacao.dICMSSubstValor
        .dICMSValor = mvarobjNFTributacao.dICMSValor
        .dIPIBase = mvarobjNFTributacao.dIPIBase
        .dIPICredito = mvarobjNFTributacao.dIPICredito
        .dIPIValor = mvarobjNFTributacao.dIPIValor
        .dIRRFAliquota = mvarobjNFTributacao.dIRRFAliquota
        .dIRRFBase = mvarobjNFTributacao.dIRRFBase
        .dIRRFValor = mvarobjNFTributacao.dIRRFValor
        .dISSAliquota = mvarobjNFTributacao.dISSAliquota
        .dISSBase = mvarobjNFTributacao.dISSBase
        .dISSValor = mvarobjNFTributacao.dISSValor
        .iISSIncluso = mvarobjNFTributacao.iISSIncluso
        .iTipoTributacao = mvarobjNFTributacao.iTipoTributacao
        .lNumIntDoc = objNFiscal.lNumIntDoc
        .sNaturezaOpInterna = mvarobjNFTributacao.sNaturezaOp
        .dINSSBase = mvarobjNFTributacao.dINSSBase
        .dINSSDeducoes = mvarobjNFTributacao.dINSSDeducoes
        .dValorINSS = mvarobjNFTributacao.dINSSValor
        .iINSSRetido = mvarobjNFTributacao.iINSSRetido
        .dPISCredito = mvarobjNFTributacao.dPISCredito
        .dCOFINSCredito = mvarobjNFTributacao.dCOFINSCredito
        .dPISRetido = mvarobjNFTributacao.dPISRetido
        .dISSRetido = mvarobjNFTributacao.dISSRetido
        .dCOFINSRetido = mvarobjNFTributacao.dCOFINSRetido
        .dCSLLRetido = mvarobjNFTributacao.dCSLLRetido
        .dTotTrib = mvarobjNFTributacao.dTotTrib
        .iTotTribManual = mvarobjNFTributacao.iTotTribManual
        .iTotTribTipo = mvarobjNFTributacao.iTotTribTipo
        .iTotTribTipoManual = mvarobjNFTributacao.iTotTribTipoManual
        .dPISValor = mvarobjNFTributacao.dPISValor
        .dCOFINSValor = mvarobjNFTributacao.dCOFINSValor

        .iISSRetidoManual = mvarobjNFTributacao.iISSRetidoManual
        .iTipoTributacaoManual = mvarobjNFTributacao.iTipoTributacaoManual
        .iNaturezaOpManual = mvarobjNFTributacao.iNaturezaOpManual
        .iPISRetidoManual = mvarobjNFTributacao.iPISRetidoManual
        .iCOFINSRetidoManual = mvarobjNFTributacao.iCOFINSRetidoManual
        .iCSLLRetidoManual = mvarobjNFTributacao.iCSLLRetidoManual
        .iIRRFAliquotaManual = mvarobjNFTributacao.iIRRFAliquotaManual
        .iIRRFBaseManual = mvarobjNFTributacao.iIRRFBaseManual
        .iIRRFValorManual = mvarobjNFTributacao.iIRRFValorManual

        .iPISCreditoManual = mvarobjNFTributacao.iPISCreditoManual
        .iCOFINSCreditoManual = mvarobjNFTributacao.iCOFINSCreditoManual

        'nfe 3.10
        .dICMSValorIsento = mvarobjNFTributacao.dICMSValorIsento
        .dISSValorDeducao = mvarobjNFTributacao.dISSValorDeducao
        .dISSValorOutrasRet = mvarobjNFTributacao.dISSValorOutrasRet
        .dISSValorDescIncond = mvarobjNFTributacao.dISSValorDescIncond
        .dISSValorDescCond = mvarobjNFTributacao.dISSValorDescCond
        .iDestino = mvarobjNFTributacao.iDestino
        .iDestinoManual = mvarobjNFTributacao.iDestinoManual
        .iFinalidadeNFe = mvarobjNFTributacao.iFinalidadeNFe
        .iFinalidadeNFeManual = mvarobjNFTributacao.iFinalidadeNFeManual
        .iIndConsumidorFinal = mvarobjNFTributacao.iIndConsumidorFinal
        .iIndConsumidorFinalManual = mvarobjNFTributacao.iIndConsumidorFinalManual
        .iIndPresenca = mvarobjNFTributacao.iIndPresenca
        .iIndPresencaManual = mvarobjNFTributacao.iIndPresencaManual
        .iRegTribEspecial = mvarobjNFTributacao.iRegTribEspecial
        .iRegTribEspecialManual = mvarobjNFTributacao.iRegTribEspecialManual
        .dtDataPrestServico = mvarobjNFTributacao.dtDataPrestServico
        .iDataPrestServicoManual = mvarobjNFTributacao.iDataPrestServicoManual
    
        .dTotTribFed = mvarobjNFTributacao.dTotTribFed
        .iTotTribFedManual = mvarobjNFTributacao.iTotTribFedManual
        .dTotTribEst = mvarobjNFTributacao.dTotTribEst
        .iTotTribEstManual = mvarobjNFTributacao.iTotTribEstManual
        .dTotTribMunic = mvarobjNFTributacao.dTotTribMunic
        .iTotTribMunicManual = mvarobjNFTributacao.iTotTribMunicManual
        
        .dICMSInterestVlrFCPUFDest = mvarobjNFTributacao.dICMSInterestVlrFCPUFDest
        .dICMSInterestVlrUFDest = mvarobjNFTributacao.dICMSInterestVlrUFDest
        .dICMSInterestVlrUFRemet = mvarobjNFTributacao.dICMSInterestVlrUFRemet
        
        'NFe 4.00
        .dICMSVlrFCP = mvarobjNFTributacao.dICMSVlrFCP
        .dICMSVlrFCPST = mvarobjNFTributacao.dICMSVlrFCPST
        .dICMSVlrFCPSTRet = mvarobjNFTributacao.dICMSVlrFCPSTRet
        .dIPIVlrDevolvido = mvarobjNFTributacao.dIPIVlrDevolvido
    
    End With
    
    'preencher itens
    iIndice = 1
    For Each objTribItemNFAux In mvarobjNFTributacao.colItens
    
        lErro = PreencherNFItem(objTribItemNF, objTribItemNFAux)
        If lErro <> SUCESSO Then Error 22852
        
        objTribItemNF.lNumIntDoc = objNFiscal.lNumIntDoc
        
        Set objNFiscal.colItensNF(iIndice).objTributacaoItemNF = objTribItemNF
    
        iIndice = iIndice + 1
    Next
    
    If gbTemSRV Then
        'preencher itens
        iIndice = 1
        For Each objTribItemNFAux In mvarobjNFTributacao.colItensSRV
        
            lErro = PreencherNFItem(objTribItemNF, objTribItemNFAux)
            If lErro <> SUCESSO Then Error 22852
            
            objTribItemNF.lNumIntDoc = objNFiscal.lNumIntDoc
            
            Set objNFiscal.colItensSRV(iIndice).objTributacaoItemNF = objTribItemNF
        
            iIndice = iIndice + 1
        Next
    End If
    
    Set objNFiscal.objTributacaoNF = objTributacaoNF
    
    PreencherNF = SUCESSO

    Exit Function

Erro_PreencherNF:

    PreencherNF = Err

    Select Case Err

        Case 22848 To 22852
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153757)

    End Select

    Exit Function

End Function

Private Function PreencherNFItem(objTribItemNF As ClassTributacaoDocItem, objTribItemNFAux As ClassTribItemNFAux) As Long
'move os valores de objTribItemNFAux (estrutura p/uso na tela) para objTribItemNF (estrutura p/armazenar no bd)
Dim lErro As Long

On Error GoTo Erro_PreencherNFItem

    Set objTribItemNF = New ClassTributacaoDocItem
    
    With objTribItemNF
    
        .dICMSAliquota = objTribItemNFAux.dICMSAliquota
        .dICMSBase = objTribItemNFAux.dICMSBase
        .dICMSCredito = objTribItemNFAux.dICMSCredito
        .dICMSPercRedBase = objTribItemNFAux.dICMSPercRedBase
        .dICMSSubstAliquota = objTribItemNFAux.dICMSSubstAliquota
        .dICMSSubstBase = objTribItemNFAux.dICMSSubstBase
        .dICMSSubstValor = objTribItemNFAux.dICMSSubstValor
        .dICMSValor = objTribItemNFAux.dICMSValor
        .dIPIAliquota = objTribItemNFAux.dIPIAliquota
        .dIPIBaseCalculo = objTribItemNFAux.dIPIBaseCalculo
        .dIPICredito = objTribItemNFAux.dIPICredito
        .dIPIPercRedBase = objTribItemNFAux.dIPIPercRedBase
        .dIPIValor = objTribItemNFAux.dIPIValor
        .iICMSTipo = objTribItemNFAux.iICMSTipo
        .iIPITipo = objTribItemNFAux.iIPITipo
        .iTipoTributacao = objTribItemNFAux.iTipoTributacao
        .sNaturezaOp = objTribItemNFAux.sNaturezaOp
        .dPISCredito = objTribItemNFAux.dPISCredito
        .dCOFINSCredito = objTribItemNFAux.dCOFINSCredito
        .dICMSSubstPercRedBase = objTribItemNFAux.dICMSSubstPercRedBase
        .dICMSSubstPercMVA = objTribItemNFAux.dICMSSubstPercMVA
        
        .dICMSSTCobrAntBase = objTribItemNFAux.dICMSSTCobrAntBase
        .dICMSSTCobrAntValor = objTribItemNFAux.dICMSSTCobrAntValor
        
        'para nfs de importacao
        .dDespImpICMSBase = objTribItemNFAux.dDespImpICMSBase
        .dDespImpICMSValor = objTribItemNFAux.dDespImpICMSValor
        .dDespImpICMSCredito = objTribItemNFAux.dDespImpICMSCredito
        .dICMSAliquotaAdicaoDI = objTribItemNFAux.dICMSAliquotaAdicaoDI
        .dICMSPercRedBaseAdicaoDI = objTribItemNFAux.dICMSPercRedBaseAdicaoDI
        
        .dICMSSubstPercRedBase = objTribItemNFAux.dICMSSubstPercRedBase
        .dICMSSubstPercMVA = objTribItemNFAux.dICMSSubstPercMVA
        .iPISTipo = objTribItemNFAux.iPISTipo
        .iPISTipoCalculo = objTribItemNFAux.iPISTipoCalculo
        .dPISBase = objTribItemNFAux.dPISBase
        .dPISAliquota = objTribItemNFAux.dPISAliquota
        .dPISAliquotaValor = objTribItemNFAux.dPISAliquotaValor
        .dPISQtde = objTribItemNFAux.dPISQtde
        .dPISValor = objTribItemNFAux.dPISValor
        .iPISSTTipoCalculo = objTribItemNFAux.iPISSTTipoCalculo
        .dPISSTBase = objTribItemNFAux.dPISSTBase
        .dPISSTAliquota = objTribItemNFAux.dPISSTAliquota
        .dPISSTAliquotaValor = objTribItemNFAux.dPISSTAliquotaValor
        .dPISSTQtde = objTribItemNFAux.dPISSTQtde
        .dPISSTValor = objTribItemNFAux.dPISSTValor
        .iCOFINSTipo = objTribItemNFAux.iCOFINSTipo
        .iCOFINSTipoCalculo = objTribItemNFAux.iCOFINSTipoCalculo
        .dCOFINSBase = objTribItemNFAux.dCOFINSBase
        .dCOFINSAliquota = objTribItemNFAux.dCOFINSAliquota
        .dCOFINSAliquotaValor = objTribItemNFAux.dCOFINSAliquotaValor
        .dCOFINSQtde = objTribItemNFAux.dCOFINSQtde
        .dCOFINSValor = objTribItemNFAux.dCOFINSValor
        .iCOFINSSTTipoCalculo = objTribItemNFAux.iCOFINSSTTipoCalculo
        .dCOFINSSTBase = objTribItemNFAux.dCOFINSSTBase
        .dCOFINSSTAliquota = objTribItemNFAux.dCOFINSSTAliquota
        .dCOFINSSTAliquotaValor = objTribItemNFAux.dCOFINSSTAliquotaValor
        .dCOFINSSTQtde = objTribItemNFAux.dCOFINSSTQtde
        .dCOFINSSTValor = objTribItemNFAux.dCOFINSSTValor
        .sCST = Trim(objTribItemNFAux.sCST)
        .sCSOSN = Trim(objTribItemNFAux.sCSOSN)
        .sISSQN = Trim(objTribItemNFAux.sISSQN)
        .dISSBase = objTribItemNFAux.dISSBase
        .dISSAliquota = objTribItemNFAux.dISSAliquota
        .sISSCidadeIBGE = Trim(objTribItemNFAux.sISSCidadeIBGE)
        .dISSValor = objTribItemNFAux.dISSValor
        .dIIBase = objTribItemNFAux.dIIBase
        .dIIDespAduaneira = objTribItemNFAux.dIIDespAduaneira
        .dIIIOF = objTribItemNFAux.dIIIOF
        .dIIValor = objTribItemNFAux.dIIValor
        .iICMSBaseModalidade = objTribItemNFAux.iICMSBaseModalidade
        .iICMSSubstBaseModalidade = objTribItemNFAux.iICMSSubstBaseModalidade
        .sIPIEnquadramentoClasse = Trim(objTribItemNFAux.sIPIEnquadramentoClasse)
        .sIPIEnquadramentoCodigo = Trim(objTribItemNFAux.sIPIEnquadramentoCodigo)
        .sIPICNPJProdutor = Trim(objTribItemNFAux.sIPICNPJProdutor)
        .sIPISeloCodigo = Trim(objTribItemNFAux.sIPISeloCodigo)
        .lIPISeloQtde = objTribItemNFAux.lIPISeloQtde
        .iIPITipoCalculo = objTribItemNFAux.iIPITipoCalculo
        .dIPIUnidadePadraoQtde = objTribItemNFAux.dIPIUnidadePadraoQtde
        .dIPIUnidadePadraoValor = objTribItemNFAux.dIPIUnidadePadraoValor
        .dValorFreteItem = objTribItemNFAux.dValorFreteItem
        .dValorSeguroItem = objTribItemNFAux.dValorSeguroItem
        .dValorOutrasDespesasItem = objTribItemNFAux.dValorOutrasDespesasItem
        '.dValorDescontoItem = objTribItemNFAux.dValorDescontoItem
        .iOrigemMercadoria = objTribItemNFAux.iOrigemMercadoria
        .sFCI = objTribItemNFAux.sFCI
        .iExTIPI = objTribItemNFAux.iExTIPI
        .sGenero = Trim(objTribItemNFAux.sGenero)
        .iProdutoEspecifico = objTribItemNFAux.iProdutoEspecifico
        .sEAN = Trim(objTribItemNFAux.sEAN)
        .sEANTrib = Trim(objTribItemNFAux.sEANTrib)
        .dQtdTrib = objTribItemNFAux.dQtdTrib
        .sUMTrib = Trim(objTribItemNFAux.sUMTrib)
        .dValorUnitTrib = objTribItemNFAux.dValorUnitTrib
        .sIPICodProduto = Trim(objTribItemNFAux.sIPICodProduto)
        .dDescontoGrid = objTribItemNFAux.dValorDesconto
        .dPrecoTotal = objTribItemNFAux.dPrecoTotal
        
        .iICMSAliquotaManual = objTribItemNFAux.iICMSAliquotaManual
        .iICMSBaseManual = objTribItemNFAux.iICMSBaseManual
        .iICMSCreditoManual = objTribItemNFAux.iICMSCreditoManual
        .iICMSPercRedBaseManual = objTribItemNFAux.iICMSPercRedBaseManual
        .iICMSSubstAliquotaManual = objTribItemNFAux.iICMSSubstAliquotaManual
        .iICMSSubstBaseManual = objTribItemNFAux.iICMSSubstBaseManual
        .iICMSSubstValorManual = objTribItemNFAux.iICMSSubstValorManual
        .iICMSValorManual = objTribItemNFAux.iICMSValorManual
        .iIPIAliquotaManual = objTribItemNFAux.iIPIAliquotaManual
        .iIPIBaseManual = objTribItemNFAux.iIPIBaseManual
        .iIPICreditoManual = objTribItemNFAux.iIPICreditoManual
        .iIPIPercRedBaseManual = objTribItemNFAux.iIPIPercRedBaseManual
        .iIPIValorManual = objTribItemNFAux.iIPIValorManual
        .iICMSTipoManual = objTribItemNFAux.iICMSTipoManual
        .iIPITipoManual = objTribItemNFAux.iIPITipoManual
        .iTipoTributacaoManual = objTribItemNFAux.iTipoTributacaoManual
        .iNaturezaOpManual = objTribItemNFAux.iNaturezaOpManual
        .iPISCreditoManual = objTribItemNFAux.iPISCreditoManual
        .iCOFINSCreditoManual = objTribItemNFAux.iCOFINSCreditoManual
        .iICMSSubstPercRedBaseManual = objTribItemNFAux.iICMSSubstPercRedBaseManual
        .iICMSSubstPercMVAManual = objTribItemNFAux.iICMSSubstPercMVAManual
        .iDespImpICMSBaseManual = objTribItemNFAux.iDespImpICMSBaseManual
        .iDespImpICMSValorManual = objTribItemNFAux.iDespImpICMSValorManual
        .iDespImpICMSCreditoManual = objTribItemNFAux.iDespImpICMSCreditoManual
        .iICMSAliquotaAdicaoDIManual = objTribItemNFAux.iICMSAliquotaAdicaoDIManual
        .iICMSPercRedBaseAdicaoDIManual = objTribItemNFAux.iICMSPercRedBaseAdicaoDIManual
        .iICMSSubstPercRedBaseManual = objTribItemNFAux.iICMSSubstPercRedBaseManual
        .iICMSSubstPercMVAManual = objTribItemNFAux.iICMSSubstPercMVAManual
        .iPISTipoManual = objTribItemNFAux.iPISTipoManual
        .iPISTipoCalculoManual = objTribItemNFAux.iPISTipoCalculoManual
        .iPISBaseManual = objTribItemNFAux.iPISBaseManual
        .iPISAliquotaManual = objTribItemNFAux.iPISAliquotaManual
        .iPISAliquotaValorManual = objTribItemNFAux.iPISAliquotaValorManual
        .iPISQtdeManual = objTribItemNFAux.iPISQtdeManual
        .iPISValorManual = objTribItemNFAux.iPISValorManual
        .iPISSTTipoCalculoManual = objTribItemNFAux.iPISSTTipoCalculoManual
        .iPISSTBaseManual = objTribItemNFAux.iPISSTBaseManual
        .iPISSTAliquotaManual = objTribItemNFAux.iPISSTAliquotaManual
        .iPISSTAliquotaValorManual = objTribItemNFAux.iPISSTAliquotaValorManual
        .iPISSTQtdeManual = objTribItemNFAux.iPISSTQtdeManual
        .iPISSTValorManual = objTribItemNFAux.iPISSTValorManual
        .iCOFINSTipoManual = objTribItemNFAux.iCOFINSTipoManual
        .iCOFINSTipoCalculoManual = objTribItemNFAux.iCOFINSTipoCalculoManual
        .iCOFINSBaseManual = objTribItemNFAux.iCOFINSBaseManual
        .iCOFINSAliquotaManual = objTribItemNFAux.iCOFINSAliquotaManual
        .iCOFINSAliquotaValorManual = objTribItemNFAux.iCOFINSAliquotaValorManual
        .iCOFINSQtdeManual = objTribItemNFAux.iCOFINSQtdeManual
        .iCOFINSValorManual = objTribItemNFAux.iCOFINSValorManual
        .iCOFINSSTTipoCalculoManual = objTribItemNFAux.iCOFINSSTTipoCalculoManual
        .iCOFINSSTBaseManual = objTribItemNFAux.iCOFINSSTBaseManual
        .iCOFINSSTAliquotaManual = objTribItemNFAux.iCOFINSSTAliquotaManual
        .iCOFINSSTAliquotaValorManual = objTribItemNFAux.iCOFINSSTAliquotaValorManual
        .iCOFINSSTQtdeManual = objTribItemNFAux.iCOFINSSTQtdeManual
        .iCOFINSSTValorManual = objTribItemNFAux.iCOFINSSTValorManual
        .iCSTManual = objTribItemNFAux.iCSTManual
        .iISSQNManual = objTribItemNFAux.iISSQNManual
        .iISSBaseManual = objTribItemNFAux.iISSBaseManual
        .iISSAliquotaManual = objTribItemNFAux.iISSAliquotaManual
        .iISSCidadeIBGEManual = objTribItemNFAux.iISSCidadeIBGEManual
        .iISSValorManual = objTribItemNFAux.iISSValorManual
        .iIIBaseManual = objTribItemNFAux.iIIBaseManual
        .iIIDespAduaneiraManual = objTribItemNFAux.iIIDespAduaneiraManual
        .iIIIOFManual = objTribItemNFAux.iIIIOFManual
        .iIIValorManual = objTribItemNFAux.iIIValorManual
        .iICMSBaseModalidadeManual = objTribItemNFAux.iICMSBaseModalidadeManual
        .iICMSSubstBaseModalidadeManual = objTribItemNFAux.iICMSSubstBaseModalidadeManual
        .iIPIEnquadramentoClasseManual = objTribItemNFAux.iIPIEnquadramentoClasseManual
        .iIPIEnquadramentoCodigoManual = objTribItemNFAux.iIPIEnquadramentoCodigoManual
        .iIPICNPJProdutorManual = objTribItemNFAux.iIPICNPJProdutorManual
        .iIPISeloCodigoManual = objTribItemNFAux.iIPISeloCodigoManual
        .iIPISeloQtdeManual = objTribItemNFAux.iIPISeloQtdeManual
        .iIPITipoCalculoManual = objTribItemNFAux.iIPITipoCalculoManual
        .iIPIUnidadePadraoQtdeManual = objTribItemNFAux.iIPIUnidadePadraoQtdeManual
        .iIPIUnidadePadraoValorManual = objTribItemNFAux.iIPIUnidadePadraoValorManual
        .iValorFreteItemManual = objTribItemNFAux.iValorFreteItemManual
        .iValorSeguroItemManual = objTribItemNFAux.iValorSeguroItemManual
        .iValorOutrasDespesasItemManual = objTribItemNFAux.iValorOutrasDespesasItemManual
        .iValorDescontoItemManual = objTribItemNFAux.iValorDescontoItemManual
        .iOrigemMercadoriaManual = objTribItemNFAux.iOrigemMercadoriaManual
        .iFCIManual = objTribItemNFAux.iFCIManual
        
        .iICMSSTCobrAntBaseManual = objTribItemNFAux.iICMSSTCobrAntBaseManual
        .iICMSSTCobrAntValorManual = objTribItemNFAux.iICMSSTCobrAntValorManual
                
        .dICMSpercBaseOperacaoPropria = objTribItemNFAux.dICMSpercBaseOperacaoPropria
        .iICMSpercBaseOperacaoPropriaManual = objTribItemNFAux.iICMSpercBaseOperacaoPropriaManual
        .sICMSUFDevidoST = objTribItemNFAux.sICMSUFDevidoST
        .iICMSUFDevidoSTManual = objTribItemNFAux.iICMSUFDevidoSTManual
        .dICMSvBCSTRet = objTribItemNFAux.dICMSvBCSTRet
        .iICMSvBCSTRetManual = objTribItemNFAux.iICMSvBCSTRetManual
        .dICMSvICMSSTRet = objTribItemNFAux.dICMSvICMSSTRet
        .iICMSvICMSSTRetManual = objTribItemNFAux.iICMSvICMSSTRetManual
        .dICMSvBCSTDest = objTribItemNFAux.dICMSvBCSTDest
        .iICMSvBCSTDestManual = objTribItemNFAux.iICMSvBCSTDestManual
        .dICMSvICMSSTDest = objTribItemNFAux.dICMSvICMSSTDest
        .iICMSvICMSSTDestManual = objTribItemNFAux.iICMSvICMSSTDestManual
        .dICMSpCredSN = objTribItemNFAux.dICMSpCredSN
        .iICMSpCredSNManual = objTribItemNFAux.iICMSpCredSNManual
        .dICMSvCredSN = objTribItemNFAux.dICMSvCredSN
        .iICMSvCredSNManual = objTribItemNFAux.iICMSvCredSNManual
        .dICMSValorIsento = objTribItemNFAux.dICMSValorIsento
        .iICMSValorIsentoManual = objTribItemNFAux.iICMSValorIsentoManual
        .iICMSMotivo = objTribItemNFAux.iICMSMotivo
        .iICMSMotivoManual = objTribItemNFAux.iICMSMotivoManual
        .iICMSSimplesTipo = objTribItemNFAux.iICMSSimplesTipo
        .iICMSSimplesTipoManual = objTribItemNFAux.iICMSSimplesTipoManual
        .iISSTipo = objTribItemNFAux.iISSTipo
        .iISSTipoManual = objTribItemNFAux.iISSTipoManual
        .iRegimeTributario = objTribItemNFAux.iRegimeTributario
        .iRegimeTributarioManual = objTribItemNFAux.iRegimeTributarioManual
        .sISSCST = objTribItemNFAux.sISSCST
        
        .sNatBCCred = objTribItemNFAux.sNatBCCred
        .iNatBCCredManual = objTribItemNFAux.iNatBCCredManual
        .dTotTrib = objTribItemNFAux.dTotTrib
        .iTotTribManual = objTribItemNFAux.iTotTribManual
        .iTotTribTipo = objTribItemNFAux.iTotTribTipo
        .iTotTribTipoManual = objTribItemNFAux.iTotTribTipoManual
        
        'nfe 3.10
        .dICMS51ValorOp = objTribItemNFAux.dICMS51ValorOp
        .dICMSPercDifer = objTribItemNFAux.dICMSPercDifer
        .dICMSValorDif = objTribItemNFAux.dICMSValorDif
        .iICMS51ValorOpManual = objTribItemNFAux.iICMS51ValorOpManual
        .iICMSPercDiferManual = objTribItemNFAux.iICMSPercDiferManual
        .iICMSValorDifManual = objTribItemNFAux.iICMSValorDifManual
        .dISSValorDeducao = objTribItemNFAux.dISSValorDeducao
        .dISSValorOutrasRet = objTribItemNFAux.dISSValorOutrasRet
        .dISSValorDescIncond = objTribItemNFAux.dISSValorDescIncond
        .dISSValorDescCond = objTribItemNFAux.dISSValorDescCond
        .dISSValorRet = objTribItemNFAux.dISSValorRet
        .iISSIndExigibilidade = objTribItemNFAux.iISSIndExigibilidade
        .sISSCodServ = objTribItemNFAux.sISSCodServ
        .sISSMunicIncidImp = objTribItemNFAux.sISSMunicIncidImp
        .iISSCodPais = objTribItemNFAux.iISSCodPais
        .sISSNumProcesso = objTribItemNFAux.sISSNumProcesso
        .iISSIndIncentivo = objTribItemNFAux.iISSIndIncentivo
        .iISSValorDeducaoManual = objTribItemNFAux.iISSValorDeducaoManual
        .iISSValorOutrasRetManual = objTribItemNFAux.iISSValorOutrasRetManual
        .iISSValorDescIncondManual = objTribItemNFAux.iISSValorDescIncondManual
        .iISSValorDescCondManual = objTribItemNFAux.iISSValorDescCondManual
        .iISSValorRetManual = objTribItemNFAux.iISSValorRetManual
        .iISSIndExigibilidadeManual = objTribItemNFAux.iISSIndExigibilidadeManual
        .iISSCodServManual = objTribItemNFAux.iISSCodServManual
        .iISSMunicIncidImpManual = objTribItemNFAux.iISSMunicIncidImpManual
        .iISSCodPaisManual = objTribItemNFAux.iISSCodPaisManual
        .iISSNumProcessoManual = objTribItemNFAux.iISSNumProcessoManual
        .iISSIndIncentivoManual = objTribItemNFAux.iISSIndIncentivoManual
    
        .dTotTribAliqFed = objTribItemNFAux.dTotTribAliqFed
        .iTotTribAliqFedManual = objTribItemNFAux.iTotTribAliqFedManual
        .dTotTribAliqEst = objTribItemNFAux.dTotTribAliqEst
        .iTotTribAliqEstManual = objTribItemNFAux.iTotTribAliqEstManual
        .dTotTribAliqMunic = objTribItemNFAux.dTotTribAliqMunic
        .iTotTribAliqMunicManual = objTribItemNFAux.iTotTribAliqMunicManual
        
        ''NT 2015.003 (ICMS Interestadual)
        .sCEST = objTribItemNFAux.sCEST
        .dICMSInterestBCUFDest = objTribItemNFAux.dICMSInterestBCUFDest
        .iICMSInterestBCUFDestManual = objTribItemNFAux.iICMSInterestBCUFDestManual
        .dICMSInterestPercFCPUFDest = objTribItemNFAux.dICMSInterestPercFCPUFDest
        .iICMSInterestPercFCPUFDestManual = objTribItemNFAux.iICMSInterestPercFCPUFDestManual
        .dICMSInterestAliqUFDest = objTribItemNFAux.dICMSInterestAliqUFDest
        .iICMSInterestAliqUFDestManual = objTribItemNFAux.iICMSInterestAliqUFDestManual
        .dICMSInterestAliq = objTribItemNFAux.dICMSInterestAliq
        .iICMSInterestAliqManual = objTribItemNFAux.iICMSInterestAliqManual
        .dICMSInterestPercPartilha = objTribItemNFAux.dICMSInterestPercPartilha
        .iICMSInterestPercPartilhaManual = objTribItemNFAux.iICMSInterestPercPartilhaManual
        .dICMSInterestVlrFCPUFDest = objTribItemNFAux.dICMSInterestVlrFCPUFDest
        .iICMSInterestVlrFCPUFDestManual = objTribItemNFAux.iICMSInterestVlrFCPUFDestManual
        .dICMSInterestVlrUFDest = objTribItemNFAux.dICMSInterestVlrUFDest
        .iICMSInterestVlrUFDestManual = objTribItemNFAux.iICMSInterestVlrUFDestManual
        .dICMSInterestVlrUFRemet = objTribItemNFAux.dICMSInterestVlrUFRemet
        .iICMSInterestVlrUFRemetManual = objTribItemNFAux.iICMSInterestVlrUFRemetManual
        
        'UA. Tributos Devolvidos (para o item da NF-e)
        .dpDevol = objTribItemNFAux.dpDevol
        .ipDevolManual = objTribItemNFAux.ipDevolManual
        .dIPIVlrDevolvido = objTribItemNFAux.dIPIVlrDevolvido
        .iIPIVlrDevolvidoManual = objTribItemNFAux.iIPIVlrDevolvidoManual
        
        'NFe 4.00
        .dICMSInterestBCFCPUFDest = objTribItemNFAux.dICMSInterestBCFCPUFDest
        .iICMSInterestBCFCPUFDestManual = objTribItemNFAux.iICMSInterestBCFCPUFDestManual
        .dICMSvBCFCPST = objTribItemNFAux.dICMSvBCFCPST
        .iICMSvBCFCPSTManual = objTribItemNFAux.iICMSvBCFCPSTManual
        .dICMSpFCPST = objTribItemNFAux.dICMSpFCPST
        .iICMSpFCPSTManual = objTribItemNFAux.iICMSpFCPSTManual
        .dICMSvFCPST = objTribItemNFAux.dICMSvFCPST
        .iICMSvFCPSTManual = objTribItemNFAux.iICMSvFCPSTManual
        .dICMSvBCFCPSTRet = objTribItemNFAux.dICMSvBCFCPSTRet
        .iICMSvBCFCPSTRetManual = objTribItemNFAux.iICMSvBCFCPSTRetManual
        .dICMSpFCPSTRet = objTribItemNFAux.dICMSpFCPSTRet
        .iICMSpFCPSTRetManual = objTribItemNFAux.iICMSpFCPSTRetManual
        .dICMSvFCPSTRet = objTribItemNFAux.dICMSvFCPSTRet
        .iICMSvFCPSTRetManual = objTribItemNFAux.iICMSvFCPSTRetManual
        .dICMSSTCobrAntAliquota = objTribItemNFAux.dICMSSTCobrAntAliquota
        .iICMSSTCobrAntAliquotaManual = objTribItemNFAux.iICMSSTCobrAntAliquotaManual
        .dICMSvBCFCP = objTribItemNFAux.dICMSvBCFCP
        .iICMSvBCFCPManual = objTribItemNFAux.iICMSvBCFCPManual
        .dICMSpFCP = objTribItemNFAux.dICMSpFCP
        .iICMSpFCPManual = objTribItemNFAux.iICMSpFCPManual
        .dICMSvFCP = objTribItemNFAux.dICMSvFCP
        .iICMSvFCPManual = objTribItemNFAux.iICMSvFCPManual
        .scBenef = objTribItemNFAux.scBenef
        .sindEscala = objTribItemNFAux.sindEscala
        .sCNPJFab = objTribItemNFAux.sCNPJFab
        
        .iICMSSTBaseDupla = objTribItemNFAux.iICMSSTBaseDupla
        .iICMSSTBaseDuplaManual = objTribItemNFAux.iICMSSTBaseDuplaManual
        .iICMSInterestBaseDupla = objTribItemNFAux.iICMSInterestBaseDupla
        .iICMSInterestBaseDuplaManual = objTribItemNFAux.iICMSInterestBaseDuplaManual
        
        .icBenefManual = objTribItemNFAux.icBenefManual
    
    End With
    
    PreencherNFItem = SUCESSO

    Exit Function

Erro_PreencherNFItem:

    PreencherNFItem = Err

    Select Case Err

        Case 22853
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153759)

    End Select

    Exit Function

End Function

Private Function ExtraiNF(objNFiscal As ClassNFiscal) As Long
'carrega dados de tributacao de objNFiscal para "Me"(a propria instancia da classe ClassTribTab)

Dim lErro As Long, objTribItemNFAux As ClassTribItemNFAux

On Error GoTo Erro_ExtraiNF

    'obter o tipo da nota
    
    With objNFiscal.objTributacaoNF
    
        '.sNaturezaOpInterna = objNFiscal.sNaturezaOp
    
        mvarobjNFTributacao.dICMSBase = .dICMSBase
        mvarobjNFTributacao.dICMSCredito = .dICMSCredito
        mvarobjNFTributacao.dICMSSubstBase = .dICMSSubstBase
        mvarobjNFTributacao.dICMSSubstValor = .dICMSSubstValor
        mvarobjNFTributacao.dICMSValor = .dICMSValor
        mvarobjNFTributacao.dIPIBase = .dIPIBase
        mvarobjNFTributacao.dIPICredito = .dIPICredito
        mvarobjNFTributacao.dIPIValor = .dIPIValor
        mvarobjNFTributacao.dIRRFAliquota = .dIRRFAliquota
        mvarobjNFTributacao.dIRRFBase = .dIRRFBase
        mvarobjNFTributacao.dIRRFValor = .dIRRFValor
        mvarobjNFTributacao.dISSAliquota = .dISSAliquota
        mvarobjNFTributacao.dISSBase = .dISSBase
        mvarobjNFTributacao.dISSValor = .dISSValor
        mvarobjNFTributacao.iISSIncluso = .iISSIncluso
        mvarobjNFTributacao.iTipoTributacao = .iTipoTributacao
        mvarobjNFTributacao.sNaturezaOp = .sNaturezaOpInterna
        mvarobjNFTributacao.dValorDesconto = objNFiscal.dValorDesconto
        mvarobjNFTributacao.dValorFrete = objNFiscal.dValorFrete
        mvarobjNFTributacao.dValorOutrasDespesas = objNFiscal.dValorOutrasDespesas
        mvarobjNFTributacao.dValorSeguro = objNFiscal.dValorSeguro
        mvarobjNFTributacao.iFilialCli = objNFiscal.iFilialCli
        mvarobjNFTributacao.iFilialEmpresa = objNFiscal.iFilialEmpresa
        mvarobjNFTributacao.iFilialEntrega = objNFiscal.iFilialEntrega
        mvarobjNFTributacao.iFilialForn = objNFiscal.iFilialForn
        mvarobjNFTributacao.dTotTrib = .dTotTrib
        mvarobjNFTributacao.iTotTribTipo = .iTotTribTipo
        
        mvarobjNFTributacao.iICMSBaseManual = IIf(.iICMSBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iICMSBaseManual)
        mvarobjNFTributacao.iICMSSubstBaseManual = IIf(.iICMSSubstBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iICMSSubstBaseManual)
        mvarobjNFTributacao.iICMSSubstValorManual = IIf(.iICMSSubstValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iICMSSubstValorManual)
        mvarobjNFTributacao.iICMSValorManual = IIf(.iICMSValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iICMSValorManual)
        mvarobjNFTributacao.iIPIBaseManual = IIf(.iIPIBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iIPIBaseManual)
        mvarobjNFTributacao.iIPIValorManual = IIf(.iIPIValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iIPIValorManual)
        
        mvarobjNFTributacao.iIRRFAliquotaManual = IIf(.iIRRFAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iIRRFAliquotaManual)
        mvarobjNFTributacao.iIRRFValorManual = IIf(.iIRRFValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iIRRFValorManual)
        mvarobjNFTributacao.iISSAliquotaManual = IIf(.iISSAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iISSAliquotaManual)
        mvarobjNFTributacao.iISSValorManual = IIf(.iISSValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iISSValorManual)
        'mvarobjNFTributacao.iISSInclusoManual = VAR_PREENCH_MANUAL
        mvarobjNFTributacao.iTipoTributacaoManual = IIf(.iTipoTributacaoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iTipoTributacaoManual)
        mvarobjNFTributacao.iNaturezaOpManual = IIf(.iNaturezaOpManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iNaturezaOpManual)
        
        mvarobjNFTributacao.lCliente = objNFiscal.lCliente
        mvarobjNFTributacao.lFornecedor = objNFiscal.lFornecedor
        mvarobjNFTributacao.sNatOpExterna = objNFiscal.sNaturezaOp
        
        mvarobjNFTributacao.dINSSBase = .dINSSValorBase
        mvarobjNFTributacao.iINSSBaseManual = IIf(.iINSSValorBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iINSSValorBaseManual)
        mvarobjNFTributacao.dINSSDeducoes = .dINSSValorDeducoes
        mvarobjNFTributacao.iINSSDeducoesManual = IIf(.iINSSValorDeducoesManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iINSSValorDeducoesManual)
        mvarobjNFTributacao.dINSSValor = .dValorINSS
        mvarobjNFTributacao.iINSSValorManual = IIf(.iValorINSSManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iValorINSSManual)
        mvarobjNFTributacao.iINSSRetido = .iINSSRetido
        mvarobjNFTributacao.iINSSRetidoManual = IIf(.iINSSRetidoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iINSSRetidoManual)
        mvarobjNFTributacao.dPISCredito = .dPISCredito
        'mvarobjNFTributacao.iPISCreditoManual = IIf(.iPISCreditoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iPISCreditoManual)
        mvarobjNFTributacao.iPISCreditoManual = VAR_PREENCH_AUTOMATICO 'Campo que não pode ser alterado em tela tem que ser automático
        mvarobjNFTributacao.dCOFINSCredito = .dCOFINSCredito
        'mvarobjNFTributacao.iCOFINSCreditoManual = IIf(.iCOFINSCreditoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iCOFINSCreditoManual)
        mvarobjNFTributacao.iCOFINSCreditoManual = VAR_PREENCH_AUTOMATICO
        
        mvarobjNFTributacao.dPISRetido = .dPISRetido
        mvarobjNFTributacao.dISSRetido = .dISSRetido
        mvarobjNFTributacao.dCOFINSRetido = .dCOFINSRetido
        mvarobjNFTributacao.dCSLLRetido = .dCSLLRetido
        mvarobjNFTributacao.iPISRetidoManual = IIf(.iPISRetidoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iPISRetidoManual)
        mvarobjNFTributacao.iISSRetidoManual = IIf(.iISSRetidoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iISSRetidoManual)
        mvarobjNFTributacao.iCOFINSRetidoManual = IIf(.iCOFINSRetidoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iCOFINSRetidoManual)
        mvarobjNFTributacao.iCSLLRetidoManual = IIf(.iCSLLRetidoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iCSLLRetidoManual)
        
        mvarobjNFTributacao.dtDataEmissao = objNFiscal.dtDataEmissao
        mvarobjNFTributacao.dtDataEntrada = objNFiscal.dtDataEntrada
        mvarobjNFTributacao.sSerie = objNFiscal.sSerie
        
        mvarobjNFTributacao.dTotTrib = .dTotTrib
        mvarobjNFTributacao.iTotTribManual = IIf(.iTotTribManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iTotTribManual)
        mvarobjNFTributacao.iTotTribTipo = IIf(.iTotTribTipo = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iTotTribTipo)
        mvarobjNFTributacao.iTotTribTipoManual = IIf(.iTotTribTipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iTotTribTipoManual)
    
        mvarobjNFTributacao.dPISValor = .dPISValor
        mvarobjNFTributacao.dCOFINSValor = .dCOFINSValor
        
        'nfe 3.10
        mvarobjNFTributacao.dICMSValorIsento = .dICMSValorIsento
        mvarobjNFTributacao.dISSValorDeducao = .dISSValorDeducao
        mvarobjNFTributacao.dISSValorOutrasRet = .dISSValorOutrasRet
        mvarobjNFTributacao.dISSValorDescIncond = .dISSValorDescIncond
        mvarobjNFTributacao.dISSValorDescCond = .dISSValorDescCond
        mvarobjNFTributacao.iDestino = .iDestino
        mvarobjNFTributacao.iDestinoManual = IIf(.iDestinoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iDestinoManual)
        mvarobjNFTributacao.iFinalidadeNFe = .iFinalidadeNFe
        mvarobjNFTributacao.iFinalidadeNFeManual = IIf(.iFinalidadeNFeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iFinalidadeNFeManual)
        mvarobjNFTributacao.iIndConsumidorFinal = .iIndConsumidorFinal
        mvarobjNFTributacao.iIndConsumidorFinalManual = IIf(.iIndConsumidorFinalManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iIndConsumidorFinalManual)
        mvarobjNFTributacao.iIndPresenca = .iIndPresenca
        mvarobjNFTributacao.iIndPresencaManual = IIf(.iIndPresencaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iIndPresencaManual)
        mvarobjNFTributacao.iRegTribEspecial = .iRegTribEspecial
        mvarobjNFTributacao.iRegTribEspecialManual = IIf(.iRegTribEspecialManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iRegTribEspecialManual)
        mvarobjNFTributacao.dtDataPrestServico = .dtDataPrestServico
        mvarobjNFTributacao.iDataPrestServicoManual = IIf(.iDataPrestServicoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iDataPrestServicoManual)
    
        mvarobjNFTributacao.dTotTribFed = .dTotTribFed
        mvarobjNFTributacao.iTotTribFedManual = IIf(.iTotTribFedManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iTotTribFedManual)
        mvarobjNFTributacao.dTotTribEst = .dTotTribEst
        mvarobjNFTributacao.iTotTribEstManual = IIf(.iTotTribEstManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iTotTribEstManual)
        mvarobjNFTributacao.dTotTribMunic = .dTotTribMunic
        mvarobjNFTributacao.iTotTribMunicManual = IIf(.iTotTribMunicManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, .iTotTribMunicManual)
    
        mvarobjNFTributacao.dICMSInterestVlrFCPUFDest = .dICMSInterestVlrFCPUFDest
        mvarobjNFTributacao.dICMSInterestVlrUFDest = .dICMSInterestVlrUFDest
        mvarobjNFTributacao.dICMSInterestVlrUFRemet = .dICMSInterestVlrUFRemet
        
        'NFe 4.00
        mvarobjNFTributacao.dICMSVlrFCP = .dICMSVlrFCP
        mvarobjNFTributacao.dICMSVlrFCPST = .dICMSVlrFCPST
        mvarobjNFTributacao.dICMSVlrFCPSTRet = .dICMSVlrFCPSTRet
        mvarobjNFTributacao.dIPIVlrDevolvido = .dIPIVlrDevolvido
    
    End With

    ExtraiNF = SUCESSO

    Exit Function

Erro_ExtraiNF:

    ExtraiNF = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153760)

    End Select

    Exit Function

End Function

Private Function ExtraiNFItem(objNFiscal As ClassNFiscal, objItemNF As ClassItemNF, Optional ByVal bSRV As Boolean = False) As Long

Dim lErro As Long, objTribItemNFAux As New ClassTribItemNFAux, objTributoDocItem As ClassTributoDocItem

On Error GoTo Erro_ExtraiNFItem

    With objItemNF.objTributacaoItemNF
    
        objTribItemNFAux.dICMSAliquota = .dICMSAliquota
        objTribItemNFAux.dICMSBase = .dICMSBase
        objTribItemNFAux.dICMSCredito = .dICMSCredito
        objTribItemNFAux.dICMSPercRedBase = .dICMSPercRedBase
        objTribItemNFAux.dICMSSubstAliquota = .dICMSSubstAliquota
        objTribItemNFAux.dICMSSubstBase = .dICMSSubstBase
        objTribItemNFAux.dICMSSubstValor = .dICMSSubstValor
        objTribItemNFAux.dICMSValor = .dICMSValor
        objTribItemNFAux.dIPIAliquota = .dIPIAliquota
        objTribItemNFAux.dIPIBaseCalculo = .dIPIBaseCalculo
        objTribItemNFAux.dIPICredito = .dIPICredito
        objTribItemNFAux.dIPIPercRedBase = .dIPIPercRedBase
        objTribItemNFAux.dIPIValor = .dIPIValor
        objTribItemNFAux.iICMSTipo = .iICMSTipo
        objTribItemNFAux.iIPITipo = .iIPITipo
        objTribItemNFAux.iTipoTributacao = .iTipoTributacao
        objTribItemNFAux.sNaturezaOp = .sNaturezaOp
        objTribItemNFAux.dPISCredito = .dPISCredito
        objTribItemNFAux.dCOFINSCredito = .dCOFINSCredito
        objTribItemNFAux.dICMSSubstPercRedBase = .dICMSSubstPercRedBase
        objTribItemNFAux.dICMSSubstPercMVA = .dICMSSubstPercMVA
        objTribItemNFAux.sIPICodProduto = .sIPICodProduto
        
        objTribItemNFAux.dICMSSTCobrAntValor = .dICMSSTCobrAntValor
        objTribItemNFAux.dICMSSTCobrAntBase = .dICMSSTCobrAntBase
        
        'nf de importacao
        objTribItemNFAux.dDespImpICMSBase = .dDespImpICMSBase
        objTribItemNFAux.dDespImpICMSValor = .dDespImpICMSValor
        objTribItemNFAux.dDespImpICMSCredito = .dDespImpICMSCredito
        objTribItemNFAux.dICMSAliquotaAdicaoDI = .dICMSAliquotaAdicaoDI
        objTribItemNFAux.dICMSPercRedBaseAdicaoDI = .dICMSPercRedBaseAdicaoDI
        
        'XXXX
        'INÍCIO NOVOS CAMPOS
        objTribItemNFAux.dICMSSubstPercRedBase = .dICMSSubstPercRedBase
        objTribItemNFAux.dICMSSubstPercMVA = .dICMSSubstPercMVA
        objTribItemNFAux.iPISTipo = .iPISTipo
        objTribItemNFAux.iPISTipoCalculo = .iPISTipoCalculo
        objTribItemNFAux.dPISBase = .dPISBase
        objTribItemNFAux.dPISAliquota = .dPISAliquota
        objTribItemNFAux.dPISAliquotaValor = .dPISAliquotaValor
        objTribItemNFAux.dPISQtde = .dPISQtde
        objTribItemNFAux.dPISValor = .dPISValor
        objTribItemNFAux.iPISSTTipoCalculo = .iPISSTTipoCalculo
        objTribItemNFAux.dPISSTBase = .dPISSTBase
        objTribItemNFAux.dPISSTAliquota = .dPISSTAliquota
        objTribItemNFAux.dPISSTAliquotaValor = .dPISSTAliquotaValor
        objTribItemNFAux.dPISSTQtde = .dPISSTQtde
        objTribItemNFAux.dPISSTValor = .dPISSTValor
        objTribItemNFAux.iCOFINSTipo = .iCOFINSTipo
        objTribItemNFAux.iCOFINSTipoCalculo = .iCOFINSTipoCalculo
        objTribItemNFAux.dCOFINSBase = .dCOFINSBase
        objTribItemNFAux.dCOFINSAliquota = .dCOFINSAliquota
        objTribItemNFAux.dCOFINSAliquotaValor = .dCOFINSAliquotaValor
        objTribItemNFAux.dCOFINSQtde = .dCOFINSQtde
        objTribItemNFAux.dCOFINSValor = .dCOFINSValor
        objTribItemNFAux.iCOFINSSTTipoCalculo = .iCOFINSSTTipoCalculo
        objTribItemNFAux.dCOFINSSTBase = .dCOFINSSTBase
        objTribItemNFAux.dCOFINSSTAliquota = .dCOFINSSTAliquota
        objTribItemNFAux.dCOFINSSTAliquotaValor = .dCOFINSSTAliquotaValor
        objTribItemNFAux.dCOFINSSTQtde = .dCOFINSSTQtde
        objTribItemNFAux.dCOFINSSTValor = .dCOFINSSTValor
        objTribItemNFAux.sCST = .sCST
        objTribItemNFAux.sCSOSN = .sCSOSN
        objTribItemNFAux.sISSQN = .sISSQN
        objTribItemNFAux.dISSBase = .dISSBase
        objTribItemNFAux.dISSAliquota = .dISSAliquota
        objTribItemNFAux.sISSCidadeIBGE = .sISSCidadeIBGE
        objTribItemNFAux.dISSValor = .dISSValor
        objTribItemNFAux.dIIBase = .dIIBase
        objTribItemNFAux.dIIDespAduaneira = .dIIDespAduaneira
        objTribItemNFAux.dIIIOF = .dIIIOF
        objTribItemNFAux.dIIValor = .dIIValor
        objTribItemNFAux.iICMSBaseModalidade = .iICMSBaseModalidade
        objTribItemNFAux.iICMSSubstBaseModalidade = .iICMSSubstBaseModalidade
        objTribItemNFAux.sIPIEnquadramentoClasse = .sIPIEnquadramentoClasse
        objTribItemNFAux.sIPIEnquadramentoCodigo = .sIPIEnquadramentoCodigo
        objTribItemNFAux.sIPICNPJProdutor = .sIPICNPJProdutor
        objTribItemNFAux.sIPISeloCodigo = .sIPISeloCodigo
        objTribItemNFAux.lIPISeloQtde = .lIPISeloQtde
        objTribItemNFAux.iIPITipoCalculo = .iIPITipoCalculo
        objTribItemNFAux.dIPIUnidadePadraoQtde = .dIPIUnidadePadraoQtde
        objTribItemNFAux.dIPIUnidadePadraoValor = .dIPIUnidadePadraoValor
        objTribItemNFAux.dValorFreteItem = .dValorFreteItem
        objTribItemNFAux.dValorSeguroItem = .dValorSeguroItem
        objTribItemNFAux.dValorOutrasDespesasItem = .dValorOutrasDespesasItem
        objTribItemNFAux.dValorDescontoItem = .dValorDescontoItem
        objTribItemNFAux.iOrigemMercadoria = .iOrigemMercadoria
        objTribItemNFAux.sFCI = .sFCI
        
        objTribItemNFAux.iExTIPI = .iExTIPI
        objTribItemNFAux.sGenero = .sGenero
        objTribItemNFAux.iProdutoEspecifico = .iProdutoEspecifico
        objTribItemNFAux.sEAN = .sEAN
        objTribItemNFAux.sEANTrib = .sEANTrib
        objTribItemNFAux.dQtdTrib = .dQtdTrib
        objTribItemNFAux.sUMTrib = .sUMTrib
        objTribItemNFAux.dValorUnitTrib = .dValorUnitTrib

        objTribItemNFAux.iICMSAliquotaManual = .iICMSAliquotaManual
        objTribItemNFAux.iICMSBaseManual = .iICMSBaseManual
        objTribItemNFAux.iICMSCreditoManual = .iICMSCreditoManual
        objTribItemNFAux.iICMSPercRedBaseManual = .iICMSPercRedBaseManual
        objTribItemNFAux.iICMSSubstAliquotaManual = .iICMSSubstAliquotaManual
        objTribItemNFAux.iICMSSubstBaseManual = .iICMSSubstBaseManual
        objTribItemNFAux.iICMSSubstValorManual = .iICMSSubstValorManual
        objTribItemNFAux.iICMSValorManual = .iICMSValorManual
        objTribItemNFAux.iIPIAliquotaManual = .iIPIAliquotaManual
        objTribItemNFAux.iIPIBaseManual = .iIPIBaseCalculoManual
        objTribItemNFAux.iIPICreditoManual = .iIPICreditoManual
        objTribItemNFAux.iIPIPercRedBaseManual = .iIPIPercRedBaseManual
        objTribItemNFAux.iIPIValorManual = .iIPIValorManual
        objTribItemNFAux.iICMSTipoManual = .iICMSTipoManual
        objTribItemNFAux.iIPITipoManual = .iIPITipoManual
        objTribItemNFAux.iTipoTributacaoManual = .iTipoTributacaoManual
        objTribItemNFAux.iNaturezaOpManual = .iNaturezaOpManual
        objTribItemNFAux.iPISCreditoManual = .iPISCreditoManual
        objTribItemNFAux.iCOFINSCreditoManual = .iCOFINSCreditoManual
        objTribItemNFAux.iICMSSubstPercRedBaseManual = .iICMSSubstPercRedBaseManual
        objTribItemNFAux.iICMSSubstPercMVAManual = .iICMSSubstPercMVAManual
        objTribItemNFAux.iDespImpICMSBaseManual = .iDespImpICMSBaseManual
        objTribItemNFAux.iDespImpICMSValorManual = .iDespImpICMSValorManual
        objTribItemNFAux.iDespImpICMSCreditoManual = .iDespImpICMSCreditoManual
        objTribItemNFAux.iICMSAliquotaAdicaoDIManual = .iICMSAliquotaAdicaoDIManual
        objTribItemNFAux.iICMSSubstPercRedBaseManual = .iICMSSubstPercRedBaseManual
        objTribItemNFAux.iICMSSubstPercMVAManual = .iICMSSubstPercMVAManual
        objTribItemNFAux.iPISTipoManual = .iPISTipoManual
        objTribItemNFAux.iPISTipoCalculoManual = .iPISTipoCalculoManual
        objTribItemNFAux.iPISBaseManual = .iPISBaseManual
        objTribItemNFAux.iPISAliquotaManual = .iPISAliquotaManual
        objTribItemNFAux.iPISAliquotaValorManual = .iPISAliquotaValorManual
        objTribItemNFAux.iPISQtdeManual = .iPISQtdeManual
        objTribItemNFAux.iPISValorManual = .iPISValorManual
        objTribItemNFAux.iPISSTTipoCalculoManual = .iPISSTTipoCalculoManual
        objTribItemNFAux.iPISSTBaseManual = .iPISSTBaseManual
        objTribItemNFAux.iPISSTAliquotaManual = .iPISSTAliquotaManual
        objTribItemNFAux.iPISSTAliquotaValorManual = .iPISSTAliquotaValorManual
        objTribItemNFAux.iPISSTQtdeManual = .iPISSTQtdeManual
        objTribItemNFAux.iPISSTValorManual = .iPISSTValorManual
        objTribItemNFAux.iCOFINSTipoManual = .iCOFINSTipoManual
        objTribItemNFAux.iCOFINSTipoCalculoManual = .iCOFINSTipoCalculoManual
        objTribItemNFAux.iCOFINSBaseManual = .iCOFINSBaseManual
        objTribItemNFAux.iCOFINSAliquotaManual = .iCOFINSAliquotaManual
        objTribItemNFAux.iCOFINSAliquotaValorManual = .iCOFINSAliquotaValorManual
        objTribItemNFAux.iCOFINSQtdeManual = .iCOFINSQtdeManual
        objTribItemNFAux.iCOFINSValorManual = .iCOFINSValorManual
        objTribItemNFAux.iCOFINSSTTipoCalculoManual = .iCOFINSSTTipoCalculoManual
        objTribItemNFAux.iCOFINSSTBaseManual = .iCOFINSSTBaseManual
        objTribItemNFAux.iCOFINSSTAliquotaManual = .iCOFINSSTAliquotaManual
        objTribItemNFAux.iCOFINSSTAliquotaValorManual = .iCOFINSSTAliquotaValorManual
        objTribItemNFAux.iCOFINSSTQtdeManual = .iCOFINSSTQtdeManual
        objTribItemNFAux.iCOFINSSTValorManual = .iCOFINSSTValorManual
        objTribItemNFAux.iCSTManual = .iCSTManual
        objTribItemNFAux.iISSQNManual = .iISSQNManual
        objTribItemNFAux.iISSBaseManual = .iISSBaseManual
        objTribItemNFAux.iISSAliquotaManual = .iISSAliquotaManual
        objTribItemNFAux.iISSCidadeIBGEManual = .iISSCidadeIBGEManual
        objTribItemNFAux.iISSValorManual = .iISSValorManual
        objTribItemNFAux.iIIBaseManual = .iIIBaseManual
        objTribItemNFAux.iIIDespAduaneiraManual = .iIIDespAduaneiraManual
        objTribItemNFAux.iIIIOFManual = .iIIIOFManual
        objTribItemNFAux.iIIValorManual = .iIIValorManual
        objTribItemNFAux.iICMSBaseModalidadeManual = .iICMSBaseModalidadeManual
        objTribItemNFAux.iICMSSubstBaseModalidadeManual = .iICMSSubstBaseModalidadeManual
        objTribItemNFAux.iIPIEnquadramentoClasseManual = .iIPIEnquadramentoClasseManual
        objTribItemNFAux.iIPIEnquadramentoCodigoManual = .iIPIEnquadramentoCodigoManual
        objTribItemNFAux.iIPICNPJProdutorManual = .iIPICNPJProdutorManual
        objTribItemNFAux.iIPISeloCodigoManual = .iIPISeloCodigoManual
        objTribItemNFAux.iIPISeloQtdeManual = .iIPISeloQtdeManual
        objTribItemNFAux.iIPITipoCalculoManual = .iIPITipoCalculoManual
        objTribItemNFAux.iIPIUnidadePadraoQtdeManual = .iIPIUnidadePadraoQtdeManual
        objTribItemNFAux.iIPIUnidadePadraoValorManual = .iIPIUnidadePadraoValorManual
        objTribItemNFAux.iValorFreteItemManual = .iValorFreteItemManual
        objTribItemNFAux.iValorSeguroItemManual = .iValorSeguroItemManual
        objTribItemNFAux.iValorOutrasDespesasItemManual = .iValorOutrasDespesasItemManual
        objTribItemNFAux.iValorDescontoItemManual = .iValorDescontoItemManual
        objTribItemNFAux.iOrigemMercadoriaManual = .iOrigemMercadoriaManual
        objTribItemNFAux.iFCIManual = .iFCIManual
    
        objTribItemNFAux.iICMSSTCobrAntBaseManual = .iICMSSTCobrAntBaseManual
        objTribItemNFAux.iICMSSTCobrAntValorManual = .iICMSSTCobrAntValorManual
        
        objTribItemNFAux.dICMSpercBaseOperacaoPropria = .dICMSpercBaseOperacaoPropria
        objTribItemNFAux.iICMSpercBaseOperacaoPropriaManual = .iICMSpercBaseOperacaoPropriaManual
        objTribItemNFAux.sICMSUFDevidoST = .sICMSUFDevidoST
        objTribItemNFAux.iICMSUFDevidoSTManual = .iICMSUFDevidoSTManual
        objTribItemNFAux.dICMSvBCSTRet = .dICMSvBCSTRet
        objTribItemNFAux.iICMSvBCSTRetManual = .iICMSvBCSTRetManual
        objTribItemNFAux.dICMSvICMSSTRet = .dICMSvICMSSTRet
        objTribItemNFAux.iICMSvICMSSTRetManual = .iICMSvICMSSTRetManual
        objTribItemNFAux.dICMSvBCSTDest = .dICMSvBCSTDest
        objTribItemNFAux.iICMSvBCSTDestManual = .iICMSvBCSTDestManual
        objTribItemNFAux.dICMSvICMSSTDest = .dICMSvICMSSTDest
        objTribItemNFAux.iICMSvICMSSTDestManual = .iICMSvICMSSTDestManual
        objTribItemNFAux.dICMSpCredSN = .dICMSpCredSN
        objTribItemNFAux.iICMSpCredSNManual = .iICMSpCredSNManual
        objTribItemNFAux.dICMSvCredSN = .dICMSvCredSN
        objTribItemNFAux.iICMSvCredSNManual = .iICMSvCredSNManual
        objTribItemNFAux.dICMSValorIsento = .dICMSValorIsento
        objTribItemNFAux.iICMSValorIsentoManual = .iICMSValorIsentoManual
        objTribItemNFAux.iICMSMotivo = .iICMSMotivo
        objTribItemNFAux.iICMSMotivoManual = .iICMSMotivoManual
        objTribItemNFAux.iICMSSimplesTipo = .iICMSSimplesTipo
        objTribItemNFAux.iICMSSimplesTipoManual = .iICMSSimplesTipoManual
        objTribItemNFAux.iISSTipo = .iISSTipo
        objTribItemNFAux.iISSTipoManual = .iISSTipoManual
        objTribItemNFAux.iRegimeTributario = .iRegimeTributario
        objTribItemNFAux.iRegimeTributarioManual = .iRegimeTributarioManual
        objTribItemNFAux.sISSCST = .sISSCST
        
        objTribItemNFAux.sNatBCCred = .sNatBCCred
        objTribItemNFAux.iNatBCCredManual = .iNatBCCredManual
        objTribItemNFAux.dTotTrib = .dTotTrib
        objTribItemNFAux.iTotTribManual = .iTotTribManual
        objTribItemNFAux.iTotTribTipo = .iTotTribTipo
        objTribItemNFAux.iTotTribTipoManual = .iTotTribTipoManual
    
        'nfe 3.10
        objTribItemNFAux.dICMS51ValorOp = .dICMS51ValorOp
        objTribItemNFAux.dICMSPercDifer = .dICMSPercDifer
        objTribItemNFAux.dICMSValorDif = .dICMSValorDif
        objTribItemNFAux.iICMS51ValorOpManual = .iICMS51ValorOpManual
        objTribItemNFAux.iICMSPercDiferManual = .iICMSPercDiferManual
        objTribItemNFAux.iICMSValorDifManual = .iICMSValorDifManual
        objTribItemNFAux.dISSValorDeducao = .dISSValorDeducao
        objTribItemNFAux.dISSValorOutrasRet = .dISSValorOutrasRet
        objTribItemNFAux.dISSValorDescIncond = .dISSValorDescIncond
        objTribItemNFAux.dISSValorDescCond = .dISSValorDescCond
        objTribItemNFAux.dISSValorRet = .dISSValorRet
        objTribItemNFAux.iISSIndExigibilidade = .iISSIndExigibilidade
        objTribItemNFAux.sISSCodServ = .sISSCodServ
        objTribItemNFAux.sISSMunicIncidImp = .sISSMunicIncidImp
        objTribItemNFAux.iISSCodPais = .iISSCodPais
        objTribItemNFAux.sISSNumProcesso = .sISSNumProcesso
        objTribItemNFAux.iISSIndIncentivo = .iISSIndIncentivo
        objTribItemNFAux.iISSValorDeducaoManual = .iISSValorDeducaoManual
        objTribItemNFAux.iISSValorOutrasRetManual = .iISSValorOutrasRetManual
        objTribItemNFAux.iISSValorDescIncondManual = .iISSValorDescIncondManual
        objTribItemNFAux.iISSValorDescCondManual = .iISSValorDescCondManual
        objTribItemNFAux.iISSValorRetManual = .iISSValorRetManual
        objTribItemNFAux.iISSIndExigibilidadeManual = .iISSIndExigibilidadeManual
        objTribItemNFAux.iISSCodServManual = .iISSCodServManual
        objTribItemNFAux.iISSMunicIncidImpManual = .iISSMunicIncidImpManual
        objTribItemNFAux.iISSCodPaisManual = .iISSCodPaisManual
        objTribItemNFAux.iISSNumProcessoManual = .iISSNumProcessoManual
        objTribItemNFAux.iISSIndIncentivoManual = .iISSIndIncentivoManual
    
        objTribItemNFAux.dTotTribAliqFed = .dTotTribAliqFed
        objTribItemNFAux.iTotTribAliqFedManual = .iTotTribAliqFedManual
        objTribItemNFAux.dTotTribAliqEst = .dTotTribAliqEst
        objTribItemNFAux.iTotTribAliqEstManual = .iTotTribAliqEstManual
        objTribItemNFAux.dTotTribAliqMunic = .dTotTribAliqMunic
        objTribItemNFAux.iTotTribAliqMunicManual = .iTotTribManual
    
        'NT 2015.003 (ICMS Interestadual)
        objTribItemNFAux.sCEST = .sCEST
        objTribItemNFAux.dICMSInterestBCUFDest = .dICMSInterestBCUFDest
        objTribItemNFAux.iICMSInterestBCUFDestManual = .iICMSInterestBCUFDestManual
        objTribItemNFAux.dICMSInterestPercFCPUFDest = .dICMSInterestPercFCPUFDest
        objTribItemNFAux.iICMSInterestPercFCPUFDestManual = .iICMSInterestPercFCPUFDestManual
        objTribItemNFAux.dICMSInterestAliqUFDest = .dICMSInterestAliqUFDest
        objTribItemNFAux.iICMSInterestAliqUFDestManual = .iICMSInterestAliqUFDestManual
        objTribItemNFAux.dICMSInterestAliq = .dICMSInterestAliq
        objTribItemNFAux.iICMSInterestAliqManual = .iICMSInterestAliqManual
        objTribItemNFAux.dICMSInterestPercPartilha = .dICMSInterestPercPartilha
        objTribItemNFAux.iICMSInterestPercPartilhaManual = .iICMSInterestPercPartilhaManual
        objTribItemNFAux.dICMSInterestVlrFCPUFDest = .dICMSInterestVlrFCPUFDest
        objTribItemNFAux.iICMSInterestVlrFCPUFDestManual = .iICMSInterestVlrFCPUFDestManual
        objTribItemNFAux.dICMSInterestVlrUFDest = .dICMSInterestVlrUFDest
        objTribItemNFAux.iICMSInterestVlrUFDestManual = .iICMSInterestVlrUFDestManual
        objTribItemNFAux.dICMSInterestVlrUFRemet = .dICMSInterestVlrUFRemet
        objTribItemNFAux.iICMSInterestVlrUFRemetManual = .iICMSInterestVlrUFRemetManual
        
        'UA. Tributos Devolvidos (para o item da NF-e)
        objTribItemNFAux.dpDevol = .dpDevol
        objTribItemNFAux.ipDevolManual = .ipDevolManual
        objTribItemNFAux.dIPIVlrDevolvido = .dIPIVlrDevolvido
        objTribItemNFAux.iIPIVlrDevolvidoManual = .iIPIVlrDevolvidoManual
        
        'NFe 4.00
        objTribItemNFAux.dICMSInterestBCFCPUFDest = .dICMSInterestBCFCPUFDest
        objTribItemNFAux.iICMSInterestBCFCPUFDestManual = .iICMSInterestBCFCPUFDestManual
        objTribItemNFAux.dICMSvBCFCPST = .dICMSvBCFCPST
        objTribItemNFAux.iICMSvBCFCPSTManual = .iICMSvBCFCPSTManual
        objTribItemNFAux.dICMSpFCPST = .dICMSpFCPST
        objTribItemNFAux.iICMSpFCPSTManual = .iICMSpFCPSTManual
        objTribItemNFAux.dICMSvFCPST = .dICMSvFCPST
        objTribItemNFAux.iICMSvFCPSTManual = .iICMSvFCPSTManual
        objTribItemNFAux.dICMSvBCFCPSTRet = .dICMSvBCFCPSTRet
        objTribItemNFAux.iICMSvBCFCPSTRetManual = .iICMSvBCFCPSTRetManual
        objTribItemNFAux.dICMSpFCPSTRet = .dICMSpFCPSTRet
        objTribItemNFAux.iICMSpFCPSTRetManual = .iICMSpFCPSTRetManual
        objTribItemNFAux.dICMSvFCPSTRet = .dICMSvFCPSTRet
        objTribItemNFAux.iICMSvFCPSTRetManual = .iICMSvFCPSTRetManual
        objTribItemNFAux.dICMSSTCobrAntAliquota = .dICMSSTCobrAntAliquota
        objTribItemNFAux.iICMSSTCobrAntAliquotaManual = .iICMSSTCobrAntAliquotaManual
        objTribItemNFAux.dICMSvBCFCP = .dICMSvBCFCP
        objTribItemNFAux.iICMSvBCFCPManual = .iICMSvBCFCPManual
        objTribItemNFAux.dICMSpFCP = .dICMSpFCP
        objTribItemNFAux.iICMSpFCPManual = .iICMSpFCPManual
        objTribItemNFAux.dICMSvFCP = .dICMSvFCP
        objTribItemNFAux.iICMSvFCPManual = .iICMSvFCPManual
        objTribItemNFAux.scBenef = .scBenef
        objTribItemNFAux.sindEscala = .sindEscala
        objTribItemNFAux.sCNPJFab = .sCNPJFab
        
        objTribItemNFAux.iICMSSTBaseDupla = .iICMSSTBaseDupla
        objTribItemNFAux.iICMSSTBaseDuplaManual = .iICMSSTBaseDuplaManual
        objTribItemNFAux.iICMSInterestBaseDupla = .iICMSInterestBaseDupla
        objTribItemNFAux.iICMSInterestBaseDuplaManual = .iICMSInterestBaseDuplaManual
        objTribItemNFAux.icBenefManual = .icBenefManual
        
    End With
            
    With objTribItemNFAux
        .dValorDesconto = objItemNF.dValorDesconto
        .dPrecoTotal = objItemNF.dValorTotal
        .dPrecoUnitario = objItemNF.dPrecoUnitario
        .dPrecoUnitarioMoeda = objItemNF.dPrecoUnitarioMoeda
        .dQuantidade = objItemNF.dQuantidade
        .iItem = objItemNF.iItem
        .sProduto = objItemNF.sProduto
        .sProdutoDescricao = objItemNF.sDescricaoItem
        .sUnidadeMed = objItemNF.sUnidadeMed
        
        'nf de importacao
        .dDespImpValorRateado = objItemNF.dDespImpValorRateado
    End With
    
    If objItemNF.lNumIntDocOrig <> 0 Then
    
        Set objTribItemNFAux.objItemNFOrig = New ClassItemNF
        objTribItemNFAux.objItemNFOrig.lNumIntDoc = objItemNF.lNumIntDocOrig
        
        lErro = CF("ItemNFiscal_Le", objTribItemNFAux.objItemNFOrig)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
           
        Set objTribItemNFAux.objItemNFOrig.objTributacao = New ClassTributacaoDocItem
        objTribItemNFAux.objItemNFOrig.objTributacao.iTipoDoc = TIPODOC_TRIB_NF
        objTribItemNFAux.objItemNFOrig.objTributacao.lNumIntDocItem = objTribItemNFAux.objItemNFOrig.lNumIntDoc
        lErro = CF("TributacaoDocItem_Le", objTribItemNFAux.objItemNFOrig.objTributacao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    Call TribItemNFAux_Manual(objTribItemNFAux, objItemNF.objTributacaoItemNF)
    
    If bSRV Then
        lErro = mvarobjNFTributacao.IncluirItemSRV(objTribItemNFAux)
    Else
        lErro = mvarobjNFTributacao.IncluirItem(objTribItemNFAux)
    End If
    If lErro <> SUCESSO Then gError 22883

    Set objTributoDocItem = objTribItemNFAux
    lErro = objTributoDocItem.Ativar(mvarobjNFTributacao)
    If lErro <> SUCESSO Then gError 19306

    ExtraiNFItem = SUCESSO

    Exit Function

Erro_ExtraiNFItem:

    ExtraiNFItem = gErr

    Select Case gErr

        Case 22883, 19306
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153762)

    End Select

    Exit Function

End Function

Private Sub TribItemNFAux_Manual(objTribItemNFAux As ClassTribItemNFAux, objTribItemNF As Object)

    With objTribItemNFAux
    
        .iICMSAliquotaManual = IIf(objTribItemNF.iICMSAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSAliquotaManual)
        .iICMSBaseManual = IIf(objTribItemNF.iICMSBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSBaseManual)
        .iICMSPercRedBaseManual = IIf(objTribItemNF.iICMSPercRedBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSPercRedBaseManual)
        .iICMSSubstAliquotaManual = IIf(objTribItemNF.iICMSSubstAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSubstAliquotaManual)
        .iICMSSubstBaseManual = IIf(objTribItemNF.iICMSSubstBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSubstBaseManual)
        .iICMSSubstValorManual = IIf(objTribItemNF.iICMSSubstValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSubstValorManual)
        .iICMSTipoManual = IIf(objTribItemNF.iICMSTipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSTipoManual)
        .iICMSValorManual = IIf(objTribItemNF.iICMSValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSValorManual)
        .iIPIAliquotaManual = IIf(objTribItemNF.iIPIAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIAliquotaManual)
        .iIPIBaseManual = IIf(objTribItemNF.iIPIBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIBaseManual)
        .iIPIPercRedBaseManual = IIf(objTribItemNF.iIPIPercRedBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIPercRedBaseManual)
        .iIPITipoManual = IIf(objTribItemNF.iIPITipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPITipoManual)
        .iIPIValorManual = IIf(objTribItemNF.iIPIValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIValorManual)
        .iNaturezaOpManual = IIf(objTribItemNF.iNaturezaOpManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iNaturezaOpManual)
        .iTipoTributacaoManual = IIf(objTribItemNF.iTipoTributacaoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iTipoTributacaoManual)
        '.iPISCreditoManual = IIf(objTribItemNF.iPISCreditoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISCreditoManual)
        '.iCOFINSCreditoManual = IIf(objTribItemNF.iCOFINSCreditoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSCreditoManual)
        .iPISCreditoManual = VAR_PREENCH_AUTOMATICO 'Campo que não pode ser alterado em tela tem que ser automático
        .iCOFINSCreditoManual = VAR_PREENCH_AUTOMATICO
        
        .iICMSAliquotaAdicaoDIManual = IIf(objTribItemNF.iICMSAliquotaAdicaoDIManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSAliquotaAdicaoDIManual)
        .iICMSPercRedBaseAdicaoDIManual = IIf(objTribItemNF.iICMSPercRedBaseAdicaoDIManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSPercRedBaseAdicaoDIManual)
        .iDespImpICMSBaseManual = IIf(objTribItemNF.iDespImpICMSBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iDespImpICMSBaseManual)
        .iDespImpICMSValorManual = IIf(objTribItemNF.iDespImpICMSValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iDespImpICMSValorManual)
        .iDespImpICMSCreditoManual = IIf(objTribItemNF.iDespImpICMSCreditoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iDespImpICMSCreditoManual)
        .iICMSSubstPercRedBaseManual = IIf(objTribItemNF.iICMSSubstPercRedBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSubstPercRedBaseManual)
        .iICMSSubstPercMVAManual = IIf(objTribItemNF.iICMSSubstPercMVAManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSubstPercMVAManual)
        .iPISTipoManual = IIf(objTribItemNF.iPISTipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISTipoManual)
        .iPISTipoCalculoManual = IIf(objTribItemNF.iPISTipoCalculoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISTipoCalculoManual)
        .iPISBaseManual = IIf(objTribItemNF.iPISBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISBaseManual)
        .iPISAliquotaManual = IIf(objTribItemNF.iPISAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISAliquotaManual)
        .iPISAliquotaValorManual = IIf(objTribItemNF.iPISAliquotaValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISAliquotaValorManual)
        .iPISQtdeManual = IIf(objTribItemNF.iPISQtdeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISQtdeManual)
        .iPISValorManual = IIf(objTribItemNF.iPISValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISValorManual)
        .iPISSTTipoCalculoManual = IIf(objTribItemNF.iPISSTTipoCalculoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISSTTipoCalculoManual)
        .iPISSTBaseManual = IIf(objTribItemNF.iPISSTBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISSTBaseManual)
        .iPISSTAliquotaManual = IIf(objTribItemNF.iPISSTAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISSTAliquotaManual)
        .iPISSTAliquotaValorManual = IIf(objTribItemNF.iPISSTAliquotaValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISSTAliquotaValorManual)
        .iPISSTQtdeManual = IIf(objTribItemNF.iPISSTQtdeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISSTQtdeManual)
        .iPISSTValorManual = IIf(objTribItemNF.iPISSTValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iPISSTValorManual)
        .iCOFINSTipoManual = IIf(objTribItemNF.iCOFINSTipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSTipoManual)
        .iCOFINSTipoCalculoManual = IIf(objTribItemNF.iCOFINSTipoCalculoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSTipoCalculoManual)
        .iCOFINSBaseManual = IIf(objTribItemNF.iCOFINSBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSBaseManual)
        .iCOFINSAliquotaManual = IIf(objTribItemNF.iCOFINSAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSAliquotaManual)
        .iCOFINSAliquotaValorManual = IIf(objTribItemNF.iCOFINSAliquotaValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSAliquotaValorManual)
        .iCOFINSQtdeManual = IIf(objTribItemNF.iCOFINSQtdeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSQtdeManual)
        .iCOFINSValorManual = IIf(objTribItemNF.iCOFINSValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSValorManual)
        .iCOFINSSTTipoCalculoManual = IIf(objTribItemNF.iCOFINSSTTipoCalculoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSSTTipoCalculoManual)
        .iCOFINSSTBaseManual = IIf(objTribItemNF.iCOFINSSTBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSSTBaseManual)
        .iCOFINSSTAliquotaManual = IIf(objTribItemNF.iCOFINSSTAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSSTAliquotaManual)
        .iCOFINSSTAliquotaValorManual = IIf(objTribItemNF.iCOFINSSTAliquotaValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSSTAliquotaValorManual)
        .iCOFINSSTQtdeManual = IIf(objTribItemNF.iCOFINSSTQtdeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSSTQtdeManual)
        .iCOFINSSTValorManual = IIf(objTribItemNF.iCOFINSSTValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCOFINSSTValorManual)
        .iCSTManual = IIf(objTribItemNF.iCSTManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iCSTManual)
        .iISSQNManual = IIf(objTribItemNF.iISSQNManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSQNManual)
        .iISSBaseManual = IIf(objTribItemNF.iISSBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSBaseManual)
        .iISSAliquotaManual = IIf(objTribItemNF.iISSAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSAliquotaManual)
        .iISSCidadeIBGEManual = IIf(objTribItemNF.iISSCidadeIBGEManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSCidadeIBGEManual)
        .iISSValorManual = IIf(objTribItemNF.iISSValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSValorManual)
        .iIIBaseManual = IIf(objTribItemNF.iIIBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIIBaseManual)
        .iIIDespAduaneiraManual = IIf(objTribItemNF.iIIDespAduaneiraManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIIDespAduaneiraManual)
        .iIIIOFManual = IIf(objTribItemNF.iIIIOFManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIIIOFManual)
        .iIIValorManual = IIf(objTribItemNF.iIIValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIIValorManual)
        .iICMSBaseModalidadeManual = IIf(objTribItemNF.iICMSBaseModalidadeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSBaseModalidadeManual)
        .iICMSSubstBaseModalidadeManual = IIf(objTribItemNF.iICMSSubstBaseModalidadeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSubstBaseModalidadeManual)
        .iIPIEnquadramentoClasseManual = IIf(objTribItemNF.iIPIEnquadramentoClasseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIEnquadramentoClasseManual)
        .iIPIEnquadramentoCodigoManual = IIf(objTribItemNF.iIPIEnquadramentoCodigoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIEnquadramentoCodigoManual)
        .iIPICNPJProdutorManual = IIf(objTribItemNF.iIPICNPJProdutorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPICNPJProdutorManual)
        .iIPISeloCodigoManual = IIf(objTribItemNF.iIPISeloCodigoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPISeloCodigoManual)
        .iIPISeloQtdeManual = IIf(objTribItemNF.iIPISeloQtdeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPISeloQtdeManual)
        .iIPITipoCalculoManual = IIf(objTribItemNF.iIPITipoCalculoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPITipoCalculoManual)
        .iIPIUnidadePadraoQtdeManual = IIf(objTribItemNF.iIPIUnidadePadraoQtdeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIUnidadePadraoQtdeManual)
        .iIPIUnidadePadraoValorManual = IIf(objTribItemNF.iIPIUnidadePadraoValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIUnidadePadraoValorManual)
        .iValorFreteItemManual = IIf(objTribItemNF.iValorFreteItemManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iValorFreteItemManual)
        .iValorSeguroItemManual = IIf(objTribItemNF.iValorSeguroItemManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iValorSeguroItemManual)
        .iValorOutrasDespesasItemManual = IIf(objTribItemNF.iValorOutrasDespesasItemManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iValorOutrasDespesasItemManual)
        .iValorDescontoItemManual = IIf(objTribItemNF.iValorDescontoItemManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iValorDescontoItemManual)
        .iOrigemMercadoriaManual = IIf(objTribItemNF.iOrigemMercadoriaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iOrigemMercadoriaManual)
        .iFCIManual = IIf(objTribItemNF.iFCIManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iFCIManual)
        .iICMSSTCobrAntBaseManual = IIf(objTribItemNF.iICMSSTCobrAntBaseManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSTCobrAntBaseManual)
        .iICMSSTCobrAntValorManual = IIf(objTribItemNF.iICMSSTCobrAntValorManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSTCobrAntValorManual)
        
        
        .iICMSpercBaseOperacaoPropriaManual = IIf(objTribItemNF.iICMSpercBaseOperacaoPropriaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSpercBaseOperacaoPropriaManual)
        .iICMSUFDevidoSTManual = IIf(objTribItemNF.iICMSUFDevidoSTManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSUFDevidoSTManual)
        .iICMSvBCSTRetManual = IIf(objTribItemNF.iICMSvBCSTRetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvBCSTRetManual)
        .iICMSvICMSSTRetManual = IIf(objTribItemNF.iICMSvICMSSTRetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvICMSSTRetManual)
        .iICMSvBCSTDestManual = IIf(objTribItemNF.iICMSvBCSTDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvBCSTDestManual)
        .iICMSvICMSSTDestManual = IIf(objTribItemNF.iICMSvICMSSTDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvICMSSTDestManual)
        .iICMSpCredSNManual = IIf(objTribItemNF.iICMSpCredSNManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSpCredSNManual)
        .iICMSvCredSNManual = IIf(objTribItemNF.iICMSvCredSNManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvCredSNManual)
        .iICMSValorIsentoManual = IIf(objTribItemNF.iICMSValorIsentoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSValorIsentoManual)
        .iICMSMotivoManual = IIf(objTribItemNF.iICMSMotivoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSMotivoManual)
        .iICMSSimplesTipoManual = IIf(objTribItemNF.iICMSSimplesTipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSimplesTipoManual)
        .iISSTipoManual = IIf(objTribItemNF.iISSTipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSTipoManual)
       
        .iNatBCCredManual = IIf(objTribItemNF.iNatBCCredManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iNatBCCredManual)
        .iTotTribManual = IIf(objTribItemNF.iTotTribManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iTotTribManual)
        .iTotTribAliqFedManual = IIf(objTribItemNF.iTotTribAliqFedManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iTotTribAliqFedManual)
        .iTotTribAliqEstManual = IIf(objTribItemNF.iTotTribAliqEstManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iTotTribAliqEstManual)
        .iTotTribAliqMunicManual = IIf(objTribItemNF.iTotTribAliqMunicManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iTotTribAliqMunicManual)
        .iTotTribTipoManual = IIf(objTribItemNF.iTotTribTipoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iTotTribTipoManual)
       
       'nfe 3.10
        .iICMS51ValorOpManual = IIf(objTribItemNF.iICMS51ValorOpManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMS51ValorOpManual)
        .iICMSpCredSNManual = IIf(objTribItemNF.iICMSpCredSNManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSpCredSNManual)
        .iICMSValorDifManual = IIf(objTribItemNF.iICMSValorDifManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSValorDifManual)
        .iISSValorDeducaoManual = IIf(objTribItemNF.iISSValorDeducaoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSValorDeducaoManual)
        .iISSValorOutrasRetManual = IIf(objTribItemNF.iISSValorOutrasRetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSValorOutrasRetManual)
        .iISSValorDescIncondManual = IIf(objTribItemNF.iISSValorDescIncondManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSValorDescIncondManual)
        .iISSValorDescCondManual = IIf(objTribItemNF.iISSValorDescCondManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSValorDescCondManual)
        .iISSValorRetManual = IIf(objTribItemNF.iISSValorRetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSValorRetManual)
        .iISSIndExigibilidadeManual = IIf(objTribItemNF.iISSIndExigibilidadeManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSIndExigibilidadeManual)
        .iISSIndIncentivoManual = IIf(objTribItemNF.iISSIndIncentivoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSIndIncentivoManual)
        .iISSMunicIncidImpManual = IIf(objTribItemNF.iISSMunicIncidImpManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSMunicIncidImpManual)
        .iISSCodServManual = IIf(objTribItemNF.iISSCodServManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSCodServManual)
        .iISSNumProcessoManual = IIf(objTribItemNF.iISSNumProcessoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSNumProcessoManual)
        .iISSCodPaisManual = IIf(objTribItemNF.iISSCodPaisManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iISSCodPaisManual)
        
        'NT 2015.003 (ICMS Interestadual)
        .iICMSInterestBCUFDestManual = IIf(objTribItemNF.iICMSInterestBCUFDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestBCUFDestManual)
        .iICMSInterestPercFCPUFDestManual = IIf(objTribItemNF.iICMSInterestPercFCPUFDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestPercFCPUFDestManual)
        .iICMSInterestAliqUFDestManual = IIf(objTribItemNF.iICMSInterestAliqUFDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestAliqUFDestManual)
        .iICMSInterestAliqManual = IIf(objTribItemNF.iICMSInterestAliqManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestAliqManual)
        .iICMSInterestPercPartilhaManual = IIf(objTribItemNF.iICMSInterestPercPartilhaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestPercPartilhaManual)
        .iICMSInterestVlrFCPUFDestManual = IIf(objTribItemNF.iICMSInterestVlrFCPUFDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestVlrFCPUFDestManual)
        .iICMSInterestVlrUFDestManual = IIf(objTribItemNF.iICMSInterestVlrUFDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestVlrUFDestManual)
        .iICMSInterestVlrUFRemetManual = IIf(objTribItemNF.iICMSInterestVlrUFRemetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestVlrUFRemetManual)
        
        'UA. Tributos Devolvidos (para o item da NF-e)
        .ipDevolManual = IIf(objTribItemNF.ipDevolManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.ipDevolManual)
        .iIPIVlrDevolvidoManual = IIf(objTribItemNF.iIPIVlrDevolvidoManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iIPIVlrDevolvidoManual)
        
        'NFe 4.00
        .iICMSInterestBCFCPUFDestManual = IIf(objTribItemNF.iICMSInterestBCFCPUFDestManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestBCFCPUFDestManual)
        .iICMSvBCFCPSTManual = IIf(objTribItemNF.iICMSvBCFCPSTManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvBCFCPSTManual)
        .iICMSpFCPSTManual = IIf(objTribItemNF.iICMSpFCPSTManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSpFCPSTManual)
        .iICMSvFCPSTManual = IIf(objTribItemNF.iICMSvFCPSTManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvFCPSTManual)
        .iICMSvBCFCPSTRetManual = IIf(objTribItemNF.iICMSvBCFCPSTRetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvBCFCPSTRetManual)
        .iICMSpFCPSTRetManual = IIf(objTribItemNF.iICMSpFCPSTRetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSpFCPSTRetManual)
        .iICMSvFCPSTRetManual = IIf(objTribItemNF.iICMSvFCPSTRetManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvFCPSTRetManual)
        .iICMSSTCobrAntAliquotaManual = IIf(objTribItemNF.iICMSSTCobrAntAliquotaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSTCobrAntAliquotaManual)
        .iICMSvBCFCPManual = IIf(objTribItemNF.iICMSvBCFCPManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvBCFCPManual)
        .iICMSpFCPManual = IIf(objTribItemNF.iICMSpFCPManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSpFCPManual)
        .iICMSvFCPManual = IIf(objTribItemNF.iICMSvFCPManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSvFCPManual)
       
        .iICMSSTBaseDuplaManual = IIf(objTribItemNF.iICMSSTBaseDuplaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSSTBaseDuplaManual)
        .iICMSInterestBaseDuplaManual = IIf(objTribItemNF.iICMSInterestBaseDuplaManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.iICMSInterestBaseDuplaManual)
        .icBenefManual = IIf(objTribItemNF.icBenefManual = VAR_PREENCH_VAZIO, VAR_PREENCH_MANUAL, objTribItemNF.icBenefManual)
       
    End With

End Sub

Public Function Valida_TributacaoNF_Gravacao(Optional objAux As Object) As Long

Dim lErro As Long
Dim objItem As Object
Dim iIndice As Integer
Dim dtDataRef As Date
Dim sCFOP As String
Dim objTribDocItem As New ClassTributacaoDocItem

On Error GoTo Erro_Valida_TributacaoNF_Gravacao

    If iTipoTela = TIPOTELA_NF Then If Len(Trim(gobjDocTribTributacao(objAux).sNaturezaOpInterna)) = 0 Then gError 51090
    If gobjDocTribTributacao(objAux).iTipoTributacao = 0 Then gError 51091
    
    If mvariNFEntrada = 1 Then
        dtDataRef = objAux.dtDataEntrada
        If dtDataRef = DATA_NULA Then dtDataRef = objAux.dtDataEmissao
    Else
        dtDataRef = objAux.dtDataEmissao
    End If
    
    iIndice = 0
    
    For Each objItem In gcolDocItem(objAux)
        
        iIndice = iIndice + 1
        If Len(Trim(gobjDocTribItemTributacao(objItem).sNaturezaOp)) = 0 Then gError 51092
        If gobjDocTribItemTributacao(objItem).iTipoTributacao = 0 Then gError 51093
        If Natop_ErroTamanho(dtDataRef, gobjDocTribItemTributacao(objItem).sNaturezaOp) Then gError 32282
    
    Next
    
    If gbTemSRV Then
    
         For Each objItem In gcolDocItem(objAux, True)
            
            iIndice = iIndice + 1
            If Len(Trim(gobjDocTribItemTributacao(objItem).sNaturezaOp)) = 0 Then gError 51092
            If gobjDocTribItemTributacao(objItem).iTipoTributacao = 0 Then gError 51093
            If Natop_ErroTamanho(dtDataRef, gobjDocTribItemTributacao(objItem).sNaturezaOp) Then gError 32282
        
        Next
    End If

    If Natop_ErroTamanho(dtDataRef, objAux.sNaturezaOp) Then gError 32281
    
    'VALIDA DE ACORDO COM gobjVersao
    'FALTA FAZER O CÓDIGO
    '1 - CFOP
    '2 - Tipo ICMS
    '3 - Tipo PIS
    '4 - Tipo Cofins
    'Se é uma NF interna
    If mvariEmitenteNF = DOCINFO_EMPRESA And gobjCRFAT.iUsaNFe = MARCADO Then
    
        sCFOP = Trim(gobjDocTribTributacao(objAux).sNaturezaOpInterna)
        
        lErro = CF("CFOP_Testa_Versao", sCFOP, gobjVersao.sColunaCFOP, gobjVersao.sDescricao)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        For Each objItem In gcolDocItem(objAux)
            
            objTribDocItem.iICMSTipo = gobjDocTribItemTributacao(objItem).iICMSTipo
            objTribDocItem.iICMSSimplesTipo = gobjDocTribItemTributacao(objItem).iICMSSimplesTipo
            objTribDocItem.iIPITipo = gobjDocTribItemTributacao(objItem).iIPITipo
            objTribDocItem.iPISTipo = gobjDocTribItemTributacao(objItem).iPISTipo
            objTribDocItem.iCOFINSTipo = gobjDocTribItemTributacao(objItem).iCOFINSTipo
            objTribDocItem.iISSTipo = gobjDocTribItemTributacao(objItem).iISSTipo
            objTribDocItem.sNaturezaOp = Trim(gobjDocTribTributacao(objAux).sNaturezaOp)
            
            lErro = Valida_TributacaoNF_Versao(objTribDocItem)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        Next
        
        If gbTemSRV Then
        
             For Each objItem In gcolDocItem(objAux, True)
                
                objTribDocItem.iICMSTipo = gobjDocTribItemTributacao(objItem).iICMSTipo
                objTribDocItem.iICMSSimplesTipo = gobjDocTribItemTributacao(objItem).iICMSSimplesTipo
                objTribDocItem.iIPITipo = gobjDocTribItemTributacao(objItem).iIPITipo
                objTribDocItem.iPISTipo = gobjDocTribItemTributacao(objItem).iPISTipo
                objTribDocItem.iCOFINSTipo = gobjDocTribItemTributacao(objItem).iCOFINSTipo
                objTribDocItem.iISSTipo = gobjDocTribItemTributacao(objItem).iISSTipo
                objTribDocItem.sNaturezaOp = Trim(gobjDocTribTributacao(objAux).sNaturezaOp)
                
                lErro = Valida_TributacaoNF_Versao(objTribDocItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            Next
        End If
    
    
    End If

    Valida_TributacaoNF_Gravacao = SUCESSO

    Exit Function
    
Erro_Valida_TributacaoNF_Gravacao:

    Valida_TributacaoNF_Gravacao = gErr

    Select Case gErr
    
        Case 51090
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
    
        Case 51091
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_NAO_PREENCHIDO", gErr)
        
        Case 51092
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_ITEM_TRIBUTACAO_NAO_PREENCHIDA", gErr, iIndice)
        
        Case 51093
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_ITEM_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 51094
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESCONTO_NAO_PRENCHIDA", gErr)
        
        Case 51095
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESCONTO_NAO_PREENCHIDO", gErr)
        
        Case 51096
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_FRETE_NAO_PRENCHIDA", gErr)
        
        Case 51097
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_FRETE_NAO_PREENCHIDO", gErr)

        Case 51098
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_DESPESAS_NAO_PRENCHIDA", gErr)
        
        Case 51099
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_DESPESAS_NAO_PREENCHIDO", gErr)
    
        Case 51094
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_SEGURO_NAO_PRENCHIDA", gErr)
        
        Case 51095
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_TRIBUTACAO_SEGURO_NAO_PREENCHIDO", gErr)
            
        Case 32281, 32282
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_TAMANHO_INCORRETO", gErr)
            
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153763)
    
    End Select

    Exit Function
    
End Function

Sub NFiscal_ObterTrib_Pedido(objPedidoVenda As ClassPedidoDeVenda, objNFiscal As ClassNFiscal)
'Transfere os dados da tributação do pedido de venda passado para objNFiscal
Dim objFilialEmpPedido As New AdmFiliais, lErro As Long
Dim objItemNF As ClassItemNF

On Error GoTo Erro_NFiscal_ObterTrib_Pedido

    Set objNFiscal.objTributacaoNF = New ClassTributacaoDoc
    
    Call objNFiscal.objTributacaoNF.Copia(objPedidoVenda.objTributacaoPV)
    
    objNFiscal.objTributacaoNF.sNaturezaOpInterna = objPedidoVenda.sNaturezaOp
    objNFiscal.objTributacaoNF.iNaturezaOpManual = objPedidoVenda.iNaturezaOpManual
    
''    If objPedidoVenda.iFilialEmpresa <> objPedidoVenda.iFilialEmpresaFaturamento Then
''
''        objFilialEmpPedido.iCodFilial = objPedidoVenda.iFilialEmpresa
''        lErro = CF("FilialEmpresa_Le", objFilialEmpPedido)
''
''        If objFilialEmpPedido.iSuperSimples <> objFilialEmpresa.iSuperSimples Then
''
''            For Each objItemNF In objNFiscal.colItens
''
''                objItemNF.objTributacaoItemNF.iRegimeTributarioManual = VAR_PREENCH_AUTOMATICO
''
''            Next
''
''        End If
''
''    End If

    Exit Sub
    
Erro_NFiscal_ObterTrib_Pedido:

    Exit Sub

End Sub

Sub ItemNFiscal_ObterTrib_ItemPedido(objItemPV As ClassItemPedido, objItemTributacaoNF As ClassTributacaoDocItem)

Dim lErro As Long, dFator As Double, dQtde As Double, objProduto As New ClassProduto

On Error GoTo Erro_ItemNFiscal_ObterTrib_ItemPedido

    Set objItemTributacaoNF = New ClassTributacaoDocItem
    
    Call objItemTributacaoNF.Copia(objItemPV.objTributacaoItemPV)
    
    If objItemPV.dQuantFaturada <> 0 Then
    
        dQtde = objItemPV.dQuantidade - objItemPV.dQuantCancelada - objItemPV.dQuantFaturada
        
        If UCase(objItemPV.objTributacaoItemPV.sUMTrib) <> UCase(objItemPV.sUnidadeMed) And dQtde <> 0 And Len(Trim(objItemPV.sProduto)) <> 0 Then
            
            objProduto.sCodigo = objItemPV.sProduto
            
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then Error 27811
            If lErro <> SUCESSO Then Error 27811
            
            'Faz a conversão da UM da tela para a UM de estoque
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objItemPV.sUnidadeMed, objItemTributacaoNF.sUMTrib, dFator)
            If lErro <> SUCESSO Then Error 27811
                
            dQtde = dQtde * dFator
            
        End If
        
        If dQtde < 0 Then dQtde = 0
        objItemTributacaoNF.dQtdTrib = dQtde
        
    End If
    
    Exit Sub

Erro_ItemNFiscal_ObterTrib_ItemPedido:

    Select Case Err

        Case 27811

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 211316)

    End Select

    Exit Sub
    
End Sub

Private Sub IPI_Frame_Habilita(bHabilitado As Boolean)

     TabTrib_Controle("FrameTabTribDet")(3).Enabled = bHabilitado
    
End Sub
    
Private Sub ICMS_Frame_Habilita(bHabilitado As Boolean)

     TabTrib_Controle("FrameTabTribDet")(2).Enabled = bHabilitado
    
End Sub

Private Sub ISS_Frame_Habilita(bHabilitado As Boolean)

     TabTrib_Controle("FrameTabTribDet")(6).Enabled = bHabilitado
    
End Sub

Sub NatOpInterna_GotFocus(iAlterado As Integer)
Dim iNaturezaAux As Integer

    iNaturezaAux = mvariNatOpInternaAlterada
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("NatOpInterna"), iAlterado)
    mvariNatOpInternaAlterada = iNaturezaAux

End Sub

Sub TipoTributacao_GotFocus(iAlterado As Integer)
Dim iTipoTribAux As Integer

    iTipoTribAux = mvariTipoTributacaoAlterado
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("TipoTributacao"), iAlterado)
    mvariTipoTributacaoAlterado = iTipoTribAux
    
End Sub

Sub NaturezaOpItem_GotFocus(iAlterado As Integer)
Dim iNaturezaAux As Integer

    iNaturezaAux = mvariNatOpItemAlterado
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("NaturezaOpItem"), iAlterado)
    mvariNatOpItemAlterado = iNaturezaAux
    
End Sub

Sub TipoTributacaoItem_GotFocus(iAlterado As Integer)
Dim iTipoTribAux As Integer

    iTipoTribAux = mvariTipoTributacaoItemAlterado
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("TipoTributacaoItem"), iAlterado)
    mvariTipoTributacaoItemAlterado = iTipoTribAux
    
End Sub

Private Function TributacaoNFOrig_Reset(objNFiscal As ClassNFiscal) As Long
'cria ou atualiza mvarobjNFTributacao com dados de nf original p/nf de devolucao

Dim lErro As Long, objTributoDoc As ClassTributoDoc, objItemNF As ClassItemNF
Dim iIndice As Integer
Dim iTipoNFiscal As Integer, sSiglaTipoNFiscal As String, iEmitenteNF As Integer
Dim dtDataEmissao As Date, dtDataEntrada As Date
Dim sSerie As String
Dim objTributoDocItemAux As Object

On Error GoTo Erro_TributacaoNFOrig_Reset

    'se mvarobjNFTributacao já foi inicializado
    If Not (mvarobjNFTributacao Is Nothing) Then
        
        'salva valores anteriores
        iTipoNFiscal = mvarobjNFTributacao.iTipoNFiscal
        sSiglaTipoNFiscal = mvarobjNFTributacao.sSiglaTipoNFiscal
        iEmitenteNF = mvariEmitenteNF
        dtDataEmissao = mvarobjNFTributacao.dtDataEmissao
        dtDataEntrada = mvarobjNFTributacao.dtDataEntrada
        sSerie = mvarobjNFTributacao.sSerie
        
        Set objTributoDoc = mvarobjNFTributacao
        lErro = objTributoDoc.Desativar
        If lErro <> SUCESSO Then Error 59304
        
        Set mvarobjNFTributacao = New ClassTribNFAux
        
        'retorna os valores salvos acima
        mvarobjNFTributacao.iTipoNFiscal = iTipoNFiscal
        mvarobjNFTributacao.sSiglaTipoNFiscal = sSiglaTipoNFiscal
        mvariEmitenteNF = iEmitenteNF
        mvarobjNFTributacao.dtDataEmissao = dtDataEmissao
        mvarobjNFTributacao.dtDataEntrada = dtDataEntrada
        mvarobjNFTributacao.sSerie = sSerie
        
    Else
    
        Set mvarobjNFTributacao = New ClassTribNFAux
        mvarobjNFTributacao.dtDataEmissao = gdtDataAtual
        mvarobjNFTributacao.dtDataEntrada = gdtDataAtual
        mvarobjNFTributacao.sSerie = "" '???

    End If
    
    Set objTributoDoc = mvarobjNFTributacao
    lErro = objTributoDoc.Ativar
    If lErro <> SUCESSO Then Error 59305

    For iIndice = 1 To objNFiscal.colItensNF.Count
        lErro = Inclusao_Item_Grid(iIndice, objNFiscal.colItensNF.Item(iIndice).sProduto)
        If lErro <> SUCESSO Then Error 59306
        
        Set objTributoDocItemAux = gobjDocTribItem(iIndice, False)
        Set objTributoDocItemAux.objItemNFOrig = objNFiscal.colItensNF.Item(iIndice)
        
    Next
    
    With objNFiscal.objTributacaoNF
    
        mvarobjNFTributacao.dValorDesconto = objNFiscal.dValorDesconto
        mvarobjNFTributacao.dValorFrete = objNFiscal.dValorFrete
        mvarobjNFTributacao.dValorOutrasDespesas = objNFiscal.dValorOutrasDespesas
        mvarobjNFTributacao.dValorSeguro = objNFiscal.dValorSeguro
        mvarobjNFTributacao.iFilialCli = objNFiscal.iFilialCli
        mvarobjNFTributacao.iFilialEmpresa = objNFiscal.iFilialEmpresa
        mvarobjNFTributacao.iFilialEntrega = objNFiscal.iFilialEntrega
        mvarobjNFTributacao.iFilialForn = objNFiscal.iFilialForn
        
        mvarobjNFTributacao.lCliente = objNFiscal.lCliente
        mvarobjNFTributacao.lFornecedor = objNFiscal.lFornecedor
        
        mvarobjNFTributacao.dtDataEmissao = objNFiscal.dtDataEmissao
        'mvarobjNFTributacao.dtDataEntrada = objNFiscal.dtDataEntrada
        mvarobjNFTributacao.sSerie = objNFiscal.sSerie
        
    End With
    
    Set gobjDocTribTributacao.colTiposTribICMS = mvarcolTiposTribICMS
    Set gobjDocTribTributacao.colTiposTribIPI = mvarcolTiposTribIPI
    Set gobjDocTribTributacao.colTiposTribPISCOFINS = mvarcolTiposTribPISCOFINS
    Set gobjDocTribTributacao.colTiposTribISS = mvarcolTiposTribISS
    Set gobjDocTribTributacao.colTiposTribICMSSimples = mvarcolTiposTribICMSSimples
    
    mvariNatOpInternaAlterada = 0
    mvariTipoTributacaoAlterado = 0
    mvariISSAliquotaAlterada = 0
    mvariISSValorAlterado = 0
    mvariValorIRRFAlterado = 0
    mvariPISRetidoAlterado = 0
    mvariPISValorAlterado = 0
    mvariISSRetidoAlterado = 0
    mvariCOFINSRetidoAlterado = 0
    mvariCOFINSValorAlterado = 0
    mvariCSLLRetidoAlterado = 0
    mvariAliqIRAlterada = 0

    mvariINSSBaseAlterada = 0
    mvariINSSDeducoesAlterada = 0
    mvariINSSValorAlterado = 0
'    mvariINSSRetidoAlterado = 0
    mvariIRBaseAlterado = 0
    
    mvariTotTribAlterado = 0
    mvariTotTribFedAlterado = 0
    mvariTotTribEstAlterado = 0
    mvariTotTribMunicAlterado = 0
    mvariTotTribTipoAlterado = 0
    
    'nfe 3.10
'    mvariIndConsumidorFinalAlterado = 0
'    mvariIndPresencaAlterado = 0
    mvariDataPrestServicoAlterado = 0
    mvariNatOpItemAlterado = 0
    mvariTipoTributacaoItemAlterado = 0
    mvariICMSBaseItemAlterado = 0
    mvariICMSPercRedBaseItemAlterado = 0
    mvariICMSAliquotaItemAlterado = 0
    mvariICMSValorItemAlterado = 0
    mvariICMSSubstBaseItemAlterado = 0
    mvariICMSSubstAliquotaItemAlterado = 0
    mvariICMSSubstValorItemAlterado = 0
    mvariIPIBaseItemAlterado = 0
    mvariIPIPercRedBaseItemAlterado = 0
    mvariIPIAliquotaItemAlterado = 0
    mvariIPIValorItemAlterado = 0
    mvariICMSSTCobrAntBaseItemAlterado = 0
    mvariICMSSTCobrAntValorItemAlterado = 0
    
    mvariCOFINSSTAliquotaValorItemAlterado = 0
    mvariCOFINSSTAliquotaItemAlterado = 0
    mvariCOFINSSTValorItemAlterado = 0
    mvariCOFINSSTBaseItemAlterado = 0
    mvariCOFINSSTQtdeItemAlterado = 0
    mvariCOFINSAliquotaValorItemAlterado = 0
    mvariCOFINSAliquotaItemAlterado = 0
    mvariCOFINSValorItemAlterado = 0
    mvariCOFINSBaseItemAlterado = 0
    mvariCOFINSQtdeItemAlterado = 0
    mvariPISAliquotaValorItemAlterado = 0
    mvariPISAliquotaItemAlterado = 0
    mvariPISValorItemAlterado = 0
    mvariPISBaseItemAlterado = 0
    mvariPISQtdeItemAlterado = 0
    mvariPISSTAliquotaValorItemAlterado = 0
    mvariPISSTAliquotaItemAlterado = 0
    mvariPISSTValorItemAlterado = 0
    mvariPISSTBaseItemAlterado = 0
    mvariPISSTQtdeItemAlterado = 0
    mvariISSCidadeIBGEItemAlterado = 0
    mvariISSAliquotaItemAlterado = 0
    mvariISSValorItemAlterado = 0
    mvariISSBaseItemAlterado = 0
    mvariValorFreteItemAlterado = 0
    mvariValorSeguroItemAlterado = 0
    mvariValorDescontoItemAlterado = 0
    mvariValorOutrasDespesasItemAlterado = 0
    mvariIIIOFItemAlterado = 0
    mvariIIBaseItemAlterado = 0
    mvariIIValorItemAlterado = 0
    mvariIIDespAduaneiraItemAlterado = 0
    mvariICMSSubstPercRedBaseItemAlterado = 0
    mvariICMSSubstPercMVAItemAlterado = 0
    mvariIPIUnidadePadraoValorItemAlterado = 0
    mvariIPIUnidadePadraoQtdeItemAlterado = 0
    mvariFCIAlterado = 0
    
    mvariICMSpercBaseOperacaoPropriaAlterado = 0
    mvariICMSUFDevidoSTAlterado = 0
    mvariICMSvBCSTRetAlterado = 0
    mvariICMSvICMSSTRetAlterado = 0
    mvariICMSvBCSTDestAlterado = 0
    mvariICMSvICMSSTDestAlterado = 0
    mvariICMSpCredSNAlterado = 0
    mvariICMSvCredSNAlterado = 0
    mvariICMSValorIsentoAlterado = 0
    mvariICMSMotivoAlterado = 0
    mvariICMSSimplesTipoAlterado = 0
    mvariISSTipoAlterado = 0
    
    mvariTotTribItemAlterado = 0
    mvariTotTribFedItemAlterado = 0
    mvariTotTribEstItemAlterado = 0
    mvariTotTribMunicItemAlterado = 0
    mvariTotTribItemTipoAlterado = 0
    
    'NT 2015.003 (ICMS Interestadual)
    mvariICMSInterestPercFCPUFDestItemAlterado = 0
    mvariICMSInterestBCUFDestItemAlterado = 0
    mvariICMSInterestAliqUFDestItemAlterado = 0
    mvariICMSInterestVlrUFDestItemAlterado = 0
    mvariICMSInterestVlrUFRemetItemAlterado = 0
    mvariICMSInterestVlrFCPUFDestItemAlterado = 0
    
    'nfe 4.0
    mvariICMSInterestBCFCPUFDestItemAlterado = 0
    mvariICMSpFCPSTItemAlterado = 0
    mvariICMSvFCPSTItemAlterado = 0
    mvariICMSvBCFCPSTRetItemAlterado = 0
    mvariICMSpFCPSTRetItemAlterado = 0
    mvariICMSvFCPSTRetItemAlterado = 0
    mvariICMSSTCobrAntAliquotaItemAlterado = 0
    mvariICMSvBCFCPItemAlterado = 0
    mvariICMSpFCPItemAlterado = 0
    mvariICMSvFCPItemAlterado = 0
    
    TributacaoNFOrig_Reset = SUCESSO

    Exit Function

Erro_TributacaoNFOrig_Reset:

    TributacaoNFOrig_Reset = Err

    Select Case Err

        Case 59304, 59305, 59306

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153764)

    End Select

    Exit Function

End Function

Public Sub INSSBase_Change()
    mvariINSSBaseAlterada = 1
End Sub

Public Sub INSSBase_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_INSSBase_Validate

    If mvariINSSBaseAlterada = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("INSSBase").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("INSSBase").Text)
        If lErro <> SUCESSO Then gError 81598
    
    End If
    
    Call AtualizarTabTrib
        
    mvariINSSBaseAlterada = 0

    Exit Sub
    
Erro_INSSBase_Validate:

    Cancel = True
    
    Select Case gErr

        Case 81598

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153765)

    End Select

    Exit Sub

End Sub

Public Sub INSSDeducoes_Change()
    mvariINSSDeducoesAlterada = 1
End Sub

Public Sub INSSDeducoes_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_INSSDeducoes_Validate

    If mvariINSSDeducoesAlterada = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("INSSDeducoes").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("INSSDeducoes").Text)
        If lErro <> SUCESSO Then gError 81599
    
    End If
    
    Call AtualizarTabTrib
        
    mvariINSSDeducoesAlterada = 0

    Exit Sub
    
Erro_INSSDeducoes_Validate:

    Cancel = True
    
    Select Case gErr

        Case 81599

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153766)

    End Select

    Exit Sub

End Sub

Public Sub INSSRetido_Click()
Dim lErro As Long
On Error GoTo Erro_INSSRetido_Click

    If mvariTrazendoTribTela = 0 Then Call AtualizarTabTrib

    Exit Sub

Erro_INSSRetido_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153767)

    End Select

    Exit Sub

End Sub

Public Sub INSSValor_Change()
    mvariINSSValorAlterado = 1
End Sub

Public Sub INSSValor_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_INSSValor_Validate

    If mvariINSSValorAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("INSSValor").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("INSSValor").Text)
        If lErro <> SUCESSO Then gError 81600
    
    End If
    
    Call AtualizarTabTrib
        
    mvariINSSValorAlterado = 0

    Exit Sub
    
Erro_INSSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 81600

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153768)

    End Select

    Exit Sub

End Sub

Sub NFiscal_ObterTrib_Orcamento(objOrcamentoVenda As ClassOrcamentoVenda, objNFiscal As ClassNFiscal)
'Transfere os dados da tributação do orcamento de venda passado para objNFiscal

    Set objNFiscal.objTributacaoNF = New ClassTributacaoDoc
    
    Call objNFiscal.objTributacaoNF.Copia(objOrcamentoVenda.objTributacaoOV)
    
    objNFiscal.objTributacaoNF.sNaturezaOpInterna = objOrcamentoVenda.sNaturezaOp
    objNFiscal.objTributacaoNF.iNaturezaOpManual = objOrcamentoVenda.iNaturezaOpManual
    
End Sub

Sub ItemNFiscal_ObterTrib_ItemOrcamento(objItemOV As ClassItemOV, objItemTributacaoNF As ClassTributacaoDocItem)

    Set objItemTributacaoNF = New ClassTributacaoDocItem
    
    Call objItemTributacaoNF.Copia(objItemOV.objTributacaoItemOV)
    
End Sub

Sub DataEmissao_Alterada(ByVal dtDataEmissao As Date)
'deve ser chamada qdo a data de emissao for alterada

    If gobjDocTrib.dtDataEmissao <> dtDataEmissao Then

        gobjDocTrib.dtDataEmissao = dtDataEmissao
        
        Call gobjDocTrib2.Alterado

        Call mvarobjForm.ValorTotal_Calcula
        
    End If
    
End Sub

Sub DataEntrada_Alterada(ByVal dtDataEntrada As Date)
'deve ser chamada qdo a data de entrada for alterada

    If mvarobjNFTributacao.dtDataEntrada <> dtDataEntrada Then
    
        mvarobjNFTributacao.dtDataEntrada = dtDataEntrada
        
        Call gobjDocTrib2.Alterado

        Call mvarobjForm.ValorTotal_Calcula
        
    End If
    
End Sub

Function NatOp_VerificaTamanho(ByVal sNatOp As String) As Long
    
Dim lErro As Long
Dim dtDataRef As Date

On Error GoTo Erro_NatOp_VerificaTamanho

    If Len(Trim(TabTrib_Controle("DataEmissao").ClipText)) > 0 Then
        dtDataRef = MaskedParaDate(TabTrib_Controle("DataEmissao"))
    Else
        dtDataRef = DATA_NULA
    End If
        
    If sNatOp < NATUREZA_SAIDA_COD_INICIAL Then
            
        If dtDataRef < DATA_INICIO_CFOP4 Then
        
            If Len(Trim(TabTrib_Controle("DataEntrada").ClipText)) > 0 Then
                dtDataRef = MaskedParaDate(TabTrib_Controle("DataEntrada"))
            End If
        
        End If
            
    End If
    
    If Natop_ErroTamanho(dtDataRef, sNatOp) Then gError 106550
    
    NatOp_VerificaTamanho = SUCESSO
     
    Exit Function
    
Erro_NatOp_VerificaTamanho:

    NatOp_VerificaTamanho = gErr
     
    Select Case gErr
          
        Case 106550
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_TAMANHO_INCORRETO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153769)
     
    End Select
     
    Exit Function

End Function

Sub DataEmissaoEntrada_Alterada(ByVal dtDataEmissao As Date, ByVal dtDataEntrada As Date)
'deve ser chamada qdo a data de emissao ou de entrada for alterada

Dim bMudou As Boolean

    bMudou = False
    
    If mvarobjNFTributacao.dtDataEmissao <> dtDataEmissao Then

        mvarobjNFTributacao.dtDataEmissao = dtDataEmissao
        bMudou = True
        
    End If
    
    If mvarobjNFTributacao.dtDataEntrada <> dtDataEntrada Then
    
        mvarobjNFTributacao.dtDataEntrada = dtDataEntrada
        bMudou = True
        
    End If

    If bMudou Then
    
        Call gobjDocTrib2.Alterado

        Call mvarobjForm.ValorTotal_Calcula
        
    End If
    
End Sub

Sub Serie_Alterada(ByVal sSerie As String)
'deve ser chamada qdo a serie for alterada

    If Not ISSerieEletronica(sSerie) Then
        If Not (mvarobjForm Is Nothing) Then sSerie = Converte_Serie_Eletronica(sSerie, TabTrib_Controle("EletronicaFed").Value)
    End If

    If mvarobjNFTributacao.sSerie <> sSerie Then

        mvarobjNFTributacao.sSerie = sSerie
    
        Call gobjDocTrib2.Alterado

        Call mvarobjForm.ValorTotal_Calcula
        
    End If
    
End Sub

Sub PISRetido_Change()

    mvariPISRetidoAlterado = REGISTRO_ALTERADO

End Sub

Sub ISSRetido_Change()

    mvariISSRetidoAlterado = REGISTRO_ALTERADO

End Sub

Sub COFINSRetido_Change()

    mvariCOFINSRetidoAlterado = REGISTRO_ALTERADO

End Sub

Sub CSLLRetido_Change()

    mvariCSLLRetidoAlterado = REGISTRO_ALTERADO

End Sub

Sub TotTrib_Change()

    mvariTotTribAlterado = REGISTRO_ALTERADO

End Sub

Sub PISRetido_Validate(Cancel As Boolean)

Dim dPISRetido As Double
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_PISRetido_Validate

    If mvariPISRetidoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISRetido").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISRetido").Text)
        If lErro <> SUCESSO Then Error 27797

        dPISRetido = CDbl(TabTrib_Controle("PISRetido").Text)

        'este teste é necessário pois nas nfs externas o valor total é uma masked pois é informado pelo usuario p/efeito de conferencia
        If mvariNFEntrada = NF_SAIDA Then
            If Len(Trim(TabTrib_Controle("ValorTotal").Caption)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("ValorTotal").Caption)
        Else
            If Len(Trim(TabTrib_Controle("Total").Text)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("Total").Text)
        End If

        If dPISRetido > dValorTotal Then Error 27798

        TabTrib_Controle("PISRetido").Text = Format(dPISRetido, "Standard")

    End If

    mvariPISRetidoAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_PISRetido_Validate:

    Cancel = True
    
    Select Case Err

        Case 26654, 27677, 27678, 27797
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case 27798
            Call Rotina_Erro(vbOKOnly, "ERRO_PIS_RETIDO_MAIOR_VALOR_TOTAL", Err, dPISRetido, dValorTotal)
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153770)

    End Select

    Exit Sub

End Sub

Sub COFINSRetido_Validate(Cancel As Boolean)

Dim dCOFINSRetido As Double
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_COFINSRetido_Validate

    If mvariCOFINSRetidoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSRetido").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSRetido").Text)
        If lErro <> SUCESSO Then Error 27797

        dCOFINSRetido = CDbl(TabTrib_Controle("COFINSRetido").Text)

        'este teste é necessário pois nas nfs externas o valor total é uma masked pois é informado pelo usuario p/efeito de conferencia
        If mvariNFEntrada = NF_SAIDA Then
            If Len(Trim(TabTrib_Controle("ValorTotal").Caption)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("ValorTotal").Caption)
        Else
            If Len(Trim(TabTrib_Controle("Total").Text)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("Total").Text)
        End If

        If dCOFINSRetido > dValorTotal Then Error 27798

        TabTrib_Controle("COFINSRetido").Text = Format(dCOFINSRetido, "Standard")

    End If

    mvariCOFINSRetidoAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_COFINSRetido_Validate:

    Cancel = True
    
    Select Case Err

        Case 26654, 27677, 27678, 27797
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case 27798
            Call Rotina_Erro(vbOKOnly, "ERRO_COFINS_RETIDO_MAIOR_VALOR_TOTAL", Err, dCOFINSRetido, dValorTotal)
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153771)

    End Select

    Exit Sub

End Sub

Sub CSLLRetido_Validate(Cancel As Boolean)

Dim dCSLLRetido As Double
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_CSLLRetido_Validate

    If mvariCSLLRetidoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("CSLLRetido").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("CSLLRetido").Text)
        If lErro <> SUCESSO Then Error 27797

        dCSLLRetido = CDbl(TabTrib_Controle("CSLLRetido").Text)

        'este teste é necessário pois nas nfs externas o valor total é uma masked pois é informado pelo usuario p/efeito de conferencia
        If mvariNFEntrada = NF_SAIDA Then
            If Len(Trim(TabTrib_Controle("ValorTotal").Caption)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("ValorTotal").Caption)
        Else
            If Len(Trim(TabTrib_Controle("Total").Text)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("Total").Text)
        End If

        If dCSLLRetido > dValorTotal Then Error 27798

        TabTrib_Controle("CSLLRetido").Text = Format(dCSLLRetido, "Standard")

    End If

    mvariCSLLRetidoAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_CSLLRetido_Validate:

    Cancel = True
    
    Select Case Err

        Case 26654, 27677, 27678, 27797
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case 27798
            Call Rotina_Erro(vbOKOnly, "ERRO_CSLL_RETIDO_MAIOR_VALOR_TOTAL", Err, dCSLLRetido, dValorTotal)
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153772)

    End Select

    Exit Sub

End Sub

Function Alteracao_Item_Grid_ItemPC(ByVal iIndice As Integer, ByVal objItemPC As Object) As Long
'trata a alteracao de uma linha de item no grid usando dados de um item de pedido de compra

Dim lErro As Long, sProduto As String, iPreenchido As Integer
Dim objItem As ClassTribItemNFAux

On Error GoTo Erro_Alteracao_Item_Grid_ItemPC

    '???? aguardar configuracao em fatconfig
    If False Then
    
        Set objItem = mvarobjNFTributacao.colItens.Item(iIndice)
    
        objItem.iICMSAliquotaManual = VAR_PREENCH_MANUAL
        objItem.dICMSAliquota = objItemPC.dAliquotaICMS
        objItem.iIPIAliquotaManual = VAR_PREENCH_MANUAL
        objItem.dIPIAliquota = objItemPC.dAliquotaIPI
    
    End If
    
    Alteracao_Item_Grid_ItemPC = SUCESSO

    Exit Function

Erro_Alteracao_Item_Grid_ItemPC:

    Alteracao_Item_Grid_ItemPC = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153773)

    End Select

    Exit Function

End Function

Sub NatOpInterna_Automatica()
    mvarobjNFTributacao.iNaturezaOpManual = VAR_PREENCH_AUTOMATICO
End Sub

Sub ISSRetido_Validate(Cancel As Boolean)

Dim dISSRetido As Double
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_ISSRetido_Validate

    If mvariISSRetidoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ISSRetido").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSRetido").Text)
        If lErro <> SUCESSO Then Error 27797

        dISSRetido = CDbl(TabTrib_Controle("ISSRetido").Text)

        'este teste é necessário pois nas nfs externas o valor total é uma masked pois é informado pelo usuario p/efeito de conferencia
        If mvariNFEntrada = NF_SAIDA Then
            If Len(Trim(TabTrib_Controle("ValorTotal").Caption)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("ValorTotal").Caption)
        Else
            If Len(Trim(TabTrib_Controle("Total").Text)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("Total").Text)
        End If

        If dISSRetido > dValorTotal Then Error 27798

        TabTrib_Controle("ISSRetido").Text = Format(dISSRetido, "Standard")

    End If

    mvariISSRetidoAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_ISSRetido_Validate:

    Cancel = True
    
    Select Case Err

        Case 26654, 27677, 27678, 27797
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case 27798
            Call Rotina_Erro(vbOKOnly, "ERRO_ISS_RETIDO_MAIOR_VALOR_TOTAL", Err, dISSRetido, dValorTotal)
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153770)

    End Select

    Exit Sub

End Sub

Sub PISValor_Change()

    mvariPISValorAlterado = REGISTRO_ALTERADO

End Sub

Sub COFINSValor_Change()

    mvariCOFINSValorAlterado = REGISTRO_ALTERADO

End Sub

Sub PISValor_Validate(Cancel As Boolean)

Dim dPISValor As Double
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_PISValor_Validate

    If mvariPISValorAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISValor").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISValor").Text)
        If lErro <> SUCESSO Then Error 27797

        dPISValor = CDbl(TabTrib_Controle("PISValor").Text)

        'este teste é necessário pois nas nfs externas o valor total é uma masked pois é informado pelo usuario p/efeito de conferencia
        If mvariNFEntrada = NF_SAIDA Then
            If Len(Trim(TabTrib_Controle("ValorTotal").Caption)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("ValorTotal").Caption)
        Else
            If Len(Trim(TabTrib_Controle("Total").Text)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("Total").Text)
        End If

        If dPISValor > dValorTotal Then Error 27798

        TabTrib_Controle("PISValor").Text = Format(dPISValor, "Standard")

    End If

    mvariPISValorAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_PISValor_Validate:

    Cancel = True
    
    Select Case Err

        Case 26654, 27677, 27678, 27797
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case 27798
            Call Rotina_Erro(vbOKOnly, "ERRO_PIS_Valor_MAIOR_VALOR_TOTAL", Err, dPISValor, dValorTotal)
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153770)

    End Select

    Exit Sub

End Sub

Sub COFINSValor_Validate(Cancel As Boolean)

Dim dCOFINSValor As Double
Dim lErro As Long
Dim dValorTotal As Double

On Error GoTo Erro_COFINSValor_Validate

    If mvariCOFINSValorAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSValor").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSValor").Text)
        If lErro <> SUCESSO Then Error 27797

        dCOFINSValor = CDbl(TabTrib_Controle("COFINSValor").Text)

        'este teste é necessário pois nas nfs externas o valor total é uma masked pois é informado pelo usuario p/efeito de conferencia
        If mvariNFEntrada = NF_SAIDA Then
            If Len(Trim(TabTrib_Controle("ValorTotal").Caption)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("ValorTotal").Caption)
        Else
            If Len(Trim(TabTrib_Controle("Total").Text)) > 0 Then dValorTotal = CDbl(TabTrib_Controle("Total").Text)
        End If

        If dCOFINSValor > dValorTotal Then Error 27798

        TabTrib_Controle("COFINSValor").Text = Format(dCOFINSValor, "Standard")

    End If

    mvariCOFINSValorAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_COFINSValor_Validate:

    Cancel = True
    
    Select Case Err

        Case 26654, 27677, 27678, 27797
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case 27798
            Call Rotina_Erro(vbOKOnly, "ERRO_COFINS_Valor_MAIOR_VALOR_TOTAL", Err, dCOFINSValor, dValorTotal)
''???            If TabTrib_Controle("ValorIRRF").Visible Then TabTrib_Controle("ValorIRRF").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153771)

    End Select

    Exit Sub

End Sub

Function NFImportacao_AtualizaItem(ByVal iIndice As Integer, ByVal dDespImpValorRateado As Double, ByVal dICMSAliquotaAdicaoDI As Double, ByVal dPISAliquotaAdicaoDI As Double, ByVal dCOFINSAliquotaAdicaoDI As Double, ByVal dIIBase As Double, ByVal dIIDespAduaneira As Double, ByVal dIIIOF As Double, ByVal dIIValor As Double, ByVal objItemAdicaoDIItemNF As ClassItemAdicaoDIItemNF) As Long
Dim objItem As ClassTribItemNFAux, dAux As Double
    
    Set objItem = mvarobjNFTributacao.colItens.Item(iIndice)
    objItem.dDespImpValorRateado = dDespImpValorRateado
    objItem.dICMSAliquotaAdicaoDI = dICMSAliquotaAdicaoDI
    objItem.dICMSPercRedBaseAdicaoDI = objItemAdicaoDIItemNF.dICMSPercRedBaseAdicaoDI
    objItem.dPISAliquotaAdicaoDI = dPISAliquotaAdicaoDI
    objItem.dCOFINSAliquotaAdicaoDI = dCOFINSAliquotaAdicaoDI
    objItem.dIPIAliquotaAdicaoDI = objItemAdicaoDIItemNF.dIPIAliquotaAdicaoDI
    objItem.dIPIValorUnitAdicaoDI = objItemAdicaoDIItemNF.dIPIUnidadePadraoValor
    
    objItem.dIIBase = dIIBase
    objItem.dIIDespAduaneira = dIIDespAduaneira
    objItem.dIIIOF = dIIIOF
    objItem.dIIValor = dIIValor
    
    dAux = 0
    
    For Each objItem In mvarobjNFTributacao.colItens
    
        dAux = Arredonda_Moeda(dAux + Arredonda_Moeda(objItem.dIIValor))
    
    Next
    
    mvarobjNFTributacao.dIIValor = dAux
    
End Function

Function ExtraiNF_Pack(objNFiscal As ClassNFiscal) As Long
'carrega dados de objNFiscal para "Me"(a propria instancia da classe ClassTribTab)
'para notas com pack (Pharmakin)

Dim lErro As Long, objTribItemNFAux As ClassTribItemNFAux

On Error GoTo Erro_ExtraiNF_Pack

    'obter o tipo da nota
    
    With objNFiscal.objTributacaoNF
    
        mvarobjNFTributacao.dValorDesconto = objNFiscal.dValorDesconto
        mvarobjNFTributacao.dValorFrete = objNFiscal.dValorFrete
        mvarobjNFTributacao.dValorOutrasDespesas = objNFiscal.dValorOutrasDespesas
        mvarobjNFTributacao.dValorSeguro = objNFiscal.dValorSeguro
        mvarobjNFTributacao.iFilialCli = objNFiscal.iFilialCli
        mvarobjNFTributacao.iFilialEmpresa = objNFiscal.iFilialEmpresa
        mvarobjNFTributacao.iFilialEntrega = objNFiscal.iFilialEntrega
        mvarobjNFTributacao.iFilialForn = objNFiscal.iFilialForn
        mvarobjNFTributacao.lCliente = objNFiscal.lCliente
        mvarobjNFTributacao.lFornecedor = objNFiscal.lFornecedor
        mvarobjNFTributacao.dtDataEmissao = objNFiscal.dtDataEmissao
        mvarobjNFTributacao.dtDataEntrada = objNFiscal.dtDataEntrada
        mvarobjNFTributacao.sSerie = objNFiscal.sSerie
        
    End With
    
    ExtraiNF_Pack = SUCESSO

    Exit Function

Erro_ExtraiNF_Pack:

    ExtraiNF_Pack = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153760)

    End Select

    Exit Function

End Function

Private Function TabTrib_Controle(ByVal sControle As String) As Object

On Error GoTo Erro_TabTrib_Controle
    
    If mvarobjTabTrib Is Nothing Then
        Set TabTrib_Controle = mvarobjForm.Controls(sControle)
        Exit Function
    End If
    Set TabTrib_Controle = mvarobjTabTrib.object.Controls(sControle)  '   mvarobjTabTrib.Controle(sControle)
    Exit Function
Erro_TabTrib_Controle:
    Set TabTrib_Controle = mvarobjForm.Controls(sControle)
    Exit Function
End Function

Public Function ValorTotal_Calcula(dValor As Double, Optional ByVal iTesta As Integer = 0) As Long

Dim lErro As Long
Dim dICMSSubstValor1 As Double
Dim dIPIValor1 As Double
Dim dValorProdutos As Double
Dim dValorTotal As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorDespesas As Double
Dim dValorIRRF As Double
Dim dValorAposIR As Double
Dim dValorISS As Double, dICMSValorIsento As Double
Dim dTotalAux As Double, objItem As Object
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ValorTotal_Calcula

    If Not mvarobjForm.gbCarregandoTela Then
        'Faz a atualização dos valores da tributação
        lErro = AtualizarTributacao()
        If lErro <> SUCESSO Then gError 42176
    End If

    'Recolhe os valores da tela
    If mvariNFEntrada = NF_SAIDA Then
        dValorProdutos = StrParaDbl(TabTrib_Controle("ValorProdutos").Caption)
    Else
        dValorProdutos = StrParaDbl(TabTrib_Controle("SubTotal").Caption)
    End If
    
    dICMSValorIsento = StrParaDbl(TabTrib_Controle("ICMSDesonerado").Caption)
    dICMSSubstValor1 = StrParaDbl(TabTrib_Controle("ICMSSubstValor1").Caption)
    dIPIValor1 = StrParaDbl(TabTrib_Controle("IPIValor1").Caption)
    dValorFrete = StrParaDbl(TabTrib_Controle("ValorFrete").Text)
    dValorSeguro = StrParaDbl(TabTrib_Controle("ValorSeguro").Text)
    dValorDespesas = StrParaDbl(TabTrib_Controle("ValorDespesas").Text)
    dValorIRRF = StrParaDbl(TabTrib_Controle("ValorIRRF").Text)
    If TabTrib_Controle("ISSIncluso").Value = vbUnchecked Then dValorISS = StrParaDbl(TabTrib_Controle("ISSValor").Caption)

    'Calcula o valor total
    dValorTotal = dValorProdutos + dICMSSubstValor1 + dIPIValor1 + dValorFrete + dValorSeguro + dValorDespesas + dValorISS - IIf(gobjCRFAT.iNFImportacaoTribFlag13 = 0 Or mvariEmitenteNF <> DOCINFO_EMPRESA, dICMSValorIsento, 0)
    
    If mvariNFEntrada = NF_ENTRADA Then
        'se for nota de importacao pode ter que incluir o valor do icms
        If TabTrib_Controle("TipoNFiscal").ListIndex <> -1 Then
            If TabTrib_Controle("TipoNFiscal").ItemData(TabTrib_Controle("TipoNFiscal").ListIndex) = DOCINFO_NFIEIMP Or TabTrib_Controle("TipoNFiscal").ItemData(TabTrib_Controle("TipoNFiscal").ListIndex) = DOCINFO_NFIEIMPSE Or TabTrib_Controle("TipoNFiscal").ItemData(TabTrib_Controle("TipoNFiscal").ListIndex) = DOCINFO_NFIEIMPPC Then
                If StrParaDbl(TabTrib_Controle("ICMSBase1").Caption) <> 0 And StrParaDbl(TabTrib_Controle("ICMSBase1").Caption) > dValorTotal And gobjCRFAT.iNFImportacaoTribFlag03 = DESMARCADO Then
                    dValorTotal = StrParaDbl(TabTrib_Controle("ICMSBase1").Caption)
                Else
                    If gobjCRFAT.iNFImportacaoTribFlag04 = DESMARCADO Then dValorTotal = dValorTotal + StrParaDbl(TabTrib_Controle("ICMSValor1").Caption)
                    If gobjCRFAT.iNFImportacaoTribFlag08 = MARCADO Then dValorTotal = dValorTotal - dValorDespesas
                    If Not (gcolDocTribItem Is Nothing) Then
                        For Each objItem In gcolDocTribItem
                            If gobjCRFAT.iNFImportacaoTribFlag05 = MARCADO Then dValorTotal = dValorTotal + objItem.dIIValor
                        Next
                    End If
                    If gobjCRFAT.iNFImportacaoTribFlag12 = MARCADO Then dValorTotal = dValorTotal + dPISValor + dCOFINSValor
                End If
            End If
        End If
        
        If iTesta = 1 Then
            dTotalAux = StrParaDbl(TabTrib_Controle("Total").Text)
            If Abs(dValorTotal - dTotalAux) > DELTA_VALORMONETARIO Then
                If Abs(dValorTotal - dICMSSubstValor1 - dTotalAux) <= DELTA_VALORMONETARIO Then
                    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_ICMS_ST_NAO_COBRADA")
                    If vbMsgRes = vbYes Then
                        dValorTotal = Arredonda_Moeda(dValorTotal - dICMSSubstValor1)
                    Else
                        gError 43786
                    End If
                Else
                    gError 43786
                End If
            End If
        End If
    Else
        dValorAposIR = dValorTotal - dValorIRRF
    
        If dValorAposIR < -DELTA_VALORMONETARIO And dValorIRRF > DELTA_VALORMONETARIO Then
        
            Call Rotina_Aviso(vbOKOnly, "AVISO_IR_FONTE_MAIOR_VALOR_TOTAL", dValorIRRF, dValorTotal)
            
            TabTrib_Controle("ValorIRRF").Text = ""
            Call ValorIRRF_Validate(bSGECancelDummy)
            
            'Faz a atualização dos valores da tributação
            lErro = AtualizarTributacao()
            If lErro <> SUCESSO Then gError 56884
        
        End If
        
    End If

    dValor = dValorTotal
    
    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr
    
        Case 42176, 56884

        Case 43786
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTAL_DIFERENTE", gErr, dTotalAux, dValorTotal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153760)

    End Select

    Exit Function
    
End Function

Public Function ValorPagar_Calcula(dValor As Double) As Long

Dim lErro As Long
Dim dValorPagar As Double
Dim dValorTotal As Double, dValorIRRF As Double, dValorINSSRetido As Double, dPISRetido As Double, dISSRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
Dim X As New ClassFatGlobal
Dim dPISSTValor As Double, dCOFINSSTValor As Double
Dim objTribItem As Object, dValorAux As Double, objItemNF As ClassItemNF
Dim bFat As Boolean, iItem As Integer, dValorDesconto As Double

On Error GoTo Erro_ValorPagar_Calcula

    'Recolhe os valores da tela
    dValorIRRF = StrParaDbl(TabTrib_Controle("ValorIRRF").Text)
    If TabTrib_Controle("INSSRetido").Value = vbChecked Then dValorINSSRetido = StrParaDbl(TabTrib_Controle("INSSValor").Text)
    dPISRetido = StrParaDbl(TabTrib_Controle("PISRetido").Text)
    dISSRetido = StrParaDbl(TabTrib_Controle("ISSRetido").Text)
    dCOFINSRetido = StrParaDbl(TabTrib_Controle("COFINSRetido").Text)
    dCSLLRetido = StrParaDbl(TabTrib_Controle("CSLLRetido").Text)
    
    If mvariNFEntrada = DESMARCADO Then
        dValorDesconto = StrParaDbl(TabTrib_Controle("ValorDescontoTit").Text)
    Else
        dValorDesconto = 0
    End If
    
    dCOFINSSTValor = 0
    dPISSTValor = 0
    dValorAux = 0
    iItem = 0
    For Each objTribItem In gcolDocTribItem
        
        iItem = iItem + 1
    
        dPISSTValor = dPISSTValor + gobjDocTribItemTributacao(objTribItem).dPISSTValor
        dCOFINSSTValor = dCOFINSSTValor + gobjDocTribItemTributacao(objTribItem).dCOFINSSTValor
        
        If iTipoTela = TIPOTELA_NF Then

            If Not (mvarobjForm.gobjNFiscal Is Nothing) Then
            
                If mvarobjForm.gobjNFiscal.colItens.Count >= iItem Then
                
                    Set objItemNF = mvarobjForm.gobjNFiscal.colItens.Item(iItem)
                    
                    If objItemNF.objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                    
                        lErro = Testa_Item_Fat(iItem, bFat)
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
                        If Not bFat Then
                            dValorAux = dValorAux + objTribItem.dPrecoTotal
                        End If
        
                    End If

                End If

            End If

        End If
        
    Next

    'Valor a Pagar
    If mvariNFEntrada = NF_SAIDA Then
        dValorTotal = StrParaDbl(TabTrib_Controle("ValorTotal").Caption)
    Else
        dValorTotal = StrParaDbl(TabTrib_Controle("Total").Text)
    End If
    If X.gobjFAT.iAbateICMSSubstVlrCobr = DESMARCADO Or mvariNFEntrada = NF_ENTRADA Then
        dValorPagar = dValorTotal - (dValorIRRF + dPISRetido + dISSRetido + dCOFINSRetido + dCSLLRetido) - dValorINSSRetido + dPISSTValor + dCOFINSSTValor
    Else
        dValorPagar = dValorTotal - (dValorIRRF + dPISRetido + dISSRetido + dCOFINSRetido + dCSLLRetido) - dValorINSSRetido - StrParaDbl(TabTrib_Controle("ICMSSubstValor1").Caption) + dPISSTValor + dCOFINSSTValor
    End If

    dValor = Arredonda_Moeda(dValorPagar - dValorAux, 2)
    
    If dValor - dValorDesconto > DELTA_VALORMONETARIO Then
        dValor = dValor - dValorDesconto
    Else
        dValor = 0
        If dValorDesconto > DELTA_VALORMONETARIO Then gError 209184
    End If
    
    ValorPagar_Calcula = SUCESSO

    Exit Function

Erro_ValorPagar_Calcula:

    ValorPagar_Calcula = gErr

    Select Case gErr
    
        Case 209184
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCTIT_MAIOR_VALORTIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153760)

    End Select

    Exit Function
    
End Function

Public Function Testa_Item_Fat(ByVal iItem As Integer, bFat As Boolean, Optional bMovEst As Boolean) As Long

Dim lErro As Long
Dim objProduto As ClassProduto
Dim colItens As Collection
Dim objTribItem As Object
Dim objTrib As Object
Dim bMexeu As Boolean
Dim sNatOp As String

On Error GoTo Erro_Testa_Item_Fat

    bMexeu = False
    bMovEst = True
    Set colItens = gcolDocTribItem
    Set objTribItem = colItens(iItem)

    Select Case gobjDocTribItemTributacao(objTribItem).sNaturezaOp
        Case "5901", "5902", "6901", "6902", "6925", "5925", "5910", "6910"
            bFat = False
        Case Else
            bFat = True 'Por padrão considera na fatura
    End Select
    
    If DOCINFO_NFEEFCNT = mvarobjNFTributacao.iTipoNFiscal Or DOCINFO_NFEEFCNT3 = mvarobjNFTributacao.iTipoNFiscal Then
        If Len(Trim(objTribItem.sProduto)) <> 0 Then
            Set objProduto = New ClassProduto
            objProduto.sCodigo = objTribItem.sProduto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then
                bFat = False
            End If
        End If
    End If
    
    'NF Conjulgada Ma Benef e Rem. Simbólica
    'Não cobra o que for Remessa Simbólica
    If DOCINFO_NFFBFRS = mvarobjNFTributacao.iTipoNFiscal Then
    
        Select Case gobjDocTribItemTributacao(objTribItem).sNaturezaOp
            Case "5901", "5902", "6901", "6902", "6925", "5925", "5910", "6910"
                bFat = False
                bMovEst = False
            Case Else
                'Gambiarra para evitar trabalho de ir trocando o CFOP,
                'Considera que o primeiro é o material benefiado e os demais
                'a remessa simbólica. Se for diferente aí sim troca na aba de tributação
                Set objTrib = gobjDocTribItemTributacao(objTribItem)
                If objTrib.iNaturezaOpManual <> VAR_PREENCH_MANUAL Then
                    If iItem > 1 Then
                        If gobjDocTribItemTributacao(colItens(1)).sNaturezaOp = "5124" Then
                            If objTrib.sNaturezaOp <> "5902" Then
                                objTrib.sNaturezaOp = "5902"
                                objTrib.iTipoTributacao = 32
                                bMexeu = True
                                bFat = False
                                bMovEst = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        ElseIf gobjDocTribItemTributacao(colItens(1)).sNaturezaOp = "6124" Then
                            If objTrib.sNaturezaOp <> "6902" Then
                                objTrib.sNaturezaOp = "6902"
                                objTrib.iTipoTributacao = 32
                                bMexeu = True
                                bFat = False
                                bMovEst = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        ElseIf gobjDocTribItemTributacao(colItens(1)).sNaturezaOp = "5125" Then
                            If objTrib.sNaturezaOp <> "5925" Then
                                objTrib.sNaturezaOp = "5925"
                                objTrib.iTipoTributacao = 32
                                bMexeu = True
                                bFat = False
                                bMovEst = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        ElseIf gobjDocTribItemTributacao(colItens(1)).sNaturezaOp = "6125" Then
                            If objTrib.sNaturezaOp <> "6925" Then
                                objTrib.sNaturezaOp = "6925"
                                objTrib.iTipoTributacao = 32
                                bMexeu = True
                                bFat = False
                                bMovEst = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        End If
                    End If
                    
                    If Not mvarobjForm.gbCarregandoTela And bMexeu Then
                        'Faz a atualização dos valores da tributação
                        lErro = AtualizarTributacao()
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    End If
                End If
            
        End Select
    End If
    
    If DOCINFO_NFEFBFRS = mvarobjNFTributacao.iTipoNFiscal Then
    
        Select Case gobjDocTribItemTributacao(objTribItem).sNaturezaOp
            Case "1902", "2902", "2925", "1925"
                bFat = False
            Case Else
                'Gambiarra para evitar trabalho de ir trocando o CFOP,
                'Considera que o primeiro é o material benefiado e os demais
                'a remessa simbólica. Se for diferente aí sim troca na aba de tributação
                Set objTrib = gobjDocTribItemTributacao(objTribItem)
                If objTrib.iNaturezaOpManual <> VAR_PREENCH_MANUAL Then
                    If iItem > 1 Then
                        If mvarobjForm.NaturezaOp.Text = "5124" Then
                            If objTrib.sNaturezaOp <> "1902" Then
                                objTrib.sNaturezaOp = "1902"
                                objTrib.iTipoTributacao = 22
                                bMexeu = True
                                bFat = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        ElseIf mvarobjForm.NaturezaOp.Text = "6124" Then
                            If objTrib.sNaturezaOp <> "2902" Then
                                objTrib.sNaturezaOp = "2902"
                                objTrib.iTipoTributacao = 22
                                bMexeu = True
                                bFat = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        ElseIf mvarobjForm.NaturezaOp.Text = "5125" Then
                            If objTrib.sNaturezaOp <> "1925" Then
                                objTrib.sNaturezaOp = "1925"
                                objTrib.iTipoTributacao = 22
                                bMexeu = True
                                bFat = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        ElseIf mvarobjForm.NaturezaOp.Text = "6125" Then
                            If objTrib.sNaturezaOp <> "2925" Then
                                objTrib.sNaturezaOp = "2925"
                                objTrib.iTipoTributacao = 22
                                bMexeu = True
                                bFat = False
                                objTrib.iNaturezaOpManual = VAR_PREENCH_MANUAL
                                objTrib.iTipoTributacaoManual = VAR_PREENCH_MANUAL
                            End If
                        End If
                    Else
                        If Len(Trim(mvarobjForm.NaturezaOp.Text)) > 0 Then
                            sNatOp = Format(CLng(mvarobjForm.NaturezaOp.Text) - 4000, "0000")
                        Else
                            sNatOp = ""
                        End If
                    
                        If objTrib.sNaturezaOp <> sNatOp Then
                        
                            TabTrib_Controle("NatOpInterna").Text = sNatOp
                            Call NatOpInterna_Validate(bSGECancelDummy)

                        End If
                    End If
                    
                    If Not mvarobjForm.gbCarregandoTela And bMexeu Then
                        'Faz a atualização dos valores da tributação
                        lErro = AtualizarTributacao()
                        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    End If
                End If
            
        End Select
        
    End If
        
    Testa_Item_Fat = SUCESSO

    Exit Function

Erro_Testa_Item_Fat:

    Testa_Item_Fat = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153760)

    End Select

    Exit Function

End Function

Public Function Valida_Dados() As Long

Dim lErro As Long
Dim dValorIRRF As Double
Dim dValorTotal As Double
Dim dValorINSSRetido As Double

Dim dValorProdL As Double
Dim dValorProdB As Double
Dim dValorDesconto As Double
Dim dValorProdLAux As Double
Dim dValorProdBAux As Double
Dim dValorDescontoAux As Double
Dim objItem As Object

On Error GoTo Erro_Valida_Dados

    If TabTrib_Controle("INSSRetido").Value = vbChecked Then
        dValorINSSRetido = StrParaDbl(TabTrib_Controle("INSSValor").Text)
    End If

    dValorIRRF = StrParaDbl(TabTrib_Controle("ValorIRRF").Text)
    
    If mvariNFEntrada = NF_SAIDA Then
        dValorTotal = StrParaDbl(TabTrib_Controle("ValorTotal").Caption)
    Else
        dValorTotal = StrParaDbl(TabTrib_Controle("Total").Text)
    End If
    
    If dValorTotal < (dValorIRRF + dValorINSSRetido) Then gError 56905
    
    If gcolDocTribItem.Count + gcolDocTribItem(, gbTemSRV).Count = 0 Then gError 205029
    
    If mvariNFEntrada = NF_SAIDA Then
        
        'Testa a soma dos itens
        dValorProdL = StrParaDbl(TabTrib_Controle("ValorProdutos").Caption)
        dValorProdB = StrParaDbl(TabTrib_Controle("ValorProdutos2").Caption)
        dValorDesconto = StrParaDbl(TabTrib_Controle("ValorDescontoItens").Text)
        
        For Each objItem In gcolDocItem(gobjDocTrib)
            dValorDescontoAux = Arredonda_Moeda(dValorDescontoAux + objItem.dValorDesconto)
            dValorProdBAux = Arredonda_Moeda(dValorProdBAux + objItem.dPrecoUnitario * IIf(objItem.dQuantidade = 0, 1, objItem.dQuantidade))
        Next
        For Each objItem In gcolDocItem(gobjDocTrib, True)
            dValorDescontoAux = Arredonda_Moeda(dValorDescontoAux + objItem.dValorDesconto)
            dValorProdBAux = Arredonda_Moeda(dValorProdBAux + objItem.dPrecoUnitario * IIf(objItem.dQuantidade = 0, 1, objItem.dQuantidade))
        Next
        dValorProdLAux = dValorProdBAux - dValorDescontoAux
        
        If Abs(dValorProdB - dValorProdBAux) > DELTA_VALORMONETARIO Then gError 216130
        If Abs(dValorProdL - dValorProdLAux) > DELTA_VALORMONETARIO Then gError 216131
        If Abs(dValorDesconto - dValorDescontoAux) > DELTA_VALORMONETARIO Then gError 216132
        
    End If
    
    Valida_Dados = SUCESSO

    Exit Function

Erro_Valida_Dados:

    Valida_Dados = gErr

    Select Case gErr

        Case 56905
            Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)

        Case 56906
        
        Case 205029
            Call Rotina_Erro(vbOKOnly, "ERRO_DOC_SEM_ITENS", gErr)
            
        Case 216130
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_PROD_BRUTO_INCONSISTENTE", gErr, dValorProdB, dValorProdBAux)

        Case 216131
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_PROD_LIQUIDO_INCONSISTENTE", gErr, dValorProdL, dValorProdLAux)

        Case 216132
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_TOTAL_DESCONTO_INCONSISTENTE", gErr, dValorDesconto, dValorDescontoAux)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153760)

    End Select

    Exit Function
    
End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
 
        Case ISS_VALOR
            objMnemonicoValor.colValor.Add StrParaDbl(TabTrib_Controle("ISSValor").Text)
        
        Case ISS_INCLUSO
            objMnemonicoValor.colValor.Add TabTrib_Controle("ISSIncluso").Value

        Case VALOR_IRRF
            objMnemonicoValor.colValor.Add StrParaDbl(TabTrib_Controle("ValorIRRF").Text)

        Case PIS_RETIDO
            objMnemonicoValor.colValor.Add StrParaDbl(TabTrib_Controle("PISRetido").Text)

        Case ISS_RETIDO
            objMnemonicoValor.colValor.Add StrParaDbl(TabTrib_Controle("ISSRetido").Text)

        Case COFINS_RETIDO
            objMnemonicoValor.colValor.Add StrParaDbl(TabTrib_Controle("COFINSRetido").Text)

        Case CSLL_RETIDO
            objMnemonicoValor.colValor.Add StrParaDbl(TabTrib_Controle("CSLLRetido").Text)

        Case VALOR_INSS
            objMnemonicoValor.colValor.Add StrParaDbl(TabTrib_Controle("INSSValor").Text)

        Case INSS_RETIDO
            objMnemonicoValor.colValor.Add TabTrib_Controle("INSSRetido").Value
            
        Case Else
        
            lErro = Calcula_Mnemonico_MN(objMnemonicoValor)
            If lErro = CONTABIL_MNEMONICO_NAO_ENCONTRADO Then gError 39675

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39675
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157258)

    End Select

    Exit Function

End Function

Function Preenche_Itens_Projeto_NF(ByVal colItensPRJ As Collection, Optional ByVal bSemDesconto As Boolean = False) As Long

On Error GoTo Erro_Preenche_Itens_Projeto_NF

    If iSubTipoTela = SUBTIPOTELA_CONSULTA Then
        Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "FR", "Frete", StrParaDbl(TabTrib_Controle("ValorFrete").Caption), 1, 0, "")
        Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "SG", "Seguro", StrParaDbl(TabTrib_Controle("ValorSeguro").Caption), 1, 0, "")
        Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "OD", "Outras Despesas", StrParaDbl(TabTrib_Controle("ValorDespesas").Caption), 1, 0, "")
    Else
        Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "FR", "Frete", StrParaDbl(TabTrib_Controle("ValorFrete").Text), 1, 0, "")
        Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "SG", "Seguro", StrParaDbl(TabTrib_Controle("ValorSeguro").Text), 1, 0, "")
        Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "OD", "Outras Despesas", StrParaDbl(TabTrib_Controle("ValorDespesas").Text), 1, 0, "")
    End If
    Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "ICMS", "ICMS Substituto", StrParaDbl(TabTrib_Controle("ICMSSubstValor1").Caption), 1, 0, "")
    Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "IPI", "IPI", StrParaDbl(TabTrib_Controle("IPIValor1").Caption), 1, 0, "")
    Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "ISS", "ISS", StrParaDbl(TabTrib_Controle("ISSValor").Caption), 1, 0, "")
    
    If iSubTipoTela = SUBTIPOTELA_CONSULTA Then
        If Not bSemDesconto Then Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "DESC", "Desconto", StrParaDbl(TabTrib_Controle("ValorDesconto").Caption), 1, 1, "")
    Else
        If Not bSemDesconto Then Call mvarobjForm.Preenche_Itens_Projeto_Aux(colItensPRJ, "DESC", "Desconto", StrParaDbl(TabTrib_Controle("ValorDesconto").Text), 1, 1, "")
    End If
    
    Preenche_Itens_Projeto_NF = SUCESSO

    Exit Function

Erro_Preenche_Itens_Projeto_NF:

    Preenche_Itens_Projeto_NF = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function

End Function

Public Property Get sNatOpInterna() As String
    sNatOpInterna = TabTrib_Controle("NatOpInterna").Text
End Property

Public Property Let sNatOpInterna(ByVal vData As String)
    TabTrib_Controle("NatOpInterna").Text = vData
End Property

Public Property Get dPISValor() As Double
    dPISValor = StrParaDbl(TabTrib_Controle("PISValor").Caption)
End Property

Public Property Get dCOFINSValor() As Double
    dCOFINSValor = StrParaDbl(TabTrib_Controle("COFINSValor").Caption)
End Property

Public Property Get gobjDocTrib() As Object
    If iTipoTela = TIPOTELA_NF Then
        Set gobjDocTrib = mvarobjNFTributacao
    ElseIf iTipoTela = TIPOTELA_PV Then
        Set gobjDocTrib = mvarobjPV
    ElseIf iTipoTela = TIPOTELA_OV Then
        Set gobjDocTrib = mvarobjOV
    Else
        Set gobjDocTrib = mvarobjDocTribOutros
    End If
End Property

Public Property Get gobjDocTrib2() As ClassTributoDoc
    If iTipoTela = TIPOTELA_NF Then
        Set gobjDocTrib2 = mvarobjNFTributacao
    ElseIf iTipoTela = TIPOTELA_PV Then
        Set gobjDocTrib2 = mvarobjPV
    ElseIf iTipoTela = TIPOTELA_OV Then
        Set gobjDocTrib2 = mvarobjOV
    Else
        Set gobjDocTrib2 = mvarobjDocTribOutros
    End If
End Property

Public Property Get gcolDocTribItem(Optional ByVal objAux As Object = Nothing, Optional ByVal bSRV As Boolean = False) As Collection
    If bSRV Then
        If Not (objAux Is Nothing) Then
            Set gcolDocTribItem = objAux.colItensSRV
        Else
            Set gcolDocTribItem = gobjDocTrib.colItensSRV
        End If
    Else
        If Not (objAux Is Nothing) Then
            Set gcolDocTribItem = objAux.colItens
        Else
            Set gcolDocTribItem = gobjDocTrib.colItens
        End If
    End If
End Property

Public Property Get gcolDocItem(ByVal objAux As Object, Optional ByVal bSRV As Boolean = False) As Object
    If bSRV Then
        Set gcolDocItem = objAux.colItensSRV
    Else
        Set gcolDocItem = objAux.colItens
    End If
End Property

Public Property Let gobjDocTrib(ByVal vData As Object)
    If iTipoTela = TIPOTELA_NF Then
        Set mvarobjNFTributacao = gobjDocTrib
    ElseIf iTipoTela = TIPOTELA_PV Then
        Set mvarobjPV = gobjDocTrib
    ElseIf iTipoTela = TIPOTELA_OV Then
        Set mvarobjOV = gobjDocTrib
    Else
        Set mvarobjDocTribOutros = gobjDocTrib
    End If
End Property

Public Property Get gobjDocTribTributacao(Optional ByVal objAux As Object = Nothing) As Object
    If iTipoTela = TIPOTELA_NF Then
        If Not (objAux Is Nothing) Then
            Set gobjDocTribTributacao = objAux.objTributacaoNF
        Else
            Set gobjDocTribTributacao = mvarobjNFTributacao
        End If
    Else
        If Not (objAux Is Nothing) Then
            Set gobjDocTribTributacao = objAux.objTributacao
        Else
            Set gobjDocTribTributacao = gobjDocTrib.objTributacao
        End If
    End If
End Property

Public Property Get gobjDocTribItemTributacao(ByVal objAux As Object) As Object
    If UCase(TypeName(objAux)) = "CLASSTRIBITEMNFAUX" Or UCase(TypeName(objAux)) = "CLASSTRIBNFAUX" Then
        Set gobjDocTribItemTributacao = objAux
    Else
        Set gobjDocTribItemTributacao = objAux.objTributacao
    End If
End Property

Public Property Get gobjDocTribItem(ByVal iIndice As Integer, Optional ByVal bSRV As Boolean = False) As Object
    If bSRV Then
        If iTipoTela = TIPOTELA_NF Then
            If iIndice <= gobjDocTrib.colItensSRV.Count Then
                Set gobjDocTribItem = mvarobjNFTributacao.colItensSRV.Item(iIndice)
            Else
                Set gobjDocTribItem = mvarobjNFTributacao.colItens.Item(iIndice - gobjDocTrib.colItensSRV.Count)
            End If
        Else
            If iIndice <= gobjDocTrib.colItensSRV.Count Then
                Set gobjDocTribItem = gobjDocTrib.colItensSRV.Item(iIndice).objTributacao
            Else
                Set gobjDocTribItem = gobjDocTrib.colItens.Item(iIndice - gobjDocTrib.colItensSRV.Count).objTributacao
            End If
        End If
    Else
        If iTipoTela = TIPOTELA_NF Then
            Set gobjDocTribItem = mvarobjNFTributacao.colItens.Item(iIndice)
        Else
            Set gobjDocTribItem = gobjDocTrib.colItens.Item(iIndice).objTributacao
        End If
    End If
End Property

Public Property Get dQuantidadeTela(ByVal iIndice As Integer, Optional ByVal bSRV As Boolean = False) As Double
    If bSRV Then
        If iTipoTela <> TIPOTELA_PV Then
            dQuantidadeTela = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_QuantidadeSRV_Col))
        Else
            dQuantidadeTela = StrParaDbl(TabTrib_Controle("GridItensSRV").TextMatrix(iIndice, mvarobjForm.iGrid_QuantPedidaSRV_Col)) - IIf(mvarobjForm.iGrid_QuantCancelSRV_Col <> 0, StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_QuantCancelSRV_Col)), 0)
        End If
    Else
        If iTipoTela <> TIPOTELA_PV Then
            dQuantidadeTela = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_Quantidade_Col))
        Else
            dQuantidadeTela = StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_QuantPedida_Col)) - IIf(mvarobjForm.iGrid_QuantCancel_Col <> 0, StrParaDbl(TabTrib_Controle("GridItens").TextMatrix(iIndice, mvarobjForm.iGrid_QuantCancel_Col)), 0)
        End If
    End If
End Property

Sub PedidoVenda_ObterTrib_Orcamento(objPedidoVenda As ClassPedidoDeVenda, objOrcamentoVenda As ClassOrcamentoVenda)

    Set objPedidoVenda.objTributacaoPV = New ClassTributacaoDoc
    
    Call objPedidoVenda.objTributacaoPV.Copia(objOrcamentoVenda.objTributacaoOV)
    
    objPedidoVenda.sNaturezaOp = objOrcamentoVenda.sNaturezaOp
    objPedidoVenda.iNaturezaOpManual = VAR_PREENCH_MANUAL 'objOrcamentoVenda.iNaturezaOpManual
    
End Sub

Sub ItemPedidoVenda_ObterTrib_ItemOrcamento(objItemPV As ClassItemPedido, objItemOV As ClassItemOV)

    Set objItemPV = New ClassItemPedido
    
    Call objItemPV.objTributacaoItemPV.Copia(objItemOV.objTributacaoItemOV)
    
End Sub

'##################################################################
'Novo código de Tributação NFe - 19/06/2009
'Procurar por XXXX ->Pontos a mexer fora daqui

Private Function Carrega_Outros_Controles() As Long

Dim lErro As Long, sCodigo As String
Dim objTipoTribPISCOFINS As ClassTipoTribPISCOFINS
Dim objTipoTribISS As ClassTipoTribISS
Dim objTipoTribICMSSimples As ClassTipoTribICMSSimples
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Carrega_Outros_Controles

    mvariFrameAtualDet = 1
    mvariFrameAtualRes = 1
    
    Set mvarcolTiposTribPISCOFINS = New Collection
    
    'Tipo de PIS e Tipo de COFINS
    lErro = CF("TiposTribPISCOFINS_Le_Todos", mvarcolTiposTribPISCOFINS)
    If lErro <> SUCESSO Then gError 202835

    TabTrib_Controle("ComboPISTipo").Clear
    TabTrib_Controle("ComboCOFINSTipo").Clear
    
    For Each objTipoTribPISCOFINS In mvarcolTiposTribPISCOFINS
        sCodigo = CStr(objTipoTribPISCOFINS.iTipo) & SEPARADOR & objTipoTribPISCOFINS.sDescricao
        If Not (mvarobjForm Is Nothing) Then
            TabTrib_Controle("ComboPISTipo").AddItem (sCodigo)
            TabTrib_Controle("ComboPISTipo").ItemData(TabTrib_Controle("ComboPISTipo").NewIndex) = objTipoTribPISCOFINS.iTipo
            TabTrib_Controle("ComboCOFINSTipo").AddItem (sCodigo)
            TabTrib_Controle("ComboCOFINSTipo").ItemData(TabTrib_Controle("ComboPISTipo").NewIndex) = objTipoTribPISCOFINS.iTipo
        End If
    Next
    
    'Origem da Mercadoria
    TabTrib_Controle("OrigemMercadoria").Clear
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_NACIONAL & SEPARADOR & PRODUTO_ORIGEM_NACIONAL_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_NACIONAL
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_IMP_DIRETA & SEPARADOR & PRODUTO_ORIGEM_IMP_DIRETA_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_IMP_DIRETA
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_EST_MERC_INTERNO & SEPARADOR & PRODUTO_ORIGEM_EST_MERC_INTERNO_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_EST_MERC_INTERNO
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_NAC_COM_IMP & SEPARADOR & PRODUTO_ORIGEM_NAC_COM_IMP_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_NAC_COM_IMP
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_NAC_PPB & SEPARADOR & PRODUTO_ORIGEM_NAC_PPB_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_NAC_PPB
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_NAC_POUCO_IMP & SEPARADOR & PRODUTO_ORIGEM_NAC_POUCO_IMP_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_NAC_POUCO_IMP
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_IMP_DIR_SEM_SIMILAR & SEPARADOR & PRODUTO_ORIGEM_IMP_DIR_SEM_SIMILAR_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_IMP_DIR_SEM_SIMILAR
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_IMP_MERC_INT_SEM_SIMILAR & SEPARADOR & PRODUTO_ORIGEM_IMP_MERC_INT_SEM_SIMILAR_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_IMP_MERC_INT_SEM_SIMILAR
    
    TabTrib_Controle("OrigemMercadoria").AddItem PRODUTO_ORIGEM_NAC_COM_IMP2 & SEPARADOR & PRODUTO_ORIGEM_NAC_COM_IMP2_TEXTO
    TabTrib_Controle("OrigemMercadoria").ItemData(TabTrib_Controle("OrigemMercadoria").NewIndex) = PRODUTO_ORIGEM_NAC_COM_IMP2
    
    'Modalidade do ICMS
    TabTrib_Controle("ICMSModalidade").Clear
    
    TabTrib_Controle("ICMSModalidade").AddItem ICMS_MODALIDADE_MARGEM & SEPARADOR & ICMS_MODALIDADE_MARGEM_TEXTO
    TabTrib_Controle("ICMSModalidade").ItemData(TabTrib_Controle("ICMSModalidade").NewIndex) = ICMS_MODALIDADE_MARGEM

    TabTrib_Controle("ICMSModalidade").AddItem ICMS_MODALIDADE_PAUTA & SEPARADOR & ICMS_MODALIDADE_PAUTA_TEXTO
    TabTrib_Controle("ICMSModalidade").ItemData(TabTrib_Controle("ICMSModalidade").NewIndex) = ICMS_MODALIDADE_PAUTA

    TabTrib_Controle("ICMSModalidade").AddItem ICMS_MODALIDADE_TABELA & SEPARADOR & ICMS_MODALIDADE_TABELA_TEXTO
    TabTrib_Controle("ICMSModalidade").ItemData(TabTrib_Controle("ICMSModalidade").NewIndex) = ICMS_MODALIDADE_TABELA

    TabTrib_Controle("ICMSModalidade").AddItem ICMS_MODALIDADE_VALOR & SEPARADOR & ICMS_MODALIDADE_VALOR_TEXTO
    TabTrib_Controle("ICMSModalidade").ItemData(TabTrib_Controle("ICMSModalidade").NewIndex) = ICMS_MODALIDADE_VALOR

    'Modalidade do ICMSST
    TabTrib_Controle("ICMSSTModalidade").Clear

    TabTrib_Controle("ICMSSTModalidade").AddItem ICMS_ST_MODALIDADE_TABELA & SEPARADOR & ICMS_ST_MODALIDADE_TABELA_TEXTO
    TabTrib_Controle("ICMSSTModalidade").ItemData(TabTrib_Controle("ICMSSTModalidade").NewIndex) = ICMS_ST_MODALIDADE_TABELA

    TabTrib_Controle("ICMSSTModalidade").AddItem ICMS_ST_MODALIDADE_LISTA_NEG & SEPARADOR & ICMS_ST_MODALIDADE_LISTA_NEG_TEXTO
    TabTrib_Controle("ICMSSTModalidade").ItemData(TabTrib_Controle("ICMSSTModalidade").NewIndex) = ICMS_ST_MODALIDADE_LISTA_NEG

    TabTrib_Controle("ICMSSTModalidade").AddItem ICMS_ST_MODALIDADE_LISTA_POS & SEPARADOR & ICMS_ST_MODALIDADE_LISTA_POS_TEXTO
    TabTrib_Controle("ICMSSTModalidade").ItemData(TabTrib_Controle("ICMSSTModalidade").NewIndex) = ICMS_ST_MODALIDADE_LISTA_POS

    TabTrib_Controle("ICMSSTModalidade").AddItem ICMS_ST_MODALIDADE_LISTA_NEU & SEPARADOR & ICMS_ST_MODALIDADE_LISTA_NEU_TEXTO
    TabTrib_Controle("ICMSSTModalidade").ItemData(TabTrib_Controle("ICMSSTModalidade").NewIndex) = ICMS_ST_MODALIDADE_LISTA_NEU

    TabTrib_Controle("ICMSSTModalidade").AddItem ICMS_ST_MODALIDADE_MARGEM & SEPARADOR & ICMS_ST_MODALIDADE_MARGEM_TEXTO
    TabTrib_Controle("ICMSSTModalidade").ItemData(TabTrib_Controle("ICMSSTModalidade").NewIndex) = ICMS_ST_MODALIDADE_MARGEM

    TabTrib_Controle("ICMSSTModalidade").AddItem ICMS_ST_MODALIDADE_PAUTA & SEPARADOR & ICMS_ST_MODALIDADE_PAUTA_TEXTO
    TabTrib_Controle("ICMSSTModalidade").ItemData(TabTrib_Controle("ICMSSTModalidade").NewIndex) = ICMS_ST_MODALIDADE_PAUTA

    'Tipo de Cálculo do IPI
    TabTrib_Controle("IPITipoCalculo").Clear

    TabTrib_Controle("IPITipoCalculo").AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    TabTrib_Controle("IPITipoCalculo").ItemData(TabTrib_Controle("IPITipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_VALOR

    TabTrib_Controle("IPITipoCalculo").AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    TabTrib_Controle("IPITipoCalculo").ItemData(TabTrib_Controle("IPITipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    'Tipo de Cálculo do PIS
    TabTrib_Controle("PISTipoCalculo").Clear

    TabTrib_Controle("PISTipoCalculo").AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    TabTrib_Controle("PISTipoCalculo").ItemData(TabTrib_Controle("PISTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_VALOR

    TabTrib_Controle("PISTipoCalculo").AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    TabTrib_Controle("PISTipoCalculo").ItemData(TabTrib_Controle("PISTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    'Tipo de Cálculo do PISST
    TabTrib_Controle("PISSTTipoCalculo").Clear

    TabTrib_Controle("PISSTTipoCalculo").AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    TabTrib_Controle("PISSTTipoCalculo").ItemData(TabTrib_Controle("PISSTTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_VALOR

    TabTrib_Controle("PISSTTipoCalculo").AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    TabTrib_Controle("PISSTTipoCalculo").ItemData(TabTrib_Controle("PISSTTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    'Tipo de Cálculo do COFINS
    TabTrib_Controle("COFINSTipoCalculo").Clear

    TabTrib_Controle("COFINSTipoCalculo").AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    TabTrib_Controle("COFINSTipoCalculo").ItemData(TabTrib_Controle("COFINSTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_VALOR

    TabTrib_Controle("COFINSTipoCalculo").AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    TabTrib_Controle("COFINSTipoCalculo").ItemData(TabTrib_Controle("COFINSTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    'Tipo de Cálculo do COFINSST
    TabTrib_Controle("COFINSSTTipoCalculo").Clear

    TabTrib_Controle("COFINSSTTipoCalculo").AddItem TRIB_TIPO_CALCULO_VALOR & SEPARADOR & TRIB_TIPO_CALCULO_VALOR_TEXTO
    TabTrib_Controle("COFINSSTTipoCalculo").ItemData(TabTrib_Controle("COFINSSTTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_VALOR

    TabTrib_Controle("COFINSSTTipoCalculo").AddItem TRIB_TIPO_CALCULO_PERCENTUAL & SEPARADOR & TRIB_TIPO_CALCULO_PERCENTUAL_TEXTO
    TabTrib_Controle("COFINSSTTipoCalculo").ItemData(TabTrib_Controle("COFINSSTTipoCalculo").NewIndex) = TRIB_TIPO_CALCULO_PERCENTUAL

    'Lista de Serviços
    lErro = CF("Carrega_Combo", TabTrib_Controle("ISSListaServ"), "ISSQN", "Codigo", TIPO_STR, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then gError 202836
       
    Call ComboPISTipo_Click
    
    Call ComboCOFINSTipo_Click
    
    lErro = CF("TiposTribISS_Le_Todos", mvarcolTiposTribISS)
    If lErro <> SUCESSO Then gError 202835

    lErro = CF("TiposTribICMSSimples_Le_Todos", mvarcolTiposTribICMSSimples)
    If lErro <> SUCESSO Then gError 202835
    
    TabTrib_Controle("ComboISSTipo").Clear
    TabTrib_Controle("ComboICMSSimplesTipo").Clear
    
    For Each objTipoTribISS In mvarcolTiposTribISS
        sCodigo = objTipoTribISS.sDescricao
        If Not (mvarobjForm Is Nothing) Then
            TabTrib_Controle("ComboISSTipo").AddItem (sCodigo)
            TabTrib_Controle("ComboISSTipo").ItemData(TabTrib_Controle("ComboISSTipo").NewIndex) = objTipoTribISS.iTipo
        End If
    Next
    
    For Each objTipoTribICMSSimples In mvarcolTiposTribICMSSimples
        sCodigo = Format(objTipoTribICMSSimples.iCSOSN, "000") & SEPARADOR & objTipoTribICMSSimples.sDescricao
        If Not (mvarobjForm Is Nothing) Then
            TabTrib_Controle("ComboICMSSimplesTipo").AddItem (sCodigo)
            TabTrib_Controle("ComboICMSSimplesTipo").ItemData(TabTrib_Controle("ComboICMSSimplesTipo").NewIndex) = objTipoTribICMSSimples.iTipo
        End If
    Next
    
    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    TabTrib_Controle("ICMSUFDevidoST").Clear
    
    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    For Each vCodigo In colCodigo
        TabTrib_Controle("ICMSUFDevidoST").AddItem vCodigo
    Next
    
    lErro = CF("Carrega_Combo", TabTrib_Controle("ICMSMotivo"), "TribICMSMotivos", "Codigo", TIPO_INT, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then gError 202836

    lErro = CF("Carrega_Combo", TabTrib_Controle("NatBCCred"), "NatBCCred", "Codigo", TIPO_STR, "Descricao", TIPO_STR)
    If lErro <> SUCESSO Then gError 202836
    
    TabTrib_Controle("NatBCCred").AddItem " "
    TabTrib_Controle("NatBCCred").ItemData(TabTrib_Controle("NatBCCred").NewIndex) = 0
    
    'Lê cada código e descrição da tabela Países
    lErro = CF("Cod_Nomes_Le", "Paises", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche cada ComboBox País com os objetos da coleção colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        If objCodigoDescricao.iCodigo <> PAIS_BRASIL Then

            TabTrib_Controle("ISSCodPaisItem").AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
            TabTrib_Controle("ISSCodPaisItem").ItemData(TabTrib_Controle("ISSCodPaisItem").NewIndex) = objCodigoDescricao.iCodigo
            
        Else
        
            TabTrib_Controle("ISSCodPaisItem").AddItem ""
            TabTrib_Controle("ISSCodPaisItem").ItemData(TabTrib_Controle("ISSCodPaisItem").NewIndex) = 0
        
        End If

    Next
    
    Carrega_Outros_Controles = SUCESSO

    Exit Function

Erro_Carrega_Outros_Controles:

    Carrega_Outros_Controles = gErr

    Select Case gErr

        Case 202835, 202836, ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202837)

    End Select

    Exit Function

End Function

Public Sub OpcaoTribDet_Click()

Dim lErro As Long, objAux As Object

On Error GoTo Erro_OpcaoTribDet_Click

    'Se frame selecionado não for o atual
    If TabTrib_Controle("OpcaoTribDet").SelectedItem.Index <> mvariFrameAtualDet Then

        If TabStrip_PodeTrocarTab(mvariFrameAtualDet, TabTrib_Controle("OpcaoTribDet"), mvarobjForm) <> SUCESSO Then Exit Sub
        
        'Esconde o frame atual, mostra o novo
        Set objAux = TabTrib_Controle("FrameTabTribDet")
        objAux(TabTrib_Controle("OpcaoTribDet").SelectedItem.Index).Visible = True
        objAux(mvariFrameAtualDet).Visible = False
        'Armazena novo valor de mvariFrameAtual
        mvariFrameAtualDet = TabTrib_Controle("OpcaoTribDet").SelectedItem.Index
    
    End If

    Exit Sub

Erro_OpcaoTribDet_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202834)

    End Select

    Exit Sub

End Sub

Public Sub IPITipoCalculo_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    Call TipoCalculo_Click(Codigo_Extrai(TabTrib_Controle("IPITipoCalculo")), "IPI", bAtualizaTrib)
End Sub

Public Sub PISTipoCalculo_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    Call TipoCalculo_Click(Codigo_Extrai(TabTrib_Controle("PISTipoCalculo")), "PIS", bAtualizaTrib)
End Sub

Public Sub PISSTTipoCalculo_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    Call TipoCalculo_Click(Codigo_Extrai(TabTrib_Controle("PISSTTipoCalculo")), "PISST", bAtualizaTrib)
End Sub

Public Sub COFINSTipoCalculo_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    Call TipoCalculo_Click(Codigo_Extrai(TabTrib_Controle("COFINSTipoCalculo")), "COFINS", bAtualizaTrib)
End Sub

Public Sub COFINSSTTipoCalculo_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    Call TipoCalculo_Click(Codigo_Extrai(TabTrib_Controle("COFINSSTTipoCalculo")), "COFINSST", bAtualizaTrib)
End Sub

Private Sub TipoCalculo_Click(ByVal iTipo As Integer, ByVal sImposto As String, Optional ByVal bAtualizaTrib As Boolean = True)

On Error GoTo Erro_TipoCalculo_Click

    '-2 = Respeita o que estã no tipo
    If iTipo = -2 Then
        If Len(Trim(TabTrib_Controle(sImposto & "TipoCalculo"))) > 0 Then
            iTipo = Codigo_Extrai(TabTrib_Controle(sImposto & "TipoCalculo"))
        Else
            iTipo = -1
        End If
    End If

    Select Case iTipo
   
        Case -1
            TabTrib_Controle("Frame" & sImposto & "TipoCalculoBase").Enabled = False
            TabTrib_Controle(sImposto & "TipoCalculo").ListIndex = -1
            TabTrib_Controle("Frame" & sImposto & "TipoCalculoQtd").Enabled = False
            TabTrib_Controle(sImposto & "ValorItem").Enabled = False
        
        Case TRIB_TIPO_CALCULO_VALOR
            TabTrib_Controle("Frame" & sImposto & "TipoCalculoBase").Enabled = False
            TabTrib_Controle("Frame" & sImposto & "TipoCalculoQtd").Enabled = True
            TabTrib_Controle(sImposto & "ValorItem").Enabled = True
        
        Case TRIB_TIPO_CALCULO_PERCENTUAL
            TabTrib_Controle("Frame" & sImposto & "TipoCalculoBase").Enabled = True
            TabTrib_Controle("Frame" & sImposto & "TipoCalculoQtd").Enabled = False
            TabTrib_Controle(sImposto & "ValorItem").Enabled = True
    
    End Select
    
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
    
    Exit Sub

Erro_TipoCalculo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202838)

    End Select

    Exit Sub
    
End Sub

Public Sub ComboPISTipo_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    Call ComboPISCOFINSTipo_Click(Codigo_Extrai(TabTrib_Controle("ComboPISTipo")), "PIS", bAtualizaTrib)
End Sub

Public Sub ComboCOFINSTipo_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    Call ComboPISCOFINSTipo_Click(Codigo_Extrai(TabTrib_Controle("ComboCOFINSTipo")), "COFINS", bAtualizaTrib)
End Sub

Private Sub ComboPISCOFINSTipo_Click(ByVal iTipo As Integer, ByVal sImposto As String, Optional ByVal bAtualizaTrib As Boolean = True)

Dim objTipoTribPISCOFINS As ClassTipoTribPISCOFINS

On Error GoTo Erro_ComboPISCOFINSTipo_Click

    For Each objTipoTribPISCOFINS In mvarcolTiposTribPISCOFINS
        If objTipoTribPISCOFINS.iTipo = iTipo Then Exit For
    Next
    If Not (objTipoTribPISCOFINS Is Nothing) Then
    
        Select Case objTipoTribPISCOFINS.iTipoCalculo
            
            Case TIPO_TRIB_TIPO_CALCULO_DESABILITADO
                Call TipoCalculo_Click(-1, sImposto, bAtualizaTrib)
                TabTrib_Controle(sImposto & "TipoCalculo").Enabled = False
    
            Case TIPO_TRIB_TIPO_CALCULO_PERCENTUAL
                Call TipoCalculo_Click(TRIB_TIPO_CALCULO_PERCENTUAL, sImposto, bAtualizaTrib)
                TabTrib_Controle(sImposto & "TipoCalculo").Enabled = False
    
            Case TIPO_TRIB_TIPO_CALCULO_VALOR
                Call TipoCalculo_Click(TRIB_TIPO_CALCULO_VALOR, sImposto, bAtualizaTrib)
                TabTrib_Controle(sImposto & "TipoCalculo").Enabled = False
        
            Case TIPO_TRIB_TIPO_CALCULO_ESCOLHA, TIPO_TRIB_TIPO_CALCULO_ESCOLHA_MANUAL
                Call TipoCalculo_Click(-2, sImposto, bAtualizaTrib)
                TabTrib_Controle(sImposto & "TipoCalculo").Enabled = True
        
        End Select
        
    End If
        
    Exit Sub

Erro_ComboPISCOFINSTipo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202839)

    End Select

    Exit Sub
    
End Sub

Sub ISSValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariISSValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ISSValorItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSValorItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariISSValorItemAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ISSAliquotaItem_Change()
    mvariISSAliquotaItemAlterado = 1
End Sub

Sub ISSAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSAliquotaItem_Validate

    If mvariISSAliquotaItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("ISSAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ISSAliquotaItem").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariISSAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSAliquotaItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub ISSBaseItem_Change()

    mvariISSBaseItemAlterado = 1
    
End Sub

Sub ISSBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSBaseItem_Validate

    If mvariISSBaseItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("ISSBaseItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSBaseItem").Text)
            If lErro <> SUCESSO Then gError 56875
        
        End If
    
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariISSBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSBaseItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56875

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153750)

    End Select

    Exit Sub

End Sub

Sub ISSValorItem_Change()
    mvariISSValorItemAlterado = 1
End Sub

Sub PISValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISValor_Validate

    If mvariPISValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISValorItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISValorItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariPISValorItemAlterado = 0

    Exit Sub
    
Erro_PISValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub PISAliquotaItem_Change()
    mvariPISAliquotaItemAlterado = 1
End Sub

Sub PISAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISAliquotaItem_Validate

    If mvariPISAliquotaItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("PISAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("PISAliquotaItem").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariPISAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_PISAliquotaItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub PISBaseItem_Change()

    mvariPISBaseItemAlterado = 1
    
End Sub

Sub PISBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISBaseItem_Validate

    If mvariPISBaseItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("PISBaseItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISBaseItem").Text)
            If lErro <> SUCESSO Then gError 56875
        
        End If
    
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariPISBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_PISBaseItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56875

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153750)

    End Select

    Exit Sub

End Sub

Sub PISValorItem_Change()
    mvariPISValorItemAlterado = 1
End Sub

Sub PISAliquotaRSItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISAliquotaRS_Validate

    If mvariPISAliquotaValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISAliquotaRSItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISAliquotaRSItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariPISAliquotaValorItemAlterado = 0

    Exit Sub
    
Erro_PISAliquotaRS_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub PISAliquotaRSItem_Change()
    mvariPISAliquotaValorItemAlterado = 1
End Sub

Sub PISQtdeItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISQtde_Validate

    If mvariPISQtdeItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISQtdeItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISQtdeItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariPISQtdeItemAlterado = 0

    Exit Sub
    
Erro_PISQtde_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub PISQtdeItem_Change()
    mvariPISQtdeItemAlterado = 1
End Sub

Sub PISSTValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISSTValor_Validate

    If mvariPISSTValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISSTValorItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISSTValorItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariPISSTValorItemAlterado = 0

    Exit Sub
    
Erro_PISSTValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub PISSTAliquotaItem_Change()
    mvariPISSTAliquotaItemAlterado = 1
End Sub

Sub PISSTAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISSTAliquotaItem_Validate

    If mvariPISSTAliquotaItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("PISSTAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("PISSTAliquotaItem").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariPISSTAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_PISSTAliquotaItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub PISSTBaseItem_Change()

    mvariPISSTBaseItemAlterado = 1
    
End Sub

Sub PISSTBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISSTBaseItem_Validate

    If mvariPISSTBaseItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("PISSTBaseItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISSTBaseItem").Text)
            If lErro <> SUCESSO Then gError 56875
        
        End If
    
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariPISSTBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_PISSTBaseItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56875

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153750)

    End Select

    Exit Sub

End Sub

Sub PISSTValorItem_Change()
    mvariPISSTValorItemAlterado = 1
End Sub

Sub PISSTAliquotaRSItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISSTAliquotaRS_Validate

    If mvariPISSTAliquotaValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISSTAliquotaRSItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISSTAliquotaRSItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariPISSTAliquotaValorItemAlterado = 0

    Exit Sub
    
Erro_PISSTAliquotaRS_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub PISSTAliquotaRSItem_Change()
    mvariPISSTAliquotaValorItemAlterado = 1
End Sub

Sub PISSTQtdeItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISSTQtde_Validate

    If mvariPISSTQtdeItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("PISSTQtdeItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("PISSTQtdeItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariPISSTQtdeItemAlterado = 0

    Exit Sub
    
Erro_PISSTQtde_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub PISSTQtdeItem_Change()
    mvariPISSTQtdeItemAlterado = 1
End Sub

Sub COFINSValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSValor_Validate

    If mvariCOFINSValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSValorItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSValorItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariCOFINSValorItemAlterado = 0

    Exit Sub
    
Erro_COFINSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub COFINSAliquotaItem_Change()
    mvariCOFINSAliquotaItemAlterado = 1
End Sub

Sub COFINSAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSAliquotaItem_Validate

    If mvariCOFINSAliquotaItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("COFINSAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("COFINSAliquotaItem").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariCOFINSAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_COFINSAliquotaItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub COFINSBaseItem_Change()

    mvariCOFINSBaseItemAlterado = 1
    
End Sub

Sub COFINSBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSBaseItem_Validate

    If mvariCOFINSBaseItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("COFINSBaseItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSBaseItem").Text)
            If lErro <> SUCESSO Then gError 56875
        
        End If
    
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariCOFINSBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_COFINSBaseItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56875

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153750)

    End Select

    Exit Sub

End Sub

Sub COFINSValorItem_Change()
    mvariCOFINSValorItemAlterado = 1
End Sub

Sub COFINSAliquotaRSItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSAliquotaRS_Validate

    If mvariCOFINSAliquotaValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSAliquotaRSItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSAliquotaRSItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariCOFINSAliquotaValorItemAlterado = 0

    Exit Sub
    
Erro_COFINSAliquotaRS_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub COFINSAliquotaRSItem_Change()
    mvariCOFINSAliquotaValorItemAlterado = 1
End Sub

Sub COFINSQtdeItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSQtdeItem_Validate

    If mvariCOFINSQtdeItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSQtdeItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSQtdeItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariCOFINSQtdeItemAlterado = 0

    Exit Sub
    
Erro_COFINSQtdeItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub COFINSQtdeItem_Change()
    mvariCOFINSQtdeItemAlterado = 1
End Sub

Sub COFINSSTValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSSTValorItem_Validate

    If mvariCOFINSSTValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSSTValorItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSSTValorItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariCOFINSSTValorItemAlterado = 0

    Exit Sub
    
Erro_COFINSSTValorItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub COFINSSTAliquotaItem_Change()
    mvariCOFINSSTAliquotaItemAlterado = 1
End Sub

Sub COFINSSTAliquotaItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSSTAliquotaItem_Validate

    If mvariCOFINSSTAliquotaItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("COFINSSTAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("COFINSSTAliquotaItem").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariCOFINSSTAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_COFINSSTAliquotaItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub COFINSSTBaseItem_Change()

    mvariCOFINSSTBaseItemAlterado = 1
    
End Sub

Sub COFINSSTBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSSTBaseItem_Validate

    If mvariCOFINSSTBaseItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("COFINSSTBaseItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSSTBaseItem").Text)
            If lErro <> SUCESSO Then gError 56875
        
        End If
    
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariCOFINSSTBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_COFINSSTBaseItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56875

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153750)

    End Select

    Exit Sub

End Sub

Sub COFINSSTValorItem_Change()
    mvariCOFINSSTValorItemAlterado = 1
End Sub

Sub COFINSSTAliquotaRSItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSSTAliquotaRS_Validate

    If mvariCOFINSSTAliquotaValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSSTAliquotaRSItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSSTAliquotaRSItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariCOFINSSTAliquotaValorItemAlterado = 0

    Exit Sub
    
Erro_COFINSSTAliquotaRS_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub COFINSSTAliquotaRSItem_Change()
    mvariCOFINSSTAliquotaValorItemAlterado = 1
End Sub

Sub COFINSSTQtdeItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSSTQtde_Validate

    If mvariCOFINSSTQtdeItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("COFINSSTQtdeItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("COFINSSTQtdeItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariCOFINSSTQtdeItemAlterado = 0

    Exit Sub
    
Erro_COFINSSTQtde_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub COFINSSTQtdeItem_Change()
    mvariCOFINSSTQtdeItemAlterado = 1
End Sub

Sub ISSCodIBGE_Change()
    mvariISSCidadeIBGEItemAlterado = 1
End Sub

Sub ISSCodIBGE_GotFocus()
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("ISSCodIBGE"), mvariISSCidadeIBGEItemAlterado)
End Sub

Sub ISSCodIBGE_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)

Dim lErro As Long
Dim sDescIBGE As String

On Error GoTo Erro_ISSCodIBGE_Validate

    If Len(Trim(TabTrib_Controle("ISSCodIBGE").Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(TabTrib_Controle("ISSCodIBGE").Text)
        If lErro <> SUCESSO Then gError 57975
        
        lErro = CF("Le_Campo_Tabela", "IBGEMunicipios", "DescIBGE", TIPO_STR, "CodIBGE", TabTrib_Controle("ISSCodIBGE").Text, sDescIBGE, TabTrib_Controle("ISSDescIBGE"))
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 200981
        
        If lErro <> SUCESSO Then gError 200982
    Else
    
        TabTrib_Controle("ISSDescIBGE").Caption = ""
        
    End If
    
    If bAtualizaTrib And mvariISSCidadeIBGEItemAlterado = 1 And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
    
    mvariISSCidadeIBGEItemAlterado = 0
    
    Exit Sub
    
Erro_ISSCodIBGE_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 57975, 200981 'Erro tratado na rotina chamada
        
        Case 200982
            Call Rotina_Erro(vbOKOnly, "ERRO_MUNIC_CODIBGE_NAO_CADASTRADO", gErr, TabTrib_Controle("ISSCodIBGE.Text"))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 164299)
    
    End Select
    
    Exit Sub

End Sub

Sub ISSLabelCodIBGE_Click()

Dim colSelecao As New Collection
Dim objIBGEMunicipios As New ClassIBGEMunicipios

On Error GoTo Erro_ISSLabelCodIBGE_Click
    
    'Chama Tela TituloReceberLista
    Call Chama_Tela("IBGEMunicipiosLista", colSelecao, objIBGEMunicipios, mvarobjEventoCidadeIBGE)

    Exit Sub

Erro_ISSLabelCodIBGE_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200973)

    End Select

    Exit Sub

End Sub

Private Sub mvarobjEventoCidadeIBGE_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objIBGEMunicipios As ClassIBGEMunicipios

On Error GoTo Erro_mvarobjEventoCidadeIBGE_evSelecao

    Set objIBGEMunicipios = obj1
    
    TabTrib_Controle("ISSCodIBGE").PromptInclude = False
    TabTrib_Controle("ISSCodIBGE").Text = objIBGEMunicipios.sCodUF & SEPARADOR & objIBGEMunicipios.sCodMunic
    TabTrib_Controle("ISSCodIBGE").PromptInclude = True
    Call ISSCodIBGE_Validate(bSGECancelDummy)
    
    Exit Sub

Erro_mvarobjEventoCidadeIBGE_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200974)

    End Select

    Exit Sub

End Sub

Sub IIBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IIBaseItem_Validate

    If mvariIIBaseItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IIBaseItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IIBaseItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariIIBaseItemAlterado = 0

    Exit Sub
    
Erro_IIBaseItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub IIBaseItem_Change()
    mvariIIBaseItemAlterado = 1
End Sub

Sub IIValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IIValorItem_Validate

    If mvariIIValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IIValorItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IIValorItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariIIValorItemAlterado = 0

    Exit Sub
    
Erro_IIValorItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub IIValorItem_Change()
    mvariIIValorItemAlterado = 1
End Sub

Sub IIIOFItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IIIOFItem_Validate

    If mvariIIIOFItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IIIOFItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IIIOFItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariIIIOFItemAlterado = 0

    Exit Sub
    
Erro_IIIOFItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub IIIOFItem_Change()
    mvariIIIOFItemAlterado = 1
End Sub

Sub IIDespAduItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IIDespAduItem_Validate

    If mvariIIDespAduaneiraItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IIDespAduItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IIDespAduItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariIIDespAduaneiraItemAlterado = 0

    Exit Sub
    
Erro_IIDespAduItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub IIDespAduItem_Change()
    mvariIIDespAduaneiraItemAlterado = 1
End Sub

Sub ValorFreteItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorFreteItem_Validate

    If mvariValorFreteItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ValorFreteItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ValorFreteItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariValorFreteItemAlterado = 0

    Exit Sub
    
Erro_ValorFreteItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ValorFreteItem_Change()
    mvariValorFreteItemAlterado = 1
End Sub

Sub ValorSeguroItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorSeguroItem_Validate

    If mvariValorSeguroItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ValorSeguroItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ValorSeguroItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariValorSeguroItemAlterado = 0

    Exit Sub
    
Erro_ValorSeguroItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ValorSeguroItem_Change()
    mvariValorSeguroItemAlterado = 1
End Sub

Sub ValorDescontoItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDescontoItem_Validate

    If mvariValorDescontoItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ValorDescontoItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ValorDescontoItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariValorDescontoItemAlterado = 0

    Exit Sub
    
Erro_ValorDescontoItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ValorDescontoItem_Change()
    mvariValorDescontoItemAlterado = 1
End Sub

Sub ValorDespesasItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorDespesasItem_Validate

    If mvariValorOutrasDespesasItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ValorDespesasItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ValorDespesasItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariValorOutrasDespesasItemAlterado = 0

    Exit Sub
    
Erro_ValorDespesasItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ValorDespesasItem_Change()
    mvariValorOutrasDespesasItemAlterado = 1
End Sub

Sub ICMSModalidade_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub

Sub ICMSSTModalidade_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub

Sub OrigemMercadoria_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If TabTrib_Controle("OrigemMercadoria").ListIndex = -1 Then
        TabTrib_Controle("OrigemMercadoria").ToolTipText = ""
    Else
        TabTrib_Controle("OrigemMercadoria").ToolTipText = TabTrib_Controle("OrigemMercadoria").Text
    End If
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub

Sub FCI_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer, sChar As String

On Error GoTo Erro_FCI_Validate

    If mvariFCIAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("FCI").Text)) > 0 Then
        
            If Len(Trim(TabTrib_Controle("FCI").Text)) <> 36 Then gError 201573
            
            For iIndice = 1 To Len(Trim(TabTrib_Controle("FCI").Text))
                sChar = Mid(TabTrib_Controle("FCI").Text, iIndice, 1)
                If sChar <> "0" And sChar <> "1" And sChar <> "2" And sChar <> "3" And _
                sChar <> "4" And sChar <> "5" And sChar <> "6" And sChar <> "7" And _
                sChar <> "8" And sChar <> "9" And sChar <> "A" And sChar <> "B" And _
                sChar <> "C" And sChar <> "D" And sChar <> "E" And sChar <> "F" And _
                sChar <> "-" Then gError 201573
            Next
        
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariFCIAlterado = 0
    
    End If

    Exit Sub
    
Erro_FCI_Validate:

    Cancel = True
    
    Select Case gErr

        Case 201573
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMATO_FCI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201572)

    End Select

    Exit Sub

End Sub

Sub FCI_Change()
    mvariFCIAlterado = 1
End Sub

Sub ICMSSubstPercRedBaseItem_Change()
    mvariICMSSubstPercRedBaseItemAlterado = 1
End Sub

Sub ICMSSubstPercRedBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstPercRedBaseItem_Validate

    If mvariICMSSubstPercRedBaseItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("ICMSSubstPercRedBaseItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSSubstPercRedBaseItem").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariICMSSubstPercRedBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSubstPercRedBaseItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub ICMSSubstPercMVAItem_Change()
    mvariICMSSubstPercMVAItemAlterado = 1
End Sub

Sub ICMSSubstPercMVAItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSubstPercMVAItem_Validate

    If mvariICMSSubstPercMVAItemAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("ICMSSubstPercMVAItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSSubstPercMVAItem").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariICMSSubstPercMVAItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSubstPercMVAItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub IPICodEnq_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)
Dim lErro As Long
On Error GoTo Erro_IPICodEnq_Validate
    lErro = IPICodEnq_Valida
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
    Exit Sub
Erro_IPICodEnq_Validate:
    Cancel = True
End Sub

Sub IPIClasseEnq_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)
    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Sub IPICodSelo_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)
    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Sub IPISeloQtde_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)
    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Sub IPICNPJProdutor_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)

Dim lErro As Long

On Error GoTo Erro_IPICNPJProdutor_Validate

    If Len(Trim(TabTrib_Controle("IPICNPJProdutor").Text)) = 0 Then Exit Sub

    Select Case Len(Trim(TabTrib_Controle("IPICNPJProdutor").Text))
    
        Case STRING_CGC  'CGC
    
            lErro = Cgc_Critica(TabTrib_Controle("IPICNPJProdutor").Text)
            If lErro <> SUCESSO Then gError 12781
    
            TabTrib_Controle("IPICNPJProdutor").Format = "00\.000\.000\/0000-00; ; ; "
            TabTrib_Controle("IPICNPJProdutor").Text = TabTrib_Controle("IPICNPJProdutor").Text
    
    Case Else

        gError 12782

    End Select

    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem

    Exit Sub

Erro_IPICNPJProdutor_Validate:

    Cancel = True

    Select Case gErr

        Case 12780, 12781

        Case 12782
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160598)

    End Select

    Exit Sub

End Sub

Sub IPIUnidadePadraoValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIUnidadePadraoValorItem_Validate

    If mvariIPIUnidadePadraoValorItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IPIUnidadePadraoValorItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IPIUnidadePadraoValorItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariIPIUnidadePadraoValorItemAlterado = 0

    Exit Sub
    
Erro_IPIUnidadePadraoValorItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub IPIUnidadePadraoValorItem_Change()
    mvariIPIUnidadePadraoValorItemAlterado = 1
End Sub

Sub IPIUnidadePadraoQtdeItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIUnidadePadraoQtdeItem_Validate

    If mvariIPIUnidadePadraoQtdeItemAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IPIUnidadePadraoQtdeItem").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IPIUnidadePadraoQtdeItem").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariIPIUnidadePadraoQtdeItemAlterado = 0

    Exit Sub
    
Erro_IPIUnidadePadraoQtdeItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub IPIUnidadePadraoQtdeItem_Change()
    mvariIPIUnidadePadraoQtdeItemAlterado = 1
End Sub

Sub ISSListaServ_Click()
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Sub IRBase_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IRBase_Validate

    If mvariIRBaseAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("IRBase").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IRBase").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    Call AtualizarTabTrib
        
    mvariIRBaseAlterado = 0

    Exit Sub
    
Erro_IRBase_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub IRBase_Change()
    mvariIRBaseAlterado = 1
End Sub

Sub ICMSSTCobrAntBaseItem_Change()

    mvariICMSSTCobrAntBaseItemAlterado = 1
    
End Sub

Sub ICMSSTCobrAntBaseItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSTCobrAntBaseItem_Validate

    If mvariICMSSTCobrAntBaseItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSSTCobrAntBaseItem").Text)) > 0 Then
    
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSSTCobrAntBaseItem").Text)
            If lErro <> SUCESSO Then Error 56872
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSSTCobrAntBaseItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSTCobrAntBaseItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56872
''???            If TabTrib_Controle("ICMSSTCobrAntBaseItem").Visible Then TabTrib_Controle("ICMSSTCobrAntBaseItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153746)

    End Select

    Exit Sub

End Sub

Sub ICMSSTCobrAntValorItem_Change()

    mvariICMSSTCobrAntValorItemAlterado = 1

End Sub

Sub ICMSSTCobrAntValorItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSSTCobrAntValorItem_Validate

    If mvariICMSSTCobrAntValorItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSSTCobrAntValorItem").Text)) > 0 Then
    
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSSTCobrAntValorItem").Text)
            If lErro <> SUCESSO Then Error 56873
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSSTCobrAntValorItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSTCobrAntValorItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56873
''???            If TabTrib_Controle("ICMSSTCobrAntValorItem").Visible Then TabTrib_Controle("ICMSSTCobrAntValorItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153747)

    End Select

    Exit Sub

End Sub
'##################################################################

Public Function AtualizarMsg() As Long

Dim lErro As Long
Dim objMsg As ClassMensagensRegra
Dim colMsg As New Collection
Dim sMsg(0 To 1) As String
Dim objTransacao As Object

On Error GoTo Erro_AtualizarMsg

    Set objTransacao = Me

    lErro = CF("RegrasMsg_Calcula_Regras", objTransacao, REGRAMSG_TIPODOC_NF, colMsg)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objMsg In colMsg
        sMsg(objMsg.iTipoMsg) = sMsg(objMsg.iTipoMsg) & IIf(Len(Trim(sMsg(objMsg.iTipoMsg))) = 0, "", vbNewLine) & objMsg.sMensagem
    Next
    
    TabTrib_Controle("Mensagem").Text = sMsg(REGRAMSG_TIPOMSG_NORMAL)
    TabTrib_Controle("MensagemCorpo").Text = sMsg(REGRAMSG_TIPOMSG_CORPO)
    
    AtualizarMsg = SUCESSO

    Exit Function
    
Erro_AtualizarMsg:

    AtualizarMsg = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 208270)

    End Select

    Exit Function
    
End Function

Public Function Inicializa_Calculo_RegrasMsg() As Long
    Set gcolRetRegras = New Collection
End Function

Public Function Inicializa_Calculo_RegrasMsg2(ByVal colRegras As Collection) As Long
    Set gcolRetRegras = New Collection
    Set gcolRegras = colRegras
End Function

Public Function Retorno_Calculo_RegrasMsg(ByVal bFlag As Boolean) As Long
    gcolRetRegras.Add bFlag
End Function

Public Function Finaliza_Calculo_RegrasMsg() As Long
    Set gcolRetRegras = Nothing
    Set gcolRegras = Nothing
End Function

Public Function Calcula_Mnemonico_MN(objMnemonicoValor As ClassMnemonicoValor) As Long
'************
'IMPORTANTE
'************
'Qualquer alteração nessa rotina deverá ser replicada em ClassGeracaoNFiscal

Dim lErro As Long, bAchou As Boolean
Dim iLinha As Integer, sCat As String, sCatItem As String
Dim objTributoDocItem As ClassTributoDocItem, objTributoDocItemAux As Object
Dim objProduto As ClassProduto, objProdCat As ClassProdutoCategoria
Dim objCli As ClassCliente, objFilCli As ClassFilialCliente, objTransp As ClassTransportadora
Dim objForn As ClassFornecedor, objFilForn As ClassFilialFornecedor
Dim objEnd As ClassEndereco, objPais As ClassPais, sAux As String, dValorBase As Double
Dim dAcumSemSubst As Double, dAcumComSubst As Double, dValorNaoFat As Double, dAliquota As Double, dValor As Double
Dim objDASAliquotas As ClassDASAliquotas, objPV As ClassPedidoDeVenda, objMensagem As ClassMensagem
Dim colCatItem As Collection, objFilCliCat As ClassFilialCliCategoria, objFilFornCat As ClassFilialFornCategoria
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objRegra As ClassRegrasMsg, objRegraAux As ClassRegrasMsg, iRegra As Integer
Dim objNFiscalOriginal As ClassNFiscal, lNumIntNotaOriginal As Long, iItem As Integer
Dim X As ClassFAT, dValorBaseAux As Double, iAliqDif As Integer
Dim sCGCTranspRedesp As String, bPedSrv As Boolean, sListaNFOrigDet As String
Dim colListaNFOrigDet As Collection, objOV As ClassOrcamentoVenda
Dim objItemNF As ClassItemNF, objItemNFOrig As ClassItemNF

On Error GoTo Erro_Calcula_Mnemonico_MN

    Select Case objMnemonicoValor.sMnemonico
 
        Case MN_EXISTE_PRODUTO
            bAchou = False
            iLinha = 0
            For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, False)
                iLinha = iLinha + 1
                Set objTributoDocItemAux = gobjDocTribItem(iLinha, False)
                If UCase(objTributoDocItemAux.sProduto) = UCase(objMnemonicoValor.vParam(1)) Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, True)
                    iLinha = iLinha + 1
                    Set objTributoDocItemAux = gobjDocTribItem(iLinha, True)
                    If UCase(objTributoDocItemAux.sProduto) = UCase(objMnemonicoValor.vParam(1)) Then
                        bAchou = True
                        Exit For
                    End If
                Next
            End If
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_ENTRADA
            If mvariNFEntrada = MARCADO Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_SAIDA
            If mvariNFEntrada = MARCADO Then
                objMnemonicoValor.colValor.Add False
            Else
                objMnemonicoValor.colValor.Add True
            End If
            
        Case MN_TIPOTRIB
            objMnemonicoValor.colValor.Add gobjDocTribTributacao.iTipoTributacao

        Case MN_TOTTRIB
            If gobjDocTribTributacao.dTotTrib > DELTA_VALORMONETARIO Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTrib, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIB
            If gobjDocTribTributacao.dTotTrib > DELTA_VALORMONETARIO And StrParaDbl(TabTrib_Controle("ValorTotal").Caption) > 0 Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTrib / StrParaDbl(TabTrib_Controle("ValorTotal").Caption), "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBFED
            If gobjDocTribTributacao.dTotTribFed > DELTA_VALORMONETARIO Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTribFed, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIBFED
            If gobjDocTribTributacao.dTotTribFed > DELTA_VALORMONETARIO And StrParaDbl(TabTrib_Controle("ValorTotal").Caption) > 0 Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTribFed / StrParaDbl(TabTrib_Controle("ValorTotal").Caption), "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBEST
            If gobjDocTribTributacao.dTotTribEst > DELTA_VALORMONETARIO Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTribEst, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIBEST
            If gobjDocTribTributacao.dTotTribEst > DELTA_VALORMONETARIO And StrParaDbl(TabTrib_Controle("ValorTotal").Caption) > 0 Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTribEst / StrParaDbl(TabTrib_Controle("ValorTotal").Caption), "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TOTTRIBMUNIC
            If gobjDocTribTributacao.dTotTribMunic > DELTA_VALORMONETARIO Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTribMunic, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_PERCTOTTRIBMUNIC
            If gobjDocTribTributacao.dTotTribMunic > DELTA_VALORMONETARIO And StrParaDbl(TabTrib_Controle("ValorTotal").Caption) > 0 Then
                objMnemonicoValor.colValor.Add Format(gobjDocTribTributacao.dTotTribMunic / StrParaDbl(TabTrib_Controle("ValorTotal").Caption), "PERCENT")
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_TIPONF
            objMnemonicoValor.colValor.Add mvarobjNFTributacao.iTipoNFiscal
            
        Case MN_NUMPEDIDOVENDA
            
            If mvarobjNFTributacao.iTipoNFiscal <> 0 Then
            
                objTipoDocInfo.iCodigo = mvarobjNFTributacao.iTipoNFiscal
            
                'Lê o Tipo de Documento
                lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
                'Se achou o Tipo de Documento
                If lErro = SUCESSO Then
                
                    Select Case objTipoDocInfo.sNomeTelaNFiscal
                    
                        Case "NFiscalFaturaPedido", "NFiscalPedido", "NFiscalRemPedido"
                            objMnemonicoValor.colValor.Add StrParaLong(TabTrib_Controle("Pedido").Text)
                            
                        Case Else
                            objMnemonicoValor.colValor.Add 0
                            
                    End Select
            
                Else
                
                    objMnemonicoValor.colValor.Add 0
                    
                End If
            
            Else
            
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case MN_REGRA
            iLinha = objMnemonicoValor.vParam(1)
'            If iLinha > gcolRetRegras.Count Then gError 208272 'Não pode referenciar uma regra posterior
'            objMnemonicoValor.colValor.Add gcolRetRegras.Item(iLinha)
            iRegra = gcolRetRegras.Count + 1
            bAchou = False
            iItem = 0
            Set objRegra = gcolRegras.Item(iRegra)
            For Each objRegraAux In gcolRegras
                iItem = iItem + 1
                If iItem = iRegra Then Exit For
                If objRegraAux.iTipoDoc = objRegra.iTipoDoc And objRegraAux.iTipoMsg = objRegra.iTipoMsg And objRegraAux.iSeq = iLinha Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then gError 208272
            objMnemonicoValor.colValor.Add gcolRetRegras.Item(iItem)
            
        Case MN_FILEMPCGC
            objMnemonicoValor.colValor.Add objFilialEmpresa.sCgc
            
        Case MN_FILEMP
            objMnemonicoValor.colValor.Add objFilialEmpresa.iCodFilial
            
        Case MN_EXISTE_CATPRODITEM
            bAchou = False
            iLinha = 0
            sCat = objMnemonicoValor.vParam(1)
            sCatItem = objMnemonicoValor.vParam(2)
            For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, False)
                iLinha = iLinha + 1
                Set objTributoDocItemAux = gobjDocTribItem(iLinha, False)
                
                Set objProduto = New ClassProduto
                
                objProduto.sCodigo = objTributoDocItemAux.sProduto
                
                'Carrega a colCategoriaItem do objProduto
                lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For Each objProdCat In objProduto.colCategoriaItem
                    If UCase(objProdCat.sCategoria) = UCase(sCat) And UCase(objProdCat.sItem) = UCase(sCatItem) Then
                        bAchou = True
                        Exit For
                    End If
                Next
            Next
            If Not bAchou Then
                For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, True)
                    iLinha = iLinha + 1
                    Set objTributoDocItemAux = gobjDocTribItem(iLinha, True)
                    
                    Set objProduto = New ClassProduto
                    
                    objProduto.sCodigo = objTributoDocItemAux.sProduto
                    
                    'Carrega a colCategoriaItem do objProduto
                    lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    For Each objProdCat In objProduto.colCategoriaItem
                        If UCase(objProdCat.sCategoria) = UCase(sCat) And UCase(objProdCat.sItem) = UCase(sCatItem) Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                Next
            End If
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_EXISTE_CATPROD
            bAchou = False
            iLinha = 0
            sCat = objMnemonicoValor.vParam(1)
            For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, False)
                iLinha = iLinha + 1
                Set objTributoDocItemAux = gobjDocTribItem(iLinha, False)
                
                Set objProduto = New ClassProduto
                
                objProduto.sCodigo = objTributoDocItemAux.sProduto
                
                'Carrega a colCategoriaItem do objProduto
                lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For Each objProdCat In objProduto.colCategoriaItem
                    If UCase(objProdCat.sCategoria) = UCase(sCat) Then
                        bAchou = True
                        Exit For
                    End If
                Next
            Next
            If Not bAchou Then
                For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, True)
                    iLinha = iLinha + 1
                    Set objTributoDocItemAux = gobjDocTribItem(iLinha, True)
                    
                    Set objProduto = New ClassProduto
                    
                    objProduto.sCodigo = objTributoDocItemAux.sProduto
                    
                    'Carrega a colCategoriaItem do objProduto
                    lErro = CF("Produto_Le_Categorias", objProduto, objProduto.colCategoriaItem)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    For Each objProdCat In objProduto.colCategoriaItem
                        If UCase(objProdCat.sCategoria) = UCase(sCat) Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                Next
            End If
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_EXISTE_TIPOPROD
            bAchou = False
            iLinha = 0
            For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, False)
                iLinha = iLinha + 1
                Set objTributoDocItemAux = gobjDocTribItem(iLinha, False)
                
                Set objProduto = New ClassProduto
                
                objProduto.sCodigo = objTributoDocItemAux.sProduto

                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                
                If objProduto.iTipo = objMnemonicoValor.vParam(1) Then
                    bAchou = True
                    Exit For
                End If
                
            Next
            If Not bAchou Then
                For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, True)
                    iLinha = iLinha + 1
                    Set objTributoDocItemAux = gobjDocTribItem(iLinha, True)
                    
                    Set objProduto = New ClassProduto
                    
                    objProduto.sCodigo = objTributoDocItemAux.sProduto
    
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
                    
                    If objProduto.iTipo = objMnemonicoValor.vParam(1) Then
                        bAchou = True
                        Exit For
                    End If
                Next
            End If
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_CODBACENDEST, MN_UFDEST, MN_ENDREFDEST
            
            Set objEnd = New ClassEndereco
            
            If mvarobjNFTributacao.lCliente > 0 And mvarobjNFTributacao.iFilialCli > 0 Then
            
                Set objFilCli = New ClassFilialCliente
                
                objFilCli.lCodCliente = mvarobjNFTributacao.lCliente
                objFilCli.iCodFilial = mvarobjNFTributacao.iFilialCli
            
                lErro = CF("FilialCliente_Le", objFilCli)
                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
                            
                objEnd.lCodigo = objFilCli.lEndereco
            
            ElseIf mvarobjNFTributacao.lFornecedor > 0 And mvarobjNFTributacao.iFilialForn > 0 Then
            
                Set objFilForn = New ClassFilialFornecedor
                
                objFilForn.lCodFornecedor = mvarobjNFTributacao.lFornecedor
                objFilForn.iCodFilial = mvarobjNFTributacao.iFilialForn
            
                lErro = CF("FilialFornecedor_Le", objFilForn)
                If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
                        
                objEnd.lCodigo = objFilForn.lEndereco
            
            End If
                
            If objEnd.lCodigo > 0 Then
            
                lErro = CF("Endereco_Le", objEnd)
                If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoValor.sMnemonico = MN_UFDEST Then
                    objMnemonicoValor.colValor.Add objEnd.sSiglaEstado
                ElseIf objMnemonicoValor.sMnemonico = MN_ENDREFDEST Then
                    objMnemonicoValor.colValor.Add objEnd.sReferencia
                Else
                    Set objPais = New ClassPais
                    
                    objPais.iCodigo = objEnd.iCodigoPais
                    
                    lErro = CF("Paises_Le", objPais)
                    If lErro <> SUCESSO And lErro <> 47876 Then gError ERRO_SEM_MENSAGEM
                    
                    objMnemonicoValor.colValor.Add objPais.iCodBacen
                End If
            
            Else
            
                Select Case objMnemonicoValor.sMnemonico
                
                    Case MN_CODBACENDEST
                        objMnemonicoValor.colValor.Add 1058 'codbacen do Brasil
                    
                    Case MN_UFDEST
                        objMnemonicoValor.colValor.Add "SP"
                    
                    Case MN_ENDREFDEST
                        objMnemonicoValor.colValor.Add ""
                    
                End Select
            
            End If
            
        Case MN_BASECREDICMS, MN_ALIQUOTACREDICMS, MN_VALORCREDICMS
        
            bAchou = False
            iLinha = 0
            dValorBase = 0
            dValorBaseAux = 0
            dAliquota = 0
            iAliqDif = 0
            dValor = 0
            
            For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, False)
                iLinha = iLinha + 1
                Set objTributoDocItemAux = gobjDocTribItem(iLinha, False)
                'Acomula base e valor do itens com crédito
                If objTributoDocItemAux.dICMSvCredSN > 0 Then
                    'Se utilizar o total do item como base deixar os valores inconsitentes utiliza o valor calculado
                    If Abs((objTributoDocItemAux.dPrecoTotal * objTributoDocItemAux.dICMSpCredSN) - objTributoDocItemAux.dICMSvCredSN) > DELTA_VALORMONETARIO Then
                        dValorBaseAux = Arredonda_Moeda(objTributoDocItemAux.dICMSvCredSN / objTributoDocItemAux.dICMSpCredSN)
                    Else
                        dValorBaseAux = objTributoDocItemAux.dPrecoTotal
                    End If
                    dValorBase = dValorBase + dValorBaseAux
                    dValor = dValor + objTributoDocItemAux.dICMSvCredSN
                    If dAliquota = 0 Then
                        dAliquota = objTributoDocItemAux.dICMSpCredSN
                    ElseIf Abs(dAliquota - objTributoDocItemAux.dICMSpCredSN) > DELTA_VALORMONETARIO2 Then
                    'Se existe mais de uma alíquota vai exibir a média
                        iAliqDif = MARCADO
                    End If
                End If
            Next
                       
            If objMnemonicoValor.sMnemonico = MN_BASECREDICMS Then
                objMnemonicoValor.colValor.Add Format(dValorBase, "STANDARD")
            ElseIf objMnemonicoValor.sMnemonico = MN_ALIQUOTACREDICMS Then
                If iAliqDif = MARCADO Then
                    objMnemonicoValor.colValor.Add Format(dValor / dValorBase, "PERCENT")
                Else
                    objMnemonicoValor.colValor.Add Format(dAliquota, "PERCENT")
                End If
            Else
                objMnemonicoValor.colValor.Add Format(dValor, "STANDARD")
            End If
        
        Case MN_MSGPV
        
            sAux = ""
            If mvariNFEntrada <> MARCADO Then
            
                If Not (mvarobjForm.gobjNFiscal Is Nothing) Then
            
                    If mvarobjForm.gobjNFiscal.lNumPedidoVenda <> 0 Then
                    
                        bPedSrv = False
                        If mvarobjNFTributacao.iTipoNFiscal <> 0 Then
                            objTipoDocInfo.iCodigo = mvarobjNFTributacao.iTipoNFiscal
                            'Lê o Tipo de Documento
                            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
                            'Se achou o Tipo de Documento
                            If lErro = SUCESSO Then If objTipoDocInfo.sNomeTelaNFiscal = "NFiscalFatPedSRV" Then bPedSrv = True
                        End If
                                                
                        Set objPV = New ClassPedidoDeVenda
                    
                        objPV.lCodigo = mvarobjForm.gobjNFiscal.lNumPedidoVenda
                        objPV.iFilialEmpresa = mvarobjForm.gobjNFiscal.iFilialPedido
                    
                        If Not bPedSrv Then
                            
                            lErro = CF("PedidoDeVenda_Le", objPV)
                            If lErro <> SUCESSO And lErro <> 26509 Then gError ERRO_SEM_MENSAGEM
                            
                            If lErro <> SUCESSO Then
                            
                                lErro = CF("PedidoVendaBaixado_Le", objPV)
                                If lErro <> SUCESSO And lErro <> 46135 Then gError ERRO_SEM_MENSAGEM
                            
                            End If
                            
                        Else
                        
                            lErro = CF("PedidoServico_Le_Todos_Completo", objPV)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
                        End If
                    
                        sAux = objPV.sMensagemNota
                        
                    ElseIf mvarobjForm.gobjNFiscal.lNumOrcamentoVenda <> 0 Then
                    
                        bPedSrv = False
                        If mvarobjNFTributacao.iTipoNFiscal <> 0 Then
                            objTipoDocInfo.iCodigo = mvarobjNFTributacao.iTipoNFiscal
                            'Lê o Tipo de Documento
                            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
                            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
                            'Se achou o Tipo de Documento
                            If lErro = SUCESSO Then If objTipoDocInfo.sNomeTelaNFiscal = "NFiscalSRV" Then bPedSrv = True
                        End If
                                                
                        Set objOV = New ClassOrcamentoVenda
                    
                        objOV.lCodigo = mvarobjForm.gobjNFiscal.lNumOrcamentoVenda
                        objOV.iFilialEmpresa = mvarobjForm.gobjNFiscal.iFilialOrcamento
                    
                        If Not bPedSrv Then
                            
                            lErro = CF("OrcamentoVenda_Le_Completo", objOV)
                            If lErro <> SUCESSO And lErro <> 101234 Then gError ERRO_SEM_MENSAGEM
                           
                        Else
                        
                            lErro = CF("OrcamentoSRV_Le_Completo", objOV)
                            If lErro <> SUCESSO And lErro <> 193343 Then gError ERRO_SEM_MENSAGEM
                        
                        End If
                    
                        sAux = objOV.sMensagemNota
                    
                    End If
                    
                End If
                
            End If
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_MSGFILCLI
            
            sAux = ""
            If mvarobjNFTributacao.lCliente > 0 And mvarobjNFTributacao.iFilialCli > 0 Then
            
                Set objFilCli = New ClassFilialCliente
                
                objFilCli.lCodCliente = mvarobjNFTributacao.lCliente
                objFilCli.iCodFilial = mvarobjNFTributacao.iFilialCli
            
                lErro = CF("FilialCliente_Le", objFilCli)
                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
 
                If objFilCli.iCodMensagem > 0 Then
                
                    Set objMensagem = New ClassMensagem
                
                    objMensagem.iCodigo = objFilCli.iCodMensagem
                    
                    lErro = CF("Mensagem_Le", objMensagem)
                    If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM
            
                    If lErro = SUCESSO Then sAux = objMensagem.sDescricao
                    
                Else
                
                    Set objCli = New ClassCliente
                    
                    objCli.lCodigo = mvarobjNFTributacao.lCliente
            
                    lErro = CF("Cliente_Le", objCli)
                    If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                   
                   If objCli.iCodMensagem > 0 Then
                   
                       Set objMensagem = New ClassMensagem
                   
                       objMensagem.iCodigo = objCli.iCodMensagem
                       
                       lErro = CF("Mensagem_Le", objMensagem)
                       If lErro <> SUCESSO And lErro <> 19234 Then gError ERRO_SEM_MENSAGEM
               
                       If lErro = SUCESSO Then sAux = objMensagem.sDescricao
                       
                    End If
            
                End If

            End If
            
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_ENDENTREGA
            
            sAux = ""
            
            'If mvarobjNFTributacao.lCliente > 0 And mvarobjNFTributacao.iFilialCli > 0 Then
            
'                Set objFilCli = New ClassFilialCliente
'                Set objEnd = New ClassEndereco
'
'                objFilCli.lCodCliente = mvarobjNFTributacao.lCliente
'                objFilCli.iCodFilial = mvarobjNFTributacao.iFilialCli
'
'                lErro = CF("FilialCliente_Le", objFilCli)
'                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM

'               objEnd.lCodigo = objFilCli.lEnderecoEntrega
'
'               lErro = CF("Endereco_Le", objEnd)
'               If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
'
'               If Len(Trim(objEnd.sEndereco)) > 0 Then
'                   sAux = objEnd.sEndereco & " - " & objEnd.sBairro & " - " & objEnd.sCidade & " - " & objEnd.sSiglaEstado & " - " & Format(objEnd.sCEP, "#####-###")
'               End If
                   
                If Not (gobjInfoAdic Is Nothing) Then
                    If Not (gobjInfoAdic.objRetEnt Is Nothing) Then
                        If Not (gobjInfoAdic.objRetEnt.objEnderecoEnt Is Nothing) Then
                            Set objEnd = gobjInfoAdic.objRetEnt.objEnderecoEnt
                            If Len(Trim(objEnd.sEndereco)) = 0 Then
                                objEnd.sEndereco = objEnd.sTipoLogradouro
                                objEnd.sEndereco = objEnd.sEndereco & IIf(Len(Trim(objEnd.sEndereco)) <> 0, " ", "") & objEnd.sLogradouro
                                objEnd.sEndereco = objEnd.sEndereco & IIf(objEnd.lNumero <> 0, ", " & CStr(objEnd.lNumero), "")
                                objEnd.sEndereco = objEnd.sEndereco & IIf(Len(Trim(objEnd.sComplemento)) <> 0, " - " & objEnd.sComplemento, "")
                            End If
                            If Len(Trim(objEnd.sEndereco)) > 0 Then
                                sAux = objEnd.sEndereco & " - " & objEnd.sBairro & " - " & objEnd.sCidade & " - " & objEnd.sSiglaEstado & " - " & Format(objEnd.sCEP, "#####-###")
                            End If
                        End If
                    End If
                End If

            'End If
            
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_ENDTRANSREDESP
        
            sAux = ""
            If mvariNFEntrada <> MARCADO Then
                If Codigo_Extrai(TabTrib_Controle("TranspRedespacho")) <> 0 Then
           
                    Set objTransp = New ClassTransportadora
                    Set objEnd = New ClassEndereco
           
                    objTransp.iCodigo = Codigo_Extrai(TabTrib_Controle("TranspRedespacho"))
               
                    lErro = CF("Transportadora_Le", objTransp)
                    If lErro <> SUCESSO And lErro <> 19250 Then gError ERRO_SEM_MENSAGEM
               
                    objEnd.lCodigo = objTransp.lEndereco
    
                    lErro = CF("Endereco_Le", objEnd)
                    If lErro <> SUCESSO And lErro <> 12309 Then gError ERRO_SEM_MENSAGEM
                   
                    Select Case Len(Trim(objTransp.sCgc))
                   
                        Case STRING_CPF
                            sCGCTranspRedesp = " CPF: " & Format(objTransp.sCgc, "000\.000\.000-00; ; ; ")
                        Case STRING_CGC
                            sCGCTranspRedesp = " CNPJ: " & Format(objTransp.sCgc, "00\.000\.000\/0000-00; ; ; ")
                        Case Else
                            sCGCTranspRedesp = ""
                    End Select
               
                    If Len(Trim(objEnd.sEndereco)) > 0 Then
                        sAux = objTransp.sNomeReduzido & sCGCTranspRedesp & " - " & objEnd.sTelefone1 & " - " & objEnd.sEndereco & " - " & objEnd.sBairro & " - " & objEnd.sCidade & " - " & objEnd.sSiglaEstado & " - " & Format(objEnd.sCEP, "#####-###")
                    End If
                   
                End If
            End If
            objMnemonicoValor.colValor.Add sAux
                
        Case MN_FILCLIFORNINSCSUF, MN_FILCLIFORNINSCEST, MN_FILCLIFORNINSCMUN

            sAux = ""
            If mvarobjNFTributacao.lCliente > 0 And mvarobjNFTributacao.iFilialCli > 0 Then
            
                Set objFilCli = New ClassFilialCliente
                
                objFilCli.lCodCliente = mvarobjNFTributacao.lCliente
                objFilCli.iCodFilial = mvarobjNFTributacao.iFilialCli
            
                lErro = CF("FilialCliente_Le", objFilCli)
                If lErro <> SUCESSO And lErro <> 12567 Then gError ERRO_SEM_MENSAGEM
                            
                If objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCSUF Then
                    If Len(Trim(objFilCli.sInscricaoSuframa)) <> 0 Then sAux = Format(objFilCli.sInscricaoSuframa, "00\.0000\-00\-0")
                ElseIf objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCEST Then
                    sAux = objFilCli.sInscricaoEstadual
                Else
                    sAux = objFilCli.sInscricaoMunicipal
                End If
            
            ElseIf mvarobjNFTributacao.lFornecedor > 0 And mvarobjNFTributacao.iFilialForn > 0 Then
            
                Set objFilForn = New ClassFilialFornecedor
                
                objFilForn.lCodFornecedor = mvarobjNFTributacao.lFornecedor
                objFilForn.iCodFilial = mvarobjNFTributacao.iFilialForn
            
                lErro = CF("FilialFornecedor_Le", objFilForn)
                If lErro <> SUCESSO And lErro <> 12929 Then gError ERRO_SEM_MENSAGEM
            
                If objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCSUF Then
                    If Len(Trim(objFilForn.sInscricaoSuframa)) <> 0 Then sAux = Format(objFilForn.sInscricaoSuframa, "00\.0000\-00\-0")
                ElseIf objMnemonicoValor.sMnemonico = MN_FILCLIFORNINSCEST Then
                    sAux = objFilForn.sInscricaoEstadual
                Else
                    sAux = objFilForn.sInscricaoMunicipal
                End If
                
            End If
            objMnemonicoValor.colValor.Add sAux
            
        Case MN_NUMPEDIDOTERC
            If mvariNFEntrada <> MARCADO Then
                objMnemonicoValor.colValor.Add TabTrib_Controle("NumPedidoTerc").Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case MN_CLIFORN
            If mvarobjNFTributacao.lCliente > 0 Then
                objMnemonicoValor.colValor.Add mvarobjNFTributacao.lCliente
            Else
                objMnemonicoValor.colValor.Add mvarobjNFTributacao.lFornecedor
            End If
                
        Case MN_CFOP
            objMnemonicoValor.colValor.Add mvarobjNFTributacao.sNaturezaOp
             
        Case MN_TIPOCLIFORN, MN_CLIFORNNOMERED
             
            bAchou = False
            If mvarobjNFTributacao.lCliente > 0 Then
            
                Set objCli = New ClassCliente
                
                objCli.lCodigo = mvarobjNFTributacao.lCliente
            
                lErro = CF("Cliente_Le", objCli)
                If lErro <> SUCESSO And lErro <> 12293 Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoValor.sMnemonico = MN_TIPOCLIFORN Then
                    objMnemonicoValor.colValor.Add objCli.iTipo
                Else
                    objMnemonicoValor.colValor.Add objCli.sNomeReduzido
                End If

                bAchou = True
            
            ElseIf mvarobjNFTributacao.lFornecedor > 0 Then
            
                Set objForn = New ClassFornecedor
                
                objForn.lCodigo = mvarobjNFTributacao.lFornecedor
            
                lErro = CF("Fornecedor_Le", objForn)
                If lErro <> SUCESSO And lErro <> 12729 Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoValor.sMnemonico = MN_TIPOCLIFORN Then
                    objMnemonicoValor.colValor.Add objForn.iTipo
                Else
                    objMnemonicoValor.colValor.Add objForn.sNomeReduzido
                End If
                
                bAchou = True
            
            End If
            
            If bAchou = False Then
                If objMnemonicoValor.sMnemonico = MN_TIPOCLIFORN Then
                    objMnemonicoValor.colValor.Add 0
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            End If
        
        Case MN_POSSUI_CATCLIFORITEM, MN_POSSUI_CATCLIFORN
            
            bAchou = False
            sCat = objMnemonicoValor.vParam(1)
            If mvarobjNFTributacao.lCliente > 0 And mvarobjNFTributacao.iFilialCli > 0 Then
            
                Set objFilCli = New ClassFilialCliente
                Set colCatItem = New Collection
                
                objFilCli.lCodCliente = mvarobjNFTributacao.lCliente
                objFilCli.iCodFilial = mvarobjNFTributacao.iFilialCli
                
                lErro = CF("FilialCliente_Le_Categorias", objFilCli, colCatItem)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                For Each objFilCliCat In colCatItem
                    
                    If UCase(objFilCliCat.sCategoria) = UCase(sCat) Then
                        If MN_POSSUI_CATCLIFORN = objMnemonicoValor.sMnemonico Then
                            bAchou = True
                            Exit For
                        End If
                        sCatItem = objMnemonicoValor.vParam(2)
                        If UCase(objFilCliCat.sItem) = UCase(sCatItem) Then
                            bAchou = True
                            Exit For
                        End If
                    End If
                Next
            
            ElseIf mvarobjNFTributacao.lFornecedor > 0 And mvarobjNFTributacao.iFilialForn > 0 Then
            
                Set objFilForn = New ClassFilialFornecedor
                Set colCatItem = New Collection
                
                objFilForn.lCodFornecedor = mvarobjNFTributacao.lFornecedor
                objFilForn.iCodFilial = mvarobjNFTributacao.iFilialForn
                
                lErro = CF("FilialFornecedorCategorias_Le_FornFilial", objFilForn, colCatItem)
                If lErro <> SUCESSO And lErro <> 127085 Then gError ERRO_SEM_MENSAGEM
                
                For Each objFilFornCat In colCatItem
                    
                    If UCase(objFilFornCat.sCategoria) = UCase(sCat) Then
                        If MN_POSSUI_CATCLIFORN = objMnemonicoValor.sMnemonico Then
                            bAchou = True
                            Exit For
                        End If
                        sCatItem = objMnemonicoValor.vParam(2)
                        If UCase(objFilFornCat.sItem) = UCase(sCatItem) Then
                            bAchou = True
                            Exit For
                        End If
                    End If
                Next
                
            End If
            objMnemonicoValor.colValor.Add bAchou
            
        Case MN_POSSUI_ICMSST
            If mvarobjNFTributacao.dICMSSubstValor > 0 Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_POSSUI_ICMSRED
        
            bAchou = False
            iLinha = 0
            
            For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, False)
                
                iLinha = iLinha + 1
                Set objTributoDocItemAux = gobjDocTribItem(iLinha, False)

                If objTributoDocItemAux.dICMSPercRedBase <> 0 Then
                
                    bAchou = True
                    Exit For
                    
                End If
                
            Next
            
            If bAchou Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_INTERNA
            If mvariEmitenteNF = DOCINFO_EMPRESA Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_EXTERNA
            If mvariEmitenteNF = DOCINFO_EMPRESA Then
                objMnemonicoValor.colValor.Add False
            Else
                objMnemonicoValor.colValor.Add True
            End If
            
        Case MN_SERIE
            objMnemonicoValor.colValor.Add mvarobjNFTributacao.sSerie
            
        Case MN_VLRDESCONTONF
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dValorDesconto, "STANDARD")
                
        Case MN_VLRTOTALNF
            If mvariNFEntrada = NF_SAIDA Then
                objMnemonicoValor.colValor.Add Format(StrParaDbl(TabTrib_Controle("ValorTotal").Caption), "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Format(StrParaDbl(TabTrib_Controle("Total").Text), "STANDARD")
            End If
        
        Case MN_VLRFRETENF
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dValorFrete, "STANDARD")
                
        Case MN_VLRSEGURONF
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dValorSeguro, "STANDARD")
                
        Case MN_VLRDESPESASNF
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dValorOutrasDespesas, "STANDARD")
                
        Case MN_VLRPRODUTONF
            objMnemonicoValor.colValor.Add Format(StrParaDbl(TabTrib_Controle("ValorProdutos").Caption), "STANDARD")

        Case MN_DEVOLUCAO
        
            Set objTipoDocInfo = New ClassTipoDocInfo
        
            objTipoDocInfo.iCodigo = mvarobjNFTributacao.iTipoNFiscal
        
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM

            If objTipoDocInfo.sNomeTelaNFiscal = "NFiscalDev" Or objTipoDocInfo.sNomeTelaNFiscal = "NFiscalEntDev" Then
                objMnemonicoValor.colValor.Add True
            Else
                objMnemonicoValor.colValor.Add False
            End If
            
        Case MN_NUMNFORIGDET

            Call Obter_NFOrig_Tela(lNumIntNotaOriginal)
                
            Call Obter_NFOrigDet_Tela(sListaNFOrigDet)
            
            If lNumIntNotaOriginal > 0 And Len(Trim(sListaNFOrigDet)) = 0 Then
            'Se não tem detalhamento vai funcionar igual a MN_NUMNFORIG
        
                Set objNFiscalOriginal = New ClassNFiscal
                objNFiscalOriginal.lNumIntDoc = lNumIntNotaOriginal
                
                'Lâ Nota Fiscal Original
                lErro = CF("NFiscal_Le", objNFiscalOriginal)
                If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
                
                objMnemonicoValor.colValor.Add CStr(objNFiscalOriginal.lNumNotaFiscal)
                
            ElseIf Len(Trim(sListaNFOrigDet)) = 0 Then
                objMnemonicoValor.colValor.Add ""
            Else
                objMnemonicoValor.colValor.Add left(sListaNFOrigDet, 250)
            End If
            
        Case MN_NUMNFORIGDETBENEF
        
            iLinha = objMnemonicoValor.vParam(1)
                
            Call Obter_NFOrigDet_Tela(sListaNFOrigDet)
            
            Set colListaNFOrigDet = New Collection
                        
            Call NFOrigDet_Quebra(sListaNFOrigDet, colListaNFOrigDet)
            
            sListaNFOrigDet = ""
            If iLinha <= colListaNFOrigDet.Count Then
                sListaNFOrigDet = colListaNFOrigDet.Item(iLinha)
            End If
            
            objMnemonicoValor.colValor.Add sListaNFOrigDet
            
        Case MN_TIPONFORIG, MN_NUMNFORIG, MN_CFOPNFORIG, MN_EMISSAONFORIG, MN_CLIFORNNFORIG, MN_NUMNFORIGVINC, MN_NUMNFORIGDET

            Call Obter_NFOrig_Tela(lNumIntNotaOriginal)
            
            If lNumIntNotaOriginal > 0 Then
        
                Set objNFiscalOriginal = New ClassNFiscal
                objNFiscalOriginal.lNumIntDoc = lNumIntNotaOriginal
                
                'Lâ Nota Fiscal Original
                lErro = CF("NFiscal_Le", objNFiscalOriginal)
                If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
                
                If objMnemonicoValor.sMnemonico = MN_TIPONFORIG Then
                    objMnemonicoValor.colValor.Add objNFiscalOriginal.iTipoNFiscal
                ElseIf objMnemonicoValor.sMnemonico = MN_CLIFORNNFORIG Then
                    If objNFiscalOriginal.lCliente > 0 Then
                        objMnemonicoValor.colValor.Add objNFiscalOriginal.lCliente
                    Else
                        objMnemonicoValor.colValor.Add objNFiscalOriginal.lFornecedor
                    End If
                ElseIf objMnemonicoValor.sMnemonico = MN_NUMNFORIG Then
                    objMnemonicoValor.colValor.Add CStr(objNFiscalOriginal.lNumNotaFiscal)
                ElseIf objMnemonicoValor.sMnemonico = MN_CFOPNFORIG Then
                    objMnemonicoValor.colValor.Add objNFiscalOriginal.sNaturezaOp
                ElseIf objMnemonicoValor.sMnemonico = MN_EMISSAONFORIG Then
                    objMnemonicoValor.colValor.Add Format(objNFiscalOriginal.dtDataEmissao, "dd/mm/yyyy")
                ElseIf objMnemonicoValor.sMnemonico = MN_NUMNFORIGVINC Then
                    Set X = New ClassFAT
                    Call X.Inicializa
                    If objNFiscalOriginal.iTipoNFiscal = DOCINFO_NFISSRGOV Then
                        objMnemonicoValor.colValor.Add CStr(objNFiscalOriginal.lNumNotaFiscal + 1)
                    ElseIf objNFiscalOriginal.lCliente = X.lCodCliGov And objNFiscalOriginal.lCliente <> 0 Then
                        objMnemonicoValor.colValor.Add CStr(objNFiscalOriginal.lNumNotaFiscal - 1)
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                End If
                
            Else
        
                If objMnemonicoValor.sMnemonico = MN_TIPONFORIG Or objMnemonicoValor.sMnemonico = MN_CLIFORNNFORIG Then
                    objMnemonicoValor.colValor.Add 0
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            
            End If
            
        Case MN_VLRICMS
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dICMSValor, "STANDARD")
            
        Case MN_VLRICMSINTERESTFCP
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dICMSInterestVlrFCPUFDest, "STANDARD")
            
        Case MN_VLRICMSINTERESTUFDEST
             objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dICMSInterestVlrUFDest, "STANDARD")
           
        Case MN_VLRICMSINTERESTUFREMET
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dICMSInterestVlrUFRemet, "STANDARD")
            
        Case MN_VLRICMSST
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dICMSSubstValor, "STANDARD")
            
        Case MN_BASEICMSST
            objMnemonicoValor.colValor.Add Format(mvarobjNFTributacao.dICMSSubstBase, "STANDARD")
                        
        Case MN_VLRICMSSTCOBRANT, MN_BASEICMSSTCOBRANT
        
            dValorBase = 0
            dValor = 0
            iLinha = 0
            For Each objTributoDocItem In gcolDocTribItem(gobjDocTrib, False)
                iLinha = iLinha + 1
                Set objTributoDocItemAux = gobjDocTribItem(iLinha, False)
                dValorBase = dValorBase + objTributoDocItemAux.dICMSSTCobrAntBase
                dValor = dValor + objTributoDocItemAux.dICMSSTCobrAntValor
            Next
            
            If objMnemonicoValor.sMnemonico = MN_BASEICMSSTCOBRANT Then
                objMnemonicoValor.colValor.Add Format(dValorBase, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Format(dValor, "STANDARD")
            End If
            
        Case MN_BASEICMSSTNFORIG, MN_VLRICMSSTNFORIG
            dValorBase = 0
            dValor = 0
            
            If iTipoTela = TIPOTELA_NF Then
            
                For Each objTributoDocItemAux In gcolDocItem(gobjDocTrib, False)
                
                    If Not (objTributoDocItemAux.objItemNFOrig Is Nothing) Then
                    
                        Set objItemNFOrig = objTributoDocItemAux.objItemNFOrig
                           
                        dValorBase = dValorBase + (objItemNFOrig.objTributacao.dICMSSubstBase / objItemNFOrig.dQuantidade) * objTributoDocItemAux.dQuantidade
                        dValor = dValor + (objItemNFOrig.objTributacao.dICMSSubstValor / objItemNFOrig.dQuantidade) * objTributoDocItemAux.dQuantidade
                    
                    End If
                    
                Next
            
            End If
        
            If objMnemonicoValor.sMnemonico = MN_BASEICMSSTNFORIG Then
                objMnemonicoValor.colValor.Add Format(dValorBase, "STANDARD")
            Else
                objMnemonicoValor.colValor.Add Format(dValor, "STANDARD")
            End If
        
        Case Else
             gError 208271

    End Select

    Calcula_Mnemonico_MN = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico_MN:

    Calcula_Mnemonico_MN = gErr

    Select Case gErr

        Case 208271
            Calcula_Mnemonico_MN = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case 208272
            Call Rotina_Erro(vbOKOnly, "ERRO_REGRASMSG_MN_REGRA_LIN_POS", gErr)

        Case 208273
            Call Rotina_Erro(vbOKOnly, "ERRO_DASALIQUOTAS_NAO_CADASTRADO", gErr, objDASAliquotas.iAno, objDASAliquotas.iMes)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208274)

    End Select

    Exit Function

End Function

Function Pesos_Recalcula() As Long
'trata a alteracao de uma linha de item no grid

Dim lErro As Long, iIndice As Integer
Dim objItem As Object, objProduto As ClassProduto
Dim dFator As Double, dPesoBruto As Double, dPesoLiquido As Double

On Error GoTo Erro_Pesos_Recalcula

    If mvariEmitenteNF = DOCINFO_EMPRESA Then
    
        For Each objItem In gcolDocItem(gobjDocTrib)
        
            Set objProduto = New ClassProduto
    
            objProduto.sCodigo = objItem.sProduto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItem.sUnidadeMed, objProduto.sSiglaUMVenda, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            If objProduto.dPesoBruto < objProduto.dPesoLiq Then objProduto.dPesoBruto = objProduto.dPesoLiq
            
            'Acumula o peso Bruto e o peso líquido
            dPesoBruto = dPesoBruto + objProduto.dPesoBruto * dFator * objItem.dQuantidade
            dPesoLiquido = dPesoLiquido + objProduto.dPesoLiq * dFator * objItem.dQuantidade
        
        Next
        
    End If
        
    TabTrib_Controle("PesoBruto").Text = Format(dPesoBruto, TabTrib_Controle("PesoBruto").Format) 'Alterado por Wagner
    TabTrib_Controle("PesoLiquido").Text = Format(dPesoLiquido, TabTrib_Controle("PesoLiquido").Format) 'Alterado por Wagner
        
    Pesos_Recalcula = SUCESSO

    Exit Function

Erro_Pesos_Recalcula:

    Pesos_Recalcula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 211317)

    End Select

    Exit Function

End Function

Function Obter_NFOrig_Tela(lNumIntNotaOriginal As Long) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Obter_NFOrig_Tela

    lNumIntNotaOriginal = 0
        
    If Existe_Controle_Tela("SerieNFiscalOriginal") And Existe_Controle_Tela("NFiscalOriginal") Then
    
        objNFiscal.sSerie = TabTrib_Controle("SerieNFiscalOriginal").Text
        objNFiscal.lNumNotaFiscal = StrParaLong(TabTrib_Controle("NFiscalOriginal").Text)
        objNFiscal.iFilialEmpresa = giFilialEmpresa
    
        If Len(Trim(objNFiscal.sSerie)) > 0 And objNFiscal.lNumNotaFiscal <> 0 Then
    
            objTipoDocInfo.iCodigo = mvarobjNFTributacao.iTipoNFiscal
        
            'Lê o Tipo de Documento
            lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
            If lErro <> SUCESSO And lErro <> 31415 Then gError ERRO_SEM_MENSAGEM
            
            Select Case objTipoDocInfo.sNomeTelaNFiscal
            
                Case "NFiscalEntDev"
                    lErro = CF("NFiscalInternaSaida_Le_Numero", objNFiscal)
                    If lErro <> SUCESSO And lErro <> 30765 Then gError ERRO_SEM_MENSAGEM
                    If lErro = SUCESSO Then lNumIntNotaOriginal = objNFiscal.lNumIntDoc
                    
                Case "NFiscalDev"
                    If TabTrib_Controle("NFiscalPropria").Value = vbChecked Then
                        lErro = CF("NFiscalInternaEntNaoDev_Le_Numero", objNFiscal)
                        If lErro <> SUCESSO And lErro <> 39101 Then gError ERRO_SEM_MENSAGEM
                        If lErro = SUCESSO Then lNumIntNotaOriginal = objNFiscal.lNumIntDoc
                    Else
                        objNFiscal.lCliente = mvarobjNFTributacao.lCliente
                        objNFiscal.lFornecedor = mvarobjNFTributacao.lFornecedor
                        objNFiscal.iFilialCli = mvarobjNFTributacao.iFilialCli
                        objNFiscal.iFilialForn = mvarobjNFTributacao.iFilialForn
                        lErro = CF("NFiscalExternaNaoDev_Le_Numero", objNFiscal)
                        If lErro <> SUCESSO And lErro <> 39116 Then gError ERRO_SEM_MENSAGEM
                        If lErro = SUCESSO Then lNumIntNotaOriginal = objNFiscal.lNumIntDoc
                    End If
                Case "NFiscalFatura"
                    objNFiscal.lCliente = mvarobjNFTributacao.lCliente
                    objNFiscal.iFilialCli = mvarobjNFTributacao.iFilialCli
                    objNFiscal.iFilialEmpresa = mvarobjNFTributacao.iFilialEmpresa
                    lErro = CF("NFiscalFatura_Le_NFOriginal", objNFiscal)
                    If lErro <> SUCESSO And lErro <> 45347 Then gError ERRO_SEM_MENSAGEM
                    If lErro = SUCESSO Then lNumIntNotaOriginal = objNFiscal.lNumIntDoc
                    
                Case Else
                    If objNFiscal.lFornecedor > 0 Then
                        objNFiscal.lFornecedor = mvarobjNFTributacao.lFornecedor
                        objNFiscal.iFilialForn = mvarobjNFTributacao.iFilialForn
                        lErro = CF("NFiscal_Le_Num_Forn", objNFiscal, objTipoDocInfo.iTipo)
                        If lErro <> SUCESSO And lErro <> 31467 Then gError ERRO_SEM_MENSAGEM
                        If lErro = SUCESSO Then lNumIntNotaOriginal = objNFiscal.lNumIntDoc
                    Else
                        objNFiscal.lCliente = mvarobjNFTributacao.lCliente
                        objNFiscal.iFilialCli = mvarobjNFTributacao.iFilialCli
                        lErro = CF("NFiscal_Le_Num_Cli", objNFiscal, objTipoDocInfo.iTipo)
                        If lErro <> SUCESSO And lErro <> 35956 Then gError ERRO_SEM_MENSAGEM
                        If lErro = SUCESSO Then lNumIntNotaOriginal = objNFiscal.lNumIntDoc
                    End If

            End Select

        End If

    End If
    
    Obter_NFOrig_Tela = SUCESSO

    Exit Function

Erro_Obter_NFOrig_Tela:

    Obter_NFOrig_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211318)

    End Select

    Exit Function

End Function

Private Function Existe_Controle_Tela(ByVal sControle As String) As Boolean

Dim objControle As Object

On Error GoTo Erro_Existe_Controle_Tela
    
    Set objControle = mvarobjForm.Controls(sControle)
    
    Existe_Controle_Tela = True
    
    Exit Function
    
Erro_Existe_Controle_Tela:
    
    Existe_Controle_Tela = False
    
    Exit Function
    
End Function

Sub ComboISSTipo_Click()

    If TabTrib_Controle("ComboISSTipo").ListIndex = -1 Then Exit Sub

    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem

End Sub

Sub ComboICMSSimplesTipo_Click()

    If TabTrib_Controle("ComboICMSSimplesTipo").ListIndex = -1 Then Exit Sub

    'If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
    
    'Ajusta o ICMS normal de acordo com o simples para tratar corretamente os flags em ClassTributos nas alterações manuais
    Select Case TabTrib_Controle("ComboICMSSimplesTipo").ItemData(TabTrib_Controle("ComboICMSSimplesTipo").ListIndex)
    
        Case 1, 2   'Tributada com(1)\sem(2) permissão de crédito
            Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSTipo"), 1) 'Integral
    
        Case 3, 7, 8 'Isenção do ICMS para faixa de receita bruta\Imune\Não tributada
            Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSTipo"), 2) 'Isenta
    
        Case 4, 5  'Trib. c\(4)\s\(5) permissão de crédito e c\ cobrança por ST
            Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSTipo"), 6) 'Tributado com substituição
    
        Case 6   'Isenção do ICMS e c\ cobrança por ST
            Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSTipo"), 9) 'Isento com cobrança por subst.
    
        Case 9   'ICMS cobrado anteriormente por ST ou antecipação
            Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSTipo"), 8) 'Cobrado anteriormente por subst.
    
        Case 10   'Outros
            Call Combo_Seleciona_ItemData(TabTrib_Controle("ComboICMSTipo"), 99) 'Outras
    
    End Select

    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
    
End Sub

Sub ICMSpercBaseOperacaoPropria_Change()
    mvariICMSpercBaseOperacaoPropriaAlterado = 1
End Sub

Sub ICMSpercBaseOperacaoPropria_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariICMSpercBaseOperacaoPropriaAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ICMSpercBaseOperacaoPropria").Text)) > 0 Then

        lErro = Porcentagem_Critica_Nao_Zero(TabTrib_Controle("ICMSpercBaseOperacaoPropria").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariICMSpercBaseOperacaoPropriaAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ICMSvBCSTRet_Change()
    mvariICMSvBCSTRetAlterado = 1
End Sub

Sub ICMSvBCSTRet_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariICMSvBCSTRetAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ICMSvBCSTRet").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvBCSTRet").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariICMSvBCSTRetAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ICMSvICMSSTRet_Change()
    mvariICMSvICMSSTRetAlterado = 1
End Sub

Sub ICMSvICMSSTRet_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariICMSvICMSSTRetAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ICMSvICMSSTRet").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvICMSSTRet").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariICMSvICMSSTRetAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ICMSvBCSTDest_Change()
    mvariICMSvBCSTDestAlterado = 1
End Sub

Sub ICMSvBCSTDest_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariICMSvBCSTDestAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ICMSvBCSTDest").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvBCSTDest").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariICMSvBCSTDestAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ICMSvICMSSTDest_Change()
    mvariICMSvICMSSTDestAlterado = 1
End Sub

Sub ICMSvICMSSTDest_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariICMSvICMSSTDestAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ICMSvICMSSTDest").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvICMSSTDest").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariICMSvICMSSTDestAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ICMSvCredSN_Change()
    mvariICMSvCredSNAlterado = 1
End Sub

Sub ICMSvCredSN_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariICMSvCredSNAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ICMSvCredSN").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvCredSN").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariICMSvCredSNAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ICMSValorIsento_Change()
    mvariICMSValorIsentoAlterado = 1
End Sub

Sub ICMSValorIsento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValor_Validate

    If mvariICMSValorIsentoAlterado = 0 Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("ICMSValorIsento").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSValorIsento").Text)
        If lErro <> SUCESSO Then gError 202840
    
    End If
    
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
    mvariICMSValorIsentoAlterado = 0

    Exit Sub
    
Erro_ISSValor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 202840

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202841)

    End Select

    Exit Sub

End Sub

Sub ICMSpCredSN_Change()
    mvariICMSpCredSNAlterado = 1
End Sub

Sub ICMSpCredSN_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSpCredSN_Validate

    If mvariICMSpCredSNAlterado = 1 Then
    
        If Len(Trim(TabTrib_Controle("ICMSpCredSN").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSpCredSN").Text)
            If lErro <> SUCESSO Then gError 56868
            
        End If
        
        If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariICMSpCredSNAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSpCredSN_Validate:

    Cancel = True
    
    Select Case gErr

        Case 56868

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Sub IPICNPJProdutor_GotFocus(iAlterado As Integer)
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("IPICNPJProdutor"), iAlterado)
End Sub

Sub IPICodSelo_GotFocus(iAlterado As Integer)
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("IPICodSelo"), iAlterado)
End Sub

Sub IPISeloQtde_GotFocus(iAlterado As Integer)
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("IPISeloQtde"), iAlterado)
End Sub

Sub ICMSUFDevidoST_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)
    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Sub ICMSUFDevidoST_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Sub ICMSMotivo_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)
    If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Public Function Valida_TributacaoNF_Versao(ByVal objTribItem As ClassTributacaoDocItem) As Long

Dim lErro As Long, bAchou As Boolean
Dim objTipoTribICMS As ClassTipoTribICMS
Dim objTipoTribICMSSimples As ClassTipoTribICMSSimples
Dim objTipoTribIPI As ClassTipoTribIPI
Dim objTipoTribPIS As ClassTipoTribPISCOFINS
Dim objTipoTribCOFINS As ClassTipoTribPISCOFINS
Dim objTipoTribISS As ClassTipoTribISS

On Error GoTo Erro_Valida_TributacaoNF_Versao

    bAchou = False
    For Each objTipoTribICMS In mvarcolTiposTribICMS
        If objTipoTribICMS.iTipo = objTribItem.iICMSTipo Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou Then
        If objTipoTribICMS.iVersaoNFE > gobjCRFAT.iVersaoNFE Then gError 209177
    End If
    
    bAchou = False
    For Each objTipoTribICMSSimples In mvarcolTiposTribICMSSimples
        If objTipoTribICMSSimples.iTipo = objTribItem.iICMSSimplesTipo Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou Then
        If objTipoTribICMSSimples.iVersaoNFE > gobjCRFAT.iVersaoNFE Then gError 209178
    End If
    
    bAchou = False
    For Each objTipoTribIPI In mvarcolTiposTribIPI
        If objTipoTribIPI.iTipo = objTribItem.iIPITipo Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou Then
        If objTipoTribIPI.iVersaoNFE > gobjCRFAT.iVersaoNFE Then gError 209179
    End If
    
    bAchou = False
    For Each objTipoTribPIS In mvarcolTiposTribPISCOFINS
        If objTipoTribPIS.iTipo = objTribItem.iPISTipo Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou Then
        If objTipoTribPIS.iVersaoNFE > gobjCRFAT.iVersaoNFE Then gError 209180
    End If
    
    bAchou = False
    For Each objTipoTribCOFINS In mvarcolTiposTribPISCOFINS
        If objTipoTribCOFINS.iTipo = objTribItem.iCOFINSTipo Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou Then
        If objTipoTribCOFINS.iVersaoNFE > gobjCRFAT.iVersaoNFE Then gError 209181
    End If
    
    bAchou = False
    For Each objTipoTribISS In mvarcolTiposTribISS
        If objTipoTribISS.iTipo = objTribItem.iISSTipo Then
            bAchou = True
            Exit For
        End If
    Next
    If bAchou Then
        If objTipoTribISS.iVersaoNFE > gobjCRFAT.iVersaoNFE Then gError 209182
    End If

    lErro = CF("CFOP_Testa_Versao", objTribItem.sNaturezaOp, gobjVersao.sColunaCFOP, gobjVersao.sDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Valida_TributacaoNF_Versao = SUCESSO

    Exit Function
    
Erro_Valida_TributacaoNF_Versao:

    Valida_TributacaoNF_Versao = gErr

    Select Case gErr
    
        Case 209177
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIB_ST_ICMS_INVALIDO_VERSAO", gErr, Format(objTipoTribICMS.iTipoTribCST, "00") & SEPARADOR & objTipoTribICMS.sDescricao, gobjVersao.sDescricao)
        
        Case 209178
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIB_ST_ICMS_SIMPLES_INVALIDO_VERSAO", gErr, Format(objTipoTribICMSSimples.iCSOSN, "000") & SEPARADOR & objTipoTribICMSSimples.sDescricao, gobjVersao.sDescricao)
        
        Case 209179
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIB_ST_IPI_INVALIDO_VERSAO", gErr, Format(IIf(mvariNFEntrada = MARCADO, objTipoTribIPI.iCSTEntrada, objTipoTribIPI.iCSTSaida), "00") & SEPARADOR & objTipoTribIPI.sDescricao, gobjVersao.sDescricao)
        
        Case 209180
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIB_ST_PIS_INVALIDO_VERSAO", gErr, CStr(objTipoTribPIS.iTipo) & SEPARADOR & objTipoTribPIS.sDescricao, gobjVersao.sDescricao)
        
        Case 209181
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIB_ST_COFINS_INVALIDO_VERSAO", gErr, CStr(objTipoTribCOFINS.iTipo) & SEPARADOR & objTipoTribCOFINS.sDescricao, gobjVersao.sDescricao)
        
        Case 209182
            Call Rotina_Erro(vbOKOnly, "ERRO_TRIB_ST_ISS_INVALIDO_VERSAO", gErr, objTipoTribISS.sDescricao, gobjVersao.sDescricao)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209183)
    
    End Select

    Exit Function
    
End Function

Public Sub ValorDespesas_Obter(dValor As Double)
Dim objTribItem As Object
            
        dValor = 0
        If Not (gcolDocTribItem Is Nothing) Then
        
            For Each objTribItem In gcolDocTribItem
                dValor = Arredonda_Moeda(dValor + objTribItem.dValorOutrasDespesasItem)
            Next

        End If

End Sub

Sub NatBCCred_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub

Public Function ForcaICMSTipoManual() As Long

Dim objTributacaoItem As Object, iIndice As Integer

On Error GoTo Erro_ForcaICMSTipoManual

    For iIndice = 1 To gcolDocTribItem.Count

        Set objTributacaoItem = gobjDocTribItem(iIndice)
        
        objTributacaoItem.iICMSTipoManual = VAR_PREENCH_MANUAL
        
    Next

    ForcaICMSTipoManual = SUCESSO

    Exit Function
    
Erro_ForcaICMSTipoManual:

    ForcaICMSTipoManual = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209183)
    
    End Select

    Exit Function
    
End Function

Public Function ValorPagarPrimParc_Calcula(dValorPrimParc As Double) As Long
'retorna o valor que nao deve ser rateado entre as parcelas, sendo cobrado integralmente na 1a.
'Alguns clientes preferem cobrar a ST, IPI,...na 1a parcela

Dim lErro As Long
Dim dValorPagar As Double
Dim dValorTotal As Double, dValorIRRF As Double, dValorINSSRetido As Double, dPISRetido As Double, dISSRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
Dim X As New ClassFatGlobal
Dim dPISSTValor As Double, dCOFINSSTValor As Double
Dim objTribItem As Object, dValorAux As Double
Dim bFat As Boolean, iItem As Integer, dValorDesconto As Double

On Error GoTo Erro_ValorPagarPrimParc_Calcula

    dCOFINSSTValor = 0
    dPISSTValor = 0
    
    iItem = 0
    For Each objTribItem In gcolDocTribItem
        
        iItem = iItem + 1
    
        dPISSTValor = dPISSTValor + gobjDocTribItemTributacao(objTribItem).dPISSTValor
        dCOFINSSTValor = dCOFINSSTValor + gobjDocTribItemTributacao(objTribItem).dCOFINSSTValor
        
    Next

    dValorPrimParc = 0
    
    If mvariNFEntrada <> NF_ENTRADA Then
        
        If X.gobjFAT.iPrimParcPISST Then dValorPrimParc = dValorPrimParc + dPISSTValor
        If X.gobjFAT.iPrimParcCOFINSST Then dValorPrimParc = dValorPrimParc + dCOFINSSTValor
        If X.gobjFAT.iPrimParcICMSST Then dValorPrimParc = dValorPrimParc + StrParaDbl(TabTrib_Controle("ICMSSubstValor1"))
        If X.gobjFAT.iPrimParcIPI Then dValorPrimParc = dValorPrimParc + StrParaDbl(TabTrib_Controle("IPIValor1"))
    
        dValorPrimParc = Arredonda_Moeda(dValorPrimParc, 2)
    
    End If

    ValorPagarPrimParc_Calcula = SUCESSO

    Exit Function

Erro_ValorPagarPrimParc_Calcula:

    ValorPagarPrimParc_Calcula = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 153760)

    End Select

    Exit Function
    
End Function

Function Obter_NFOrigDet_Tela(sListaNFOrigDet As String) As Long

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objItemNFRem As ClassItemNF
Dim objItemNFOrig As ClassItemNF
Dim objItemNFRemOrig As ClassItemNF
Dim colNFOrig As New Collection
Dim colItensNFOrig As New Collection
Dim objNFOrig As ClassNFiscal, bAchou As Boolean
Dim sAux As String, sAuxItem As String

On Error GoTo Erro_Obter_NFOrigDet_Tela

    sListaNFOrigDet = ""
    
    'Por enquanto vai tratar apenas da devolução simbólica de insumo utilizado no beneficiamento
    'Depois pode ser melhorada para exibir os dados a nível de grid para os demais casos
    If mvarobjNFTributacao.iTipoNFiscal = DOCINFO_NFISDBFNE Then

        If Not (mvarobjForm.gobjNFiscal Is Nothing) Then
    
            'Monta a coleção de itens de NF Original
            For Each objItemNF In mvarobjForm.gobjNFiscal.colItens
                       
                If Not (objItemNF.colItensNFDevBenefRem Is Nothing) Then
                
                    If Len(Trim(sAux)) <> 0 Then sAux = sAux & "; "
                    sAux = sAux & CStr(objItemNF.iItem) & "("
                    sAuxItem = ""
    
                    For Each objItemNFRem In objItemNF.colItensNFDevBenefRem
                    
                        Set objItemNFRemOrig = New ClassItemNF
                        
                        objItemNFRemOrig.lNumIntDoc = objItemNFRem.lNumIntDoc
                        
                        lErro = CF("ItemNFiscal_Le", objItemNFRemOrig)
                        If lErro <> SUCESSO And lErro <> 35225 Then gError ERRO_SEM_MENSAGEM
                        
                        objItemNFRemOrig.dQuantidade = objItemNFRem.dQuantidade
                        
                        '-------
                        Set objNFOrig = New ClassNFiscal
                        objNFOrig.lNumIntDoc = objItemNFRemOrig.lNumIntNF
    
                        lErro = CF("NFiscal_Le", objNFOrig)
                        If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
                        
                        If Len(Trim(sAuxItem)) <> 0 Then sAuxItem = sAuxItem & "/"
                        sAuxItem = sAuxItem & CStr(objNFOrig.lNumNotaFiscal)
                        
'                        colItensNFOrig.Add objItemNFRemOrig
                    
                    Next
                    
                    sAux = sAux & sAuxItem & ")"
                
                End If
            
            Next
            
'            'Organiza por NF
'            For Each objItemNFOrig In colItensNFOrig
'
'                bAchou = False
'                For Each objNFOrig In colNFOrig
'                    If objNFOrig.lNumIntDoc = objItemNFOrig.lNumIntNF Then
'                        bAchou = True
'                        Exit For
'                    End If
'                Next
'                If Not bAchou Then
'
'                    Set objNFOrig = New ClassNFiscal
'                    objNFOrig.lNumIntDoc = objItemNFOrig.lNumIntNF
'                    colNFOrig.Add objNFOrig
'
'                    lErro = CF("NFiscal_Le", objNFOrig)
'                    If lErro <> SUCESSO And lErro <> 31442 Then gError ERRO_SEM_MENSAGEM
'
'                End If
'
'                objNFOrig.colItensNF.Add1 objItemNFOrig
'
'            Next
'
'            'Monta o texto
'            sAux = ""
'            For Each objNFOrig In colNFOrig
'
'                If Len(Trim(sAux)) <> 0 Then sAux = sAux & "; "
'                sAux = sAux & "NF " & CStr(objNFOrig.lNumNotaFiscal) & " ("
'
'                sAuxItem = ""
'                For Each objItemNFOrig In objNFOrig.colItens
'
'                    If Len(Trim(sAuxItem)) <> 0 Then sAuxItem = sAuxItem & " - "
'                    sAuxItem = sAuxItem & "Item " & objItemNFOrig.iItem & ": " & Formata_Estoque(objItemNFOrig.dQuantidade) & " " & objItemNFOrig.sUnidadeMed
'
'                Next
'
'                sAux = sAux & sAuxItem & ")"
'
'            Next

        End If

    End If
    
    sListaNFOrigDet = sAux
    
    Obter_NFOrigDet_Tela = SUCESSO

    Exit Function

Erro_Obter_NFOrigDet_Tela:

    Obter_NFOrigDet_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211319)

    End Select

    Exit Function

End Function

Private Function NFOrigDet_Quebra(ByVal sListaNFOrigDet As String, ByVal colListaNFOrigDet As Collection) As Long

Dim iPosNewLine As Integer, iPosEspaco As Integer, iPosAnt As Integer
Dim sMsgAux As String, sLinha As String, iTamMax As Integer

On Error GoTo Erro_NFOrigDet_Quebra

    iTamMax = 210
    sMsgAux = sListaNFOrigDet
    
    Do While Len(Trim(sMsgAux)) <> 0

        If Len(sMsgAux) <= iTamMax Then
            sLinha = sMsgAux
        Else
            sLinha = left(sMsgAux, iTamMax)
            'Busca o último espaço antes do corte
            iPosAnt = 0
            iPosEspaco = InStr(1, sLinha, " ")
            Do While iPosEspaco <> 0
                iPosAnt = iPosEspaco
                iPosEspaco = InStr(iPosAnt + 1, sLinha, " ")
            Loop
            'Se achou algum espaço Corta a mensagem até ele
            If iPosAnt <> 0 Then
                sLinha = left(sLinha, iPosAnt - 1)
            End If
        End If
    
        colListaNFOrigDet.Add Trim(sLinha)
        If Len(sLinha) < Len(sMsgAux) Then
            sMsgAux = Mid(sMsgAux, Len(sLinha) + 1)
        Else
            sMsgAux = ""
        End If
        
    Loop

    NFOrigDet_Quebra = SUCESSO

    Exit Function

Erro_NFOrigDet_Quebra:

    NFOrigDet_Quebra = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 211320)

    End Select

    Exit Function
    
End Function

Sub TotTribItemTipo_Click()

    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem

End Sub

Sub TotTribTipo_Click()

    If mvariTrazendoTribTela = 0 Then Call AtualizarTabTrib

End Sub

Sub TotTribItem_Change()

    mvariTotTribItemAlterado = 1
    
End Sub

Sub TotTribItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TotTribItem_Validate

    If mvariTotTribItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("TotTribItem").Text)) > 0 Then
    
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("TotTribItem").Text)
            If lErro <> SUCESSO Then Error 56876
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariTotTribItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_TotTribItem_Validate:

    Cancel = True
    
    Select Case Err

        Case 56876
''???            If TabTrib_Controle("TotTribItem").Visible Then TabTrib_Controle("TotTribItem").SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153752)

    End Select

    Exit Sub

End Sub

Sub TotTrib_Validate(Cancel As Boolean)

Dim dTotTrib As Double
Dim lErro As Long

On Error GoTo Erro_TotTrib_Validate

    If mvariTotTribAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("TotTrib").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("TotTrib").Text)
        If lErro <> SUCESSO Then Error 27797

        dTotTrib = CDbl(TabTrib_Controle("TotTrib").Text)

        TabTrib_Controle("TotTrib").Text = Format(dTotTrib, "Standard")

    End If

    mvariTotTribAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_TotTrib_Validate:

    Cancel = True
    
    Select Case Err

        Case 27797

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153772)

    End Select

    Exit Sub

End Sub

'--------------------------------NFe 3.10 --------------------------------------------

'--********************************************RESUMO********************************************

Public Sub OpcaoResumo_Click()

Dim lErro As Long, objAux As Object

On Error GoTo Erro_OpcaoResumo_Click

    'Se frame selecionado não for o atual
    If TabTrib_Controle("OpcaoResumo").SelectedItem.Index <> mvariFrameAtualRes Then

        If TabStrip_PodeTrocarTab(mvariFrameAtualRes, TabTrib_Controle("OpcaoResumo"), mvarobjForm) <> SUCESSO Then Exit Sub
        
        'Esconde o frame atual, mostra o novo
        Set objAux = TabTrib_Controle("FrameResumo")
        objAux(TabTrib_Controle("OpcaoResumo").SelectedItem.Index).Visible = True
        objAux(mvariFrameAtualRes).Visible = False
        'Armazena novo valor de mvariFrameAtual
        mvariFrameAtualRes = TabTrib_Controle("OpcaoResumo").SelectedItem.Index

    End If

    Exit Sub

Erro_OpcaoResumo_Click:

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 153723)

    End Select

    Exit Sub

End Sub

Public Sub IndPresenca_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribTela = 0 And bAtualizaTrib Then Call AtualizarTabTrib
End Sub

Public Sub IndConsumidorFinal_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribTela = 0 And bAtualizaTrib Then Call AtualizarTabTrib
End Sub

'--ICMSDesonerado-Caption = Total do ICMS Isento
'--ISSValorDeducao-Caption = Total das Deduções de ISS
'--ISSValorOutrasRet-Caption = Total do ISS Outras Retenções
'--ISSValorDescIncond-Caption = Total do ISS Desconto Incondicionado
'--ISSValorDescCond-Caption = Total do ISS Desconto Condicionado

Public Sub DataPrestServico_Change()
    mvariDataPrestServicoAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataPrestServico_GotFocus(iAlterado As Integer)
 
Dim iDataPrestServicoAux As Integer

    iDataPrestServicoAux = mvariDataPrestServicoAlterado
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("DataPrestServico"), iAlterado)
    mvariDataPrestServicoAlterado = iDataPrestServicoAux
    
End Sub

Public Function DataPrestServico_Validate(Cancel As Boolean) As Long

Dim lErro As Long

On Error GoTo Erro_DataPrestServico_Validate
    
    If mvariDataPrestServicoAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("DataPrestServico").Text)) > 0 Then
    
            lErro = Data_Critica(TabTrib_Controle("DataPrestServico").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        End If
           
        Call AtualizarTabTrib
        
        mvariDataPrestServicoAlterado = 0
    
    End If
    
    DataPrestServico_Validate = SUCESSO
    
    Exit Function

Erro_DataPrestServico_Validate:

    Cancel = True
    
    DataPrestServico_Validate = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 213583)

    End Select

    Exit Function
  
End Function

Public Sub UpDownPrestServico_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPrestServico_DownClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(TabTrib_Controle("DataPrestServico"), DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = DataPrestServico_Validate(bSGECancelDummy)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownPrestServico_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157961)

    End Select

    Exit Sub
    
End Sub

Public Sub UpDownPrestServico_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownPrestServico_UpClick

    'Diminui a adata em um dia
    lErro = Data_Up_Down_Click(TabTrib_Controle("DataPrestServico"), AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = DataPrestServico_Validate(bSGECancelDummy)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownPrestServico_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157961)

    End Select

    Exit Sub
    
End Sub

'--PISValorST-Caption = Soma do Valor do PIS nos itens
'--COFINSValorST-Caption = Soma do Valor do COFINS nos itens
'--IIBase-Caption = Soma da Base do II nos itens
'--IIDespAdu-Caption = Soma da Despesa Aduaneira nos itens
'--IIIOF-Caption = Soma do IOF nos itens
'--IIValor-Caption = Soma do valor do II nos itens

'--********************************************ITENS********************************************

Public Sub ICMS51ValorOpItem_Change()
    mvariICMS51ValorOpItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ICMS51ValorOpItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMS51ValorOpItem_Validate

    If mvariICMS51ValorOpItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ICMS51ValorOpItem").Text)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMS51ValorOpItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMS51ValorOpItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMS51ValorOpItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSPercDiferItem_Change()
    mvariICMSPercDiferItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ICMSPercDiferItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSPercDiferItem_Validate

    If mvariICMSPercDiferItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSPercDiferItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSPercDiferItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSPercDiferItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSPercDiferItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Public Sub ICMSValorDifItem_Change()
    mvariICMSValorDifItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ICMSValorDifItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSValorDifItem_Validate

    If mvariICMSValorDifItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSValorDifItem").Text)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSValorDifItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSValorDifItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSValorDifItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

Public Sub ISSIndIncentivoItem_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub

Public Sub ISSValorDeducaoItem_Change()
    mvariISSValorDeducaoItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSValorDeducaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValorDeducaoItem_Validate

    If mvariISSValorDeducaoItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ISSValorDeducaoItem").Text)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSValorDeducaoItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariISSValorDeducaoItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSValorDeducaoItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Public Sub ISSValorOutrasRetItem_Change()
    mvariISSValorOutrasRetItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSValorOutrasRetItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValorOutrasRetItem_Validate

    If mvariISSValorOutrasRetItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ISSValorOutrasRetItem").Text)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSValorOutrasRetItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariISSValorOutrasRetItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSValorOutrasRetItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

Public Sub ISSValorDescIncondItem_Change()
    mvariISSValorDescIncondItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSValorDescIncondItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValorDescIncondItem_Validate

    If mvariISSValorDescIncondItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ISSValorDescIncondItem").Text)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSValorDescIncondItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariISSValorDescIncondItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSValorDescIncondItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

Public Sub ISSValorDescCondItem_Change()
    mvariISSValorDescCondItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSValorDescCondItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValorDescCondItem_Validate

    If mvariISSValorDescCondItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ISSValorDescCondItem").Text)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSValorDescCondItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariISSValorDescCondItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSValorDescCondItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

Public Sub ISSValorRetItem_Change()
    mvariISSValorRetItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSValorRetItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSValorRetItem_Validate

    If mvariISSValorRetItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ISSValorRetItem").Text)) > 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ISSValorRetItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariISSValorRetItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSValorRetItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

Public Sub ISSIndExigibilidadeItem_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub

Public Sub ISSCodServItem_Change()
    mvariISSCodServItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSCodServItemLabel_Click()

Dim colSelecao As New Collection
Dim objCodServMun As New ClassCodServMun

On Error GoTo Erro_ISSCodServItemLabel_Click
    
    objCodServMun.sCodServ = Trim(TabTrib_Controle("ISSCodServItem").Text)
   
    colSelecao.Add Trim(TabTrib_Controle("ISSMunicFaturamento").Text)
    
    'Chama Tela TituloReceberLista
    Call Chama_Tela("CodServMunLista", colSelecao, objCodServMun, mvarobjEventoISSCodServ)

    Exit Sub

Erro_ISSCodServItemLabel_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200973)

    End Select

    Exit Sub
    
End Sub

Private Sub mvarobjEventoISSCodServ_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCodServ As ClassCodServMun

On Error GoTo Erro_objEventoISSCodServ_evSelecao

    Set objCodServ = obj1
        
    TabTrib_Controle("ISSCodServItem").PromptInclude = False
    TabTrib_Controle("ISSCodServItem").Text = objCodServ.sCodServ
    TabTrib_Controle("ISSCodServItem").PromptInclude = True
    Call ISSCodServItem_Validate(bSGECancelDummy)

    Exit Sub

Erro_objEventoISSCodServ_evSelecao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208173)

    End Select

    Exit Sub

End Sub

Public Sub ISSCodServItem_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)

Dim lErro As Long

On Error GoTo Erro_ISSCodServItem_Validate

    If mvariISSCodServItemAlterado = REGISTRO_ALTERADO Then
    
        lErro = ISSCodServItem_Trata_Descricao
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariISSCodServItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSCodServItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

'--ISSCodServDesc-Caption=Descrição do código de ISSCodServItem

Private Function ISSCodServItem_Trata_Descricao() As Long

Dim lErro As Long
Dim objCodServ As New ClassCodServMun

On Error GoTo Erro_ISSCodServItem_Trata_Descricao

    objCodServ.sCodIBGE = Trim(TabTrib_Controle("ISSMunicFaturamento").Text)
    objCodServ.sCodServ = Trim(TabTrib_Controle("ISSCodServItem").Text)
    
    lErro = CF("CodServMun_Le", objCodServ)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro <> ERRO_LEITURA_SEM_DADOS Then

        TabTrib_Controle("ISSCodServDesc").Caption = objCodServ.sDescricao1
        
    Else
        
        TabTrib_Controle("ISSCodServDesc").Caption = ""
        
    End If
        
    ISSCodServItem_Trata_Descricao = SUCESSO

    Exit Function
    
Erro_ISSCodServItem_Trata_Descricao:

    ISSCodServItem_Trata_Descricao = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Function
    
End Function

Public Sub ISSLabelMunicIncidImp_Click()

Dim colSelecao As New Collection
Dim objIBGEMunicipios As New ClassIBGEMunicipios

On Error GoTo Erro_ISSLabelMunicIncidImp_Click
    
    'Chama Tela TituloReceberLista
    Call Chama_Tela("IBGEMunicipiosLista", colSelecao, objIBGEMunicipios, mvarobjEventoISSMunicIncidImp)

    Exit Sub

Erro_ISSLabelMunicIncidImp_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200973)

    End Select

    Exit Sub

End Sub

Private Sub mvarobjEventoISSMunicIncidImp_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objIBGEMunicipios As ClassIBGEMunicipios

On Error GoTo Erro_mvarobjEventoISSMunicIncidImp_evSelecao

    Set objIBGEMunicipios = obj1
    
    TabTrib_Controle("ISSMunicIncidImpItem").PromptInclude = False
    TabTrib_Controle("ISSMunicIncidImpItem").Text = objIBGEMunicipios.sCodUF & SEPARADOR & objIBGEMunicipios.sCodMunic
    TabTrib_Controle("ISSMunicIncidImpItem").PromptInclude = True
    Call ISSMunicIncidImpItem_Validate(bSGECancelDummy)
    
    Exit Sub

Erro_mvarobjEventoISSMunicIncidImp_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200974)

    End Select

    Exit Sub

End Sub

Public Sub ISSMunicIncidImpItem_Change()
    mvariISSMunicIncidImpItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSMunicIncidImpItem_GotFocus(iAlterado As Integer)
 
Dim iISSMunicIncidImpItemAux As Integer

    iISSMunicIncidImpItemAux = mvariISSMunicIncidImpItemAlterado
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("ISSMunicIncidImpItem"), iAlterado)
    mvariISSMunicIncidImpItemAlterado = iISSMunicIncidImpItemAux
    
End Sub

Public Sub ISSMunicIncidImpItem_Validate(Cancel As Boolean, Optional ByVal bAtualizaTrib As Boolean = True)

Dim lErro As Long
Dim sDescIBGE As String

On Error GoTo Erro_ISSMunicIncidImpItem_Validate

    If mvariISSMunicIncidImpItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ISSMunicIncidImpItem").Text)) > 0 Then

            lErro = Valor_Positivo_Critica(TabTrib_Controle("ISSMunicIncidImpItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            lErro = CF("Le_Campo_Tabela", "IBGEMunicipios", "DescIBGE", TIPO_STR, "CodIBGE", TabTrib_Controle("ISSMunicIncidImpItem").Text, sDescIBGE, TabTrib_Controle("ISSMunicIncidDesc"))
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            If lErro <> SUCESSO Then gError 200982
            
        Else
            
            TabTrib_Controle("ISSMunicIncidDesc").Caption = ""
            
        End If
        
'        lErro = ISSCodServItem_Trata_Descricao
'        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If bAtualizaTrib And mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
        
        mvariISSMunicIncidImpItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSMunicIncidImpItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 200982
            Call Rotina_Erro(vbOKOnly, "ERRO_MUNIC_CODIBGE_NAO_CADASTRADO", gErr, TabTrib_Controle("ISSMunicIncidImpItem").Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub

Public Sub ISSPaisLabel_Click()

Dim colSelecao As New Collection
Dim objPais As New ClassPais

On Error GoTo Erro_ISSPaisLabel_Click
    
    objPais.iCodigo = Codigo_Extrai(TabTrib_Controle("ISSCodPaisItem").Text)
   
    'Chama Tela TituloReceberLista
    Call Chama_Tela("PaisesLista", colSelecao, objPais, mvarobjEventoISSCodPais)

    Exit Sub

Erro_ISSPaisLabel_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200973)

    End Select

    Exit Sub
    
End Sub

Private Sub mvarobjEventoISSCodPais_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPais As New ClassPais

On Error GoTo Erro_mvarobjEventoISSCodPais_evSelecao

    Set objPais = obj1

    TabTrib_Controle("ISSCodPaisItem").Text = CStr(objPais.iCodigo)
    Call ISSCodPaisItem_Validate(bSGECancelDummy)

    Exit Sub

Erro_mvarobjEventoISSCodPais_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202968)

    End Select

    Exit Sub

End Sub

Public Sub ISSCodPaisItem_Change()
    mvariISSCodPaisItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSCodPaisItem_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    mvariISSCodPaisItemAlterado = REGISTRO_ALTERADO
    If mvariTrazendoTribItemTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub

Public Sub ISSCodPaisItem_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objPais As New ClassPais
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ISSCodPaisItem_Validate

    If mvariISSCodPaisItemAlterado = REGISTRO_ALTERADO Then
    
        lErro = Combo_Seleciona(TabTrib_Controle("ISSCodPaisItem"), iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

        'Nao existe o item com o CODIGO na List da ComboBox
        If lErro = 6730 Then
    
            objPais.iCodigo = iCodigo
    
            'Tenta ler Pais com esse codigo no BD
            lErro = CF("Paises_Le", objPais)
            If lErro <> SUCESSO And lErro <> 47876 Then gError ERRO_SEM_MENSAGEM
            If lErro <> SUCESSO Then gError 202971
    
            TabTrib_Controle("ISSCodPaisItem").Text = CStr(iCodigo) & SEPARADOR & objPais.sNome
    
        End If
    
        'Nao existe o item com a STRING na List da ComboBox
        If lErro = 6731 Then gError 202972
        
        If TabTrib_Controle("ISSCodPaisItem").ListIndex > 0 Then
        
            TabTrib_Controle("ISSCodIBGE").PromptInclude = False
            TabTrib_Controle("ISSCodIBGE").Text = ""
            TabTrib_Controle("ISSCodIBGE").PromptInclude = True
            Call ISSCodIBGE_Validate(bSGECancelDummy, False)
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariISSCodPaisItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSCodPaisItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 202971
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PAIS", objPais.iCodigo)
            If vbMsgRes = vbYes Then Call Chama_Tela("Paises", objPais)

        Case 202972
            Call Rotina_Erro(vbOKOnly, "ERRO_PAIS_NAO_CADASTRADO1", gErr, Trim(TabTrib_Controle("ISSCodPaisItem").Text))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub

End Sub

Public Sub ISSNumProcessoItem_Change()
    mvariISSNumProcessoItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSNumProcessoItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSNumProcessoItem_Validate

    If mvariISSNumProcessoItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("ISSNumProcessoItem").Text)) > 0 Then
        
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariISSNumProcessoItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ISSNumProcessoItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153749)

    End Select

    Exit Sub
    
End Sub
'--------------------------------NFe 3.10 --------------------------------------------

Public Sub NatOpInterna_AlterarTrocaTipoNF(ByVal sCFOP As String)
'alteracao de cfop no processo de troca de tipo de nota fiscal

    sNatOpInterna = sCFOP
    TabTrib_Controle("DescNatOpInterna").Caption = ""
    
End Sub

'###############################
'IBPTax 0.0.9

Sub TotTribFed_Change()

    mvariTotTribFedAlterado = REGISTRO_ALTERADO

End Sub

Sub TotTribFed_Validate(Cancel As Boolean)

Dim dTotTrib As Double
Dim lErro As Long

On Error GoTo Erro_TotTribFed_Validate

    If mvariTotTribFedAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("TotTribFed").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("TotTribFed").Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dTotTrib = CDbl(TabTrib_Controle("TotTribFed").Text)

        TabTrib_Controle("TotTribFed").Text = Format(dTotTrib, "Standard")

    End If

    mvariTotTribFedAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_TotTribFed_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153772)

    End Select

    Exit Sub

End Sub

Sub TotTribEst_Change()

    mvariTotTribEstAlterado = REGISTRO_ALTERADO

End Sub

Sub TotTribEst_Validate(Cancel As Boolean)

Dim dTotTrib As Double
Dim lErro As Long

On Error GoTo Erro_TotTribEst_Validate

    If mvariTotTribEstAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("TotTribEst").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("TotTribEst").Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dTotTrib = CDbl(TabTrib_Controle("TotTribEst").Text)

        TabTrib_Controle("TotTribEst").Text = Format(dTotTrib, "Standard")

    End If

    mvariTotTribEstAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_TotTribEst_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153772)

    End Select

    Exit Sub

End Sub

Sub TotTribMunic_Change()

    mvariTotTribMunicAlterado = REGISTRO_ALTERADO

End Sub

Sub TotTribMunic_Validate(Cancel As Boolean)

Dim dTotTrib As Double
Dim lErro As Long

On Error GoTo Erro_TotTribMunic_Validate

    If mvariTotTribMunicAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    If Len(Trim(TabTrib_Controle("TotTribMunic").Text)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("TotTribMunic").Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dTotTrib = CDbl(TabTrib_Controle("TotTribMunic").Text)

        TabTrib_Controle("TotTribMunic").Text = Format(dTotTrib, "Standard")

    End If

    mvariTotTribMunicAlterado = 0
    
    Call AtualizarTabTrib

    Exit Sub

Erro_TotTribMunic_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153772)

    End Select

    Exit Sub

End Sub

Sub TotTribAliqFedItem_Change()

    mvariTotTribFedItemAlterado = 1
    
End Sub

Sub TotTribAliqFedItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_TotTribAliqFedItem_Validate

    If mvariTotTribFedItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("TotTribAliqFedItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("TotTribAliqFedItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariTotTribFedItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_TotTribAliqFedItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub

Sub TotTribAliqEstItem_Change()

    mvariTotTribEstItemAlterado = 1
    
End Sub

Sub TotTribAliqEstItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_TotTribAliqEstItem_Validate

    If mvariTotTribEstItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("TotTribAliqEstItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("TotTribAliqEstItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariTotTribEstItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_TotTribAliqEstItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub

Sub TotTribAliqMunicItem_Change()

    mvariTotTribMunicItemAlterado = 1
    
End Sub

Sub TotTribAliqMunicItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_TotTribAliqMunicItem_Validate

    If mvariTotTribMunicItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("TotTribAliqMunicItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("TotTribAliqMunicItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariTotTribMunicItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_TotTribAliqMunicItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub
'################################

Public Sub ICMSInterestPercFCPUFDestItem_Change()
    mvariICMSInterestPercFCPUFDestItemAlterado = MARCADO
End Sub

Public Sub ICMSInterestPercFCPUFDestItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSInterestPercFCPUFDestItem_Validate

    If mvariICMSInterestPercFCPUFDestItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSInterestPercFCPUFDestItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSInterestPercFCPUFDestItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSInterestPercFCPUFDestItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSInterestPercFCPUFDestItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153742)

    End Select
    
    Exit Sub
    
End Sub

Public Sub ICMSInterestBCUFDestItem_Change()
    mvariICMSInterestBCUFDestItemAlterado = MARCADO
End Sub

Public Sub ICMSInterestBCUFDestItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSInterestBCUFDestItem_Validate

    If mvariICMSInterestBCUFDestItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSInterestBCUFDestItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSInterestBCUFDestItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSInterestBCUFDestItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSInterestBCUFDestItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSInterestAliqUFDestItem_Change()
    mvariICMSInterestAliqUFDestItemAlterado = MARCADO
End Sub

Public Sub ICMSInterestAliqUFDestItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSInterestAliqUFDestItem_Validate

    If mvariICMSInterestAliqUFDestItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSInterestAliqUFDestItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSInterestAliqUFDestItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSInterestAliqUFDestItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSInterestAliqUFDestItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153742)

    End Select
    
    Exit Sub
    
End Sub

Public Sub ICMSInterestAliqItem_Click()
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Public Sub ICMSInterestPercPartilhaItem_Click()
    If mvariTrazendoTribItemTela = 0 Then Call AtualizarTabTribItem
End Sub

Public Sub ICMSInterestVlrUFDestItem_Change()
    mvariICMSInterestVlrUFDestItemAlterado = MARCADO
End Sub

Public Sub ICMSInterestVlrUFDestItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSInterestVlrUFDestItem_Validate

    If mvariICMSInterestVlrUFDestItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSInterestVlrUFDestItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSInterestVlrUFDestItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSInterestVlrUFDestItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSInterestVlrUFDestItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSInterestVlrUFRemetItem_Change()
    mvariICMSInterestVlrUFRemetItemAlterado = MARCADO
End Sub

Public Sub ICMSInterestVlrUFRemetItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSInterestVlrUFRemetItem_Validate

    If mvariICMSInterestVlrUFRemetItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSInterestVlrUFRemetItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSInterestVlrUFRemetItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSInterestVlrUFRemetItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSInterestVlrUFRemetItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSInterestVlrFCPUFDestItem_Change()
    mvariICMSInterestVlrFCPUFDestItemAlterado = MARCADO
End Sub

Public Sub ICMSInterestVlrFCPUFDestItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSInterestVlrFCPUFDestItem_Validate

    If mvariICMSInterestVlrFCPUFDestItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSInterestVlrFCPUFDestItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSInterestVlrFCPUFDestItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSInterestVlrFCPUFDestItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSInterestVlrFCPUFDestItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Sub IPICodEnq_GotFocus(iAlterado As Integer)
    Call MaskEdBox_TrataGotFocus(TabTrib_Controle("IPICodEnq"), iAlterado)
End Sub

Private Function IPICodEnq_Valida() As Long

Dim lErro As Long
Dim objIPICodEnq As New ClassIPICodEnquadramento

On Error GoTo Erro_IPICodEnq_Valida

    If Len(Trim(TabTrib_Controle("IPICodEnq").Text)) <> 0 Then

        objIPICodEnq.sCodigo = Format(Trim(TabTrib_Controle("IPICodEnq").Text), "000")
        
        lErro = CF("IPICodEnquadramento_Le", objIPICodEnq)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 216110 'Não cadastrado
        
        If TabTrib_Controle("ComboIPITipo").ListIndex = -1 Then gError 216111 'Tem que preencher o CST do IPI
        
        If TabTrib_Controle("ComboIPITipo").ItemData(TabTrib_Controle("ComboIPITipo").ListIndex) <> objIPICodEnq.iTipoIPI Then gError 216112 'Tipo incompatível
        
        TabTrib_Controle("IPICodEnqDesc").Caption = objIPICodEnq.sDescCompleta
        
    Else
        TabTrib_Controle("IPICodEnqDesc").Caption = ""
    End If

    IPICodEnq_Valida = SUCESSO

    Exit Function

Erro_IPICodEnq_Valida:

    IPICodEnq_Valida = gErr

    Select Case gErr
    
        Case 216110
            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODENQUADRAMENTO_NAO_CADASTRADO", gErr, objIPICodEnq.sCodigo)
                
        Case 216111
            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODENQ_TIPOIPI_NAO_CADASTRADO", gErr)
        
        Case 216112
            Call Rotina_Erro(vbOKOnly, "ERRO_IPICODENQ_TIPOIPI_INCOMPATIVEL", gErr)
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216109)

    End Select
    
    If TabTrib_Controle("IPICodEnq").Visible Then Call TabTrib_Controle("IPICodEnq").SetFocus

    
    Exit Function
    
End Function

Sub IPICodEnqLabel_Click()

Dim colSelecao As New Collection
Dim objIPICodEnq As New ClassIPICodEnquadramento

    If TabTrib_Controle("ComboIPITipo").ListIndex <> -1 Then

        colSelecao.Add TabTrib_Controle("ComboIPITipo").ItemData(TabTrib_Controle("ComboIPITipo").ListIndex)
        
        objIPICodEnq.sCodigo = Trim(TabTrib_Controle("IPICodEnq").Text)
        
        'Chama a tela de browse
        Call Chama_Tela("IPICodEnquadramentoLista", colSelecao, objIPICodEnq, mvarobjEventoIPICodEnq, "TipoIPI = ?")

    End If
    
End Sub

Private Sub mvarobjEventoIPICodEnq_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objIPICodEnq As ClassIPICodEnquadramento

On Error GoTo Erro_mvarobjEventoIPICodEnq_evSelecao

    Set objIPICodEnq = obj1

    TabTrib_Controle("IPICodEnq").Text = objIPICodEnq.sCodigo
    Call IPICodEnq_Validate(bSGECancelDummy)

    Exit Sub

Erro_mvarobjEventoIPICodEnq_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 174795)

    End Select

    Exit Sub

End Sub

Private Function Trata_Comercio_Exterior() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTributoItem As Object

On Error GoTo Erro_Trata_Comercio_Exterior

    For iIndice = 1 To gcolDocItem(gobjDocTrib).Count
    
        Set objTributoItem = gobjDocTribItem(iIndice)
    
        lErro = Trata_Comercio_Exterior_Item(objTributoItem)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Next

    Trata_Comercio_Exterior = SUCESSO

    Exit Function

Erro_Trata_Comercio_Exterior:

    Trata_Comercio_Exterior = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216231)

    End Select

    Exit Function
    
End Function

Private Function Trata_Comercio_Exterior_Item(ByVal objTributacaoItem As Object) As Long

Dim lErro As Long, sNCM As String
Dim sUMTrib As String, dFator As Double
Dim objProduto As New ClassProduto
Dim objUnidade As New ClassUnidadeDeMedida

On Error GoTo Erro_Trata_Comercio_Exterior_Item

    'Se é uma nota de comércio exterior
    If left(objTributacaoItem.sNaturezaOp, 1) = "7" Or _
        objTributacaoItem.sNaturezaOp = "1501" Or _
        objTributacaoItem.sNaturezaOp = "2501" Or _
        objTributacaoItem.sNaturezaOp = "5501" Or _
        objTributacaoItem.sNaturezaOp = "5502" Or _
        objTributacaoItem.sNaturezaOp = "5504" Or _
        objTributacaoItem.sNaturezaOp = "5505" Or _
        objTributacaoItem.sNaturezaOp = "6501" Or _
        objTributacaoItem.sNaturezaOp = "6502" Or _
        objTributacaoItem.sNaturezaOp = "6504" Or _
        objTributacaoItem.sNaturezaOp = "6505" Then
        
        dFator = 1
        objProduto.sCodigo = objTributacaoItem.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        sNCM = objTributacaoItem.sIPICodProduto
        If Len(Trim(sNCM)) = 0 Then sNCM = objProduto.sIPICodigo
        
        lErro = CF("NCM_Le_UMTrib", sNCM, sUMTrib)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If UCase(objTributacaoItem.sUMTrib) <> UCase(left(sUMTrib, STRING_UM_SIGLA)) Then
        
            objUnidade.iClasse = objProduto.iClasseUM
            objUnidade.sSigla = left(sUMTrib, STRING_UM_SIGLA)
        
            'Lê objUnidade2
            lErro = CF("UM_Le", objUnidade)
            If lErro <> SUCESSO And lErro <> 23775 Then gError ERRO_SEM_MENSAGEM
        
            If lErro = 23775 Then gError 216229
            
            'Faz a conversão da UM da tela para a UM tributavel
            lErro = CF("UM_Conversao", objProduto.iClasseUM, objTributacaoItem.sUMTrib, objUnidade.sSigla, dFator)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objTributacaoItem.dQtdTrib = objTributacaoItem.dQtdTrib * dFator
            objTributacaoItem.sUMTrib = sUMTrib
            
        End If
        
        If objTributacaoItem.dPrecoTotal = 0 Or objTributacaoItem.dQtdTrib = 0 Then
            objTributacaoItem.dValorUnitTrib = 0
        Else
            objTributacaoItem.dValorUnitTrib = Arredonda_Moeda((objTributacaoItem.dPrecoTotal + objTributacaoItem.dValorDescontoItem) / objTributacaoItem.dQtdTrib, 4)
        End If
            
    End If
    
    Trata_Comercio_Exterior_Item = SUCESSO

    Exit Function

Erro_Trata_Comercio_Exterior_Item:

    Trata_Comercio_Exterior_Item = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 216229
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_UMTRIB_NAO_CADAST_CLASSEUM", gErr, objUnidade.sSigla, sNCM, objProduto.iClasseUM, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216230)

    End Select

    Exit Function
    
End Function

Public Sub cBenefItem_Change()
    mvaricBenefItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub cBenefItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_cBenefItem_Validate

    If mvaricBenefItemAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(TabTrib_Controle("cBenefItem").Text)) > 0 Then
        
            
        End If
        
        Call AtualizarTabTribItem
        
        mvaricBenefItemAlterado = 0
    
    End If
    
    Exit Sub
    
Erro_cBenefItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSInterestBCFCPUFDestItem_Change()
    mvariICMSInterestBCFCPUFDestItemAlterado = MARCADO
End Sub

Public Sub ICMSInterestBCFCPUFDestItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSInterestBCFCPUFDestItem_Validate

    If mvariICMSInterestBCFCPUFDestItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSInterestBCFCPUFDestItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSInterestBCFCPUFDestItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSInterestBCFCPUFDestItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSInterestBCFCPUFDestItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSvBCFCPItem_Change()
    mvariICMSvBCFCPItemAlterado = MARCADO
End Sub

Public Sub ICMSvBCFCPItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSvBCFCPItem_Validate

    If mvariICMSvBCFCPItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSvBCFCPItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvBCFCPItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSvBCFCPItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSvBCFCPItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSpFCPItem_Change()
    mvariICMSpFCPItemAlterado = MARCADO
End Sub

Public Sub ICMSpFCPItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_ICMSpFCPItem_Validate

    If mvariICMSpFCPItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSpFCPItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSpFCPItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSpFCPItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSpFCPItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub

Public Sub ICMSvFCPItem_Change()
    mvariICMSvFCPItemAlterado = MARCADO
End Sub

Public Sub ICMSvFCPItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSvFCPItem_Validate

    If mvariICMSvFCPItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSvFCPItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvFCPItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSvFCPItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSvFCPItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSvBCFCPSTItem_Change()
    mvariICMSvBCFCPSTItemAlterado = MARCADO
End Sub

Public Sub ICMSvBCFCPSTItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSvBCFCPSTItem_Validate

    If mvariICMSvBCFCPSTItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSvBCFCPSTItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvBCFCPSTItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSvBCFCPSTItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSvBCFCPSTItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSpFCPSTItem_Change()
    mvariICMSpFCPSTItemAlterado = MARCADO
End Sub

Public Sub ICMSpFCPSTItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_ICMSpFCPSTItem_Validate

    If mvariICMSpFCPSTItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSpFCPSTItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSpFCPSTItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSpFCPSTItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSpFCPSTItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub

Public Sub ICMSvFCPSTItem_Change()
    mvariICMSvFCPSTItemAlterado = MARCADO
End Sub

Public Sub ICMSvFCPSTItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSvFCPSTItem_Validate

    If mvariICMSvFCPSTItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSvFCPSTItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvFCPSTItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSvFCPSTItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSvFCPSTItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSvBCFCPSTRetItem_Change()
    mvariICMSvBCFCPSTRetItemAlterado = MARCADO
End Sub

Public Sub ICMSvBCFCPSTRetItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSvBCFCPSTRetItem_Validate

    If mvariICMSvBCFCPSTRetItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSvBCFCPSTRetItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvBCFCPSTRetItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSvBCFCPSTRetItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSvBCFCPSTRetItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSpFCPSTRetItem_Change()
    mvariICMSpFCPSTRetItemAlterado = MARCADO
End Sub

Public Sub ICMSpFCPSTRetItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_ICMSpFCPSTRetItem_Validate

    If mvariICMSpFCPSTRetItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSpFCPSTRetItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSpFCPSTRetItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSpFCPSTRetItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSpFCPSTRetItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub

Public Sub ICMSvFCPSTRetItem_Change()
    mvariICMSvFCPSTRetItemAlterado = MARCADO
End Sub

Public Sub ICMSvFCPSTRetItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ICMSvFCPSTRetItem_Validate

    If mvariICMSvFCPSTRetItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("ICMSvFCPSTRetItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("ICMSvFCPSTRetItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSvFCPSTRetItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSvFCPSTRetItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153743)

    End Select

    Exit Sub
    
End Sub

Public Sub ICMSSTCobrAntAliquotaItem_Change()
    mvariICMSSTCobrAntAliquotaItemAlterado = MARCADO
End Sub

Public Sub ICMSSTCobrAntAliquotaItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_ICMSSTCobrAntAliquotaItem_Validate

    If mvariICMSSTCobrAntAliquotaItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("ICMSSTCobrAntAliquotaItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("ICMSSTCobrAntAliquotaItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvariICMSSTCobrAntAliquotaItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_ICMSSTCobrAntAliquotaItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub

'UA. Tributos Devolvidos (para o item da NF-e)
Sub IPIVlrDevolvidoItem_Change()
    mvariIPIVlrDevolvidoItemAlterado = REGISTRO_ALTERADO
End Sub

Public Sub IPIVlrDevolvidoItem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_IPIVlrDevolvidoItem_Validate

    If mvariIPIVlrDevolvidoItemAlterado = MARCADO Then
    
        If Len(Trim(TabTrib_Controle("IPIVlrDevolvidoItem").Text)) <> 0 Then
        
            lErro = Valor_NaoNegativo_Critica(TabTrib_Controle("IPIVlrDevolvidoItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        End If
        
        Call AtualizarTabTribItem
        
        mvariIPIVlrDevolvidoItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_IPIVlrDevolvidoItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153743)

    End Select

    Exit Sub

End Sub

Public Sub pDevolItem_Change()
    mvaripDevolItemAlterado = MARCADO
End Sub

Public Sub pDevolItem_Validate(Cancel As Boolean)
        
Dim lErro As Long

On Error GoTo Erro_pDevolItem_Validate

    If mvaripDevolItemAlterado Then
    
        If Len(Trim(TabTrib_Controle("pDevolItem").Text)) > 0 Then
        
            lErro = Porcentagem_Critica(TabTrib_Controle("pDevolItem").Text)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
        
        Call AtualizarTabTribItem
        
        mvaripDevolItemAlterado = 0
    
    End If

    Exit Sub
    
Erro_pDevolItem_Validate:

    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 153742)

    End Select

    Exit Sub

End Sub

Public Sub ICMSSTBaseDupla_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub


Public Sub ICMSInterestBaseDupla_Click(Optional ByVal bAtualizaTrib As Boolean = True)
    If mvariTrazendoTribTela = 0 And bAtualizaTrib Then Call AtualizarTabTribItem
End Sub
