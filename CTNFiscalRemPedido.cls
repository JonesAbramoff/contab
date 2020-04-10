VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFiscalRemPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Unload()

Public gobjAnotacao As ClassAnotacoes

Public giValorDescontoManual As Integer 'Alterado por Wagner
Public gobjInfoUsu As Object 'Inserido por Wagner
Public iValorDescontoAlterado As Integer 'Inserido por Wagner

Public objCTNFiscalPedido As CTNFiscalPedido
Public objCTNFiscalRem As CTNFiscalRem

Public WithEvents objEventoPedido As AdmEvento
Attribute objEventoPedido.VB_VarHelpID = -1
Public WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Public WithEvents objEventoFornEntTerc As AdmEvento
Attribute objEventoFornEntTerc.VB_VarHelpID = -1

Public gobjRastreamento As ClassRastreamento

'armazena o último tipo de NFiscal anterior ao selecionado no momento
Dim giTipoDocInfo As Integer

'armazena os itens da NFiscal e principalmente a Alocação de almoxarifado
Dim gcolItensNf As ColItensNF
Public gobjNFiscal As ClassNFiscal
Dim giComplemento As Integer

Dim iFornEntTercAlterado As Integer

'mnemônicos
Private Const TIPO_NOTA_FISCAL As String = "Tipo_Nota_Fiscal"
Private Const NATUREZA_OP As String = "Natureza_OP"
Private Const CLIENTE1 As String = "Cliente"
Private Const CLIENTE_CODIGO As String = "Cliente_Codigo"
Private Const FILIAL1 As String = "Filial_Cliente"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const Serie1 As String = "Serie"
Private Const DATA_EMISSAO As String = "Data_Emissao"
Private Const DATA_SAIDA As String = "Data_Saida"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Medida"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const PRECO_UNITARIO As String = "Preco_Unitario"
Private Const PRECO_TOTAL As String = "Preco_Total"
Private Const DESCONTO1 As String = "Desconto_Produto"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const FORNECEDOR1 As String = "Fornecedor"
Private Const FILIAL_FORNECEDOR1 As String = "Filial_Fornecedor"
Private Const CTACONTABILEST1 As String = "ContaContabilEst"
'###########Almoxarifado############
Private Const PRODUTO_ALMOX As String = "Produto_Almox"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const QUANT_ALOCADA As String = "Quant_Alocada"
Private Const UNIDADE_MED_EST As String = "Unidade_Med_Est"
'####### Valores ####################
Private Const ICMS As String = "ICMS_Valor"

Private Const ICMS_DIFAL_REMETENTE As String = "ICMS_Difal_Remetente"
Private Const ICMS_DIFAL_DESTINO As String = "ICMS_Difal_Destino"
Private Const ICMS_FCP As String = "ICMS_FCP_Valor"

Private Const ICMSSUBST As String = "ICMSSubst_Valor"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_DESPESAS As String = "Valor_Despesas"
Private Const IPI As String = "IPI_Valor"
Private Const VALOR_DESCONTO As String = "Valor_Desconto_Glb"
Private Const VALOR_TOTAL As String = "Valor_Total"
'########### Tributação ##############
Private Const ISS_VALOR As String = "ISS_Valor"
Private Const ISS_INCLUSO As String = "ISS_Incluso"
Private Const VALOR_IRRF As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
Private Const PIS_DEBITO_TOTAL As String = "PIS_Debito_Total"
Private Const COFINS_DEBITO_TOTAL As String = "COFINS_Debito_Total"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
Private Const CONTACONTABILAPLIC As String = "ContaContabilAplic"
Private Const CONTACONTABILCATPROD As String = "ContaContabilCatProd"
Private Const PIS_VALOR_TOTAL As String = "PIS_Valor_Total"
Private Const COFINS_VALOR_TOTAL As String = "COFINS_Valor_Total"

Private Const NOME_TELA_NFISCAL_REM_PEDIDO = "NFiscalRemPedido"

Public Property Get BotaoGrade() As Object
     Set BotaoGrade = objUserControl.Controls("BotaoGrade")
End Property

Public Property Get SerieNFiscalOriginal() As Object
     Set SerieNFiscalOriginal = objUserControl.Controls("SerieNFiscalOriginal")
End Property

Public Property Get NFiscalOriginal() As Object
     Set NFiscalOriginal = objUserControl.Controls("NFiscalOriginal")
End Property

Public Property Get NFiscalOriginalLabel() As Object
     Set NFiscalOriginalLabel = objUserControl.Controls("NFiscalOriginalLabel")
End Property

Public Property Get SerieNFOriginalLabel() As Object
     Set SerieNFOriginalLabel = objUserControl.Controls("SerieNFOriginalLabel")
End Property

Public Property Get FrameFornEntTerc() As Object
     Set FrameFornEntTerc = objUserControl.Controls("FrameFornEntTerc")
End Property

Public Property Get FilialFornEntTerc() As Object
     Set FilialFornEntTerc = objUserControl.Controls("FilialFornEntTerc")
End Property

Public Property Get FornEntTerc() As Object
     Set FornEntTerc = objUserControl.Controls("FornEntTerc")
End Property

Public Property Get ValorProdutos2() As Object
     Set ValorProdutos2 = objUserControl.Controls("ValorProdutos2")
End Property

Public Property Get ValorDescontoItens() As Object
     Set ValorDescontoItens = objUserControl.Controls("ValorDescontoItens")
End Property

Public Property Get PercDescontoItens() As Object
     Set PercDescontoItens = objUserControl.Controls("PercDescontoItens")
End Property

Public Property Get ISSBase1() As Object
     Set ISSBase1 = objUserControl.Controls("ISSBase1")
End Property

Public Property Get PrecoTotalB() As Object
     Set PrecoTotalB = objUserControl.Controls("PrecoTotalB")
End Property

Public Property Get MsgAutomatica() As Object
     Set MsgAutomatica = objCTNFiscalPedido.MsgAutomatica
End Property

Public Property Get ValorProdutos() As Object
     Set ValorProdutos = objCTNFiscalPedido.ValorProdutos
End Property

Public Property Get ICMSSubstValor1() As Object
     Set ICMSSubstValor1 = objCTNFiscalPedido.ICMSSubstValor1
End Property

Public Property Get IPIValor1() As Object
     Set IPIValor1 = objCTNFiscalPedido.IPIValor1
End Property

Public Property Get ValorFrete() As Object
     Set ValorFrete = objCTNFiscalPedido.ValorFrete
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objCTNFiscalPedido.ValorSeguro
End Property

Public Property Get ValorDespesas() As Object
     Set ValorDespesas = objCTNFiscalPedido.ValorDespesas
End Property

Public Property Get gobjTribTab() As Object
     Set gobjTribTab = objCTNFiscalPedido.gobjTribTab
End Property

Public Property Get gbCarregandoTela() As Boolean
     gbCarregandoTela = objCTNFiscalPedido.gbCarregandoTela
End Property

Public Function Preenche_Itens_Projeto_Aux(ByVal colItensPRJ As Collection, ByVal sItem As String, ByVal sDescricao As String, ByVal dValor As Double, ByVal iTipo As Integer, Optional ByVal iNegativo As Integer = 0, Optional ByVal sOBS As String = "") As Long
    Preenche_Itens_Projeto_Aux = objCTNFiscalPedido.Preenche_Itens_Projeto_Aux(colItensPRJ, sItem, sDescricao, dValor, iTipo, iNegativo, sOBS)
End Function

Public Function Name() As String

    Name = "NFiscalRemPedido"

End Function

Public Sub TipoNFiscal_Click()

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_TipoNFiscal_Click

    If TipoNFiscal.ListIndex = -1 Then Exit Sub

    objCTNFiscalPedido.iAlterado = REGISTRO_ALTERADO

    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    gobjRastreamento.iCodigo = Codigo_Extrai(TipoNFiscal.Text)

    'Lê o Tipo de Documento
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 91352

    'Se não achou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 91353
    
    If objTipoDocInfo.iComplementar = DOCINFO_COMPLEMENTO Then
        giComplemento = MARCADO
    Else
        giComplemento = DESMARCADO
    End If
    
    If objTipoDocInfo.iCodigo = DOCINFO_NFISRETPV Then
        FrameFornEntTerc.Enabled = True
    Else
        FrameFornEntTerc.Enabled = False
        FornEntTerc.Text = ""
        FilialFornEntTerc.Clear
    End If

    'aquiiii --> Jones está segurando referencia para objeto anterior
    ''*** incluidos p/tratamento de tributacao *******************************
    If objTipoDocInfo.iCodigo = DOCINFO_NFISSRPVEF Then Call objCTNFiscalPedido.gobjTribTab.NatOpInterna_Automatica
    Call objCTNFiscalPedido.gobjTribTab.TipoNFiscal_Definir(objTipoDocInfo.iCodigo, objTipoDocInfo.sSigla)
    If IsDate(objCTNFiscalPedido.DataEmissao.Text) Then Call objCTNFiscalPedido.gobjTribTab.DataEmissao_Alterada(StrParaDate(objCTNFiscalPedido.DataEmissao.Text))
    '*** fim tributacao

    'Verifica a Origem do Tipo da Nota Fiscal
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
        Call Habilita_Cliente
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        Call Habilita_Fornecedor
    End If

    'Se o tipo selecionado for benef. 3º
    If objTipoDocInfo.iCodigo = DOCINFO_NFISRMB3PV And giTipoDocInfo <> DOCINFO_NFISRMB3PV Then
            
        'Limpa o Grid de Alocações
        Call Grid_Limpa(objCTNFiscalPedido.objGridAlocacoes)
        objCTNFiscalPedido.objGridAlocacoes.iLinhasExistentes = 0
        
        'Atualiza o tipo de NFiscal
        giTipoDocInfo = objTipoDocInfo.iCodigo
                        
    'Se não, se o tipo selecionado for <> de benef. 3º e o último tipo selecionado <> benef. 3º
    ElseIf objTipoDocInfo.iCodigo <> DOCINFO_NFISRMB3PV And giTipoDocInfo = DOCINFO_NFISRMB3PV Then
                            
        Set objNFiscal = New ClassNFiscal
        Set objNFiscal.ColItensNF = gcolItensNf
                                    
        'Preenche o grid alocação com os valores da coleção
        Call CF2(objCTNFiscalPedido, "NFP_Preenche_GridAlocacoes", objNFiscal)
        
        'Atualiza o tipo de NFiscal
        giTipoDocInfo = objTipoDocInfo.iCodigo
        
    End If

    'Processa a alteração do subtipo para que sejam carregas os modelos contábeis referentes ao novo TipoDocInfo selecionao
    lErro = objCTNFiscalPedido.objContabil.Contabil_Processa_Alteracao_Subtipo(objTipoDocInfo.iSubTipoContabil)
    If lErro <> SUCESSO Then gError 91354

    Exit Sub

Erro_TipoNFiscal_Click:

    Select Case gErr

        Case 91352, 91354

        Case 91353
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157615)

    End Select

    Exit Sub

End Sub

Public Sub Habilita_Cliente()

Dim bCancel As Boolean
    
    ClienteLabel.Visible = True
    Cliente.Visible = True
    FornecedorLabel.Visible = False
    Fornecedor.Visible = False
    objCTNFiscalRem.iFornecedorAtivo = FORN_INATIVO
    objCTNFiscalRem.iClienteAlterado = REGISTRO_ALTERADO

    Call CF2(Me, "NFCF_Limpa_Tela")

End Sub

Public Sub Habilita_Fornecedor()
'Desabilita o Cliente e habilita o Fornecedor
    
Dim bCancel As Boolean
    
    FornecedorLabel.Visible = True
    Fornecedor.Visible = True
    objCTNFiscalRem.iFornecedorAtivo = FORN_ATIVO
    Cliente.Visible = False
    ClienteLabel.Visible = False
    objCTNFiscalRem.iFornecedorAlterado = REGISTRO_ALTERADO

    Call CF2(Me, "NFCF_Limpa_Tela")

End Sub

Public Sub TipoNFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_TipoNFiscal_Validate

    'Verifica se o tipo está preenchido
    If Len(Trim(TipoNFiscal.Text)) = 0 Then Exit Sub

    'Verifica se foi selecionado
    If TipoNFiscal.ListIndex > -1 Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(TipoNFiscal, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 91355
    
    'Se não conseguir --> Erro
    If lErro <> SUCESSO Then gError 91356
    
    'Incluido por Leonardo em 18/01/02
                
    'Se o tipo selecionado for benef. 3º
    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFISRMB3PV Then
            
        'Limpa o Grid de Alocações
        Call Grid_Limpa(objGridAlocacoes)
        objGridAlocacoes.iLinhasExistentes = 0
                                    
        'Atualiza o tipo de NFiscal
        giTipoDocInfo = DOCINFO_NFISRMB3PV
                            
    'Se não, se o tipo selecionado for <> de benef. 3º e o último tipo selecionado = benef. 3º
    ElseIf Codigo_Extrai(TipoNFiscal.Text) <> DOCINFO_NFISRMB3PV And giTipoDocInfo = DOCINFO_NFISRMB3PV Then
                            
        Set objNFiscal = New ClassNFiscal
        Set objNFiscal.ColItensNF = gcolItensNf
                            
        'Preenche o grid alocação com os valores da coleção
        Call CF2(objCTNFiscalPedido, "NFP_Preenche_GridAlocacoes", objNFiscal)
        
        'Atualiza o tipo de NFiscal
        giTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
        
    End If
        
    'Leonardo até aqui
    
    If iCodigo = DOCINFO_NFISRETPV Then
        FrameFornEntTerc.Enabled = True
    Else
        FrameFornEntTerc.Enabled = False
        FornEntTerc.Text = ""
        FilialFornEntTerc.Clear
    End If
    
    Exit Sub

Erro_TipoNFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 91355

        Case 91356
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, TipoNFiscal.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157616)

    End Select

    Exit Sub

End Sub

Public Sub TipoNFiscal_Change()
    objCTNFiscalPedido.iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Carrega_TiposDocInfo() As Long
'Carrega na os Tipo de Documentos relacionados com a tela

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Carrega_TiposDocInfo

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
        If objTipoDocInfo.sNomeTelaNFiscal = Me.Name Then
            TipoNFiscal.AddItem CStr(objTipoDocInfo.iCodigo) & SEPARADOR & objTipoDocInfo.sNomeReduzido
            TipoNFiscal.ItemData(TipoNFiscal.NewIndex) = objTipoDocInfo.iCodigo
            'se for o tipo padrao, seleciona-o
            If objTipoDocInfo.iPadrao = Padrao Then
                giTipoDocInfo = objTipoDocInfo.iCodigo 'Incluido por Leonardo em 18/01/02
                TipoNFiscal.ListIndex = TipoNFiscal.NewIndex
            End If
        End If
    Next

    Carrega_TiposDocInfo = SUCESSO

    Exit Function

Erro_Carrega_TiposDocInfo:

    Carrega_TiposDocInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157617)
    
    End Select

    Exit Function

End Function

Public Property Get TipoNFiscal() As Object
     Set TipoNFiscal = objCTNFiscalPedido.objUserControl.Controls("TipoNFiscal")
End Property

Public Function Form_Load_Ocx() As Object

    'feito em virtude do redimensionamento da msg da nota p/ inpal
    Call CF("Redimensiona_Tam_MsgNF", objCTNFiscalPedido.Mensagem)
    
    objCTNFiscalPedido.Parent.HelpContextID = IDH_NF_PEDIDO_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    objCTNFiscalPedido.Caption = "Nota Fiscal Remessa Gerada a Partir de um Pedido"
    Call Form_Load

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    objCTNFiscalPedido.iFrameAtual = 1
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        BotaoGrade.Visible = False
        
        'sumir com abas de contabilidade e rastro
        objCTNFiscalPedido.TabStrip1.Tabs.Remove (7) 'TAB_RASTRO
        objCTNFiscalPedido.TabStrip1.Tabs.Remove (6) 'TAB_Contabilizacao
        
    End If
    
    objCTNFiscalPedido.MsgAutomatica.Value = vbChecked
    objCTNFiscalPedido.Mensagem.MaxLength = MENSAGEM_NF_MAXLENGTH
    objCTNFiscalPedido.MensagemCorpo.MaxLength = MENSAGEM_NF_MAXLENGTH

    'Carrega a Combo TipoFrete
    lErro = Carrega_TipoFrete()
    If lErro <> SUCESSO Then gError 207443


    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 207443

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157618)

    End Select

    objCTNFiscalPedido.iAlterado = 0

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
   
    Set gobjInfoUsu = Nothing
   
    'Libera as variáveis globais
    Set objEventoPedido = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoFornEntTerc = Nothing

    Set gcolItensNf = Nothing
    Set gobjNFiscal = Nothing
    Set gobjAnotacao = Nothing
    Set gobjRastreamento = Nothing
    Set gcolItensNf = Nothing
    Set objCTNFiscalPedido = Nothing
    Set objCTNFiscalRem = Nothing
    
End Sub

Public Function Trata_Parametros(Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFiscal Is Nothing) Then

        lErro = CargaPosFormLoad()
        If lErro <> SUCESSO Then gError 91357

        'Tenta ler a nota Fiscal passada por parametro
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 91358

        If lErro <> SUCESSO Then gError 91359

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal

        'Lê o Tipo de Documento com o Código de objNFiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 91360

        'Se não encontrar --> Erro
        If lErro = 31415 Then gError 91361

        If objTipoDocInfo.sNomeTelaNFiscal <> NOME_TELA_NFISCAL_REM_PEDIDO Then gError 91362

        'Traz a nota para a tela
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 91363

    End If

    objCTNFiscalPedido.iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 91358, 91360, 91363, 91357

        Case 91359
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntDoc)

        Case 91361
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objNFiscal.iTipoNFiscal)

        Case 91362
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODOC_DIFERENTE_NF_PEDIDO", gErr, objNFiscal.iTipoNFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157619)

    End Select

    objCTNFiscalPedido.iAlterado = 0

    Exit Function

End Function

Public Function CargaPosFormLoad() As Long

Dim lErro As Long
Dim iSubTipo As Integer
Dim bEscaninhoVisivel As Boolean

On Error GoTo Erro_CargaPosFormLoad

    If (objCTNFiscalPedido.giPosCargaOk = 0) Then
    
        'Inicializa as variáveis de Grid e de Browse
        Set objCTNFiscalPedido.objEventoMensagem = New AdmEvento
        Set objCTNFiscalPedido.objGridItens = New AdmGrid
        Set objCTNFiscalPedido.objGridAlocacoes = New AdmGrid

        'Grid da contabilidade
        Set objCTNFiscalPedido.objGrid1 = New AdmGrid

        Set objEventoPedido = New AdmEvento
        Set objEventoNFiscal = New AdmEvento
        Set objEventoFornEntTerc = New AdmEvento

        Set objCTNFiscalPedido.gobjTelaProjetoInfo = New ClassTelaPRJInfo
        Set objCTNFiscalPedido.gobjTelaProjetoInfo.objUserControl = objUserControl
        Set objCTNFiscalPedido.gobjTelaProjetoInfo.objTela = Me

        Set objCTNFiscalPedido.objEventoSerie = New AdmEvento
        Set objCTNFiscalPedido.objEventoNaturezaOp = New AdmEvento
        Set objCTNFiscalPedido.objEventoProduto = New AdmEvento
        Set objCTNFiscalPedido.objEventoEstoqueProd = New AdmEvento
        Set objCTNFiscalPedido.objEventoLocalizacao = New AdmEvento
        Set objCTNFiscalPedido.objEventoLote = New AdmEvento
        Set objCTNFiscalPedido.objEventoDoc = New AdmEvento
        Set objCTNFiscalRem.objEventoNFiscalOriginal = New AdmEvento
        Set objCTNFiscalRem.objEventoSerieOriginal = New AdmEvento

        Set objCTNFiscalPedido.gcolItemPedido = New Collection
        Set objCTNFiscalPedido.gcolItemQuant = New Collection
        Set objCTNFiscalPedido.gcolItemUM = New Collection
        
        'Rastreamento
        Set gobjRastreamento = New ClassRastreamento
        Set gobjRastreamento.objTela = Me
        gobjRastreamento.bSaidaMaterial = True

        Set gcolItensNf = New ColItensNF
        
        Set gobjNFiscal = New ClassNFiscal
        
        Set gobjNFiscal.ColItensNF = gcolItensNf
                                
        'p/permitir o redesenho da tela
        DoEvents

        objCTNFiscalPedido.gbCarregandoTela = True

        'Inicializa o grid de Alocações
        lErro = objCTNFiscalPedido.Inicializa_Grid_Alocacoes(objCTNFiscalPedido.objGridAlocacoes)
        If lErro <> SUCESSO Then gError 91373

        'Inicializa a parte de tributacao

        Call CF("NFP_GridItens_InicializaColunas", objCTNFiscalPedido)
        
        Set objCTNFiscalPedido.gobjTribTab = New ClassTribTab
        
        lErro = objCTNFiscalPedido.gobjTribTab.Ativar(Me)
        If lErro <> SUCESSO Then gError 91364
        
        Call objCTNFiscalPedido.gobjTribTab.TipoNFiscal_Definir(DOCINFO_NFISVPV, "NFISVPV")
        If IsDate(objCTNFiscalPedido.DataEmissao.Text) Then Call objCTNFiscalPedido.gobjTribTab.DataEmissao_Alterada(StrParaDate(objCTNFiscalPedido.DataEmissao.Text))
        
        'carregar a combo FilialPedido com as Filiais da empresa
        lErro = objCTNFiscalPedido.Carrega_FilialPedido()
        If lErro <> SUCESSO Then gError 91365

        'Carrega a combo de Tipos de Documento
        lErro = Carrega_TiposDocInfo()
        If lErro <> SUCESSO Then gError 91366

        'Carrega a combo de series
        lErro = objCTNFiscalRem.Carrega_Serie()
        If lErro <> SUCESSO Then gError 91367

        'Seleciona série padrão
        lErro = objCTNFiscalPedido.Seleciona_Serie_Padrao()
        If lErro <> SUCESSO Then gError 91368

        'Carrega a combo de tabelas de preços
        lErro = objCTNFiscalPedido.Carrega_TabelaPreco()
        If lErro <> SUCESSO Then gError 91369

        'Carrega a combo de transportadoras
        lErro = objCTNFiscalPedido.Carrega_Transportadora()
        If lErro <> SUCESSO Then gError 91370

        'Carrega a combo VolumeEspecie
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, objCTNFiscalPedido.VolumeEspecie)
        If lErro <> SUCESSO Then gError 102428
    
        'Carrega a combo VolumeMarca
        lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, objCTNFiscalPedido.VolumeMarca)
        If lErro <> SUCESSO Then gError 102429
        
        'Carrega a combo PlacaUF
        lErro = objCTNFiscalPedido.Carrega_PlacaUF()
        If lErro <> SUCESSO Then gError 91371

        'Inicializa a máscara do produto
        lErro = CF("Inicializa_Mascara_Produto_MaskEd", objCTNFiscalPedido.Produto)
        If lErro <> SUCESSO Then gError 91374

        'Inicializa o grid de itens
        lErro = Inicializa_Grid_Itens(objCTNFiscalPedido.objGridItens)
        If lErro <> SUCESSO Then gError 91372

        objCTNFiscalPedido.PrecoUnitario.Format = gobjFAT.sFormatoPrecoUnitario

        objCTNFiscalPedido.QuantAlocada.Format = FORMATO_ESTOQUE
        objCTNFiscalPedido.Quantidade.Format = FORMATO_ESTOQUE
        objCTNFiscalPedido.QuantVendida.Format = FORMATO_ESTOQUE

        'Inicializa a contabilidade carregando apenas os modelos contábeis referentes ao TipoDocInfo default
        If objCTNFiscalPedido.objContabil.objGridInt Is Nothing Then

            'Obtém o código do subtipo contábil referente ao TipoDocInfo default
            lErro = CF("Obtem_Subtipo_Contabil", "TiposDocInfo", DOCINFO_NFISSRPV, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 96088
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 96089
            
            'Inicialização da parte de contabilidade
            lErro = objCTNFiscalPedido.objContabil.Contabil_Inicializa_Contabilidade3(Me, objCTNFiscalPedido.objGrid1, objCTNFiscalPedido.objEventoLote, objCTNFiscalPedido.objEventoDoc, MODULO_FATURAMENTO, iSubTipo)
            If lErro <> SUCESSO Then gError 96090

        End If

        Set objCTNFiscalPedido.gobjRastreamento = gobjRastreamento
        
        'Inicializa o grid de Rastreamento
        lErro = gobjRastreamento.Inicializa_Grid_Rastreamento()
        If lErro <> SUCESSO Then gError 91376
        
        Call CF2(Me, "NFRP_CargaPosFormLoad")

        objCTNFiscalPedido.gbCarregandoTela = False

        objCTNFiscalPedido.iAlterado = 0
        iFornEntTercAlterado = 0
        objCTNFiscalPedido.iDescontoAlterado = 0
        
        objCTNFiscalPedido.giPosCargaOk = 1
        
        Call ValorTotal_Calcula
    
    End If

    CargaPosFormLoad = SUCESSO

    Exit Function

Erro_CargaPosFormLoad:

    objCTNFiscalPedido.gbCarregandoTela = False

    CargaPosFormLoad = gErr

    Select Case gErr

        Case 96089 'tratado na rotina chamadora

        Case 91364 To 91376, 96088, 96090, 102428, 102429
        
        Case 96079
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPODOCINFO_NAO_ENCONTRADO", gErr, DOCINFO_NFISSRPV)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157620)

    End Select

    Exit Function

End Function

Public Sub Form_Activate()

Dim lErro As Long

On Error GoTo Erro_Form_Activate

    lErro = CargaPosFormLoad
    If lErro <> SUCESSO Then gError 91377
    
    Call TelaIndice_Preenche(Me)

    Exit Sub

Erro_Form_Activate:

    Select Case gErr

        Case 91377

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157621)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_NFiscalRemPedido() As Long
'Limpa a tela de Notas Fiscais Remessa gerada a partir de um pedido

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_NFiscalRemPedido

    'Limpa os campos comuns da tela
    Call Limpa_Tela(Me)
    
    Set gobjAnotacao = Nothing
    
    Set objCTNFiscalPedido.gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set objCTNFiscalPedido.gobjTelaProjetoInfo.objUserControl = objUserControl
    Set objCTNFiscalPedido.gobjTelaProjetoInfo.objTela = Me
    
'    If giTipoVersao = VERSAO_FULL Then
        objCTNFiscalPedido.Serie.Text = ""
        Call objCTNFiscalPedido.Serie_Validate(bSGECancelDummy)
        objCTNFiscalPedido.TabelaPreco.Text = ""
'    End If
    
    Set objCTNFiscalPedido.gobjInfoAdic = Nothing
    
    'Limpa os campos não limpos pelo Limpa_Tela
    objCTNFiscalPedido.Status.Caption = ""
    objCTNFiscalPedido.NFiscal.Caption = ""
    objCTNFiscalPedido.Cliente.Caption = ""
    objCTNFiscalPedido.TipoFrete.ListIndex = 0
    objCTNFiscalPedido.Filial.Caption = ""
    objCTNFiscalPedido.FilialPedido.Text = giFilialEmpresa
    Call objCTNFiscalPedido.FilialPedido_Validate(bSGECancelDummy)
    objCTNFiscalPedido.ValorProdutos.Caption = ""
    objCTNFiscalPedido.ValorProdutos2.Caption = ""
    objCTNFiscalPedido.ValorTotal.Caption = ""
    objCTNFiscalPedido.FilialEntrega.Clear
    objCTNFiscalPedido.Transportadora.Text = ""
    objCTNFiscalPedido.TranspRedespacho.Text = ""
    objCTNFiscalPedido.PlacaUF.Text = ""
    objCTNFiscalPedido.EletronicaFed.Value = vbUnchecked
    
    'Incluído por Luiz Nogueira em 21/08/03
    objCTNFiscalPedido.VolumeMarca.Text = ""
    objCTNFiscalPedido.VolumeEspecie.Text = ""
    
    'Recarrega a combo VolumeEspecie e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEESPECIE, objCTNFiscalPedido.VolumeEspecie)

    'Recarrega a combo VolumeMarca e seleciona a opção padrão
    'Foi colocada aqui com o intuito de atualizar a combo e selecionar o padrão
    Call CF("Carrega_CamposGenericos", CAMPOSGENERICOS_VOLUMEMARCA, objCTNFiscalPedido.VolumeMarca)
    
    Call CF2(Me, "NFCF_DataSaida2", objCTNFiscalPedido)
    
    objCTNFiscalPedido.DataEmissao.PromptInclude = False
    objCTNFiscalPedido.DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    objCTNFiscalPedido.DataEmissao.PromptInclude = True
    Call objCTNFiscalPedido.DataEmissao_Validate(bSGECancelDummy)
    
    objCTNFiscalPedido.giFaturaIntegral = 0
    objCTNFiscalPedido.glCodPedido = 0
    objCTNFiscalPedido.giFilialPedido = giFilialEmpresa
    
    For iIndice = 1 To objCTNFiscalPedido.objGridItens.iLinhasExistentes
        GridItens.TextMatrix(iIndice, 0) = iIndice
    Next
    
    'Limpa os grids
    Call Grid_Limpa(objCTNFiscalPedido.objGridItens)
    Call Grid_Limpa(objCTNFiscalPedido.objGridAlocacoes)
    
    Call objCTNFiscalPedido.gobjTribTab.Limpa_Tela
        
    'Limpa a contabilidade
    Call objCTNFiscalPedido.objContabil.Contabil_Limpa_Contabilidade
        
    'Limpa o Frame de Rastreamento
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    Call Seleciona_Tipo_Padrao
    
    'Seleciona Série padrão
    lErro = objCTNFiscalPedido.Seleciona_Serie_Padrao()
    If lErro <> SUCESSO Then gError 91378

    MsgAutomatica.Value = vbChecked

    objCTNFiscalPedido.iValorDescontoAlterado = 0
    objCTNFiscalPedido.iAlterado = 0
    objCTNFiscalPedido.iValorDespesasAlterado = 0
    objCTNFiscalPedido.iValorFreteAlterado = 0
    objCTNFiscalPedido.iValorSeguroAlterado = 0
    objCTNFiscalPedido.giFilialEntregaAlterada = 0
    objCTNFiscalPedido.giFilialEntregaManual = 0
    objCTNFiscalPedido.iDescontoAlterado = 0
    objCTNFiscalPedido.dValorDescontoItensAnt = 0
    objCTNFiscalPedido.dPercDescontoItensAnt = 0
    
    FilialFornEntTerc.Clear
 
    Call CF2(Me, "NFCF_Limpa_Tela")
    Call CF2(Me, "NFRP_Limpa_Tela")
 
    giTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
 
    Exit Function
    
Erro_Limpa_Tela_NFiscalRemPedido:
    
    Limpa_Tela_NFiscalRemPedido = gErr
    
    Select Case gErr
    
        Case 91378
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157622)
    
    End Select
    
    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim objNFiscal As New ClassNFiscal
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, objCTNFiscalPedido.iAlterado)
    If lErro <> SUCESSO Then gError 91379

    lErro = Limpa_Tela_NFiscalRemPedido()
    If lErro <> SUCESSO Then gError 91380
    
    'Limpa as coleções globais
    Set objCTNFiscalPedido.gcolItemPedido = New Collection
    Set objCTNFiscalPedido.gcolItemQuant = New Collection
    Set objCTNFiscalPedido.gcolItemUM = New Collection
    
    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 91379, 91380 'tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157623)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()
  
  'Fecha a tela
   Unload Me

End Sub

Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFiscalRemPedido"
    
    'Lê os dados da Tela NFiscal
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 91381

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objNFiscal.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "TipoNFiscal", objNFiscal.iTipoNFiscal, 0, "TipoNFiscal"
    colCampoValor.Add "NaturezaOp", objNFiscal.sNaturezaOp, STRING_BUFFER_MAX_TEXTO, "NaturezaOp"
    colCampoValor.Add "Serie", objNFiscal.sSerie, STRING_BUFFER_MAX_TEXTO, "Serie"
    colCampoValor.Add "NumNotaFiscal", objNFiscal.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFiscal.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataSaida", objNFiscal.dtDataSaida, 0, "DataSaida"
    colCampoValor.Add "HoraSaida", CDbl(objNFiscal.dtHoraSaida), 0, "HoraSaida"
    colCampoValor.Add "ValorProdutos", objNFiscal.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorFrete", objNFiscal.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "ValorSeguro", objNFiscal.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorOutrasDespesas", objNFiscal.dValorOutrasDespesas, 0, "ValorOutrasDespesas"
    colCampoValor.Add "ValorDesconto", objNFiscal.dValorDesconto, 0, "ValorDesconto"
    colCampoValor.Add "ValorTotal", objNFiscal.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "Placa", objNFiscal.sPlaca, STRING_BUFFER_MAX_TEXTO, "Placa"
    colCampoValor.Add "PlacaUF", objNFiscal.sPlacaUF, STRING_BUFFER_MAX_TEXTO, "PlacaUF"
    colCampoValor.Add "VolumeQuant", objNFiscal.lVolumeQuant, 0, "VolumeQuant"
    colCampoValor.Add "VolumeEspecie", objNFiscal.lVolumeEspecie, STRING_BUFFER_MAX_TEXTO, "VolumeEspecie"
    colCampoValor.Add "VolumeMarca", objNFiscal.lVolumeMarca, STRING_BUFFER_MAX_TEXTO, "VolumeMarca"
    colCampoValor.Add "MensagemNota", objNFiscal.sMensagemNota, STRING_BUFFER_MAX_TEXTO, "MensagemNota"
    colCampoValor.Add "PesoLiq", objNFiscal.dPesoLiq, 0, "Pesoliq"
    colCampoValor.Add "PesoBruto", objNFiscal.dPesoBruto, 0, "PesoBruto"
    colCampoValor.Add "Cliente", objNFiscal.lCliente, 0, "Cliente"
    colCampoValor.Add "FilialCli", objNFiscal.iFilialCli, 0, "FilialCli"
    colCampoValor.Add "Fornecedor", objNFiscal.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "FilialForn", objNFiscal.iFilialForn, 0, "FilialForn"
    colCampoValor.Add "FilialEntrega", objNFiscal.iFilialEntrega, 0, "FilialEntrega"
    colCampoValor.Add "NumPedidoVenda", objNFiscal.lNumPedidoVenda, 0, "NumPedidoVenda"
    colCampoValor.Add "CodTransportadora", objNFiscal.iCodTransportadora, 0, "CodTransportadora"
    colCampoValor.Add "CodTranspRedesp", objNFiscal.iCodTranspRedesp, 0, "CodTranspRedesp"
    colCampoValor.Add "DetPagFrete", objNFiscal.iDetPagFrete, 0, "DetPagFrete"
    colCampoValor.Add "TabelaPreco", objNFiscal.iTabelaPreco, 0, "TabelaPreco"
    colCampoValor.Add "FilialPedido", objNFiscal.iFilialPedido, 0, "FilialPedido"
    colCampoValor.Add "VolumeNumero", objNFiscal.sVolumeNumero, STRING_BUFFER_MAX_TEXTO, "VolumeNumero"
    colCampoValor.Add "FreteRespons", objNFiscal.iFreteRespons, 0, "FreteRespons"
    colCampoValor.Add "Status", objNFiscal.iStatus, 0, "Status"
    colCampoValor.Add "SemDataSaida", objNFiscal.iSemDataSaida, 0, "SemDataSaida"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "FilialPedido", OP_DIFERENTE, 0
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_CANCELADO
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 91381

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157624)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objNFiscal As ClassNFiscal) As Long
'Recolhe os dados da tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Tela_Memoria
            
    If Len(Trim(objCTNFiscalPedido.Pedido.ClipText)) > 0 Then objNFiscal.lNumPedidoVenda = CLng(objCTNFiscalPedido.Pedido.Text)
    objNFiscal.iFilialPedido = Codigo_Extrai(objCTNFiscalPedido.FilialPedido.Text)

    If objCTNFiscalRem.iFornecedorAtivo = FORN_ATIVO Then
        
        'Verifica se o Fornecedor está preenchido
        If Len(Trim(Fornecedor.Caption)) > 0 Then 'Alterado por leo em 30/01/02

            objFornecedor.sNomeReduzido = Fornecedor.Text
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 39411
            
            'Se não encontrou o Fornecedor --> erro
            If lErro = 6681 Then gError 39412

            objNFiscal.lFornecedor = objFornecedor.lCodigo
            objNFiscal.iFilialForn = Codigo_Extrai(Filial.Text)
            
        End If
    Else
        
        'Verifica se o Cliente esta preenchido
        If Len(Trim(objCTNFiscalPedido.Cliente.Caption)) > 0 Then
    
            objCliente.sNomeReduzido = objCTNFiscalPedido.Cliente.Caption
        
            'Lê o Cliente através do Nome Reduzido
            lErro = CF("Cliente_Le_NomeReduzido", objCliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 91382
        
            'Se não achou o Cliente --> erro
            If lErro <> SUCESSO Then gError 91383
        
            'Guarda código do Cliente
            objNFiscal.lCliente = objCliente.lCodigo
        
        End If
        
    End If
    
    'Armazena o responsável pelo frete
    objNFiscal.iFreteRespons = objCTNFiscalPedido.TipoFrete.ItemData(objCTNFiscalPedido.TipoFrete.ListIndex)
    
    objNFiscal.iFilialCli = Codigo_Extrai(objCTNFiscalPedido.Filial.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(objCTNFiscalPedido.Serie.Text, objCTNFiscalPedido.EletronicaFed.Value)
    
    If Len(Trim(objCTNFiscalPedido.NFiscal.Caption)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(objCTNFiscalPedido.NFiscal.Caption)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    objNFiscal.dtDataEmissao = StrParaDate(objCTNFiscalPedido.DataEmissao.Text)
    
    lErro = CF2(objCTNFiscalPedido, "NFCF_Move_DataSaida_Memoria", objNFiscal)
    If lErro <> SUCESSO Then gError 132296
    
    objNFiscal.sNaturezaOp = objCTNFiscalPedido.gobjTribTab.sNatOpInterna
    objNFiscal.iTabelaPreco = Codigo_Extrai(objCTNFiscalPedido.TabelaPreco.Text)
    objNFiscal.dValorFrete = StrParaDbl(objCTNFiscalPedido.ValorFrete.Text)
    objNFiscal.dValorSeguro = StrParaDbl(objCTNFiscalPedido.ValorSeguro.Text)
    objNFiscal.dValorOutrasDespesas = StrParaDbl(objCTNFiscalPedido.ValorDespesas.Text)
    objNFiscal.dValorDesconto = StrParaDbl(objCTNFiscalPedido.ValorDesconto.Text)
    objNFiscal.dValorProdutos = StrParaDbl(objCTNFiscalPedido.ValorProdutos.Caption)
    objNFiscal.dValorTotal = StrParaDbl(objCTNFiscalPedido.ValorTotal.Caption)
    objNFiscal.iFilialEntrega = Codigo_Extrai(objCTNFiscalPedido.FilialEntrega.Text)
    objNFiscal.iCodTransportadora = Codigo_Extrai(objCTNFiscalPedido.Transportadora.Text)
    objNFiscal.iCodTranspRedesp = Codigo_Extrai(objCTNFiscalPedido.TranspRedespacho.Text)
    objNFiscal.iDetPagFrete = objCTNFiscalPedido.RedespachoCli.Value
    objNFiscal.sPlaca = objCTNFiscalPedido.Placa.Text
    objNFiscal.sPlacaUF = objCTNFiscalPedido.PlacaUF.Text
    objNFiscal.lVolumeQuant = StrParaLong(objCTNFiscalPedido.VolumeQuant.Text)
    objNFiscal.sVolumeNumero = objCTNFiscalPedido.VolumeNumero.Text
    objNFiscal.sMensagemNota = objCTNFiscalPedido.Mensagem.Text
    objNFiscal.sMensagemCorpoNota = objCTNFiscalPedido.MensagemCorpo.Text
    objNFiscal.dPesoLiq = StrParaDbl(objCTNFiscalPedido.PesoLiquido.Text)
    objNFiscal.dPesoBruto = StrParaDbl(objCTNFiscalPedido.PesoBruto.Text)
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.dtDataEntrada = DATA_NULA
    objNFiscal.dtDataVencimento = DATA_NULA
    objNFiscal.iStatus = STATUS_LANCADO
    objNFiscal.dtDataReferencia = DATA_NULA
    
    objNFiscal.sNumPedidoTerc = objCTNFiscalPedido.NumPedidoTerc.Text
        
    If Len(Trim(objCTNFiscalPedido.VolumeEspecie.Text)) > 0 Then objNFiscal.lVolumeEspecie = Codigo_Extrai(objCTNFiscalPedido.VolumeEspecie.Text)
    If Len(Trim(objCTNFiscalPedido.VolumeMarca.Text)) > 0 Then objNFiscal.lVolumeMarca = Codigo_Extrai(objCTNFiscalPedido.VolumeMarca.Text)
        
    lErro = objCTNFiscalPedido.gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NF)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFiscal.objProjetoInfo = objProjetoInfo
    Set objNFiscal.objInfoAdic = objCTNFiscalPedido.gobjInfoAdic
    objNFiscal.dValorItens = StrParaDbl(objCTNFiscalPedido.ValorProdutos2.Caption)
    objNFiscal.dValorDescontoItens = StrParaDbl(objCTNFiscalPedido.ValorDescontoItens.Text)

    'Verifica se o Fornecedor está preenchido
    If Len(Trim(FornEntTerc.ClipText)) > 0 Then

        Set objFornecedor = New ClassFornecedor

        objFornecedor.sNomeReduzido = FornEntTerc.Text

        'Lê o Fornecedor
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError ERRO_SEM_MENSAGEM

        objNFiscal.lFornEntTerc = objFornecedor.lCodigo
        objNFiscal.iFilialFornEntTerc = Codigo_Extrai(FilialFornEntTerc.Text)

    End If
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 91382, 132296, 181761
        
        Case 91383
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157625)
            
    End Select
            
    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Tela_Preenche

    objNFiscal.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objNFiscal.lNumIntDoc <> 0 Then

        'Evita ter que alterar essa função a cada campo inserido na tabela de NF
        lErro = CF("NFiscal_Le", objNFiscal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 180098
        
        lErro = Traz_NFiscal_Tela(objNFiscal)
        If lErro <> SUCESSO Then gError 91384

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 91384, 180098

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157626)

    End Select

    Exit Sub

End Sub

Public Function Traz_NFiscal_Tela(objNFiscal As ClassNFiscal, Optional lNumIntPedido As Long = 0) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim bAlterouCT As Boolean

On Error GoTo Erro_Traz_NFiscal_Tela
    
    objCTNFiscalPedido.gbCarregandoTela = True
    
    'Lê os demais dados da nota fiscal passada se não veio de Pedido
    If lNumIntPedido = 0 Then
        lErro = objCTNFiscalPedido.Carrega_Dados_NFiscal(objNFiscal)
        If lErro <> SUCESSO Then gError 91385
    End If

    'Limpa a tela de Nota Fiscal
    lErro = Limpa_Tela_NFiscalRemPedido()
    If lErro <> SUCESSO Then gError 91386
    
    Set gobjNFiscal = objNFiscal
    
    lErro = Traz_NFiscal_Tela2(objNFiscal, lNumIntPedido)
    If lErro <> SUCESSO Then gError 91387
    
    Set objCTNFiscalPedido.gobjNFiscal = objNFiscal
        
    objCTNFiscalPedido.iAlterado = 0

    If lNumIntPedido = 0 Then
        MsgAutomatica.Value = vbUnchecked
        objCTNFiscalPedido.gbCarregandoTela = False
    Else
        MsgAutomatica.Value = vbChecked
    End If

    Traz_NFiscal_Tela = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela:

    objCTNFiscalPedido.gbCarregandoTela = False

    Traz_NFiscal_Tela = gErr

    Select Case gErr

        Case 91385, 91386, 91387
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157627)

    End Select

    Exit Function

End Function

Private Function Traz_NFiscal_Tela2(objNFiscal As ClassNFiscal, lNumIntPedido As Long) As Long
'Traz os dados da Nota Fiscal passada em objNFiscal

Dim lErro As Long
Dim iIndice As Integer
Dim objCliente As New ClassCliente
Dim objFornecedor As New ClassFornecedor
Dim objFilialCliente As New ClassFilialCliente
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colCodigoNome As New AdmColCodigoNome
Dim dValorProdutos As Double
Dim bGerarDados As Boolean
Dim iLote As Integer
Dim dtDataContabilidade As Date
Dim bCancel As Boolean
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim iCont As Integer
Dim dPercDesc As Double
Dim objNFiscalOriginal As New ClassNFiscal

On Error GoTo Erro_Traz_NFiscal_Tela2

  If objNFiscal.iTipoNFiscal <> 0 Then

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
    
        'Lê o tipo da Nota Fiscal
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO And lErro <> 31415 Then gError 39418
    
        'Se não encontrou o tipo de documento --> erro
        If lErro = 31415 Then gError 39419
    
        'De acordo com a Origem do tipo Coloca o Cliente ou o fornecedor na tela
        If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
        
            objCliente.lCodigo = objNFiscal.lCliente
    
            Call Habilita_Cliente
        
            'Lê o cliente
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 91388
            If lErro <> SUCESSO Then gError 91389
    
            'Preenche o Cliente
            objCTNFiscalPedido.Cliente.Caption = objCliente.sNomeReduzido
            objCTNFiscalPedido.gdDesconto = objCliente.dDesconto

            objFilialCliente.iCodFilial = objNFiscal.iFilialCli
        
            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", objCliente.sNomeReduzido, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 96078

            'Coloca na tela
            Filial.Caption = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome
    
            lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
            If lErro <> SUCESSO Then gError 131898
        
        ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        
             objFornecedor.lCodigo = objNFiscal.lFornecedor
    
             Call Habilita_Fornecedor
        
            'Lê o fornecedor
            lErro = CF("Fornecedor_Le", objFornecedor)
            If lErro <> SUCESSO And lErro <> 19062 Then gError 96075
            If lErro <> SUCESSO Then gError 96076
        
            'Preenche o fornecedor
             Fornecedor.Caption = objFornecedor.sNomeReduzido
        
            objFilialFornecedor.iCodFilial = objNFiscal.iFilialForn
        
            'Pesquisa se existe filial com o codigo extraido
            lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", objFornecedor.sNomeReduzido, objFilialFornecedor)
            If lErro <> SUCESSO And lErro <> 18272 Then gError 96079

            'Coloca na tela
            Filial.Caption = objFilialFornecedor.iCodFilial & SEPARADOR & objFilialCliente.sNome

            lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
            If lErro <> SUCESSO Then gError 131899

        End If
        
    Else
        objCliente.lCodigo = objNFiscal.lCliente
        Call Habilita_Cliente
        'Lê o cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 19062 Then gError 96083
        If lErro <> SUCESSO Then gError 96084

        'Preenche o Cliente
        objCTNFiscalPedido.Cliente.Caption = objCliente.sNomeReduzido
        objCTNFiscalPedido.gdDesconto = objCliente.dDesconto

        objFilialCliente.iCodFilial = objNFiscal.iFilialCli
    
        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", objCliente.sNomeReduzido, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 96078

        'Coloca na tela
        Filial.Caption = objFilialCliente.iCodFilial & SEPARADOR & objFilialCliente.sNome

        lErro = CF2(Me, "NFCF_Cliente_Traz_NF_Tela", objNFiscal)
        If lErro <> SUCESSO Then gError 131900

    End If
    
    'Carrega a combo de filial entrega
    lErro = CF("Filial_Preenche", objCTNFiscalPedido.FilialEntrega, colCodigoNome)
       
    'coloca o pedido e filialpedido na tela
    objCTNFiscalPedido.Pedido.Text = objNFiscal.lNumPedidoVenda
    objCTNFiscalPedido.FilialPedido.Text = objNFiscal.iFilialPedido
    Call objCTNFiscalPedido.FilialPedido_Validate(bSGECancelDummy)

    objCTNFiscalPedido.glCodPedido = objNFiscal.lNumPedidoVenda
    objCTNFiscalPedido.giFilialPedido = objNFiscal.iFilialPedido

    'Preenche o Status da Nota Fiscal
    If objNFiscal.iStatus = STATUS_LANCADO Then
        objCTNFiscalPedido.Status.Caption = STRING_STATUS_LANCADO
    ElseIf objNFiscal.iStatus = STATUS_BAIXADO Then
        objCTNFiscalPedido.Status.Caption = STRING_STATUS_BAIXADO
    ElseIf objNFiscal.iStatus = STATUS_CANCELADO Then
        objCTNFiscalPedido.Status.Caption = STRING_STATUS_CANCELADO
    End If
    
    'Coloca a FilialEntrega na tela
    If objNFiscal.iFilialEntrega > 0 Then
        objCTNFiscalPedido.FilialEntrega.Text = objNFiscal.iFilialEntrega
        Call objCTNFiscalPedido.FilialEntrega_Validate(bSGECancelDummy)
    End If
    
    'Coloca TabelaPreco na Tela
    If objNFiscal.iTabelaPreco > 0 Then
        objCTNFiscalPedido.TabelaPreco.Text = objNFiscal.iTabelaPreco
        Call objCTNFiscalPedido.TabelaPreco_Validate(bSGECancelDummy)
    End If
    
    'Preenche os demais dados da tela
    If objNFiscal.lNumNotaFiscal > 0 Then objCTNFiscalPedido.NFiscal.Caption = objNFiscal.lNumNotaFiscal
    
    objCTNFiscalPedido.Serie.Text = Desconverte_Serie_Eletronica(objNFiscal.sSerie)
    If ISSerieEletronica(objNFiscal.sSerie) Then
        objCTNFiscalPedido.EletronicaFed.Value = vbChecked
    Else
        objCTNFiscalPedido.EletronicaFed.Value = vbUnchecked
    End If
    Call objCTNFiscalPedido.Serie_Validate(bSGECancelDummy)

    Call DateParaMasked(objCTNFiscalPedido.DataEmissao, objNFiscal.dtDataEmissao)
    If objNFiscal.iSemDataSaida = 0 Then
        Call DateParaMasked(objCTNFiscalPedido.DataSaida, objNFiscal.dtDataSaida)
    Else
        Call DateParaMasked(objCTNFiscalPedido.DataSaida, DATA_NULA)
    End If
    
    'horasaida
    objCTNFiscalPedido.HoraSaida.PromptInclude = False

    If objNFiscal.dtDataSaida <> DATA_NULA Then objCTNFiscalPedido.HoraSaida.Text = Format(objNFiscal.dtHoraSaida, "hh:mm:ss")
    objCTNFiscalPedido.HoraSaida.PromptInclude = True
    
    'Coloca os Valores na Tela
    If objNFiscal.dValorDesconto > 0 Then objCTNFiscalPedido.ValorDesconto.Text = Format(objNFiscal.dValorDesconto, "Standard")
    If objNFiscal.dValorOutrasDespesas > 0 Then objCTNFiscalPedido.ValorDespesas.Text = Format(objNFiscal.dValorOutrasDespesas, "Standard")
    If objNFiscal.dValorFrete > 0 Then objCTNFiscalPedido.ValorFrete.Text = Format(objNFiscal.dValorFrete, "Standard")
    If objNFiscal.dValorSeguro > 0 Then objCTNFiscalPedido.ValorSeguro.Text = Format(objNFiscal.dValorSeguro, "Standard")
    If objNFiscal.dValorTotal > 0 Then objCTNFiscalPedido.ValorTotal.Caption = Format(objNFiscal.dValorTotal, "Standard")

    If objNFiscal.iCodTransportadora <> 0 Then
        objCTNFiscalPedido.Transportadora.Text = objNFiscal.iCodTransportadora
        Call objCTNFiscalPedido.Transportadora_Validate(bSGECancelDummy)
    End If

    If objNFiscal.iCodTranspRedesp <> 0 Then
        objCTNFiscalPedido.TranspRedespacho.Text = CStr(objNFiscal.iCodTranspRedesp)
        Call objCTNFiscalPedido.TranspRedespacho_Validate(bCancel)
    End If

    For iIndice = 0 To objCTNFiscalPedido.TipoFrete.ListCount - 1
        If objCTNFiscalPedido.TipoFrete.ItemData(iIndice) = objNFiscal.iFreteRespons Then
            objCTNFiscalPedido.TipoFrete.ListIndex = iIndice
            Exit For
        End If
    Next


    objCTNFiscalPedido.Placa.Text = objNFiscal.sPlaca
    objCTNFiscalPedido.PlacaUF.Text = objNFiscal.sPlacaUF
    
    'Traz a espécie dos volumes do pedido
    If objNFiscal.lVolumeEspecie > 0 Then
        objCTNFiscalPedido.VolumeEspecie.Text = objNFiscal.lVolumeEspecie
        Call objCTNFiscalPedido.VolumeEspecie_Validate(bSGECancelDummy)
    Else
        objCTNFiscalPedido.VolumeEspecie.Text = ""
    End If
    
    'Traz a marca dos volumes do pedido
    If objNFiscal.lVolumeMarca > 0 Then
        objCTNFiscalPedido.VolumeMarca.Text = objNFiscal.lVolumeMarca
        Call objCTNFiscalPedido.VolumeMarca_Validate(bSGECancelDummy)
    Else
        objCTNFiscalPedido.VolumeMarca.Text = ""
    End If
    
    objCTNFiscalPedido.VolumeNumero = objNFiscal.sVolumeNumero
    If objNFiscal.lVolumeQuant > 0 Then objCTNFiscalPedido.VolumeQuant.Text = objNFiscal.lVolumeQuant

    objCTNFiscalPedido.NumPedidoTerc.Text = objNFiscal.sNumPedidoTerc

    objCTNFiscalPedido.Mensagem.Text = objNFiscal.sMensagemNota
    objCTNFiscalPedido.MensagemCorpo.Text = objNFiscal.sMensagemCorpoNota

    If objNFiscal.dPesoLiq > 0 Then objCTNFiscalPedido.PesoLiquido.Text = Format(objNFiscal.dPesoLiq, objCTNFiscalPedido.PesoLiquido.Format) 'Alterado por Wagner
    If objNFiscal.dPesoBruto > 0 Then objCTNFiscalPedido.PesoBruto.Text = Format(objNFiscal.dPesoBruto, objCTNFiscalPedido.PesoBruto.Format) 'Alterado por Wagner
    
    lErro = CF2(objCTNFiscalPedido, "NFP_Traz_NFiscal_Tela2", objNFiscal, lNumIntPedido)
    If lErro <> SUCESSO Then gError 46850
    
    objCTNFiscalPedido.gobjTribTab.sNatOpInterna = objNFiscal.sNaturezaOp
    
    'Preenche o Grid com os itens da Nota Fiscal
    lErro = objCTNFiscalPedido.Preenche_GridItens(objNFiscal)
    If lErro <> SUCESSO Then gError 91391
    
    lErro = CF2(objCTNFiscalPedido, "NFRP_Preeche_GridItens", objNFiscal)
    If lErro <> SUCESSO Then gError 177220

    If objNFiscal.iTipoNFiscal <> 0 Then

        objTipoDocInfo.iCodigo = objNFiscal.iTipoNFiscal
        
        'Le o Tipo para saber se tem Movimento de Estoque
        lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
        If lErro <> SUCESSO Then gError 91392
        
        If objTipoDocInfo.iTipoMovtoEstoque <> 0 Then
            'Carrega ItensNF com Rastreamentos
            lErro = gobjRastreamento.Carrega_RastroItensNF(objNFiscal)
            If lErro <> SUCESSO Then gError 91393
        End If
    
        'Preenche a combo de TipoNFiscal com o tipo lido no BD
        For iCont = 0 To TipoNFiscal.ListCount - 1
        
            If Codigo_Extrai(TipoNFiscal.List(iCont)) = objNFiscal.iTipoNFiscal Then
                TipoNFiscal.ListIndex = iCont
                Exit For
            End If
            
        Next
        
        giTipoDocInfo = Codigo_Extrai(TipoNFiscal.Text)
        
        'Leo até aqui
    
    End If
    
    'Se o código da Nota Fiscal Original estiver preenchido
    If objNFiscal.lNumIntNotaOriginal <> 0 Then
        
        objNFiscalOriginal.lNumIntDoc = objNFiscal.lNumIntNotaOriginal
        'Lê a Nota Fiscal Original
        lErro = CF("NFiscal_Le", objNFiscalOriginal)
        If lErro <> SUCESSO And lErro <> 31442 Then gError 39024
        'Se não achou a Nota Fiscal Original --> erro
        If lErro = 31442 Then gError 39025
        
        'Coloca os dados relativos a Nota Fiscal Original na Tela
        SerieNFiscalOriginal.Text = objNFiscalOriginal.sSerie
        NFiscalOriginal.Text = objNFiscalOriginal.lNumNotaFiscal
        
    End If

    'Atualiza a coleção de Itens de NF's Global a tela
    Set gcolItensNf = objNFiscal.ColItensNF
    
    Set gobjNFiscal = objNFiscal
    
    'Se a nota não de remessa de material beneficiado para terceiros
    'significa que as reservas poderão ser aproveitadas e o que não foi reservado
    'pode ser alocado automaticamente, pois o movimento de estoque é feito sobre a
    'quantidade do escaninho nossa disponível
    If Codigo_Extrai(Trim(TipoNFiscal.Text)) <> DOCINFO_NFISRMB3PV Then
    
        If lNumIntPedido > 0 Then bGerarDados = True

        'Preenche o Grid com as Alocações dos itens da Nota Fiscal
        lErro = CF2(objCTNFiscalPedido, "NFP_Preenche_GridAlocacoes", objNFiscal, bGerarDados)
        If lErro <> SUCESSO Then gError 91394
    
    If lErro <> SUCESSO Then gError 39026
    
    
    'Senão, ou seja, se a nota for de remessa de material beneficiado para terceiros,
    'significa que o movimento de estoque é feito sobre a quantidade do escaninho beneficiado de 3º´s em nosso poder
    'Nesse caso, a reserva do pedido não pode ser aproveitada(pois foi feita em cima da quant. nossa disponível)
    'e a alocação não é feita de forma automática
    Else
        
        'Limpa o Grid de Alocações
        Call Grid_Limpa(objGridAlocacoes)
        objGridAlocacoes.iLinhasExistentes = 0
    
    End If
    
    'Se a nota fiscal já existe
    If objNFiscal.lNumIntDoc <> 0 Then
    
        'Traz os dados contábeis para a tela (contabilidade)
        lErro = objCTNFiscalPedido.objContabil.Contabil_Traz_Doc_Tela(objNFiscal.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError 91396
    End If
    
    dtDataContabilidade = StrParaDate(objCTNFiscalPedido.CTBDataContabil.Text)
    iLote = StrParaInt(objCTNFiscalPedido.CTBLote.Text)
    
    'Se a data contabil está ultrapassada
    If dtDataContabilidade <> DATA_NULA And dtDataContabilidade < gdtDataAtual Then
        'Atualiza a data da contabilidade
        objCTNFiscalPedido.CTBLote.Text = ""
        objCTNFiscalPedido.CTBDataContabil.PromptInclude = False
        objCTNFiscalPedido.CTBDataContabil.Text = Format(gdtDataAtual, "dd/mm/yy")
        objCTNFiscalPedido.CTBDataContabil.PromptInclude = True
        Call objCTNFiscalPedido.CTBDataContabil_Validate(bCancel)
        If iLote > 0 Then objCTNFiscalPedido.CTBLote.Text = iLote
    End If
        
    Call ValorTotal_Calcula
     
    objCTNFiscalPedido.NumNFe.Caption = objNFiscal.sStatusNFeFed
    
    Set objCTNFiscalPedido.gobjInfoAdic = objNFiscal.objInfoAdic
    
    If objNFiscal.dValorItens = 0 Then
        For iIndice = 1 To objCTNFiscalPedido.objGridItens.iLinhasExistentes
            If objNFiscal.ColItensNF.Item(iIndice).objInfoAdicDocItem.iIncluiValorTotal = MARCADO Then
                objNFiscal.dValorItens = objNFiscal.dValorItens + StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_PrecoTotalB_Col))
                objNFiscal.dValorDescontoItens = objNFiscal.dValorDescontoItens + StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_Desconto_Col))
            End If
        Next
    End If
    
    objCTNFiscalPedido.ValorProdutos2.Caption = Format(objNFiscal.dValorItens, "Standard")
    objCTNFiscalPedido.ValorDescontoItens.Text = Format(objNFiscal.dValorDescontoItens, "Standard")
    If objNFiscal.dValorItens > 0 Then
        dPercDesc = objNFiscal.dValorDescontoItens / objNFiscal.dValorItens
    Else
        dPercDesc = 0
    End If
    
    objCTNFiscalPedido.PercDescontoItens.Text = Format(dPercDesc * 100, "Fixed")
    
    objCTNFiscalPedido.dValorDescontoItensAnt = objNFiscal.dValorDescontoItens
    objCTNFiscalPedido.dPercDescontoItensAnt = dPercDesc
    
    If objNFiscal.lFornEntTerc <> 0 Then
        FornEntTerc.Text = objNFiscal.lFornEntTerc
        Call FornEntTerc_Validate(bSGECancelDummy)
        FilialFornEntTerc.Text = objNFiscal.iFilialFornEntTerc
        Call FilialFornEntTerc_Validate(bSGECancelDummy)
    End If
     
    Traz_NFiscal_Tela2 = SUCESSO

    Exit Function

Erro_Traz_NFiscal_Tela2:

    Traz_NFiscal_Tela2 = gErr

    Select Case gErr
    
        Case 39024
        
        Case 39025
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_CADASTRADA", gErr, objNFiscal.lNumIntNotaOriginal)

        Case 91388, 91390, 91391, 91392, 91393, 91394, 91395, 91396, 96077, 96078, 96079, 96083, 131898, 131899, 131900, 177220 'Inserido por Wagner
        
        Case 91389, 96084
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO", gErr, objCliente.lCodigo)
        
        Case 96075, 96076
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157628)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 91397

    'Limpa a Tela
    lErro = Limpa_Tela_NFiscalRemPedido()
    If lErro <> SUCESSO Then gError 91398
    
    Set objCTNFiscalPedido.gcolItemPedido = New Collection
    Set objCTNFiscalPedido.gcolItemQuant = New Collection
    Set objCTNFiscalPedido.gcolItemUM = New Collection

    'fecha o comando de setas
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 91397, 91398

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 157629)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim vbMsg As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objContabil As ClassContabil
Dim snProt As String
Dim iEnviado As Integer

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se os dados de algum pedido foi trazido para a tela
    If objCTNFiscalPedido.glCodPedido = 0 Then gError 91399

    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(objCTNFiscalPedido.Pedido.Text)) = 0 Then gError 91400
    If Len(Trim(objCTNFiscalPedido.FilialPedido.Text)) = 0 Then gError 91401
'???? Testar se o numero da nota fiscal está preenchido. Mario
    'Janaina
    If Len(Trim(objCTNFiscalPedido.NFiscal.Caption)) = 0 Then gError 96080
    'Janaina
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 91403 '46581
    If Len(Trim(objCTNFiscalPedido.Serie.Text)) = 0 Then gError 91404
    If Len(Trim(objCTNFiscalPedido.DataEmissao.ClipText)) = 0 Then gError 91405
    
    '????? Verificar junta a nota fiscal de remessa para tratar fornecedor. Mario
    'Janaina
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 96085
    
    'Não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 96086

    'De acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
        If Len(Trim(Cliente.Caption)) = 0 Then gError 96081
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        If Len(Trim(Fornecedor.Caption)) = 0 Then gError 96082
    End If
    'Janaina
    
    'Verifica se o código de pedido é o mesmo que foi trazido para a tela
    If objCTNFiscalPedido.glCodPedido <> StrParaLong(objCTNFiscalPedido.Pedido.Text) Or objCTNFiscalPedido.giFilialPedido <> Codigo_Extrai(objCTNFiscalPedido.FilialPedido.Text) Then gError 91406

    objNFiscal.lNumPedidoVenda = objCTNFiscalPedido.glCodPedido
    objNFiscal.iFilialPedido = objCTNFiscalPedido.giFilialPedido
    objNFiscal.lNumNotaFiscal = StrParaLong(objCTNFiscalPedido.NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(objCTNFiscalPedido.Serie.Text, objCTNFiscalPedido.EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(objCTNFiscalPedido.DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objNFiscal.iTipoNFiscal = Codigo_Extrai(TipoNFiscal.Text)

    'se origem for de empresa é interna
    If objTipoDocInfo.iOrigem = DOCINFO_EMPRESA Then
        'Verifica se a existe nota fiscal está cadastrada
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 203032
        
        If lErro = 43676 Then gError 203033
        
'        If gobjCRFAT.iUsaNFe = MARCADO Then
'
'            lErro = CF("NFeFedLote_Le", objNFiscal.lNumIntDoc, iEnviado)
'            If lErro <> SUCESSO Then gError 203034
'
'            If iEnviado = 1 Then gError 203035
'
'        End If

    End If

    'Verifica se a existe nota fiscal está cadastrada
    lErro = CF("NFiscal_Le_Pedido", objNFiscal)
    If lErro <> SUCESSO And lErro <> 46602 Then gError 91407
    
    'se a nota não está cadastrada ==> erro
    If lErro = 46602 Then gError 91408

    'pede confirmacao
    vbMsg = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_NFISCAL", objNFiscal.lNumNotaFiscal)
    If vbMsg = vbYes Then

        'Lê os itens da nota fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 91409

        'Faz o cancelamento de uma nota fiscal de Saida
        Set objContabil = objCTNFiscalPedido.objContabil
    
        lErro = CF("NotaFiscalRemPedidoSaida_Excluir", objNFiscal, objContabil)
        If lErro <> SUCESSO Then gError 91410
    
        'Limpa a Tela
        lErro = Limpa_Tela_NFiscalRemPedido()
        If lErro <> SUCESSO Then gError 91411
        
        'Libera as coleções globais
        Set objCTNFiscalPedido.gcolItemPedido = New Collection
        Set objCTNFiscalPedido.gcolItemQuant = New Collection
        Set objCTNFiscalPedido.gcolItemUM = New Collection
    
        'fecha o comando de setas
        Call ComandoSeta_Fechar(Me.Name)

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 91399
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_TRAZIDO", gErr)
        
        Case 91400
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
            
        Case 91401
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
            
        Case 91403
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 91404
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)

        Case 91405
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 91406
             Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_TELA_DIFERE_BD", gErr, StrParaLong(objCTNFiscalPedido.Pedido.Text), Codigo_Extrai(objCTNFiscalPedido.FilialPedido.Text), objCTNFiscalPedido.glCodPedido, objCTNFiscalPedido.giFilialPedido)
       
        Case 91408
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA2", gErr, objNFiscal.lNumPedidoVenda, objNFiscal.iFilialPedido, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
        
        Case 91407, 91409, 91410, 91411, 96085, 203032, 203034

        Case 96080
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_PREENCHIDA", gErr)
        
        Case 96081
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 96082
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 96086
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)

        Case 203033
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_NAO_CADASTRADA3", gErr, objNFiscal.iFilialEmpresa, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal, objNFiscal.dtDataEmissao, objNFiscal.iTipoNFiscal)
       
        Case 203035
            Call Rotina_Erro(vbOKOnly, "ERRO_NFE_CADASTRADA_NAO_EXCLUI", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157630)
        
    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dValorTotal As Double
Dim dValorDesconto As Double
Dim objNFiscal As New ClassNFiscal
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objBloqueioPV As ClassBloqueioPV
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objContabil As ClassContabil

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os dados de algum pedido foi trazido para a tela
    If objCTNFiscalPedido.glCodPedido = 0 Then gError 91412
    
    'Verifica se os campos obrigatórios estão preenchidos
    If Len(Trim(objCTNFiscalPedido.Pedido.Text)) = 0 Then gError 91413
    If Len(Trim(objCTNFiscalPedido.FilialPedido.Text)) = 0 Then gError 91414
    If Len(Trim(TipoNFiscal.Text)) = 0 Then gError 91415
    If Len(Trim(objCTNFiscalPedido.Serie.Text)) = 0 Then gError 91416
    If Len(Trim(objCTNFiscalPedido.DataEmissao.ClipText)) = 0 Then gError 91417
    If Len(Trim(objCTNFiscalPedido.gobjTribTab.sNatOpInterna)) = 0 Then gError 91419
    If Len(Trim(objCTNFiscalPedido.FilialEntrega.Text)) = 0 Then gError 91420
    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFISRETPV Then
        If Len(Trim(FornEntTerc.Text)) = 0 Then gError 209998
        If Len(Trim(FilialFornEntTerc.Text)) = 0 Then gError 209999
    End If
    
    'Verifica se tem itens no grid
    If objCTNFiscalPedido.objGridItens.iLinhasExistentes = 0 Then gError 91421

    'Verifica se o código de pedido é o mesmo que foi trazido para a tela
    If objCTNFiscalPedido.glCodPedido <> StrParaLong(objCTNFiscalPedido.Pedido.Text) Or objCTNFiscalPedido.giFilialPedido <> Codigo_Extrai(objCTNFiscalPedido.FilialPedido.Text) Then
        
        'Avisa que os dados da tela são de outro pedido e pergunta se deseja prosseguir.
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PEDIDO_TELA_NAO_UTILIZADO", objCTNFiscalPedido.glCodPedido, objCTNFiscalPedido.giFilialPedido)
        If vbMsgRes = vbNo Then gError 91422  'Sai da rotina.
        
        objCTNFiscalPedido.Pedido.Text = objCTNFiscalPedido.glCodPedido
        objCTNFiscalPedido.FilialPedido.Text = objCTNFiscalPedido.giFilialPedido
        Call objCTNFiscalPedido.FilialPedido_Validate(bSGECancelDummy)
        
    End If
    
    objPedidoVenda.lCodigo = StrParaLong(objCTNFiscalPedido.Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(objCTNFiscalPedido.FilialPedido.Text)
    
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 46990
    
    If objPedidoVenda.dtDataEmissao > StrParaDate(objCTNFiscalPedido.DataEmissao.Text) Then gError 206615

    
    If Len(Trim(objCTNFiscalPedido.PesoLiquido.ClipText)) > 0 And Len(Trim(objCTNFiscalPedido.PesoBruto.ClipText)) > 0 Then
        If CDbl(objCTNFiscalPedido.PesoLiquido.Text) > CDbl(objCTNFiscalPedido.PesoBruto.Text) Then gError 91423
    End If
    
    'Verifica se a Data de Saida e menor que a Data de Emissão
    If Len(Trim(objCTNFiscalPedido.DataSaida.ClipText)) <> 0 Then
        If CDate(objCTNFiscalPedido.DataSaida.Text) < CDate(objCTNFiscalPedido.DataEmissao.Text) Then gError 91424
    End If
    
    dValorTotal = CDbl(IIf(Len(Trim(objCTNFiscalPedido.ValorTotal.Caption)) > 0, objCTNFiscalPedido.ValorTotal.Caption, 0))

    'Se o total for negativo --> Erro
    If dValorTotal < 0 Then gError 91425

    'Verifica se o Valor do Desconto é maior que a soma dos demais valores
    If Len(Trim(objCTNFiscalPedido.ValorDesconto.Text)) > 0 Then dValorDesconto = CDbl(objCTNFiscalPedido.ValorDesconto.Text)
    
    lErro = objCTNFiscalPedido.DescontoGlobal_Testa(dValorDesconto)
    If lErro <> SUCESSO Then gError 91426
    
    lErro = objCTNFiscalPedido.gobjTribTab.Valida_Dados
    If lErro <> SUCESSO Then gError 91426
    
    objNFiscal.lNumPedidoVenda = objCTNFiscalPedido.glCodPedido
    objNFiscal.iFilialPedido = objCTNFiscalPedido.giFilialPedido
    objNFiscal.lNumNotaFiscal = StrParaLong(objCTNFiscalPedido.NFiscal.Caption)
    objNFiscal.sSerie = Converte_Serie_Eletronica(objCTNFiscalPedido.Serie.Text, objCTNFiscalPedido.EletronicaFed.Value)
    objNFiscal.dtDataEmissao = StrParaDate(objCTNFiscalPedido.DataEmissao.Text)
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    objPedidoVenda.lCodigo = objCTNFiscalPedido.glCodPedido
    objPedidoVenda.iFilialEmpresa = objCTNFiscalPedido.giFilialPedido
    
    objTipoDocInfo.iCodigo = Codigo_Extrai(TipoNFiscal.Text)
    
    'Lê o Tipo da Nota Fiscal
    lErro = CF("TipoDocInfo_Le_Codigo", objTipoDocInfo)
    If lErro <> SUCESSO And lErro <> 31415 Then gError 91431
    
    'Não encontrou o Tipo de Documento --> erro
    If lErro = 31415 Then gError 91432
    

    'De acordo com a sua Origem verifica se o Cliente ou Fornecedor estão preenchidos
    If objTipoDocInfo.iDestinatario = DOCINFO_CLIENTE Then
        If Len(Trim(Cliente.Caption)) = 0 Then gError 91433
    ElseIf objTipoDocInfo.iDestinatario = DOCINFO_FORNECEDOR Then
        If Len(Trim(Fornecedor.Caption)) = 0 Then gError 91434
    End If
                
    lErro = objCTNFiscalPedido.Valida_Dados_GridItens()
    If lErro <> SUCESSO Then gError 91435
    
    'lErro = objCTNFiscalPedido.gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    lErro = gobjRastreamento.Valida_Rastreamento(objTipoDocInfo)
    If lErro <> SUCESSO Then gError 91436

    'Set objNFiscal.objRastreamento = objCTNFiscalPedido.gobjRastreamento
    Set objNFiscal.objRastreamento = gobjRastreamento
    
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 91437
    
    'Se é uma nota de complemento ou se é uma nota de de cobrança de material consignado
    If objNFiscal.iTipoNFiscal = DOCINFO_NFISSRPVEF Then
        
        'Recolhe os dados da Nota Fiscal Original
        lErro = Move_Dados_NFOriginal_Memoria(objNFiscal, objNFiscal.iTipoNFiscal)
        If lErro <> SUCESSO Then gError 181786
        
    End If
    
    lErro = Move_Grids_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 91438
    
    lErro = objCTNFiscalPedido.gobjTribTab.Valida_TributacaoNF_Gravacao(objNFiscal)
    If lErro <> SUCESSO Then gError 91439
    
    Call objCTNFiscalPedido.Calcula_Msg
    
    'mover a parte do rastreamento
    lErro = gobjRastreamento.Move_Rastro_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 91441
    
    '###############################################
    'Inserido por Wagner 10/08/2006
    lErro = objCTNFiscalPedido.gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    '###############################################
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objCTNFiscalPedido.objContabil.Contabil_Testa_Data(CDate(objCTNFiscalPedido.DataEmissao.Text))
    If lErro <> SUCESSO Then gError 91442
    
    Set objContabil = objCTNFiscalPedido.objContabil
    
    lErro = CF("NFiscal_Valida_Diversos", objNFiscal)
    If lErro <> SUCESSO Then gError 181786
    
    lErro = CF("NFiscalRemPedido_Grava", objNFiscal, objBloqueioPV, objContabil)
    If lErro <> SUCESSO Then gError 91443
    
    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
        gobjAnotacao.sID = CStr(objNFiscal.lNumIntDoc)
        gobjAnotacao.dtDataAlteracao = gdtDataHoje
        
        lErro = CF("Anotacoes_Grava", gobjAnotacao)
        If lErro <> SUCESSO Then gError 130406
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If Len(Trim(objCTNFiscalPedido.NFiscal.Caption)) = 0 Then vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_NOTA_GRAVADA", objNFiscal.lNumNotaFiscal)
    
'    If giTipoVersao = VERSAO_FULL Then
        If objCTNFiscalPedido.ImprimeRomaneio.Value = vbChecked Then
            Call objCTNFiscalPedido.NFiscal_Imprime_Romaneio(objNFiscal)
        End If
'    End If
    
    If ISSerieEletronica(objNFiscal.sSerie) Then
        lErro = CF("NFE_Grava1", objNFiscal.lNumIntDoc, objTipoDocInfo)
        If lErro <> SUCESSO Then gError 203000
    End If
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
       
        Case 91413
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)
        
        Case 91414
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)
        
        Case 91415, 91433
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 91415
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_PREENCHIDO", gErr)

        Case 91416
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
        
        Case 91417
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
        
        Case 91418
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_NAO_PREENCHIDA", gErr)
        
        Case 91419
            Call Rotina_Erro(vbOKOnly, "ERRO_NATUREZAOP_NAO_PREENCHIDA", gErr)
        
        Case 91421
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)
        
        Case 91412
            Call Rotina_Erro(vbOKOnly, "ERRO_NENHUM_PEDIDO_TRAZIDO", gErr)
        
        Case 91440, 91439, 91422, 91438, 91437, 91443, 91428, 91435, 130406, 181786, 203000
        
        Case 91424
            Call Rotina_Erro(vbOKOnly, "ERRO_DATASAIDA_ANTERIOR_DATAEMISSAO", gErr, objCTNFiscalPedido.DataSaida, objCTNFiscalPedido.DataEmissao)
        
        Case 91425
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NF_NEGATIVO", gErr)

        Case 91423
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO", gErr, CDbl(objCTNFiscalPedido.PesoLiquido.Text), CDbl(objCTNFiscalPedido.PesoBruto.Text))

        Case 91429
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO_CREDITO", gErr, objCTNFiscalPedido.glCodPedido)
        
        Case 91430
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO_TOTAL", gErr)
                    
        Case 91420
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALENTREGA_NAO_INFORMADA", gErr)
                    
        Case 91427
            'Call Rotina_Erro(vbOKOnly, "ERRO_IR_FONTE_MAIOR_VALOR_TOTAL", gErr, dValorIRRF, dValorTotal)

        Case 91432
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iTipo)
       
        Case 91434
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 91431, 91436, 91441, 91442, 91426

        Case 206615
            Call Rotina_Erro(vbOKOnly, "ERRO_NF_DATA_MENOR_PV", gErr, StrParaDate(objCTNFiscalPedido.DataEmissao.Text), objPedidoVenda.dtDataEmissao)

        Case 209998
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNENTTERC_NAO_PREENCHIDO", gErr)

        Case 209999
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNENTTERC_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157631)

    End Select

    Exit Function

End Function

Private Function Move_Grids_Memoria(objNFiscal As ClassNFiscal) As Long
'recolhe os dados dos grids para a memória(objNFiscal)

Dim lErro As Long
Dim iIndice As Integer
Dim objItemNF As ClassItemNF

On Error GoTo Erro_Move_Grids_Memoria

    'recolhe os dados do grid de itens
    lErro = objCTNFiscalPedido.Move_GridItens_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 91444
    
    iIndice = 0
    For Each objItemNF In objNFiscal.ColItensNF
        iIndice = iIndice + 1
        lErro = CF2(Me, "NFRP_Move_GridItem_Memoria", objNFiscal, objItemNF, iIndice)
        If lErro <> SUCESSO Then gError 177222
    Next
    
    'recolhe os dados do grid de alocações
    lErro = objCTNFiscalPedido.Move_GridAlocacao_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 91445
    
    'mover a parte de tributacao
    lErro = objCTNFiscalPedido.gobjTribTab.PreencherNF(objNFiscal)
    If lErro <> SUCESSO Then gError 91446
        
    Move_Grids_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Grids_Memoria:

    Move_Grids_Memoria = gErr
    
    Select Case gErr
    
        Case 91444, 91445, 91446, 177222
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157632)
    
    End Select
    
    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objCTNFiscalPedido.objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 91447
        
        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridItens
            Case objCTNFiscalPedido.GridItens.Name

                lErro = Saida_Celula_GridItens(objGridInt)
                If lErro <> SUCESSO Then gError 91448

            Case objCTNFiscalPedido.GridRastro.Name

                lErro = gobjRastreamento.Saida_Celula()
                If lErro <> SUCESSO Then gError 91449

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 91450

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 91448, 91450, 91447, 91449

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157633)

    End Select

    Exit Function

End Function

Public Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
    Inicializa_Grid_Itens = CF2(objCTNFiscalPedido, "NFP_Inicializa_Grid_Itens", objGridInt)
End Function

Public Function Saida_Celula_GridItens(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridItens

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Quantidade
        Case objCTNFiscalPedido.iGrid_Quantidade_Col
            lErro = Saida_Celula_Quantidade(objGridInt)
            If lErro <> SUCESSO Then gError 91451

        'PrecoUnitario
        Case objCTNFiscalPedido.iGrid_ValorUnitario_Col
            lErro = Saida_Celula_PrecoUnitario(objGridInt)
            If lErro <> SUCESSO Then gError 91452

        'PercDesc
        Case objCTNFiscalPedido.iGrid_PercDesc_Col
            lErro = Saida_Celula_PercentDesc(objGridInt)
            If lErro <> SUCESSO Then gError 91453

        'Desconto
        Case objCTNFiscalPedido.iGrid_Desconto_Col
            lErro = Saida_Celula_Desconto(objGridInt)
            If lErro <> SUCESSO Then gError 91454
            
        Case Else
            lErro = CF2(Me, "NFRP_Saida_Celula_GridItens", objGridInt)
            If lErro <> SUCESSO Then gError 46933

    End Select

    Saida_Celula_GridItens = SUCESSO

    Exit Function

Erro_Saida_Celula_GridItens:

    Saida_Celula_GridItens = gErr

    Select Case gErr

        Case 91451, 91452, 91453, 91454

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157634)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantidade As Double
Dim iSairCelula As Integer
Dim dPrecoUnitario As Double
Dim dQuantidadeAnterior As Double
Dim objProduto As New ClassProduto
Dim objKitVenda As New ClassKitVenda
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemNF As New ClassItemNF
Dim objItemPV As ClassItemPedido
Dim dQtdePV As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = objCTNFiscalPedido.Quantidade

    iSairCelula = False
    
    'Verifica o preenchimento de Quantidade
    If Len(Trim(objCTNFiscalPedido.Quantidade.ClipText)) > 0 Then

        If giComplemento = DESMARCADO Then
            lErro = Valor_Positivo_Critica(objCTNFiscalPedido.Quantidade.Text)
        Else
            lErro = Valor_NaoNegativo_Critica(objCTNFiscalPedido.Quantidade.Text)
        End If
        If lErro <> SUCESSO Then gError 91455
        
        dQuantidade = CDbl(objCTNFiscalPedido.Quantidade.Text)
        objCTNFiscalPedido.Quantidade = Formata_Estoque(dQuantidade)
        
        dQuantidadeAnterior = StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_Quantidade_Col))
                   
        Set objItemNF = objCTNFiscalPedido.gobjNFiscal.ColItensNF(GridItens.Row)
       
        If objItemNF.iRecebForaFaixa = ERRO_E_REJEITA_RECEBIMENTO Then
            Set objItemPV = New ClassItemPedido
            
            objItemPV.lNumIntDoc = objCTNFiscalPedido.gcolItemPedido(GridItens.Row)
            objItemPV.iFilialEmpresa = giFilialEmpresa
            
            lErro = CF("ItensPedido_Le_NumIntDoc1", objItemPV)
            If lErro <> SUCESSO Then gError 196964
            
            dQtdePV = objItemPV.dQuantidade - objItemPV.dQuantCancelada
                
            If dQuantidade > Arredonda_Estoque(objCTNFiscalPedido.gcolItemQuant(objCTNFiscalPedido.GridItens.Row) + (dQtdePV * objItemNF.dPercentMaisReceb)) Then gError 91456  'Alterado por Wagner
        End If
                
        If dQuantidadeAnterior = dQuantidade Then iSairCelula = True

    End If

    If Not iSairCelula Then
    
        objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_Quantidade_Col) = Formata_Estoque(dQuantidade)
        
        objProduto.sCodigo = objCTNFiscalPedido.gobjNFiscal.ColItensNF(GridItens.Row).sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 177585
        
        If objProduto.iKitVendaComp <> MARCADO Then
        
            If Codigo_Extrai(TipoNFiscal.Text) <> DOCINFO_NFISRETPV Then
                'Tenta fazer uma alocação p\ oitem de acordo com a nova UM
                lErro = CF2(Me, "NFP_Alocacao_Processa") 'Alterado por Leo em 11/01/02SELECT EstoqueProduto.Almoxarifado, EstoqueProduto.LocalizacaoFisica, EstoqueProduto.DataInventario,EstoqueProduto.SaldoInicial, EstoqueProduto.DataInicial, EstoqueProduto.QuantidadeInicial, EstoqueProduto.QuantDispNossa, EstoqueProduto.QuantReservada, EstoqueProduto.QuantEmpenhada, EstoqueProduto.QuantRecIndl, EstoqueProduto.QuantIndOutras, EstoqueProduto.QuantDefeituosa, EstoqueProduto.QuantConsig, EstoqueProduto.QuantDemo3, EstoqueProduto.QuantConsig3, EstoqueProduto.QuantDemo, EstoqueProduto.QuantConserto3, EstoqueProduto.QuantConserto, EstoqueProduto.QuantOutras3, EstoqueProduto.QuantOutras, QuantBenef, QuantBenef3, Almoxarifado.NomeReduzido FROM EstoqueProduto, Almoxarifado WHERE EstoqueProduto.Produto = ? AND EstoqueProduto.Almoxarifado = Almoxarifado.Codigo AND Almoxarifado.FilialEmpresa = ? ORDER BY Almoxarifado.NomeReduzido
                If lErro <> SUCESSO And lErro <> 23796 Then gError 91457
                If lErro = 23796 Then gError 23796
            End If
    
        Else
        
            objKitVenda.sProduto = objCTNFiscalPedido.gobjNFiscal.ColItensNF(GridItens.Row).sProduto
            objKitVenda.sUM = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)
            objKitVenda.dQuantidade = dQuantidade
            
            objCTNFiscalPedido.gobjNFiscal.ColItensNF.Item(GridItens.Row).sUnidadeMed = objKitVenda.sUM
            objCTNFiscalPedido.gobjNFiscal.ColItensNF.Item(GridItens.Row).dQuantidade = dQuantidade
        
            lErro = CF("ProdutoKitVenda_Atualiza_Romaneio", objProduto, objKitVenda, objCTNFiscalPedido.gobjNFiscal.ColItensNF(GridItens.Row).colItensRomaneioGrade)
            If lErro <> SUCESSO Then gError 177586
            
            Call objCTNFiscalPedido.Atualiza_Grid_Alocacao(objCTNFiscalPedido.gobjNFiscal.ColItensNF.Item(GridItens.Row))

        End If
           
        'Recalcula o peso bruto e o peso líquido
        lErro = objCTNFiscalPedido.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 91458
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91459
    
    lErro = CF2(objCTNFiscalPedido, "NFFP_Calcula_PesoKg", GridItens.Row)
    If lErro <> SUCESSO Then gError 177107
 
    dPrecoUnitario = StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_ValorUnitario_Col))

    If dPrecoUnitario > 0 Then Call objCTNFiscalPedido.PrecoTotal_Calcula(objCTNFiscalPedido.GridItens.Row, False)
    
    '*** incluidos p/tratamento de tributacao ****************
    If objCTNFiscalPedido.GridItens.Row <= objCTNFiscalPedido.objGridItens.iLinhasExistentes Then
        lErro = objCTNFiscalPedido.gobjTribTab.Alteracao_Item_Grid(objCTNFiscalPedido.GridItens.Row)
        If lErro <> SUCESSO Then gError 91460
    End If
    '*** fim tributacao

    Call ValorTotal_Calcula

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 23796 '??? Jones 09/02/05 passou a tratar 23796
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_SEM_ALMOX_PADRAO2", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 91456
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_FATURAR_MENOR", gErr, objCTNFiscalPedido.gcolItemQuant(objCTNFiscalPedido.GridItens.Row))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91460, 91455, 91457, 91459, 91458
            objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_Quantidade_Col) = Formata_Estoque(dQuantidadeAnterior)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 177107, 177585, 177586, 177587, 196964
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157635)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_PrecoUnitario(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Preço Unitário que está deixando de ser a corrente

Dim lErro As Long
Dim dQuantidade As Double
Dim dQuantCancelada As Double

On Error GoTo Erro_Saida_Celula_PrecoUnitario

    Set objGridInt.objControle = objCTNFiscalPedido.PrecoUnitario

    'Verifica o preenchimento de Preço Unitário
    If Len(Trim(objCTNFiscalPedido.PrecoUnitario.Text)) > 0 Then

        If giComplemento = MARCADO Then
            lErro = Valor_NaoNegativo_Critica(objCTNFiscalPedido.PrecoUnitario.Text)
        Else
            lErro = Valor_Positivo_Critica(objCTNFiscalPedido.PrecoUnitario.Text)
        End If
        If lErro <> SUCESSO Then gError 91460

        objCTNFiscalPedido.PrecoUnitario = Format(objCTNFiscalPedido.PrecoUnitario.Text, gobjFAT.sFormatoPrecoUnitario)

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91461
    
    Call objCTNFiscalPedido.PrecoTotal_Calcula(objCTNFiscalPedido.GridItens.Row, False)

''*** incluidos p/tratamento de tributacao *******************************
    lErro = objCTNFiscalPedido.gobjTribTab.Alteracao_Item_Grid(objCTNFiscalPedido.GridItens.Row)
    If lErro <> SUCESSO Then gError 91462
'*** fim tributacao

    Call ValorTotal_Calcula

    Saida_Celula_PrecoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_PrecoUnitario:

    Saida_Celula_PrecoUnitario = gErr

    Select Case gErr

        Case 91460, 91461, 91462
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157636)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_PercentDesc(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual de Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPrecoUnitario As Double
Dim dDesconto As Double
Dim dValorTotal As Double
Dim dPercentDescAnterior As Double

On Error GoTo Erro_Saida_Celula_PercentDesc
    
    Set objGridInt.objControle = objCTNFiscalPedido.PercentDesc

    dPercentDesc = StrParaDbl(objCTNFiscalPedido.PercentDesc.Text)
    
    dPercentDescAnterior = PercentParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_PercDesc_Col)) * 100

    'verifica se o percentual está preenchido
    If Len(Trim(objCTNFiscalPedido.PercentDesc.Text)) > 0 Then
        'Critica a procentagem
        lErro = Porcentagem_Critica(objCTNFiscalPedido.PercentDesc.Text)
        If lErro <> SUCESSO Then gError 91463
               
        If dPercentDesc <> dPercentDescAnterior Then
        
            'Verifica se o percentual é de 100%
            If dPercentDesc = 100 Then gError 91464
    
            objCTNFiscalPedido.PercentDesc.Text = Format(dPercentDesc, "Fixed")
    
        End If
    Else
        objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, iGrid_Desconto_Col) = ""
    End If

    If Len(Trim(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_ValorUnitario_Col))

    If Abs(dPercentDesc - dPercentDescAnterior) > DELTA_VALORMONETARIO2 Then
        objCTNFiscalPedido.iDescontoAlterado = REGISTRO_ALTERADO
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91465

    If Len(Trim(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_ValorUnitario_Col))) > 0 Then dPrecoUnitario = CDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_ValorUnitario_Col))
    If dPrecoUnitario > 0 Then
    
        Call objCTNFiscalPedido.PrecoTotal_Calcula(objCTNFiscalPedido.GridItens.Row, False)
        
''*** incluidos p/tratamento de tributacao *******************************
        lErro = objCTNFiscalPedido.gobjTribTab.Alteracao_Item_Grid(objCTNFiscalPedido.GridItens.Row)
        If lErro <> SUCESSO Then gError 91466
'*** fim tributacao

        Call ValorTotal_Calcula

    End If

    Saida_Celula_PercentDesc = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentDesc:

    Saida_Celula_PercentDesc = gErr

    Select Case gErr

        Case 91466, 91463, 91465
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91464
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_100", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157637)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPrecoUnitario As Double
Dim dQuantidade As Double
Dim dPrecoTotal As Double
Dim dDesconto As Double
Dim dPercentDesc As Double
Dim iDescontoAlterado As Integer

On Error GoTo Erro_Saida_Celula_Desconto

    Set objGridInt.objControle = objCTNFiscalPedido.Desconto

    iDescontoAlterado = False

    'Veifica o preenchimento de Desconto
    If Len(Trim(objCTNFiscalPedido.Desconto.ClipText)) > 0 Then
        'Faz a critica do Desconot
        lErro = Valor_NaoNegativo_Critica(objCTNFiscalPedido.Desconto.Text)
        If lErro <> SUCESSO Then gError 91467
        'recolhe o valor do desconto informado
        dDesconto = CDbl(objCTNFiscalPedido.Desconto.Text)
    
    End If
        
    'Verifica se o valor de desconto foi alterado
    If StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, iGrid_Desconto_Col)) <> dDesconto Then iDescontoAlterado = True
    'Se o desconto foi alterado
    If iDescontoAlterado = True Then
        objCTNFiscalPedido.iDescontoAlterado = REGISTRO_ALTERADO
        'recolhe a quantidade, preço unitário  e calcula o preço total
        dQuantidade = StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, iGrid_Quantidade_Col))
        dPrecoUnitario = StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_ValorUnitario_Col))
        dPrecoTotal = dQuantidade * dPrecoUnitario
        'Se o preço total for positivo
        If dPrecoTotal > 0 Then
            'Verifica se o desconto ultrapassa o preço total
            If dDesconto >= dPrecoTotal Then gError 91468
            'Calcula p percentual de desconto em cima do novo desconto
            dPercentDesc = dDesconto / dPrecoTotal
            'Coloca o Percentual de desconto na tela
            objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_PercDesc_Col) = Format(dPercentDesc, "Percent")

        End If
    End If
    
    If dDesconto <> StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, iGrid_Desconto_Col)) Then iDescontoAlterado = True
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 91469

    If iDescontoAlterado = True Then

        Call objCTNFiscalPedido.PrecoTotal_Calcula(objCTNFiscalPedido.GridItens.Row, False)
        
        ''*** incluidos p/tratamento de tributacao *******************************
        lErro = objCTNFiscalPedido.gobjTribTab.Alteracao_Item_Grid(objCTNFiscalPedido.GridItens.Row)
        If lErro <> SUCESSO Then gError 91470
        '*** fim tributacao
        
        Call ValorTotal_Calcula
    
    End If
    
    Saida_Celula_Desconto = SUCESSO

    Exit Function

Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = gErr

    Select Case gErr

        Case 91467, 91470
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91468
            Call Rotina_Erro(vbOKOnly, "ERRO_DESCONTO_MAIOR_OU_IGUAL_PRECO_TOTAL", gErr, objCTNFiscalPedido.GridItens.Row, dDesconto, dPrecoTotal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 91469
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157638)

    End Select

    Exit Function

End Function

Public Property Get Controls() As Object
    Set Controls = objCTNFiscalPedido.objUserControl.Controls
End Property

Public Property Get objUserControl() As Object
    Set objUserControl = objCTNFiscalPedido.objUserControl
End Property

Public Property Get objGridItens() As Object
    Set objGridItens = objCTNFiscalPedido.objGridItens
End Property

Public Property Get ClienteLabel() As Object
     Set ClienteLabel = objUserControl.Controls("ClienteLabel")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Sub Cliente_Change()

    objCTNFiscalRem.iAlterado = REGISTRO_ALTERADO
    objCTNFiscalRem.iClienteAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Cliente_Formata(lCliente As Long)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Formata

    Cliente.Text = lCliente
    
    'Busca o Cliente no BD
    lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
    If lErro <> SUCESSO Then gError 91477

    lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
    If lErro <> SUCESSO Then gError 91478
    
    'Preenche ComboBox de Filiais
    lErro = CF("Filial_Preenche", Filial, colCodigoNome)
    lErro = CF("Filial_Preenche", objCTNFiscalRem.FilialEntrega, colCodigoNome)

    objCTNFiscalRem.iClienteAlterado = 0
    
    Exit Sub

Erro_Cliente_Formata:

    Select Case gErr
    
        Case 91477, 91478
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157639)

    End Select

    Exit Sub

End Sub

Public Sub Fornecedor_Change()
    
    objCTNFiscalRem.iAlterado = REGISTRO_ALTERADO
    objCTNFiscalRem.iFornecedorAlterado = REGISTRO_ALTERADO

End Sub

Public Function ValorTotal_Calcula() As Long

Dim dValorTotal As Double
Dim lErro As Long

On Error GoTo Erro_ValorTotal_Calcula

    lErro = objCTNFiscalPedido.gobjTribTab.ValorTotal_Calcula(dValorTotal)
    If lErro <> SUCESSO Then gError 91479

    objCTNFiscalPedido.ValorTotal.Caption = Format(dValorTotal, "Standard")

    Call objCTNFiscalPedido.ValorDescontoItens_Calcula

    ValorTotal_Calcula = SUCESSO

    Exit Function

Erro_ValorTotal_Calcula:

    ValorTotal_Calcula = gErr

    Select Case gErr

        Case 91479

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157640)

    End Select

    Exit Function

End Function

Public Sub BotaoPedido_Click()

Dim lErro As Long
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim objBloqueioPV As ClassBloqueioPV

On Error GoTo Erro_BotaoPedido_Click

    'Verfifica se o Pedido e a FilialPedido foram informadas
    If Len(Trim(objCTNFiscalPedido.Pedido.Text)) = 0 Then gError 91481
    If Len(Trim(objCTNFiscalPedido.FilialPedido.Text)) = 0 Then gError 91482

    objPedidoVenda.lCodigo = CLng(objCTNFiscalPedido.Pedido.Text)
    objPedidoVenda.iFilialEmpresa = Codigo_Extrai(objCTNFiscalPedido.FilialPedido.Text)
    
    'Verifica se o pedido está baixado
    lErro = CF("PedidoVendaBaixado_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 46135 Then gError 91483
    
    If lErro = SUCESSO Then gError 91484

    'Busca o pedido na tabela de Pedidos de Venda
    lErro = CF("PedidoDeVenda_Le", objPedidoVenda)
    If lErro <> SUCESSO And lErro <> 26509 Then gError 91485
    
    If lErro <> SUCESSO Then gError 91486  'Não encontrou
    
    lErro = CF("BloqueiosPV_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 91487
    
    For Each objBloqueioPV In objPedidoVenda.colBloqueiosPV
        If objBloqueioPV.dtDataLib = DATA_NULA Then
            If objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_PARCIAL And objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_NAO_RESERVA And objBloqueioPV.iTipoDeBloqueio <> BLOQUEIO_CREDITO Then gError 91488
        End If
    Next
    
    'Verifica se a Filial de Faturamento do Pedido coincide com a filial de trabalho
    If objPedidoVenda.iFilialEmpresaFaturamento <> giFilialEmpresa Then gError 91489
    
    'Traz os dados do pedido para a tela
    lErro = Traz_Dados_Pedido_Tela(objPedidoVenda)
    If lErro <> SUCESSO Then gError 91490
    
    Exit Sub

Erro_BotaoPedido_Click:

    Select Case gErr

        Case 91481
            Call Rotina_Erro(vbOKOnly, "ERRO_CODPEDIDO_NAO_INFORMADO", gErr)

        Case 91482
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALPEDIDO_NAO_INFORMADA", gErr)

        Case 91483, 91485, 91490, 91487

        Case 91484
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDOVENDA_BAIXADO", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case 91486
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case 91488
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_BLOQUEIO", gErr, objPedidoVenda.lCodigo)

        Case 91489
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFATURAMENTO_DIFERENTE", gErr, objPedidoVenda.iFilialEmpresaFaturamento, objPedidoVenda.lCodigo, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157641)

    End Select

    Exit Sub

End Sub

Public Function Traz_Dados_Pedido_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal
Dim objCliente As New ClassCliente, bFaturadoParc As Boolean
Dim objItemNF As ClassItemNF
Dim iIndice As Integer, bEmissaoForcaRecalculo As Boolean

On Error GoTo Erro_Traz_Dados_Pedido_Tela

    'Lê todos oa dados do pedido de venda
    lErro = CF("PedidoDeVenda_Le_Completo", objPedidoVenda, MARCADO)
    If lErro <> SUCESSO And lErro <> 27633 Then gError 91491
    If lErro <> SUCESSO Then gError 91492
    
    'Lê as reservas do pedido de venda
    lErro = CF("ReservasPedido_Le", objPedidoVenda)
    If lErro <> SUCESSO Then gError 91493

    'Transfere os dados de objPedidoVenda para objNFiscal
    Call objCTNFiscalPedido.Transfere_Dados_Pedido_NFiscal(objPedidoVenda, objNFiscal, bFaturadoParc)

    Call CF2(Me, "NFCF_DataSaida", objNFiscal)
      
    'Carrega os dados de objNFiscal na tela.
    lErro = Traz_NFiscal_Tela(objNFiscal, objPedidoVenda.lNumIntDoc)
    If lErro <> SUCESSO Then gError 91494

    Set objCTNFiscalPedido.gobjNFiscal = objNFiscal
    
    'Guarda o código e a filial do Pedido
    objCTNFiscalPedido.glCodPedido = objPedidoVenda.lCodigo
    objCTNFiscalPedido.giFilialPedido = objPedidoVenda.iFilialEmpresa
    
    'Armazena o dado de fatura integral do pedido
    objCTNFiscalPedido.giFaturaIntegral = objPedidoVenda.iFaturaIntegral

    'se o faturamento é parcial
    If bFaturadoParc Then
    
        lErro = CF2(objCTNFiscalPedido, "NFFP_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177108
    
        'Recalcula o peso bruto e o peso líquido
        lErro = objCTNFiscalPedido.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 91495
    
    End If
    
    objCTNFiscalPedido.HoraSaida.PromptInclude = False
    objCTNFiscalPedido.HoraSaida.Text = ""
    objCTNFiscalPedido.HoraSaida.PromptInclude = True
    
    objCTNFiscalPedido.gbCarregandoTela = False
    
    Set objCTNFiscalPedido.gobjNFiscal = objNFiscal
    
    If (Month(objPedidoVenda.dtDataEmissao) <> Month(objNFiscal.dtDataEmissao)) Or (Year(objPedidoVenda.dtDataEmissao) <> Year(objNFiscal.dtDataEmissao)) Then
        bEmissaoForcaRecalculo = True
    Else
        bEmissaoForcaRecalculo = False
    End If
    
    If bFaturadoParc Or bEmissaoForcaRecalculo Then Call ValorTotal_Calcula
    
    Traz_Dados_Pedido_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Pedido_Tela:

    Traz_Dados_Pedido_Tela = gErr

    Select Case gErr

        Case 91491, 91493, 91494, 91495, 177108, 177220

        Case 91492
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_NAO_CADASTRADO1", gErr, objPedidoVenda.lCodigo, objPedidoVenda.iFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157642)

    End Select

    Exit Function

End Function

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objCliente As New ClassCliente
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sProdutoFormatado As String
Dim iPreenchido As Integer
Dim sContaMascarada As String
Dim dQuantidadeConvertida As Double
Dim dQuantidade As Double
Dim sProduto As String
Dim sUM As String
Dim objProdutoCategoria As New ClassProdutoCategoria
Dim objCatProdItem As ClassCategoriaProdutoItem
Dim objPlanoConta As New ClassPlanoConta
Dim objProduto As New ClassProduto
Dim objFilial As New ClassFilialCliente
Dim sContaTela As String
Dim objTribItemNFAux As ClassTribItemNFAux, dValorAux As Double

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
        
        Case QUANT_ESTOQUE
            For iLinha = 1 To objGridItens.iLinhasExistentes
            
                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    lErro = CF("UMEstoque_Conversao", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)), dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 96090

                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
       
        Case ALMOXARIFADO1
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_AlmoxAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_AlmoxAloc_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

            
        Case DATA_EMISSAO
            If Len(objCTNFiscalPedido.DataEmissao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(objCTNFiscalPedido.DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
        
        Case DATA_SAIDA
            If Len(objCTNFiscalPedido.DataSaida.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(objCTNFiscalPedido.DataSaida.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case DESCONTO1
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_Desconto_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case DESCRICAO_ITEM
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_DescProduto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_DescProduto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case FILIAL1
            If Len(Filial.Caption) > 0 And objCTNFiscalRem.iFornecedorAtivo = FORN_INATIVO Then
                objMnemonicoValor.colValor.Add Filial.Caption
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case CLIENTE1
            If Len(Cliente.Caption) > 0 Then
                objMnemonicoValor.colValor.Add Cliente.Caption
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case CLIENTE_CODIGO
            If Len(Cliente.Caption) > 0 Then
            
                objCliente.sNomeReduzido = Cliente.Caption
                
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 96091
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
            Else
                objMnemonicoValor.colValor.Add 0
             End If

        Case NATUREZA_OP
            If Len(objCTNFiscalPedido.gobjTribTab.sNatOpInterna) > 0 Then
                objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.sNatOpInterna
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case PRODUTO1
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case QUANTIDADE1
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                    sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 96097
                    
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case Serie1
            If Len(objCTNFiscalPedido.Serie.Text) > 0 Then
                objMnemonicoValor.colValor.Add objCTNFiscalPedido.Serie.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case TIPO_NOTA_FISCAL
            If Len(TipoNFiscal.Text) > 0 Then
                objMnemonicoValor.colValor.Add TipoNFiscal.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case UNIDADE_MED
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case VALOR_TOTAL
            If Len(objCTNFiscalPedido.ValorTotal.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ValorTotal.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case PRECO_UNITARIO
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_ValorUnitario_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_ValorUnitario_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case PRECO_TOTAL
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_PrecoTotal_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case ICMS
            If Len(objCTNFiscalPedido.ICMSValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ICMSValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case ICMS_DIFAL_REMETENTE
            objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.objNFTributacao.dICMSInterestVlrUFRemet

        Case ICMS_DIFAL_DESTINO
            objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.objNFTributacao.dICMSInterestVlrUFDest

        Case ICMS_FCP
            objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.objNFTributacao.dICMSInterestVlrFCPUFDest
        
        Case ICMSSUBST
            If Len(objCTNFiscalPedido.ICMSSubstValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ICMSSubstValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_FRETE
            If Len(objCTNFiscalPedido.ValorFrete.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ValorFrete.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_SEGURO
            If Len(objCTNFiscalPedido.ValorSeguro.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ValorSeguro.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_DESPESAS
            If Len(objCTNFiscalPedido.ValorDespesas.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ValorDespesas.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case IPI
            If Len(objCTNFiscalPedido.IPIValor1.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.IPIValor1.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_DESCONTO
            If Len(objCTNFiscalPedido.ValorDesconto.Text) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ValorDesconto.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case VALOR_PRODUTOS
            If Len(objCTNFiscalPedido.ValorProdutos.Caption) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(objCTNFiscalPedido.ValorProdutos.Caption)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case FORNECEDOR1
            If Len(Trim(Fornecedor.Caption)) > 0 Then
                objMnemonicoValor.colValor.Add Fornecedor.Caption
            Else
                objMnemonicoValor.colValor.Add ""
            End If
            
        Case FILIAL_FORNECEDOR1
            If Len(Trim(Filial.Caption)) > 0 And objCTNFiscalRem.iFornecedorAtivo = FORN_ATIVO Then
                objMnemonicoValor.colValor.Add Filial.Caption
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        'Janaina
        Case PRODUTO_ALMOX
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_ProdutoAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_ProdutoAloc_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        'Janaina
        Case UNIDADE_MED_EST
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_UMAloc_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_UMAloc_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

         Case QUANT_ALOCADA
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_QuantAloc_Col))) > 0 Then
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                        objMnemonicoValor.colValor.Add CDbl(GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_QuantAloc_Col))
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        'Janaina
        Case CTACONTABILEST1
            
            For iLinha = 1 To objGridAlocacoes.iLinhasExistentes
                If Len(Trim(GridAlocacao.TextMatrix(iLinha, objCTNFiscalPedido.iGrid_QuantAloc_Col))) > 0 Then
                    
                    If objMnemonicoValor.vParam(1) = objMnemonicoValor.vParam(2) Then
                            
                        objAlmoxarifado.sNomeReduzido = CStr(objMnemonicoValor.vParam(3))

                        'Lê o Nome Reduzido do Almoxarifado
                        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                        If lErro <> SUCESSO And lErro <> 25060 Then gError 64215

                        'Se não encontrou ===> Erro
                        If lErro = SUCESSO Then

                            objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                            
                            lErro = CF("Produto_Formata", objMnemonicoValor.vParam(1), sProdutoFormatado, iPreenchido)
                            If iPreenchido <> PRODUTO_PREENCHIDO Then gError 64216

                            objEstoqueProduto.sProduto = sProdutoFormatado
                            
                            'Lê a conta contábil do Produto no Almoxarifado
                            lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
                            If lErro <> SUCESSO And lErro <> 49991 Then gError 64217

                            If Len(Trim(objEstoqueProduto.sContaContabil)) > 0 Then

                                sContaMascarada = String(STRING_CONTA, 0)

                                lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                                If lErro <> SUCESSO Then gError 64218

                                objMnemonicoValor.colValor.Add sContaMascarada
                            Else
                                objMnemonicoValor.colValor.Add ""
                            End If
                        Else
                            objMnemonicoValor.colValor.Add ""
                        End If
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case PIS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.objNFTributacao.dPISCredito
        
        Case COFINS_DEBITO_TOTAL
            objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.objNFTributacao.dCOFINSCredito
            
        Case PIS_VALOR_TOTAL
            dValorAux = 0
            For Each objTribItemNFAux In objCTNFiscalPedido.gobjTribTab.objNFTributacao.colItens
                dValorAux = dValorAux + objTribItemNFAux.dPISValor
            Next
            objMnemonicoValor.colValor.Add Arredonda_Moeda(dValorAux)
        
        Case COFINS_VALOR_TOTAL
            dValorAux = 0
            For Each objTribItemNFAux In objCTNFiscalPedido.gobjTribTab.objNFTributacao.colItens
                dValorAux = dValorAux + objTribItemNFAux.dCOFINSValor
            Next
            objMnemonicoValor.colValor.Add Arredonda_Moeda(dValorAux)
            
        Case ISS_VALOR
            objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.objNFTributacao.dISSValor

        Case ISS_INCLUSO
            objMnemonicoValor.colValor.Add objCTNFiscalPedido.gobjTribTab.objNFTributacao.iISSIncluso <> 0
        
        Case CONTACONTABILAPLIC
        
            For iLinha = 1 To objGridItens.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183787
                    
                    'Lê o produto
                    objProduto.sCodigo = sProdutoFormatado
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 183788
        
                    sContaMascarada = String(STRING_CONTA, 0)

                    lErro = Mascara_MascararConta(objProduto.sContaContabil, sContaMascarada)
                    If lErro <> SUCESSO Then gError 183804

                    objMnemonicoValor.colValor.Add sContaMascarada
                    
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
        
        Case CONTACONTABILCATPROD
                
            For iLinha = 1 To objGridItens.iLinhasExistentes
                
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iPreenchido)
                    If lErro <> SUCESSO Then gError 183789
                
                    objProdutoCategoria.sCategoria = objMnemonicoValor.vParam(1)
                    
                    objProdutoCategoria.sProduto = sProdutoFormatado
                    
                    Set objCatProdItem = New ClassCategoriaProdutoItem
                    
                    lErro = CF("ProdutoCategoria_Le_Item", objProdutoCategoria, objCatProdItem)
                    If lErro <> SUCESSO And lErro <> 183548 Then gError 183790
                    
                    If lErro = SUCESSO Then
                        
                        Set objPlanoConta = New ClassPlanoConta
                    
                        'le uma conta a partir da conta simplificada fornecida em lContaSimples. Retorna os dados em objPlanoConta
                        lErro = CF("PlanoConta_Le_ContaSimples", objCatProdItem.dvalor1, objPlanoConta)
                        If lErro <> SUCESSO And lErro <> 5451 Then gError 183791
                    
                        If lErro = SUCESSO Then
                    
                            sContaMascarada = String(STRING_CONTA, 0)
        
                            lErro = Mascara_MascararConta(objPlanoConta.sConta, sContaMascarada)
                            If lErro <> SUCESSO Then gError 183805
        
                            objMnemonicoValor.colValor.Add sContaMascarada
                            
                        Else
                    
                            objMnemonicoValor.colValor.Add ""
                    
                        End If
                    Else
                        objMnemonicoValor.colValor.Add ""
                    End If
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
                
            Next
        
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 186144
                
                If objFilial.sContaContabil <> "" Then
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 186145
                Else
                    sContaTela = ""
                End If
                objMnemonicoValor.colValor.Add sContaTela
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case Else
            lErro = objCTNFiscalPedido.gobjTribTab.Calcula_Mnemonico(objMnemonicoValor)
            If lErro <> SUCESSO Then gError 96096

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 96096
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 96092, 96093, 96094, 96095, 96090, 96091, 96097, 183787 To 183791, 183804, 183805, 186144, 186145
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157643)

    End Select

    Exit Function

End Function

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Sub PedidoLabel_Click()
    
Dim objPedidoVenda As New ClassPedidoDeVenda
Dim colSelecao As New Collection

    Call Chama_Tela("PedidoVendaLista", colSelecao, objPedidoVenda, objEventoPedido)
    
End Sub

Public Sub Show()
    objCTNFiscalPedido.Parent.Show
    objCTNFiscalPedido.Parent.SetFocus
End Sub

Private Sub Class_Terminate()
'If giDebug = 1 Then MsgBox ("Saiu")
End Sub

Private Sub objEventoPedido_evSelecao(obj1 As Object)

Dim objPedidoVenda As ClassPedidoDeVenda

    Set objPedidoVenda = obj1

    objCTNFiscalPedido.Pedido.Text = objPedidoVenda.lCodigo
    objCTNFiscalPedido.FilialPedido.Text = objPedidoVenda.iFilialEmpresa
    Call objCTNFiscalPedido.FilialPedido_Validate(bSGECancelDummy)
    
    BotaoPedido_Click
    
    Me.Show

End Sub

Private Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFiscal As ClassNFiscal

On Error GoTo Erro_objEventoNFiscal_evSelecao

    Set objNFiscal = obj1

    'Evita ter que alterar os browses a cada campo inserido na tabela de NF
    lErro = CF("NFiscal_Le", objNFiscal)
    If lErro <> SUCESSO And lErro <> 31442 Then gError 180105

    'Coloca na Tela a Nota Fiscal escolhida
    lErro = Traz_NFiscal_Tela(objNFiscal)
    If lErro <> SUCESSO Then gError 91496

    objCTNFiscalPedido.glCodPedido = objNFiscal.lNumPedidoVenda
    objCTNFiscalPedido.giFilialPedido = objNFiscal.iFilialPedido

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNFiscal_evSelecao:

    Select Case gErr

        Case 91496, 180105

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157644)

    End Select

    Exit Sub

End Sub

Public Sub GridItens_LeaveCell()
    Call Saida_Celula(objCTNFiscalPedido.objGridItens)
End Sub

Public Sub NFiscalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection

    'Chama a tela de browse das notas fiscais Remessa geradas a partir de um pedido de venda
    Call Chama_Tela("NFiscalRemPedidoLista", colSelecao, objNFiscal, objEventoNFiscal)

End Sub

Public Property Get iGrid_DescProduto_Col() As Integer
    iGrid_DescProduto_Col = objCTNFiscalPedido.iGrid_DescProduto_Col
End Property

Public Property Get iGrid_PercDesc_Col() As Integer
    iGrid_PercDesc_Col = objCTNFiscalPedido.iGrid_PercDesc_Col
End Property

Public Property Get iGrid_DataEntrega_Col() As Integer
    iGrid_DataEntrega_Col = objCTNFiscalPedido.iGrid_DataEntrega_Col
End Property

Public Property Get iGrid_Produto_Col() As Integer
    iGrid_Produto_Col = objCTNFiscalPedido.iGrid_Produto_Col
End Property

Public Property Get iGrid_UnidadeMed_Col() As Integer
    iGrid_UnidadeMed_Col = objCTNFiscalPedido.iGrid_UnidadeMed_Col
End Property

Public Property Get iGrid_PrecoTotal_Col() As Integer
    iGrid_PrecoTotal_Col = objCTNFiscalPedido.iGrid_PrecoTotal_Col
End Property

Public Property Get iGrid_ValorUnitario_Col() As Integer
    iGrid_ValorUnitario_Col = objCTNFiscalPedido.iGrid_ValorUnitario_Col
End Property

Public Property Get iGrid_ValorUnitarioMoeda_Col() As Integer
    iGrid_ValorUnitarioMoeda_Col = objCTNFiscalPedido.iGrid_ValorUnitarioMoeda_Col
End Property

Public Property Get iGrid_Desconto_Col() As Integer
    iGrid_Desconto_Col = objCTNFiscalPedido.iGrid_Desconto_Col
End Property

Public Property Get iGrid_Quantidade_Col() As Integer
    iGrid_Quantidade_Col = objCTNFiscalPedido.iGrid_Quantidade_Col
End Property

Public Property Get iGrid_Item_Col() As Integer
    iGrid_Item_Col = objCTNFiscalPedido.iGrid_Item_Col
End Property

Public Property Get iGrid_ProdutoAloc_Col() As Integer
    iGrid_ProdutoAloc_Col = objCTNFiscalPedido.iGrid_ProdutoAloc_Col
End Property

Public Property Get iGrid_AlmoxAloc_Col() As Integer
    iGrid_AlmoxAloc_Col = objCTNFiscalPedido.iGrid_AlmoxAloc_Col
End Property

Public Property Get iGrid_UMAloc_Col() As Integer
    iGrid_UMAloc_Col = objCTNFiscalPedido.iGrid_UMAloc_Col
End Property

Public Property Get iGrid_QuantAloc_Col() As Integer
    iGrid_QuantAloc_Col = objCTNFiscalPedido.iGrid_QuantAloc_Col
End Property

Public Property Get iGrid_QuantVend_Col() As Integer
    iGrid_QuantVend_Col = objCTNFiscalPedido.iGrid_QuantVend_Col
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objCTNFiscalPedido.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objCTNFiscalPedido.Enabled
End Property

Public Property Get iAlterado() As Integer
    iAlterado = objCTNFiscalPedido.iAlterado
End Property

Public Property Let iAlterado(iData As Integer)
    objCTNFiscalPedido.iAlterado = iData
End Property

Public Property Get GridItens() As Object
    Set GridItens = objCTNFiscalPedido.GridItens
End Property

Public Property Get objGridAlocacoes() As Object
    Set objGridAlocacoes = objCTNFiscalPedido.objGridAlocacoes
End Property

Public Property Get GridAlocacao() As Object
    Set GridAlocacao = objCTNFiscalPedido.GridAlocacao
End Property

Public Sub ItemNFRastro_Change()
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_Change
End Sub

Public Sub ItemNFRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_GotFocus
End Sub

Public Sub ItemNFRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_KeyPress(KeyAscii)
End Sub

Public Sub ItemNFRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.ItemNFRastro_Validate(Cancel)
End Sub

Public Sub AlmoxRastro_Change()
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_Change
End Sub

Public Sub AlmoxRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_GotFocus
End Sub

Public Sub AlmoxRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_KeyPress(KeyAscii)
End Sub

Public Sub AlmoxRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.AlmoxRastro_Validate(Cancel)
End Sub

Public Sub LoteRastro_Change()
'Rastreamento
    Call gobjRastreamento.LoteRastro_Change
End Sub

Public Sub LoteRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.LoteRastro_GotFocus
End Sub

Public Sub LoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.LoteRastro_KeyPress(KeyAscii)
End Sub

Public Sub LoteRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.LoteRastro_Validate(Cancel)
End Sub

Public Sub FilialOPRastro_Change()
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_Change
End Sub

Public Sub FilialOPRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_GotFocus
End Sub

Public Sub FilialOPRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_KeyPress(KeyAscii)
End Sub

Public Sub FilialOPRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.FilialOPRastro_Validate(Cancel)
End Sub

Public Sub QuantLoteRastro_Change()
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_Change
End Sub

Public Sub QuantLoteRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_GotFocus
End Sub

Public Sub QuantLoteRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_KeyPress(KeyAscii)
End Sub

Public Sub QuantLoteRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.QuantLoteRastro_Validate(Cancel)
End Sub

Public Sub GridRastro_Click()
'Rastreamento
    Call gobjRastreamento.GridRastro_Click
End Sub

Public Sub GridRastro_EnterCell()
'Rastreamento
    Call gobjRastreamento.GridRastro_EnterCell
End Sub

Public Sub GridRastro_GotFocus()
'Rastreamento
    Call gobjRastreamento.GridRastro_GotFocus
End Sub

Public Sub GridRastro_KeyPress(KeyAscii As Integer)
'Rastreamento
    Call gobjRastreamento.GridRastro_KeyPress(KeyAscii)
End Sub

Public Sub GridRastro_LeaveCell()
'Rastreamento
    Call gobjRastreamento.GridRastro_LeaveCell
End Sub

Public Sub GridRastro_Validate(Cancel As Boolean)
'Rastreamento
    Call gobjRastreamento.GridRastro_Validate(Cancel)
End Sub

Public Sub GridRastro_RowColChange()
'Rastreamento
    Call gobjRastreamento.GridRastro_RowColChange
End Sub

Public Sub GridRastro_Scroll()
'Rastreamento
    Call gobjRastreamento.GridRastro_Scroll
End Sub

Public Sub GridRastro_KeyDown(KeyCode As Integer, Shift As Integer)
'Rastreamento
    Call gobjRastreamento.GridRastro_KeyDown(KeyCode, Shift)
End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim colSiglas As New Collection
Dim objUM As ClassUnidadeDeMedida
Dim sUM As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 46713

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
    
        'Produto
        Case objCTNFiscalPedido.Produto.Name, objCTNFiscalPedido.UnidadeMed.Name
        
        'Nas demais
        Case objCTNFiscalPedido.PrecoUnitario.Name, objCTNFiscalPedido.PercentDesc.Name, objCTNFiscalPedido.Desconto.Name

            If iProdutoPreenchido = PRODUTO_VAZIO Or objCTNFiscalPedido.gcolItemPedido.Count = 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        
        Case objCTNFiscalPedido.Quantidade.Name

            If iProdutoPreenchido = PRODUTO_VAZIO Or objCTNFiscalPedido.gcolItemPedido.Count = 0 Or objCTNFiscalPedido.giFaturaIntegral <> 0 Or left(GridItens.TextMatrix(iLinha, 0), 1) = "#" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Else
            lErro = CF2(Me, "NFRP_Rotina_Grid_Enable", iLinha, objControl, iCaminho, sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 177221

    End Select

    'Rastreamento
    lErro = gobjRastreamento.Rotina_Grid_Enable(iLinha, objControl, iCaminho)
    If lErro <> SUCESSO Then gError 83378

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 46713, 83378, 177221

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157645)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLotes_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long

On Error GoTo Erro_BotaoLotes_Click
    
    Call gobjRastreamento.BotaoLotes_Click
                    
    Exit Sub

Erro_BotaoLotes_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157646)
    
    End Select
    
    Exit Sub

End Sub

Function AlocaAlmoxarifadoPradrao(dQuantidade As Double, objProduto As ClassProduto, sUM As String) As Long
'Tenta fazer a alocação do produto no almoxarifado padrão. Caso não consiga chama a tela de Alocação de produto.

Dim lErro As Long
Dim dQuantAlocar As Double
Dim dFator As Double
Dim iAlmoxarifado As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim dQuantDisponivel As Double
Dim objItemNF As ClassItemNF
Dim colOutrosProdutos As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim iIndice As Integer
Dim objItemNFAloc As ClassItemNFAlocacao
Dim sProdutoEnxuto As String
Dim iLinhas As Integer
Dim dAcrescimo As Double
Dim iNumCasasDec As Integer
Dim objItemPV As New ClassItemPedido
Dim colReservaBD As New colReservaItem
Dim dTotal As Double
Dim objItemNFAux As ClassItemNF
Dim objItemNFAlocAux As ClassItemNFAlocacao

On Error GoTo Erro_AlocaAlmoxarifadoPradrao

    'Faz a conversão da UM da tela para a UM de estoque
    lErro = CF("UM_Conversao", objProduto.iClasseUM, sUM, objProduto.sSiglaUMEstoque, dFator)
    If lErro <> SUCESSO Then gError 94329 '46743

    'Converte a quantidade para a UM de estoque
    dQuantAlocar = dQuantidade * dFator

    'Busca o Almoxarifado padrão
    lErro = CF("AlmoxarifadoPadrao_Le", giFilialEmpresa, objProduto.sCodigo, iAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 23796 Then gError 94330 '46745
    If lErro = 23796 Then gError 23796
    
    'Se encontrou
    If iAlmoxarifado > 0 Then

        objAlmoxarifado.iCodigo = iAlmoxarifado
        'Lê o Aloxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> 25056 And lErro <> SUCESSO Then gError 94331 '46746
        If lErro = 25056 Then gError 94332 '46749
        
        objEstoqueProduto.sProduto = objProduto.sCodigo
        objEstoqueProduto.iAlmoxarifado = iAlmoxarifado
        'Le os estoques desse produto nesse almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 94333 '46744
        If lErro = 21306 Then gError 94334 '46750 'Não encontrou

        If Codigo_Extrai(TipoNFiscal.Text) <> DOCINFO_NFISRMB3PV Then

            objItemPV.lNumIntDoc = objCTNFiscalPedido.gcolItemPedido(objCTNFiscalPedido.GridItens.Row)
            objItemPV.sProduto = objProduto.sCodigo
            
            lErro = CF("ReservasItemPV_Le_NumIntOrigem", objItemPV, colReservaBD)
            If lErro <> SUCESSO And lErro <> 51601 Then gError 94335 '62095
        
                
            For iIndice = 1 To colReservaBD.Count
                If objEstoqueProduto.iAlmoxarifado = colReservaBD(iIndice).iAlmoxarifado Then
                    objEstoqueProduto.dQuantDispNossa = objEstoqueProduto.dQuantDispNossa + colReservaBD(iIndice).dQuantidade
                    Exit For
                End If
            Next
        
            dQuantDisponivel = objEstoqueProduto.dQuantDisponivel
        
        Else
            dQuantDisponivel = objEstoqueProduto.dQuantBenef3
        
        End If
            
        'Remove as alocações do produto do grid de alocações
        Call objCTNFiscalPedido.Remove_Alocacoes(objCTNFiscalPedido.GridItens.Row)

        iLinhas = objGridAlocacoes.iLinhasExistentes

        dQuantAlocar = Arredonda_Estoque(dQuantAlocar)
        
        'Verifica se a Quantidade disponível é maior que a quantidade a alocar
        If dQuantAlocar <= dQuantDisponivel Then

            iLinhas = iLinhas + 1
            'Coloca no Grid alocações a alocação para o almoxarifado padrão
            objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_Item_Col) = objCTNFiscalPedido.GridItens.Row
            objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_ProdutoAloc_Col) = objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_Produto_Col)
            objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_AlmoxAloc_Col) = objAlmoxarifado.sNomeReduzido
            objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_QuantAloc_Col) = Formata_Estoque(dQuantAlocar)
            objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_QuantVend_Col) = Formata_Estoque(dQuantAlocar)
            objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_UMAloc_Col) = objProduto.sSiglaUMEstoque
                                           
            'Atualiza a coleção de alocações em gcolItensNf
            Set objItemNFAux = gcolItensNf.Item(objCTNFiscalPedido.GridItens.Row)
            objItemNFAux.colAlocacoes.Add objAlmoxarifado.iCodigo, objAlmoxarifado.sNomeReduzido, dQuantAlocar
                                            
        'Se não for
        Else
            'Se a função não estiver sendo usada pelo carregamento da tela
            'se prepara para chamar a tela de localização
            If Not objCTNFiscalPedido.gbCarregandoTela Then
           
                Set objItemNF = New ClassItemNF
                'Recolhe os dados do item
                objItemNF.iItem = objCTNFiscalPedido.GridItens.Row
                objItemNF.sProduto = objProduto.sCodigo
                objItemNF.sDescricaoItem = objProduto.sDescricao
                objItemNF.dQuantidade = dQuantidade
                objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
                objItemNF.lNumIntItemPedVenda = objCTNFiscalPedido.gcolItemPedido(objCTNFiscalPedido.GridItens.Row)
                
                'Recolhe todos os produtos dos outros itens
                For iIndice = 1 To objCTNFiscalPedido.objGridItens.iLinhasExistentes

                    If iIndice <> objCTNFiscalPedido.GridItens.Row Then

                        lErro = CF("Produto_Formata", objCTNFiscalPedido.GridItens.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_Produto_Col), sProduto, iPreenchido)
                        If lErro <> SUCESSO Then gError 94336 '46747
                        'Adiciona na coleção de produtos
                        colOutrosProdutos.Add sProduto
                    End If
                Next
    
                'Chama a tela de Localização de Produto
                Call Chama_Tela_Modal("LocalizacaoProduto1", objItemNF, colOutrosProdutos, dQuantAlocar, Codigo_Extrai(TipoNFiscal.Text))
                If giRetornoTela = vbCancel Then gError 94337 '46752 'Se nada foi feito lá
                If giRetornoTela = vbOK Then
    
                    'Se o produto foi substituido
                    If objProduto.sCodigo <> objItemNF.sProduto Then gError 94338 '46751
                                    
                    sProdutoEnxuto = String(STRING_PRODUTO, 0)
    
                    lErro = Mascara_RetornaProdutoEnxuto(objItemNF.sProduto, sProdutoEnxuto)
                    If lErro Then gError 94339 '46748
    
                    objCTNFiscalPedido.Produto.PromptInclude = False
                    objCTNFiscalPedido.Produto.Text = sProdutoEnxuto
                    objCTNFiscalPedido.Produto.PromptInclude = True
                    
                    'Calcula o Total Alocado
                    For Each objItemNFAloc In objItemNF.colAlocacoes
                        dTotal = dTotal + objItemNFAloc.dQuantidade
                    Next
                    
                    'Para cada alocação feita para o item
                    For Each objItemNFAloc In objItemNF.colAlocacoes
    
                        iLinhas = iLinhas + 1
                        'coloca no grid de alocações os dados das alocações
                        objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_Item_Col) = objItemNF.iItem
                        objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_ProdutoAloc_Col) = objCTNFiscalPedido.Produto.Text
                        objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_AlmoxAloc_Col) = objItemNFAloc.sAlmoxarifado
                        objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_QuantAloc_Col) = Formata_Estoque(objItemNFAloc.dQuantidade)
                        objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_QuantVend_Col) = Formata_Estoque(dTotal)
                        objCTNFiscalPedido.GridAlocacao.TextMatrix(iLinhas, objCTNFiscalPedido.iGrid_UMAloc_Col) = objItemNF.sUMEstoque
     
                        'Atualiza a coleção de alocações em gcolItensNf
                        Set objItemNFAux = gcolItensNf.Item(objCTNFiscalPedido.GridItens.Row)
                        objItemNFAux.colAlocacoes.Add objItemNFAloc.iAlmoxarifado, objItemNFAloc.sAlmoxarifado, objItemNFAloc.dQuantidade
                    
                    Next
                
                    If Abs(dQuantAlocar - dTotal) > QTDE_ESTOQUE_DELTA Then
                        objCTNFiscalPedido.Quantidade.Text = Formata_Estoque(dTotal / dFator)
                    Else
                        objCTNFiscalPedido.Quantidade.Text = Formata_Estoque(StrParaDbl(objCTNFiscalPedido.Quantidade.Text))
                    End If
                    
                End If
            End If
        End If
        
        'atualiza o numero de linhas existentes no grid de alocações
        objCTNFiscalPedido.objGridAlocacoes.iLinhasExistentes = iLinhas
    End If

    AlocaAlmoxarifadoPradrao = SUCESSO

    Exit Function

Erro_AlocaAlmoxarifadoPradrao:

    AlocaAlmoxarifadoPradrao = gErr

    Select Case gErr

        Case 94329, 94333, 94330, 94331, 94336, 94339, 23796

        Case 94332
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, iAlmoxarifado)

        Case 94334
            Call Rotina_Erro(vbOKOnly, "ERRO_ESTOQUE_PRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case 94338
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO", gErr)

        Case 94337
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_LOCALIZACAO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157648)

    End Select

    Exit Function

End Function

Public Sub BotaoLocalizacao_Click()

Dim lErro As Long
Dim iIndice As Integer
Dim iPreenchido As Integer
Dim sProduto As String
Dim dQuantidade As Double
Dim ColItensNF As New ColItensNF
Dim colAlocacoes As ColAlocacoesItemNF
Dim objProduto As New ClassProduto
Dim objItemNF As New ClassItemNF
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim lbLabel As label
Dim dPrecoUnitario As Double
Dim vbMsgRes As VbMsgBoxResult
Dim objNFiscal As New ClassNFiscal
Dim colItensRomaneio As New Collection
Dim iIndice2 As Integer
Dim objItemRomaneio As ClassItemRomaneioGrade
Dim objItemRomaneioAux As ClassItemRomaneioGrade
Dim objProdutoAux As New ClassProduto
Dim bAchou  As Boolean

On Error GoTo Erro_BotaoLocalizacao_Click:
    
    'Se não existir itens no grid, sai.
    If objCTNFiscalPedido.objGridItens.iLinhasExistentes = 0 Then gError 43699
    
    'Para cada ítem do grid
    For iIndice = 1 To objCTNFiscalPedido.objGridItens.iLinhasExistentes

        Set colItensRomaneio = New Collection

        'recolhe a quantidade
        dQuantidade = StrParaDbl(objCTNFiscalPedido.GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
        'Critica o produto
        lErro = CF("Produto_Critica_Filial2", objCTNFiscalPedido.GridItens.TextMatrix(iIndice, iGrid_Produto_Col), objProduto, iPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 42309
        If lErro = 51381 Then gError 86376
                
        'Se o produto é gerencial sem ser de grade ==> Erro
        If lErro = 86295 And Len(Trim(objProduto.sGrade)) = 0 And objProduto.iKitVendaComp <> MARCADO Then
            gError 86377
        ElseIf lErro = 86295 Then
        
             If objProduto.iKitVendaComp <> MARCADO Then
    
                'Lê os filhos analíticos do produto pai de grade passado
                lErro = CF("Produto_Le_Filhos_Grade", objProduto, colItensRomaneio)
                If lErro <> SUCESSO And lErro <> 86304 Then gError 86306
                
                'O produto não tem filhos de grade ou seus filhos são analíticos
                If lErro = 86304 Then gError 86307
            
            Else
            
                lErro = CF("ProdutoKitVenda_Le_Romaneio", objProduto, colItensRomaneio)
                If lErro <> SUCESSO Then gError 177582
            
            End If
            
            'Para cada item lido
            For iIndice2 = colItensRomaneio.Count To 1 Step -1
            
                Set objItemRomaneio = colItensRomaneio(iIndice2)
                
                'Lê o Produto
                objProdutoAux.sCodigo = objItemRomaneio.sProduto
                lErro = CF("Produto_Le", objProdutoAux)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 86378
                If lErro <> SUCESSO Then gError 86379
                
                'Se o produto controlar Reserva
                If objProdutoAux.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    bAchou = False
                    For Each objItemRomaneioAux In objCTNFiscalPedido.gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade
                        If objItemRomaneio.sProduto = objItemRomaneioAux.sProduto Then
                            Call Transfere_Dados_ItensRomaneio(colItensRomaneio, objCTNFiscalPedido.gobjNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
                            bAchou = True
                            Exit For
                        End If
                    Next
                End If
                If Not bAchou Or objProdutoAux.iControleEstoque = PRODUTO_CONTROLE_SEM_ESTOQUE Then
                    colItensRomaneio.Remove iIndice2
                End If
            Next
        
        End If
        'Se o produto fizer controle de estoque
        If objProduto.iControleEstoque <> PRODUTO_CONTROLE_SEM_ESTOQUE Then

            Set colAlocacoes = New ColAlocacoesItemNF
            'Recolhe as alocações do item
            If objCTNFiscalPedido.gobjNFiscal.ColItensNF(iIndice).iPossuiGrade = DESMARCADO And objProduto.iKitVendaComp = DESMARCADO Then Call objCTNFiscalPedido.Recolhe_Alocacoes(iIndice, colAlocacoes)

            ColItensNF.Add 0, iIndice, objProduto.sCodigo, objCTNFiscalPedido.GridItens.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_UnidadeMed_Col), dQuantidade, 0, 0, 0, DATA_NULA, objCTNFiscalPedido.GridItens.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_DescProduto_Col), 0, 0, 0, 0, objCTNFiscalPedido.gcolItemPedido(iIndice), colAlocacoes, 0, "", "", 0, 0, "", 0, 0, 0, objProduto.sSiglaUMEstoque, 0, 0
                        
            ColItensNF(ColItensNF.Count).iPossuiGrade = objCTNFiscalPedido.gobjNFiscal.ColItensNF(iIndice).iPossuiGrade

            If objProduto.iKitVendaComp = MARCADO Then
                ColItensNF(ColItensNF.Count).iPossuiGrade = MARCADO
            End If

            If ColItensNF(ColItensNF.Count).iPossuiGrade = MARCADO Then
                Set ColItensNF(ColItensNF.Count).colItensRomaneioGrade = colItensRomaneio
            End If
            
        End If

    Next
    'Chama a tela de localização de produtos
    Call Chama_Tela_Modal("LocalizacaoProduto", ColItensNF, Codigo_Extrai(TipoNFiscal.Text))
    'Limpa o grid de Alocações
    Call Grid_Limpa(objGridAlocacoes)
    Call gobjRastreamento.Limpa_Tela_Rastreamento
    
    'para cada item
    For Each objItemNF In ColItensNF
        
        objProduto.sCodigo = objItemNF.sProduto
        'Lê o produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 46885
        If lErro <> SUCESSO Then gError 46889
    
        'Se as UM de estoque e de venda do itemNF não estiverem preenchidas
        If Len(Trim(objItemNF.sUMEstoque)) = 0 Or Len(Trim(objItemNF.sUnidadeMed)) = 0 Then

            'Preenche as UMs do ItemNF
            objItemNF.sUMEstoque = objProduto.sSiglaUMEstoque
            objItemNF.sUnidadeMed = objProduto.sSiglaUMVenda
        End If
        
        If objProduto.iKitVendaComp = MARCADO Then
            objItemNF.iPossuiGrade = DESMARCADO
        End If
        
        'Formata o Produto
        lErro = CF("Produto_Formata", objCTNFiscalPedido.GridItens.TextMatrix(objItemNF.iItem, objCTNFiscalPedido.iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 46886
        'Se o produto foi substituído
        If sProduto <> objItemNF.sProduto Then
            'Avisa que o produto não pode ser substituido
            lErro = Rotina_Aviso(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_SUBSTITUIDO")
            'Limpa as alocações feitas
            Set objItemNF.colAlocacoes = New ColAlocacoesItemNF
        End If
        
    Next
    
    'Passa os itens com suas alocações para o objNFiscal
    Set objNFiscal.ColItensNF = ColItensNF
    
    'Preenche o grid de Alocações
    Call CF2(objCTNFiscalPedido, "NFP_Preenche_GridAlocacoes", objNFiscal)
    
    For iIndice = 1 To ColItensNF.Count
    
        Call Transfere_Dados_ItensRomaneio(objCTNFiscalPedido.gobjNFiscal.ColItensNF(ColItensNF(iIndice).iItem).colItensRomaneioGrade, objNFiscal.ColItensNF(iIndice).colItensRomaneioGrade)
    Next
    
    Exit Sub

Erro_BotaoLocalizacao_Click:

    Select Case gErr

        Case 26777
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, sProduto)
    
        Case 43699
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_ITENS_NF", gErr)

        Case 46885, 46886, 46888, 86306, 86378, 177582
        
        Case 46889
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 86307
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_PAI_GRADE_SEM_FILHOS", gErr, sProduto)
        
        Case 86376
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case 86379
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objItemRomaneio.sProduto)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157649)

    End Select

    Exit Sub

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iItemAtual As Integer
Dim iIndice As Integer
Dim dValorProdutos As Double
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
On Error GoTo Erro_GridItens_KeyDown

    'Armazena o número de linhas existentes e o item atual selecionado
    iLinhasExistentesAnterior = objCTNFiscalPedido.objGridItens.iLinhasExistentes
    iItemAtual = objCTNFiscalPedido.GridItens.Row

    'Se a tecla for a delete e o Pedido só aceitar faturamento integral
    'não permite a exclusão da linha
    If KeyCode = vbKeyDelete And objCTNFiscalPedido.giFaturaIntegral <> 0 Then gError 46633
        
    'Se a tecla for a delete e o Grid estiver vazio, Sai da Sub
    If KeyCode = vbKeyDelete And objCTNFiscalPedido.objGridItens.iLinhasExistentes <= 0 Then Exit Sub 'Inserido por Leo em 16/01/02
        
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)

    'Se algum ítem foi excluído
    If objCTNFiscalPedido.objGridItens.iLinhasExistentes < iLinhasExistentesAnterior Then
        
        '************ grade ************
        GridItens.TextMatrix(GridItens.Row, 0) = GridItens.Row
        objCTNFiscalPedido.gobjNFiscal.ColItensNF.Remove GridItens.Row
        
        Call objCTNFiscalPedido.gobjTelaProjetoInfo.Remove_Linha(iItemAtual)
        
        For iLinhasExistentesAnterior = 1 To objGridItens.iLinhasExistentes
            If objCTNFiscalPedido.gobjNFiscal.ColItensNF(iLinhasExistentesAnterior).iPossuiGrade = MARCADO Then
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = "# " & iLinhasExistentesAnterior
            Else
                GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
            End If
        Next
        
        GridItens.TextMatrix(iLinhasExistentesAnterior, 0) = iLinhasExistentesAnterior
        '*******************************
        
        If objCTNFiscalPedido.objGridAlocacoes.iLinhasExistentes > 0 Then
            
            'Remove as alocações do Pedido
            Call objCTNFiscalPedido.Remove_Alocacoes(iItemAtual)
                                    
            'Atualiza o número do item nas outras alocações
            For iIndice = 1 To objCTNFiscalPedido.objGridAlocacoes.iLinhasExistentes
                
                If Len(Trim(objCTNFiscalPedido.GridAlocacao.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_Item_Col))) > 0 Then If StrParaInt(objCTNFiscalPedido.GridAlocacao.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_Item_Col)) > iItemAtual Then objCTNFiscalPedido.GridAlocacao.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_Item_Col) = StrParaInt(objCTNFiscalPedido.GridAlocacao.TextMatrix(iIndice, objCTNFiscalPedido.iGrid_Item_Col)) - 1
            
            Next
        
        End If
        
        objCTNFiscalPedido.gcolItemPedido.Remove iItemAtual
        objCTNFiscalPedido.gcolItemQuant.Remove iItemAtual
        objCTNFiscalPedido.gcolItemUM.Remove iItemAtual
        
        If objCTNFiscalPedido.objGridItens.iLinhasExistentes <> 0 Then
            Call objCTNFiscalPedido.PrecoTotal_Calcula(objCTNFiscalPedido.objGridItens.iLinhasExistentes, False)
        Else
            If StrParaDbl(objCTNFiscalPedido.ValorDesconto.Text) <> 0 Then
                'Avisa ao usuário
                vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_VALOR_DESCONTO_MAIOR_PRODUTOS", objCTNFiscalPedido.ValorDesconto.Text, 0)
                
                'Limpa o valor de desconto
                objCTNFiscalPedido.gdDesconto = 0
                objCTNFiscalPedido.ValorDesconto.Text = ""
                objCTNFiscalPedido.iValorDescontoAlterado = 0
                
                Call objCTNFiscalPedido.gobjTribTab.ValorDesconto_Validate(bSGECancelDummy, 0)

            End If
        End If
        
        'Recalcula o valor dos produtos
        Call objCTNFiscalPedido.SubTotal_Calcula(objCTNFiscalPedido.objGridItens, objCTNFiscalPedido.iGrid_PrecoTotal_Col, dValorProdutos)
        'coloca na tela
        objCTNFiscalPedido.ValorProdutos.Caption = Format(dValorProdutos - StrParaDbl(objCTNFiscalPedido.ValorDesconto.Text), "Standard")

        'Recalcula o valor dos produtos
        Call objCTNFiscalPedido.SubTotal_Calcula(objCTNFiscalPedido.objGridItens, objCTNFiscalPedido.iGrid_PrecoTotalB_Col, dValorProdutos)
        'coloca na tela
        objCTNFiscalPedido.ValorProdutos2.Caption = Format(dValorProdutos, "Standard")

''*** incluidos p/tratamento de tributacao *******************************
        'Ataliza a tributação
        lErro = objCTNFiscalPedido.gobjTribTab.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 46702
'*** fim tributacao

        lErro = CF2(objCTNFiscalPedido, "NFFP_Calcula_PesoKg_Total")
        If lErro <> SUCESSO Then gError 177109

        'Recalcula o valor total
        Call ValorTotal_Calcula
    
        'Recalcula o peso bruto e o peso líquido
        lErro = objCTNFiscalPedido.Pesos_Recalcula()
        If lErro <> SUCESSO Then gError 81514
    
        'Rastreamento
        lErro = objCTNFiscalPedido.gobjRastreamento.Exclusao_Item_Grid(iItemAtual)
        If lErro <> SUCESSO Then gError 83379
    
        For iIndice = 1 To objCTNFiscalPedido.gobjNFiscal.ColItensNF.Count
            objCTNFiscalPedido.gobjNFiscal.ColItensNF(iIndice).iItem = iIndice 'Incluído para ao alterar a grade ou a reserva não fazer do item errado
        Next
        
    End If

    Exit Sub

Erro_GridItens_KeyDown:

    Select Case gErr
    
        Case 46702, 81514, 83379, 177109
        
        Case 46633
            Call Rotina_Erro(vbOKOnly, "ERRO_PEDIDO_VENDA_FATURA_INTEGRAL", gErr)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157650)
            
    End Select

    Exit Sub

End Sub

Public Sub BotaoGrade_Click()
'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim lErro  As Long
Dim objRomaneioGrade As ClassRomaneioGrade
Dim objItemNF As ClassItemNF

On Error GoTo Erro_BotaoGrade_Click

    If GridItens.Row > 0 And GridItens.Row <= objGridItens.iLinhasExistentes Then
    
        Set objItemNF = objCTNFiscalPedido.gobjNFiscal.ColItensNF(GridItens.Row)
        
        If objItemNF.iPossuiGrade = MARCADO Then
            
            Set objRomaneioGrade = New ClassRomaneioGrade
            
            objRomaneioGrade.sNomeTela = Me.Name
            Set objRomaneioGrade.objObjetoTela = objItemNF
                        
            Call Chama_Tela_Modal("RomaneioGrade", objRomaneioGrade)
            If lErro <> SUCESSO Then gError 86360
        
            Call objCTNFiscalPedido.Atualiza_Grid_Itens(objItemNF)
            
            Call objCTNFiscalPedido.Atualiza_Grid_Alocacao(objItemNF)
        
            If Len(Trim(objCTNFiscalPedido.GridItens.TextMatrix(objCTNFiscalPedido.GridItens.Row, objCTNFiscalPedido.iGrid_ValorUnitario_Col))) > 0 Then
                Call objCTNFiscalPedido.PrecoTotal_Calcula(GridItens.Row)
                
                lErro = objCTNFiscalPedido.gobjTribTab.Alteracao_Item_Grid(GridItens.Row)
                If lErro <> SUCESSO Then gError 86360
                
                lErro = objCTNFiscalPedido.ValorTotal_Calcula()
                If lErro <> SUCESSO Then gError 51037
            End If
        
        End If
    
    End If
    
    Exit Sub

Erro_BotaoGrade_Click:

    Select Case gErr
    
        Case 51037, 86360
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157652)
            
    End Select
    
    Exit Sub

End Sub

Function Move_ItensGrade_Tela(colItensRomaneio As Collection, colItensRomaneioTela As Collection, Optional bTrazPedido As Boolean = False) As Long

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItem As ClassReservaItem
Dim objReservaItemTela As ClassReservaItem
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim lErro As Long, iFilialEmpresa As Integer

On Error GoTo Erro_Move_ItensGrade_Tela

    'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
    For Each objItemRomaneioGradeTela In colItensRomaneioTela
                    
        Set objItemRomaneioGrade = New ClassItemRomaneioGrade
            
        objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto
        objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
        objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
        objItemRomaneioGrade.dQuantPV = objItemRomaneioGradeTela.dQuantidade - objItemRomaneioGradeTela.dQuantCancelada - objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
        objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
        objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade - objItemRomaneioGradeTela.dQuantCancelada - objItemRomaneioGradeTela.dQuantFaturada
        objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
        objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
        If bTrazPedido Then
            objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntDoc
        Else
            objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
            objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntItemPV
        End If
                            
        If objItemRomaneioGrade.dQuantidade > 0 Then
                               
            colItensRomaneio.Add objItemRomaneioGrade
        
            'Transfere as informações de Localização
            Set objItemRomaneioGrade.colLocalizacao = New Collection
                
            For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
                
                objAlmoxarifado.iCodigo = objReservaItemTela.iAlmoxarifado
                            
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> 25056 And lErro <> SUCESSO Then gError 94331
                If lErro = 25056 Then gError 94332
                    
                iFilialEmpresa = giFilialEmpresa
                lErro = CF("FilialEmpresaAlmox_Customiza", iFilialEmpresa)
                If lErro <> SUCESSO Then gError 94331
                    
                If objAlmoxarifado.iFilialEmpresa = iFilialEmpresa Then
                
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                    objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                    objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                    objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                    objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                    objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                    objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                    
                    objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                End If
            Next
        End If
    
    Next
    
    Move_ItensGrade_Tela = SUCESSO
    
    Exit Function
    
Erro_Move_ItensGrade_Tela:

    Move_ItensGrade_Tela = gErr

    Select Case gErr
            
        Case 94331

        Case 94332
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objReservaItemTela.iAlmoxarifado)
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157653)

    End Select
    
    Exit Function

End Function

Function Transfere_Dados_ItensRomaneio(colItensRomaneio As Collection, colItensRomaneioTela As Collection) As Long

'************** FUNÇÃO CRIADA PARA TRATAR GRADE **********************

Dim objItemRomaneioGrade As ClassItemRomaneioGrade
Dim objItemRomaneioGradeTela As ClassItemRomaneioGrade
Dim objReservaItemTela As ClassReservaItem
Dim objReservaItem As ClassReservaItem

    'Para cada Item de Romaneio existente do BD (Produtos Filhos do Produto passado)
    For Each objItemRomaneioGrade In colItensRomaneio
        'Para cada Item de Romaneio vindo da tela ( Aqueles que já tem quantidade)
        For Each objItemRomaneioGradeTela In colItensRomaneioTela
            'Se encontrou o Item
            If objItemRomaneioGrade.sProduto = objItemRomaneioGradeTela.sProduto Then
                'Transfere as informações vindas da tela chamadora para essa tela
                objItemRomaneioGrade.dQuantOP = objItemRomaneioGradeTela.dQuantOP
                objItemRomaneioGrade.dQuantSC = objItemRomaneioGradeTela.dQuantSC
                objItemRomaneioGrade.sDescricao = objItemRomaneioGradeTela.sDescricao
                objItemRomaneioGrade.dQuantAFaturar = objItemRomaneioGradeTela.dQuantAFaturar
                objItemRomaneioGrade.dQuantFaturada = objItemRomaneioGradeTela.dQuantFaturada
                objItemRomaneioGrade.dQuantidade = objItemRomaneioGradeTela.dQuantidade
                objItemRomaneioGrade.dQuantReservada = objItemRomaneioGradeTela.dQuantReservada
                objItemRomaneioGrade.sUMEstoque = objItemRomaneioGradeTela.sUMEstoque
                objItemRomaneioGrade.dQuantCancelada = objItemRomaneioGradeTela.dQuantCancelada
                objItemRomaneioGrade.dQuantPV = objItemRomaneioGradeTela.dQuantPV
                objItemRomaneioGrade.lNumIntItemPV = objItemRomaneioGradeTela.lNumIntItemPV
                objItemRomaneioGrade.lNumIntDoc = objItemRomaneioGradeTela.lNumIntDoc
                
                'Transfere as informações de Localização
                Set objItemRomaneioGrade.colLocalizacao = New Collection
                    
                For Each objReservaItemTela In objItemRomaneioGradeTela.colLocalizacao
                    
                    Set objReservaItem = New ClassReservaItem
                    
                    objReservaItem.dQuantidade = objReservaItemTela.dQuantidade
                    objReservaItem.dtDataValidade = objReservaItemTela.dtDataValidade
                    objReservaItem.iAlmoxarifado = objReservaItemTela.iAlmoxarifado
                    objReservaItem.iFilialEmpresa = objReservaItemTela.iFilialEmpresa
                    objReservaItem.lNumIntDoc = objReservaItemTela.lNumIntDoc
                    objReservaItem.sAlmoxarifado = objReservaItemTela.sAlmoxarifado
                    objReservaItem.sResponsavel = objReservaItemTela.sResponsavel
                    
                    objItemRomaneioGrade.colLocalizacao.Add objReservaItem
                    
                Next
                            
            End If
        
        Next
    Next

    Exit Function

End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long, sID As String
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_Anotacao_Extrai

    'Lê os dados da Tela NFiscal
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 130404

    If objNFiscal.lNumNotaFiscal <> 0 And objNFiscal.sSerie <> "" And objNFiscal.dtDataEmissao <> DATA_NULA Then
    
        lErro = CF("NFiscal_Le_NumeroSerie", objNFiscal)
        If lErro <> SUCESSO And lErro <> 43676 Then gError 130405
        
        If lErro = SUCESSO Then sID = objNFiscal.lNumIntDoc
        
    End If
    
    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_NFISCAL
    If sID <> "" Then
        objAnotacao.sID = sID
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case 130404, 130405
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157654)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157655)
     
    End Select
     
    Exit Sub

End Sub

Private Function Seleciona_Tipo_Padrao() As Long
'seleciona na combo de tipos de nfiscal o tipo padrao desta tela

Dim lErro As Long
Dim colTipoDocInfo As New colTipoDocInfo
Dim objTipoDocInfo As New ClassTipoDocInfo

On Error GoTo Erro_Seleciona_Tipo_Padrao

    Set colTipoDocInfo = gobjCRFAT.colTiposDocInfo

    'Carrega na combo só os Tipos ligados essa tela
    For Each objTipoDocInfo In colTipoDocInfo
        If objTipoDocInfo.sNomeTelaNFiscal = Me.Name Then
            'se for o tipo padrao, seleciona-o
            If objTipoDocInfo.iPadrao = Padrao Then
                
                TipoNFiscal.Text = objTipoDocInfo.iCodigo
                Call TipoNFiscal_Validate(bSGECancelDummy)
                
            End If
        End If
    Next

    Seleciona_Tipo_Padrao = SUCESSO

    Exit Function

Erro_Seleciona_Tipo_Padrao:

    Seleciona_Tipo_Padrao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157617)
    
    End Select

    Exit Function

End Function

Public Sub BotaoSerie_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long

On Error GoTo Erro_BotaoSerie_Click
    
    Call gobjRastreamento.BotaoSerie_Click
                    
    Exit Sub

Erro_BotaoSerie_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 141800)
    
    End Select
    
    Exit Sub

End Sub

Public Function NFRP_Saida_Celula_GridItens(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    NFRP_Saida_Celula_GridItens = SUCESSO
End Function

Public Sub BotaoKitVenda_Click()
    Call objCTNFiscalPedido.BotaoKitVenda_Click
End Sub

Public Function NFRP_Preeche_GridItens(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal) As Long
    NFRP_Preeche_GridItens = SUCESSO
End Function

Public Function NFRP_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer, ByVal sProdutoFormatado As String, ByVal iProdutoPreenchido As Integer) As Long
    NFRP_Rotina_Grid_Enable = SUCESSO
End Function

Public Function NFRP_Move_GridItem_Memoria(ByVal objCT As Object, ByVal objNFiscal As ClassNFiscal, ByVal objItemNF As ClassItemNF, ByVal iLinha As Integer) As Long
    NFRP_Move_GridItem_Memoria = SUCESSO
End Function

Public Function BotaoProjetos_Click() As Long
        Call objCTNFiscalPedido.gobjTelaProjetoInfo.BotaoProjetos_Click
End Function

Public Function LabelProjeto_Click() As Long
    Call objCTNFiscalPedido.gobjTelaProjetoInfo.LabelProjeto_Click
End Function

Public Function Projeto_GotFocus() As Long
    Call MaskEdBox_TrataGotFocus(objCTNFiscalPedido.Projeto, iAlterado)
End Function

Public Function Projeto_Change() As Long
    iAlterado = REGISTRO_ALTERADO
End Function

Public Function Projeto_Validate(Cancel As Boolean) As Long
    Call objCTNFiscalPedido.gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Function

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long
    Obter_ItensPRJCR = objCTNFiscalPedido.Obter_ItensPRJCR(colItensPRJ)
End Function

Public Function NFRP_Form_Load2(ByVal objCT As Object) As Long
    NFRP_Form_Load2 = SUCESSO
End Function

Private Function Carrega_TipoFrete() As Long
'Carrega versaoNFE

Dim lErro As Long
Dim objTipoFrete As ClassTipoFrete
Dim iIndice As Integer
Dim colTipoFrete As Collection

On Error GoTo Erro_Carrega_TipoFrete
    
    'Le Código e Nome de FilialEmpresa
    lErro = CF("TipoFrete_Le", colTipoFrete)
    If lErro <> SUCESSO Then gError 207444

    For Each objTipoFrete In colTipoFrete

        'coloca na combo
        objCTNFiscalPedido.TipoFrete.AddItem objTipoFrete.sDescricao
        objCTNFiscalPedido.TipoFrete.ItemData(objCTNFiscalPedido.TipoFrete.NewIndex) = objTipoFrete.iCodigo

    Next

    objCTNFiscalPedido.TipoFrete.ListIndex = 0

    Carrega_TipoFrete = SUCESSO

    Exit Function

Erro_Carrega_TipoFrete:

    Carrega_TipoFrete = gErr

    Select Case gErr

        'Erro já tratado
        Case 207444

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 207445)

    End Select

    Exit Function

End Function

Public Sub FilialFornEntTerc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialFornEntTerc_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialFornEntTerc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_FilialFornEntTerc_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(FilialFornEntTerc.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If FilialFornEntTerc.Text = FilialFornEntTerc.List(FilialFornEntTerc.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(FilialFornEntTerc, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError ERRO_SEM_MENSAGEM

    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(FornEntTerc.Text)) = 0 Then gError 209993

        sFornecedor = FornEntTerc.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError ERRO_SEM_MENSAGEM

        If lErro = 18272 Then gError 209994

        'coloca na tela
        FilialFornEntTerc.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 209995

    Exit Sub

Erro_FilialFornEntTerc_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case 209993
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 209994
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, FornEntTerc.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 209995
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, FilialFornEntTerc.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209996)

    End Select

    Exit Sub

End Sub

Public Sub FornEntTerc_Change()
    iAlterado = REGISTRO_ALTERADO
    iFornEntTercAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FornEntTerc_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornEntTerc As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_FornEntTerc_Validate
   
    If iFornEntTercAlterado = 1 Then
        If Len(Trim(FornEntTerc.Text)) > 0 Then
    
            'Tenta ler o FornEntTerc (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(FornEntTerc, objFornEntTerc, iCodFilial)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornEntTerc, colCodigoNome)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", FilialFornEntTerc, colCodigoNome)
    
            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", FilialFornEntTerc, iCodFilial)
    
    
        ElseIf Len(Trim(FornEntTerc.Text)) = 0 Then
            FilialFornEntTerc.Clear
        End If
    
        iFornEntTercAlterado = 0
    
    End If
    
    Exit Sub

Erro_FornEntTerc_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209997)

    End Select

    Exit Sub

End Sub

Public Sub FornEntTercLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'recolhe o Nome Reduzido da tela
    objFornecedor.sNomeReduzido = FornEntTerc.Text

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornEntTerc)

    Exit Sub

End Sub

Private Sub objEventoFornEntTerc_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    FornEntTerc.Text = objFornecedor.lCodigo
    Call FornEntTerc_Validate(bSGECancelDummy)

    Me.Show

End Sub

Public Function NFRP_Limpa_Tela(ByVal objCT As Object) As Long
    NFRP_Limpa_Tela = SUCESSO
End Function

Function NFRP_CargaPosFormLoad(ByVal objCT As Object) As Long
    NFRP_CargaPosFormLoad = SUCESSO
End Function

Public Sub BotaoInfoAdicItem_Click()
    Call objCTNFiscalPedido.BotaoInfoAdicItem_Click
End Sub

Private Function Move_Dados_NFOriginal_Memoria(ByVal objNFiscal As ClassNFiscal, ByVal iTipoNF As Integer) As Long
'Move os dados da NotaFiscalOriginal para a memória

Dim lErro As Long
Dim objNFOriginal As New ClassNFiscal

On Error GoTo Erro_Move_Dados_NFOriginal_Memoria

    'Verifica se a Série e o Número da NF original estão preenchidos
    If Len(Trim(NFiscalOriginal.ClipText)) = 0 Then gError 35946
    If Len(Trim(SerieNFiscalOriginal.Text)) = 0 Then gError 35947
    
    objNFOriginal.lCliente = objNFiscal.lCliente
    objNFOriginal.iFilialCli = objNFiscal.iFilialCli
    objNFOriginal.sSerie = SerieNFiscalOriginal.Text
    objNFOriginal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)

    'Lê a Nota Fiscal Original
    lErro = CF("NFiscal_Le_Num_Cli2", objNFOriginal)
    If lErro <> SUCESSO And lErro <> 92090 Then gError ERRO_SEM_MENSAGEM
    If lErro = 92090 Then gError 35948 'Se não encontrar --> Erro
        
    'Guarda em objNfiscal o Numero Interno da NF Original
    objNFiscal.lNumIntNotaOriginal = objNFOriginal.lNumIntDoc
    
    Move_Dados_NFOriginal_Memoria = SUCESSO

    Exit Function

Erro_Move_Dados_NFOriginal_Memoria:

    Move_Dados_NFOriginal_Memoria = gErr

    Select Case gErr

        Case 35946
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NFISCAL_ORIGINAL_NAO_PREENCHIDO", gErr)

        Case 35947
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NFISCAL_ORIGINAL_NAO_PREENCHIDA", gErr)

        Case 35948
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_ORIGINAL_NAO_ASSOCIADA_CLIENTE", gErr, Cliente.Caption, SerieNFiscalOriginal.Text, NFiscalOriginal.Text)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156257)

    End Select

    Exit Function

End Function

Public Sub NFiscalOriginalLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim lErro As Long, sFiltro As String

On Error GoTo Erro_NFiscalOriginalLabel_Click

    'Se o cliente não está preenchido => erro
    If Len(Trim(Cliente.Caption)) = 0 Then gError 101090
    
    sFiltro = "Complementar = 0 AND NomeCliente = ?"
    
    If Codigo_Extrai(TipoNFiscal.Text) = DOCINFO_NFISSRPVEF Then
        sFiltro = sFiltro & " AND TipoNFiscal IN (187,188)"
    End If
        
    'Recolhe os dados da Nota Fiscal
    lErro = Move_Tela_Memoria(objNFiscal)
    If lErro <> SUCESSO Then gError 31437

    'Guarda a Serie e o Número da Nota Fiscal Original da Tela
    objNFiscal.sSerie = SerieNFiscalOriginal.Text
    If Len(Trim(NFiscalOriginal.ClipText)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscalOriginal.Text)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If
    
    'Faz o filtro por cliente
    colSelecao.Add Cliente.Caption
    colSelecao.Add DOCINFO_CLIENTE

    'Chama a Tela NFiscalNormalLista
    'Call Chama_Tela("NFiscalNormalLista", colSelecao, objNFiscal, objEventoNFiscalOriginal)
    Call Chama_Tela("NFInternaDestinatarioLista", colSelecao, objNFiscal, objCTNFiscalRem.objEventoNFiscalOriginal, sFiltro)

    Exit Sub

Erro_NFiscalOriginalLabel_Click:

    Select Case gErr

        Case 31437
        
        Case 101090
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 101091
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156221)

    End Select

    Exit Sub

End Sub
