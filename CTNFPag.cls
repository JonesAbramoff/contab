VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object 'Inserido por wagner

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Dim iFrameAtual As Integer
Dim iFornecedorAlterado As Integer
Private iEmissaoAlterada As Integer
Private iValorTotalAlterado As Integer
Private iValorIRRFAlterado As Integer
Private iValorINSSAlterado As Integer
Private iPISRetidoAlterado As Integer
Private iISSRetidoAlterado As Integer
Private iCOFINSRetidoAlterado As Integer
Private iCSLLRetidoAlterado As Integer

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

'Início contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
Private Const NUMERO1 As String = "Numero_Nota_Fiscal"
Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALORTOTAL1 As String = "Valor_da_Nota"
Private Const VALOR_ICMS As String = "Valor_ICMS"
Private Const VALOR_ICMS_SUBST As String = "Valor_ICMS_Subst"
Private Const CREDITA_ICMS As String = "Credita_ICMS"
Private Const VALOR_IR As String = "Valor_IRRF"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_IPI As String = "Valor_IPI"
Private Const CREDITA_IPI As String = "Credita_IPI"
Private Const VALOR_OUTRASDESP As String = "Valor_OutrasDesp"
Private Const CONTA_DESP_ESTOQUE As String = "Conta_Desp_Estoque"
Private Const CONTA_DESP_EST_FORN As String = "Conta_Desp_Est_Forn"

Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"
'Fim da Contabilidade

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2

'--- inicio dos properties get dos controles da tela
Public Property Get ValorINSS() As Object
     Set ValorINSS = objUserControl.Controls("ValorINSS")
End Property

Public Property Get PISRetido() As Object
     Set PISRetido = objUserControl.Controls("PISRetido")
End Property

Public Property Get ISSRetido() As Object
     Set ISSRetido = objUserControl.Controls("ISSRetido")
End Property

Public Property Get COFINSRetido() As Object
     Set COFINSRetido = objUserControl.Controls("COFINSRetido")
End Property

Public Property Get CSLLRetido() As Object
     Set CSLLRetido = objUserControl.Controls("CSLLRetido")
End Property

Public Property Get INSSRetido() As Object
     Set INSSRetido = objUserControl.Controls("INSSRetido")
End Property
Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get CreditoICMS() As Object
     Set CreditoICMS = objUserControl.Controls("CreditoICMS")
End Property

Public Property Get ValorICMS() As Object
     Set ValorICMS = objUserControl.Controls("ValorICMS")
End Property

Public Property Get ValorICMSSubst() As Object
     Set ValorICMSSubst = objUserControl.Controls("ValorICMSSubst")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get CreditoIPI() As Object
     Set CreditoIPI = objUserControl.Controls("CreditoIPI")
End Property

Public Property Get ValorIPI() As Object
     Set ValorIPI = objUserControl.Controls("ValorIPI")
End Property

Public Property Get ValorProdutos() As Object
     Set ValorProdutos = objUserControl.Controls("ValorProdutos")
End Property

Public Property Get ValorFrete() As Object
     Set ValorFrete = objUserControl.Controls("ValorFrete")
End Property

Public Property Get OutrasDespesas() As Object
     Set OutrasDespesas = objUserControl.Controls("OutrasDespesas")
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
End Property

Public Property Get ValorIRRF() As Object
     Set ValorIRRF = objUserControl.Controls("ValorIRRF")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get NumNotaFiscal() As Object
     Set NumNotaFiscal = objUserControl.Controls("NumNotaFiscal")
End Property

Public Property Get UpDownVencimento() As Object
     Set UpDownVencimento = objUserControl.Controls("UpDownVencimento")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get CTBBotaoModeloPadrao() As Object
     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
End Property

Public Property Get CTBBotaoLimparGrid() As Object
     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
End Property

Public Property Get CTBModelo() As Object
     Set CTBModelo = objUserControl.Controls("CTBModelo")
End Property

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
End Property

Public Property Get CTBFrame7() As Object
     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
End Property

Public Property Get CTBCclLabel() As Object
     Set CTBCclLabel = objUserControl.Controls("CTBCclLabel")
End Property

Public Property Get CTBLabel7() As Object
     Set CTBLabel7 = objUserControl.Controls("CTBLabel7")
End Property

Public Property Get CTBContaDescricao() As Object
     Set CTBContaDescricao = objUserControl.Controls("CTBContaDescricao")
End Property

Public Property Get CTBCclDescricao() As Object
     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
End Property

Public Property Get CTBLancAutomatico() As Object
     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
End Property

Public Property Get CTBConta() As Object
     Set CTBConta = objUserControl.Controls("CTBConta")
End Property

Public Property Get CTBDebito() As Object
     Set CTBDebito = objUserControl.Controls("CTBDebito")
End Property

Public Property Get CTBCredito() As Object
     Set CTBCredito = objUserControl.Controls("CTBCredito")
End Property

Public Property Get CTBCcl() As Object
     Set CTBCcl = objUserControl.Controls("CTBCcl")
End Property

Public Property Get CTBUpDown() As Object
     Set CTBUpDown = objUserControl.Controls("CTBUpDown")
End Property

Public Property Get CTBDataContabil() As Object
     Set CTBDataContabil = objUserControl.Controls("CTBDataContabil")
End Property

Public Property Get CTBLote() As Object
     Set CTBLote = objUserControl.Controls("CTBLote")
End Property

Public Property Get CTBDocumento() As Object
     Set CTBDocumento = objUserControl.Controls("CTBDocumento")
End Property

Public Property Get CTBGridContabil() As Object
     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
End Property

Public Property Get CTBTvwCcls() As Object
     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
End Property

Public Property Get CTBTvwContas() As Object
     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
End Property

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
End Property

Public Property Get CTBLabel21() As Object
     Set CTBLabel21 = objUserControl.Controls("CTBLabel21")
End Property

Public Property Get CTBOrigem() As Object
     Set CTBOrigem = objUserControl.Controls("CTBOrigem")
End Property

Public Property Get CTBLabel14() As Object
     Set CTBLabel14 = objUserControl.Controls("CTBLabel14")
End Property

Public Property Get CTBPeriodo() As Object
     Set CTBPeriodo = objUserControl.Controls("CTBPeriodo")
End Property

Public Property Get CTBExercicio() As Object
     Set CTBExercicio = objUserControl.Controls("CTBExercicio")
End Property

Public Property Get CTBLabel13() As Object
     Set CTBLabel13 = objUserControl.Controls("CTBLabel13")
End Property

Public Property Get CTBLabel5() As Object
     Set CTBLabel5 = objUserControl.Controls("CTBLabel5")
End Property

Public Property Get CTBLabelHistoricos() As Object
     Set CTBLabelHistoricos = objUserControl.Controls("CTBLabelHistoricos")
End Property

Public Property Get CTBLabelContas() As Object
     Set CTBLabelContas = objUserControl.Controls("CTBLabelContas")
End Property

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
End Property

Public Property Get CTBLabelTotais() As Object
     Set CTBLabelTotais = objUserControl.Controls("CTBLabelTotais")
End Property

Public Property Get CTBTotalDebito() As Object
     Set CTBTotalDebito = objUserControl.Controls("CTBTotalDebito")
End Property

Public Property Get CTBTotalCredito() As Object
     Set CTBTotalCredito = objUserControl.Controls("CTBTotalCredito")
End Property

Public Property Get CTBLabel8() As Object
     Set CTBLabel8 = objUserControl.Controls("CTBLabel8")
End Property

Public Property Get CTBLabelDoc() As Object
     Set CTBLabelDoc = objUserControl.Controls("CTBLabelDoc")
End Property

Public Property Get CTBLabelLote() As Object
     Set CTBLabelLote = objUserControl.Controls("CTBLabelLote")
End Property

Public Property Get BotaoDocOriginal() As Object
     Set BotaoDocOriginal = objUserControl.Controls("BotaoDocOriginal")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property
'--- fim dos properties get dos controles da tela

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag
Dim sTela As String
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoDocOriginal_Click

    'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número, DataEmissao)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 79321

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79322

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 79323

    objNFsPag.lFornecedor = objFornecedor.lCodigo
    objNFsPag.iFilial = Codigo_Extrai(Filial.Text)
    objNFsPag.lNumNotaFiscal = CLng(NumNotaFiscal.Text)
    objNFsPag.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objNFsPag.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objNFsPag.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = NFPag_Le_Todos(objNFsPag)
    If lErro <> SUCESSO Then gError 79324

    'Le o Nome da Tela que originou este Título
    lErro = CF("Titulo_Le_DocumentoOriginal", objNFsPag.lNumIntDoc, CPR_NF_PAGAR, objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 58942 Then gError 79325

    If lErro = SUCESSO Then
        'Chama a Tela Estoque
        Call Chama_Tela(sTela, objNFiscal)
    Else
        gError 79326
    End If

    Exit Sub

Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 79321, 79322, 79324, 79325 'Tratado na Rotina chamada

        Case 79323
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case 79326
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFPAG_SEM_DOCORIGINAL", gErr, objNFsPag.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157656)

    End Select

    Exit Sub


End Sub

Private Function Critica_CamposPreenchidos() As Long
'Verifica se os campos chaves foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposPreenchidos

    'Se Fornecedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79315

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then gError 79316

    'Verifica se o Titulo está Preenchido
    If Len(Trim(NumNotaFiscal.Text)) = 0 Then gError 79317

    'Se a data de emissão do título não está preenchida, erro
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79314
        
    
    Critica_CamposPreenchidos = SUCESSO

    Exit Function

Erro_Critica_CamposPreenchidos:

    Critica_CamposPreenchidos = gErr

    Select Case gErr

        Case 79315
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 79316
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 79317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
        
        Case 79314
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157657)

    End Select

    Exit Function

End Function
Private Function NFPag_Le_Todos(objNFsPag As ClassNFsPag) As Long
'Lê em NFPag se não encontrar procura em NFPag Baixados

Dim lErro As Long

On Error GoTo Erro_NFPag_Le_Todos

    'Lê em NFPags
    lErro = CF("NFPag_Le_Numero", objNFsPag)
    If lErro <> SUCESSO And lErro <> 18338 Then gError 79318

    'Se não encontrou
    If lErro = 18338 Then

        'Procura em Títulos a Receber Baixados
        lErro = CF("NFPagBaixada_Le_Numero", objNFsPag)
        If lErro <> SUCESSO And lErro <> 18344 Then gError 79319

        'Se não encontrou nos dois ---> ERRO
        If lErro = 18344 Then gError 79320

    End If

    NFPag_Le_Todos = SUCESSO

    Exit Function

Erro_NFPag_Le_Todos:

    NFPag_Le_Todos = gErr

    Select Case gErr

        Case 79318, 79319 'Tratado na rotina chamada

        Case 79320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFPAG_NAO_CADASTRADA1", gErr, objNFsPag.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157658)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o fornecedor foi preenchido
    If Len(Trim(Fornecedor.ClipText)) = 0 Then Error 18319
    
    'Verifica se a filial foi digitada
    If Len(Trim(Filial.Text)) = 0 Then Error 18320
    
    'Verifica se o numero da nota fiscal foi informado
    If Len(Trim(NumNotaFiscal.ClipText)) = 0 Then Error 18321
    
    'Move os dados da tela para a memória
    lErro = Move_Tela_Memoria(objNFsPag)
    If lErro <> SUCESSO Then Error 18322
    
    'Busca a nota na tabela de notas fiscais à pagar
    lErro = CF("NFPag_Le_Numero", objNFsPag)
    If lErro <> SUCESSO And lErro <> 18338 Then Error 18339
    
    'Se não encontra
    If lErro = 18338 Then
    
        'Procura na tabela de notas baixadas
        lErro = CF("NFPagBaixada_Le_Numero", objNFsPag)
        If lErro <> SUCESSO And lErro <> 18344 Then Error 18345
    
        'A nota está baixada
        If lErro = AD_SQL_SUCESSO Then Error 18346
        
        'A nota não está cadastrada
        If lErro = 18344 Then Error 18347

    End If
    
    'Pede a confirmação de exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NFPAG", objNFsPag.lNumNotaFiscal)
    
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Faz a exclusão da Nota Fiscal (inclusive dados contábeis)
    lErro = CF("NFPag_Exclui", objNFsPag, objContabil)
    If lErro <> SUCESSO Then Error 18323
    
    'Limpa a tela de Notas Fiscais
    Call Limpa_Tela_NFPag
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 18319
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
        
        Case 18320
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
        
        Case 18321
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", Err)
        
        Case 18322, 18323, 18339, 18345
                    
        Case 18346
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_BAIXADA_EXCLUSAO", Err, objNFsPag.lNumNotaFiscal)
                    
        Case 18347
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFPAG_NAO_CADASTRADA1", objNFsPag.lNumNotaFiscal)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157659)
            
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    'Grava a Nota Fiscal
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 18282
    
    'Limpa a tela
    Call Limpa_Tela_NFPag
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 18282 'Tratado na Rotina chamada
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157660)
    
    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 18326

    'Limpa a tela
    Call Limpa_Tela_NFPag

    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 18326 'Tratado na Rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157661)
            
    End Select
            
    Exit Sub
    
End Sub

Public Sub CreditoICMS_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CreditoIPI_Click()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DataVencimento_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataVencimento, iAlterado)

End Sub

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO
    
End Sub


Public Sub Form_Load()

Dim lIndice As Long
Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Visible = False
    
    End If
    
    iFrameAtual = 1

    Set objEventoFornecedor = New AdmEvento
    Set objEventoNumero = New AdmEvento
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    '################################################
    'Inserido por Wagner
    lErro = CF2(Me, "NFPAG_Form_Load")
    If lErro <> SUCESSO Then gError 132231
    '################################################
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 39589
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 39589, 132231 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157662)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub DataEmissao_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 18217
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err
        
        Case 18217

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157663)

    End Select

    Exit Sub

End Sub

Public Sub DataVencimento_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVencimento_Validate

    'Verifica se a data de vencimento foi digitada
    If Len(Trim(DataVencimento.ClipText)) = 0 Then Exit Sub

    'Faz a critica da data digitada
    lErro = Data_Critica(DataVencimento.Text)
    If lErro <> SUCESSO Then Error 18221

    'Verifica se a data de emissao foi digitada
    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    Exit Sub

Erro_DataVencimento_Validate:

    Cancel = True


    Select Case Err

        Case 18221

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157664)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
    
End Sub

Public Sub Form_Unload(Cancel As Integer)
    
Dim lErro As Long

    Set objEventoFornecedor = Nothing
    Set objEventoNumero = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    Set gobjTelaProjetoInfo = Nothing

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String
Dim vbMsgRes As VbMsgBoxResult
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub
    
    'Verifica se é uma filial selecionada
    If Filial.ListIndex >= 0 Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18263
    
    'Se nao encontra o ítem com o código informado
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 18266

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo
        
        'Pesquisa se existe filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 18273

        If lErro = 18272 Then
        
            objFornecedor.sNomeReduzido = sFornecedor
            
            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 58632
            
            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            'Sugere cadastrar nova Filial
            Error 18265
        
        End If
        
        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome
        
    End If
    
    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 18267
    
    '#############################################
    'Inserido por Wagner
    lErro = CF2(Me, "NFCF_Fornecedor_Filial_Validate", objFilialFornecedor)
    If lErro <> SUCESSO Then gError 131965
    '#############################################
   
    Exit Sub
    
Erro_Filial_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 18263, 18273, 58632 'Tratados nas Rotinas chamadas
        
        Case 18265
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)
        
            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If
        
        Case 18266
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 18267
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)
            
        Case 131965 'Inserido por Wagner
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157665)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = 1

    Call Fornecedor_Preenche

End Sub

'Le os dados de TipoFornecedor ligado a objFornecedor. Coloca na Tela.
Private Function TipoFornecedor_Dados(objFornecedor As ClassFornecedor) As Long

Dim lErro As Long
Dim objTipoFornecedor As New ClassTipoFornecedor

On Error GoTo Erro_TipoFornecedor_Dados

    objTipoFornecedor.iCodigo = objFornecedor.iTipo
    
    'Lê TipoFornecedor a partir do código
    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then Error 6658
    If lErro <> SUCESSO Then Error 6903

'    'Coloca dados de TipoFornecedor na Tela
'    CreditoIPI.Value = objTipoFornecedor.iCreditoIPI
'    CreditoICMS.Value = objTipoFornecedor.iCreditoICMS

    TipoFornecedor_Dados = SUCESSO

    Exit Function

Erro_TipoFornecedor_Dados:

    TipoFornecedor_Dados = Err

    Select Case Err

        Case 6658 'Tratado na rotina chamada

        Case 6903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO", Err, objTipoFornecedor.iCodigo)
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157666)

    End Select

    Exit Function

End Function

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 6665

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 6699

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

            'CODIGO ESPECÍFICO
            'Verifica se chave de TipoFornecedor está preenchida
            If objFornecedor.iTipo > 0 Then

                'Coloca os dados do Tipo na tela
                lErro = TipoFornecedor_Dados(objFornecedor)
                If lErro <> SUCESSO Then gError 6662

            End If
            'FIM CODIGO ESPECÍFICO

            '########################################################
            'Inserido por Wagner
            lErro = CF2(Me, "NFCF_Fornecedor_Validate", objFornecedor.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131967
            '########################################################

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            Filial.Clear

            '#############################################
            'Inserido por Wagner
            Call CF2(Me, "NFCF_Limpa_Tela")
            '#############################################

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True
    
    Select Case gErr

        Case 6699, 6662, 6665, 131967 'Inserido por Wagner

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157667)

    End Select

    Exit Sub

End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche nomeReduzido com o fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text
    
    'Chama tela de Browse de Fornecedor
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39591

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39592

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 39591

        Case 39592
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub NumeroLabel_Click()

Dim objNFsPag As New ClassNFsPag
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click

    If Len(Trim(Fornecedor.Text)) = 0 Then Error 18210

    If Len(Trim(Filial.Text)) = 0 Then Error 18211
    
    lErro = Move_Tela_Memoria(objNFsPag)
    If lErro <> SUCESSO Then Error 18358

    colSelecao.Add objNFsPag.lFornecedor
    colSelecao.Add objNFsPag.iFilial
    
    Call Chama_Tela("NFPagLista", colSelecao, objNFsPag, objEventoNumero)
    
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 18210
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 18211
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 18358 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157668)

    End Select

    Exit Sub

End Sub

Public Sub NumNotaFiscal_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub NumNotaFiscal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumNotaFiscal, iAlterado)

End Sub

Public Sub NumNotaFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumNotaFiscal_Validate

    'Verifica se está preenchido
    If Len(Trim(NumNotaFiscal.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumNotaFiscal.ClipText)
    If lErro <> SUCESSO Then Error 18212
    
    Exit Sub

Erro_NumNotaFiscal_Validate:

    Cancel = True


    Select Case Err

        Case 18212

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157669)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1

    Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Call Fornecedor_Validate(Cancel)
   
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFsPag As ClassNFsPag

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objNFsPag = obj1
    
    'Traz a NotaFiscalPagar para Tela
    lErro = Traz_NFPag_Tela(objNFsPag)
    If lErro <> SUCESSO Then Error 18359
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoNumero_evSelecao:

    Select Case Err
    
        Case 18359 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157670)
            
    End Select

    Exit Sub

End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_NOTAS_FISCAIS_PAGAR_ID
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NOTAS_FISCAIS_PAGAR_CONTABILIZACAO
                        
        End Select

        
    End If

End Sub

Function Trata_Parametros(Optional objNFsPag As ClassNFsPag) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se alguma nota foi passada por parametro
    If Not (objNFsPag Is Nothing) Then
        
        'Movido para dentro da função Traz_NFPag_Tela por Wagner 21/02/2005
'        'Tenta ler a nota Fiscal passada por parametro
'        lErro = CF("NFPag_Le_Todos2", objNFsPag)
'        If lErro <> SUCESSO And lErro <> 18257 Then Error 18253
'
'        If lErro <> SUCESSO Then Error 18428

        'Traz a nota para a tela
        lErro = Traz_NFPag_Tela(objNFsPag)
        If lErro <> SUCESSO Then Error 18209
        
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 18209 ', 18253 'Tratado na rotina chamada

'        Case 18428
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFPAG_NAO_CADASTRADA", Err, objNFsPag.lNumIntDoc)
'            Call Limpa_Tela_NFPag
'            iAlterado = 0

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157671)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub OutrasDespesas_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub OutrasDespesas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OutrasDespesas_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(OutrasDespesas.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(OutrasDespesas.Text)
    If lErro <> SUCESSO Then Error 18280

    'Põe o valor formatado na tela
    OutrasDespesas.Text = Format(OutrasDespesas.Text, "Fixed")
    
    Exit Sub

Erro_OutrasDespesas_Validate:

    Cancel = True


    Select Case Err

        Case 18280

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157672)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick
    
    'Diminui a Data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 18218

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 18218

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157673)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a Data
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 18219

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 18219 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157674)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencimento_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_DownClick
    
    'Diminui Data
    lErro = Data_Up_Down_Click(DataVencimento, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 18223

    Exit Sub

Erro_UpDownVencimento_DownClick:

    Select Case Err

        Case 18223 'Tratado na Rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157675)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencimento_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownVencimento_UpClick
    
    'Aumenta Data
    lErro = Data_Up_Down_Click(DataVencimento, AUMENTA_DATA)
    If lErro Then Error 18224

    Exit Sub

Erro_UpDownVencimento_UpClick:

    Select Case Err

        Case 18224 'Tratado na Rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157676)

    End Select

    Exit Sub

End Sub

Public Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorFrete.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
    If lErro <> SUCESSO Then Error 18278
    
    'Põe o valor formatado na tela
    ValorFrete.Text = Format(ValorFrete.Text, "Fixed")

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True


    Select Case Err

        Case 18278

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157677)

    End Select

    Exit Sub

End Sub

Public Sub ValorICMS_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorICMS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorICMS_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMS.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorICMS.Text)
    If lErro <> SUCESSO Then Error 18274
    
    'Põe o valor formatado na tela
    ValorICMS.Text = Format(ValorICMS.Text, "Fixed")

    Exit Sub

Erro_ValorICMS_Validate:

    Cancel = True


    Select Case Err

        Case 18274

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157678)

    End Select

    Exit Sub

End Sub

Public Sub ValorICMSSubst_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorICMSSubst_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_ValorICMSSubst_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMSSubst.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorICMSSubst.Text)
    If lErro <> SUCESSO Then Error 18275
    
    'Põe o valor formatado na tela
    ValorICMSSubst.Text = Format(ValorICMSSubst.Text, "Fixed")

    Exit Sub

Erro_ValorICMSSubst_Validate:

    Cancel = True


    Select Case Err

        Case 18275
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157679)

    End Select

    Exit Sub

End Sub

Public Sub ValorIPI_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorIPI_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIPI_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIPI.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorIPI.Text)
    If lErro <> SUCESSO Then Error 18281
    
    'Põe o valor formatado na tela
    ValorIPI.Text = Format(ValorIPI.Text, "Fixed")

    Exit Sub

Erro_ValorIPI_Validate:

    Cancel = True


    Select Case Err

        Case 18281
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157680)

    End Select

    Exit Sub


End Sub

Public Sub ValorIRRF_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIRRF_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIRRF.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
    If lErro <> SUCESSO Then Error 18277
    
    'Põe o valor formatado na tela
    ValorIRRF.Text = Format(ValorIRRF.Text, "Fixed")

    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True


    Select Case Err

        Case 18277
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157681)

    End Select

    Exit Sub

End Sub

Public Sub ValorProdutos_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorProdutos_Validate(Cancel As Boolean)
Dim lErro As Long

On Error GoTo Erro_ValorProdutos_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorProdutos.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(ValorProdutos.Text)
    If lErro <> SUCESSO Then Error 18276
    
    'Põe o valor formatado na tela
    ValorProdutos.Text = Format(ValorProdutos.Text, "Fixed")

    Exit Sub

Erro_ValorProdutos_Validate:

    Cancel = True


    Select Case Err

        Case 18276
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157682)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorSeguro.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
    If lErro <> SUCESSO Then Error 18279
    
    'Põe o valor formatado na tela
    ValorSeguro.Text = Format(ValorSeguro.Text, "Fixed")

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True


    Select Case Err

        Case 18279

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157683)

    End Select

    Exit Sub

End Sub

Public Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Exit Sub

    'Critica se o valor é positivo
    lErro = Valor_Positivo_Critica(ValorTotal.Text)
    If lErro <> SUCESSO Then Error 18214
    
    'Põe o valor formatado na tela
    ValorTotal.Text = Format(ValorTotal.Text, "Fixed")

    Exit Sub

Erro_ValorTotal_Validate:

    Cancel = True


    Select Case Err

        Case 18214

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157684)

    End Select

    Exit Sub

End Sub

Private Function Traz_NFPag_Tela(objNFsPag As ClassNFsPag) As Long
'Preenche a tela com os dados de objNFsPag

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean
Dim objNFiscal As New ClassNFiscal 'Inserido por Wagner

On Error GoTo Erro_Traz_NFPag_Tela

    '###################################################
    'Movido por Wagner 21/02/2005
    lErro = CF("NFPag_Le_Todos2", objNFsPag)
    If lErro <> SUCESSO And lErro <> 18257 Then gError 18253

    If lErro = SUCESSO Then
    '###################################################
        
        'Coloca os dados encontrados na tela
        Fornecedor.Text = objNFsPag.lFornecedor
        
        Call Fornecedor_Validate(bCancel)
        
        Filial.Text = objNFsPag.iFilial
        Call Filial_Validate(bCancel)
        
        '#############################################################
        'Inserido por Wagner
        objNFiscal.lFornecedor = objNFsPag.lFornecedor
        objNFiscal.iFilialForn = objNFsPag.iFilial
        
        lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
        If lErro <> SUCESSO Then gError 131966
        '#############################################################
        
        NumNotaFiscal.Text = objNFsPag.lNumNotaFiscal
        DataEmissao.PromptInclude = False
        DataEmissao.Text = IIf(objNFsPag.dtDataEmissao <> DATA_NULA, Format(objNFsPag.dtDataEmissao, "dd/mm/yy"), "")
        DataEmissao.PromptInclude = True
        DataVencimento.PromptInclude = False
        DataVencimento.Text = IIf(objNFsPag.dtDataVencimento <> DATA_NULA, Format(objNFsPag.dtDataVencimento, "dd/mm/yy"), "")
        DataVencimento.PromptInclude = True
        ValorTotal.Text = objNFsPag.dValorTotal
        ValorICMS.Text = objNFsPag.dValorICMS
        ValorICMSSubst.Text = objNFsPag.dValorICMSSubst
        CreditoICMS.Value = objNFsPag.iCreditoICMS
        ValorProdutos.Text = objNFsPag.dValorProdutos
        ValorIRRF.Text = objNFsPag.dValorIRRF
        ValorFrete.Text = objNFsPag.dValorFrete
        ValorSeguro.Text = objNFsPag.dValorSeguro
        OutrasDespesas.Text = objNFsPag.dOutrasDespesas
        ValorIPI.Text = objNFsPag.dValorIPI
        CreditoIPI.Value = objNFsPag.iCreditoIPI
        ValorINSS.Text = objNFsPag.dValorINSS
        INSSRetido.Value = objNFsPag.iINSSRetido
        PISRetido.Text = objNFsPag.dPISRetido
        ISSRetido.Text = objNFsPag.dISSRetido
        COFINSRetido.Text = objNFsPag.dCOFINSRetido
        CSLLRetido.Text = objNFsPag.dCSLLRetido
        
        'traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela(objNFsPag.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 36326 Then gError 39590
    
        '###########################################################
        'Inserido por Wagner 18/02/2005
        lErro = CF2(Me, "NFPAG_Traz_NFPag_Tela", objNFsPag)
        If lErro <> SUCESSO Then gError 132232
        '###########################################################
        
    End If

    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFsPag.objProjetoInfo, objNFsPag.lNumIntDoc, PRJ_CR_TIPO_NFPAG)
    If lErro <> SUCESSO Then gError 181676
    '###################################################

    iAlterado = 0
        
    Traz_NFPag_Tela = SUCESSO

    Exit Function

Erro_Traz_NFPag_Tela:

    Traz_NFPag_Tela = gErr
    
    Select Case gErr
    
        Case 18428
            Call Rotina_Erro(vbOKOnly, "ERRO_NFPAG_NAO_CADASTRADA", gErr, objNFsPag.lNumIntDoc)
    
        Case 39590 'Tratado na Rotina chamada
        
        Case 131966, 132232, 181676 'Inserido por Wagner
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157685)
    
    End Select
        
    Exit Function
            
End Function

Public Function Gravar_Registro() As Long
'Verifica os dados necessários da Tela. Grava a Nota Fiscal.

'alteracao Daniel - 21/09/2001
'Alterada a comparacao de valores

Dim lErro As Long
Dim dSoma As Double
Dim objNFsPag As New ClassNFsPag
Dim dValorProdutos As Double
Dim dValorICMSSubst As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorIPI As Double
Dim dOutrasDespesas As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 18283
    
    'Verifica se Filial foi preechida
    If Len(Trim(Filial.Text)) = 0 Then gError 18284
    
    'Verifica se o Nuúero da Nota foi preenchido
    If Len(Trim(NumNotaFiscal.ClipText)) = 0 Then gError 18285
    
    'Verifica se ValorTotal foi preenchido
    If Len(Trim(ValorTotal.Text)) = 0 Then gError 6786
    
    'Verifica se o Valor dos Produtos foi preenchido
    If Len(Trim(ValorProdutos.Text)) = 0 Then gError 18286
    
    'Converte valores para Double
    If Len(Trim(ValorProdutos.Text)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(ValorICMSSubst.Text)) > 0 Then dValorICMSSubst = CDbl(ValorICMSSubst)
    If Len(Trim(ValorFrete.Text)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.Text)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(OutrasDespesas.Text)) > 0 Then dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorIPI.Text)) > 0 Then dValorIPI = CDbl(ValorIPI.Text)
    
    'Calcula a soma dos valores
    dSoma = dValorProdutos + dValorICMSSubst + dValorFrete + dValorSeguro + dOutrasDespesas + dValorIPI
    
    'alteracao Daniel - 21/09/2001
    'Verifica se a soma dos valores é igual ao valor total
    If Abs(dSoma - CDbl(ValorTotal.Text)) > 0.01 Then gError 18288
    
    If Len(Trim(DataEmissao.ClipText)) > 0 And Len(Trim(DataVencimento.ClipText)) > 0 Then
        'Verifica se a data de emissao é maior que a data de vencimento
        If CDate(DataEmissao) > CDate(DataVencimento) Then gError 18289
    End If
    
    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objNFsPag)
    If lErro <> SUCESSO Then gError 18290
        
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    If objNFsPag.dtDataEmissao <> DATA_NULA Then
        lErro = objContabil.Contabil_Testa_Data(objNFsPag.dtDataEmissao)
        If lErro <> SUCESSO Then gError 20832
    End If

    '###############################################
    'Inserido por Wagner 10/08/2006
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    '###############################################
    
    'Dispara a rotina de gravação(gravando também a contabilidade)
    lErro = CF("NFPag_Grava", objNFsPag, objContabil)
    If lErro <> SUCESSO Then gError 18291

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 18283
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 18284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 18285
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NF_NAO_INFORMADA", gErr)
        
        Case 6786
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", gErr)
        
        Case 18286
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPRODUTOS_NAO_INFORMADO", gErr)
            
        Case 18288
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_INVALIDO", gErr, ValorTotal.Text, dSoma)
            
        Case 18289
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_MAIOR", gErr, DataEmissao.Text, DataVencimento.Text)
            
        Case 18290, 18291, 20832, 181786 'Tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157686)
            
    End Select
            
    Exit Function
    
End Function

Private Sub Limpa_Tela_NFPag()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBoxes e MaskedEdits
    Call Limpa_Tela(Me)
    
    'Limpa os outros controles
    Filial.Clear
    CreditoICMS.Value = 0
    CreditoIPI.Value = 0
    INSSRetido.Value = 0
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    '#############################################
    'Inserido por Wagner
    Call CF2(Me, "NFCF_Limpa_Tela")
    '#############################################
    
    '#############################################
    'Inserido por Wagner 18/02/2005
    Call CF2(Me, "NFPAG_Limpa_Tela")
    '#############################################

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    iAlterado = 0
    
End Sub

Private Function Move_Tela_Memoria(objNFsPag As ClassNFsPag) As Long
'Preenche objNFsPag com os dados da Tela

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Tela_Memoria

    objNFsPag.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(Fornecedor.Text)) > 0 Then
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
    
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 18292
        If lErro = 6681 Then gError 18429
    
        objNFsPag.lFornecedor = objFornecedor.lCodigo
        
    End If

    If Len(Trim(Filial.Text)) > 0 Then
        objNFsPag.iFilial = Codigo_Extrai(Filial.Text)
    End If
    
    If Len(Trim(NumNotaFiscal.Text)) > 0 Then objNFsPag.lNumNotaFiscal = CLng(NumNotaFiscal.Text)
    
    'Se as  datas nao forem digitadas completar com data_nula
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objNFsPag.dtDataEmissao = DATA_NULA
    Else
        objNFsPag.dtDataEmissao = CDate(DataEmissao.Text)
    End If
    
    If Len(Trim(DataVencimento.ClipText)) = 0 Then
        objNFsPag.dtDataVencimento = DATA_NULA
    Else
        objNFsPag.dtDataVencimento = CDate(DataVencimento.Text)
    End If
    
    If Len(Trim(ValorTotal.ClipText)) > 0 Then objNFsPag.dValorTotal = CDbl(ValorTotal.Text)
    If Len(Trim(ValorICMS.ClipText)) > 0 Then objNFsPag.dValorICMS = CDbl(ValorICMS.Text)
    If Len(Trim(ValorICMSSubst.ClipText)) > 0 Then objNFsPag.dValorICMSSubst = CDbl(ValorICMSSubst.Text)
    objNFsPag.iCreditoICMS = CreditoICMS.Value
    If Len(Trim(ValorProdutos.ClipText)) > 0 Then objNFsPag.dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objNFsPag.dValorSeguro = CDbl(ValorSeguro)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objNFsPag.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorIRRF.ClipText)) > 0 Then objNFsPag.dValorIRRF = CDbl(ValorIRRF.Text)
    If Len(Trim(OutrasDespesas.ClipText)) > 0 Then objNFsPag.dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorIPI.ClipText)) > 0 Then objNFsPag.dValorIPI = CDbl(ValorIPI.Text)
    objNFsPag.iCreditoIPI = CreditoIPI.Value
    objNFsPag.lNumIntTitPag = 0
    
    If Len(Trim(ValorINSS.ClipText)) > 0 Then objNFsPag.dValorINSS = CDbl(ValorINSS.Text)
    objNFsPag.iINSSRetido = INSSRetido.Value
    If Len(Trim(PISRetido.ClipText)) > 0 Then objNFsPag.dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(ISSRetido.ClipText)) > 0 Then objNFsPag.dISSRetido = CDbl(ISSRetido.Text)
    If Len(Trim(COFINSRetido.ClipText)) > 0 Then objNFsPag.dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.ClipText)) > 0 Then objNFsPag.dCSLLRetido = CDbl(CSLLRetido.Text)
        
    lErro = CF2(Me, "NFPAG_Move_Tela_Memoria", objNFsPag)
    If lErro <> SUCESSO Then gError 132233
    
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_NFPAG)
    If lErro <> SUCESSO Then gError 181761
    
    Set objNFsPag.objProjetoInfo = objProjetoInfo
        
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 18292, 132233, 181761 'Tratado na rotina chamada
        
        Case 18429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157687)
    
    End Select
    
    Exit Function
    
End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFsPag_Todas"

    'Lê os dados da Tela Notas Fiscais à Pagar
    lErro = Move_Tela_Memoria(objNFsPag)
    If lErro <> SUCESSO Then Error 18360

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", CLng(0), 0, "NumIntDoc"
    colCampoValor.Add "Fornecedor", objNFsPag.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objNFsPag.iFilial, 0, "Filial"
    colCampoValor.Add "NumNotaFiscal", objNFsPag.lNumNotaFiscal, 0, "NumNotaFiscal"
    colCampoValor.Add "DataEmissao", objNFsPag.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "DataVencimento", objNFsPag.dtDataVencimento, 0, "DataVencimento"
    colCampoValor.Add "ValorTotal", objNFsPag.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorSeguro", objNFsPag.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorFrete", objNFsPag.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "OutrasDespesas", objNFsPag.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "ValorProdutos", objNFsPag.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorIRRF", objNFsPag.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "ValorICMS", objNFsPag.dValorICMS, 0, "ValorICMS"
    colCampoValor.Add "ValorICMSSubst", objNFsPag.dValorICMSSubst, 0, "ValorICMSSubst"
    colCampoValor.Add "CreditoICMS", objNFsPag.iCreditoICMS, 0, "CreditoICMS"
    colCampoValor.Add "ValorIPI", objNFsPag.dValorIPI, 0, "ValorIPI"
    colCampoValor.Add "CreditoIPI", objNFsPag.iCreditoIPI, 0, "CreditoIPI"
    
    'Filtro para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 18360 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157688)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag

On Error GoTo Erro_Tela_Preenche

    objNFsPag.lNumIntDoc = colCampoValor.Item("NumIntdoc").vValor

    If objNFsPag.lNumIntDoc <> 0 Then

        'Carrega objNFPag com os dados passados em colCampoValor
        objNFsPag.lNumNotaFiscal = colCampoValor.Item("NumNotaFiscal").vValor
        objNFsPag.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objNFsPag.iFilial = colCampoValor.Item("Filial").vValor
        objNFsPag.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objNFsPag.dtDataVencimento = colCampoValor.Item("DataVencimento").vValor
        objNFsPag.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objNFsPag.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
        objNFsPag.dValorFrete = colCampoValor.Item("ValorFrete").vValor
        objNFsPag.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
        objNFsPag.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
        objNFsPag.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        objNFsPag.dValorICMS = colCampoValor.Item("ValorICMS").vValor
        objNFsPag.dValorIPI = colCampoValor.Item("ValorIPI").vValor
        objNFsPag.iCreditoIPI = colCampoValor.Item("CreditoIPI").vValor
        objNFsPag.iCreditoICMS = colCampoValor.Item("CreditoICMS").vValor
        objNFsPag.dValorICMSSubst = colCampoValor.Item("ValorICMSSubst").vValor
        
        'Preenche a Tela
        lErro = Traz_NFPag_Tela(objNFsPag)
        If lErro <> SUCESSO Then Error 18361

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 18361 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 157689)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'Início contabilidade
Public Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Public Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

End Sub

Public Sub CTBGridContabil_EnterCell()

    Call objContabil.Contabil_GridContabil_EnterCell

End Sub

Public Sub CTBGridContabil_GotFocus()

    Call objContabil.Contabil_GridContabil_GotFocus

End Sub

Public Sub CTBGridContabil_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)

End Sub

Public Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)

    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
    
End Sub

Public Sub CTBGridContabil_LeaveCell()

    Call objContabil.Contabil_GridContabil_LeaveCell

End Sub

Public Sub CTBGridContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_GridContabil_Validate(Cancel)

End Sub

Public Sub CTBGridContabil_RowColChange()

    Call objContabil.Contabil_GridContabil_RowColChange

End Sub

Public Sub CTBGridContabil_Scroll()

    Call objContabil.Contabil_GridContabil_Scroll

End Sub

Public Sub CTBConta_Change()

    Call objContabil.Contabil_Conta_Change

End Sub

Public Sub CTBConta_GotFocus()

    Call objContabil.Contabil_Conta_GotFocus

End Sub

Public Sub CTBConta_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)

End Sub

Public Sub CTBConta_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Conta_Validate(Cancel)

End Sub

Public Sub CTBCcl_Change()

    Call objContabil.Contabil_Ccl_Change

End Sub

Public Sub CTBCcl_GotFocus()

    Call objContabil.Contabil_Ccl_GotFocus

End Sub

Public Sub CTBCcl_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)

End Sub

Public Sub CTBCcl_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Ccl_Validate(Cancel)

End Sub

Public Sub CTBCredito_Change()

    Call objContabil.Contabil_Credito_Change

End Sub

Public Sub CTBCredito_GotFocus()

    Call objContabil.Contabil_Credito_GotFocus

End Sub

Public Sub CTBCredito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)

End Sub

Public Sub CTBCredito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Credito_Validate(Cancel)

End Sub

Public Sub CTBDebito_Change()

    Call objContabil.Contabil_Debito_Change

End Sub

Public Sub CTBDebito_GotFocus()

    Call objContabil.Contabil_Debito_GotFocus

End Sub

Public Sub CTBDebito_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)

End Sub

Public Sub CTBDebito_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Debito_Validate(Cancel)

End Sub

Public Sub CTBSeqContraPartida_Change()

    Call objContabil.Contabil_SeqContraPartida_Change

End Sub

Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

Public Sub CTBSeqContraPartida_Validate(Cancel As Boolean)

    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)

End Sub

Public Sub CTBHistorico_Change()

    Call objContabil.Contabil_Historico_Change

End Sub

Public Sub CTBHistorico_GotFocus()

    Call objContabil.Contabil_Historico_GotFocus

End Sub

Public Sub CTBHistorico_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)

End Sub

Public Sub CTBHistorico_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Historico_Validate(Cancel)

End Sub

Public Sub CTBLancAutomatico_Click()

    Call objContabil.Contabil_LancAutomatico_Click

End Sub

Public Sub CTBAglutina_Click()
    
    Call objContabil.Contabil_Aglutina_Click

End Sub

Public Sub CTBAglutina_GotFocus()

    Call objContabil.Contabil_Aglutina_GotFocus

End Sub

Public Sub CTBAglutina_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)

End Sub

Public Sub CTBAglutina_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Aglutina_Validate(Cancel)

End Sub

Public Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_NodeClick(Node)

End Sub

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

End Sub

Public Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwCcls_NodeClick(Node)

End Sub

Public Sub CTBListHistoricos_DblClick()

    Call objContabil.Contabil_ListHistoricos_DblClick

End Sub

Public Sub CTBBotaoLimparGrid_Click()

    Call objContabil.Contabil_Limpa_GridContabil

End Sub

Public Sub CTBLote_Change()

    Call objContabil.Contabil_Lote_Change

End Sub

Public Sub CTBLote_GotFocus()

    Call objContabil.Contabil_Lote_GotFocus

End Sub

Public Sub CTBLote_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)

End Sub

Public Sub CTBDataContabil_Change()

    Call objContabil.Contabil_DataContabil_Change

End Sub

Public Sub CTBDataContabil_GotFocus()

    Call objContabil.Contabil_DataContabil_GotFocus

End Sub

Public Sub CTBDataContabil_Validate(Cancel As Boolean)

    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)

End Sub

Private Sub objEventoLote_evSelecao(obj1 As Object)
'traz o lote selecionado para a tela

    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)

End Sub

Private Sub objEventoDoc_evSelecao(obj1 As Object)

    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)

End Sub

Public Sub CTBDocumento_Change()

    Call objContabil.Contabil_Documento_Change

End Sub

Public Sub CTBDocumento_GotFocus()

    Call objContabil.Contabil_Documento_GotFocus

End Sub

Public Sub CTBBotaoImprimir_Click()
    
    Call objContabil.Contabil_BotaoImprimir_Click

End Sub

Public Sub CTBUpDown_DownClick()

    Call objContabil.Contabil_UpDown_DownClick
    
End Sub

Public Sub CTBUpDown_UpClick()

    Call objContabil.Contabil_UpDown_UpClick

End Sub

Public Sub CTBLabelDoc_Click()

    Call objContabil.Contabil_LabelDoc_Click
    
End Sub

Public Sub CTBLabelLote_Click()

    Call objContabil.Contabil_LabelLote_Click
    
End Sub

'MNEMONICOS
Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor, objTipoFornecedor As New ClassTipoFornecedor
Dim objFilial As New ClassFilialFornecedor, sContaTela As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case VALORTOTAL1
            
            If Len(Trim(ValorTotal.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorTotal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case FORNECEDOR_COD
            
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 39593
                
                objMnemonicoValor.colValor.Add objFornecedor.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case FORNECEDOR_NOME
        
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 39594
            
                objMnemonicoValor.colValor.Add objFornecedor.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case FILIAL_COD
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                objMnemonicoValor.colValor.Add objFilial.iCodFilial
            
            Else
                
                objMnemonicoValor.colValor.Add 0
            
            End If
            
        Case FILIAL_NOME_RED
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39595
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39596
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 41966
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39597
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case NUMERO1
            
            If Len(Trim(NumNotaFiscal.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NumNotaFiscal.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If

        Case EMISSAO1
            If Len(Trim(DataEmissao.ClipText)) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
              objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case VALOR_ICMS
            If Len(Trim(ValorICMS.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorICMS.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case VALOR_ICMS_SUBST
            If Len(Trim(ValorICMSSubst.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorICMSSubst.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case CREDITA_ICMS
            objMnemonicoValor.colValor.Add CreditoICMS.Value
            
        Case VALOR_IR
            If Len(Trim(ValorIRRF.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIRRF.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_PRODUTOS
            If Len(Trim(ValorProdutos.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorProdutos.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case VALOR_FRETE
            If Len(Trim(ValorFrete.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorFrete.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_SEGURO
            If Len(Trim(ValorSeguro.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorSeguro.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_IPI
            If Len(Trim(ValorIPI.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIPI.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CREDITA_IPI
            objMnemonicoValor.colValor.Add CreditoIPI.Value
                                
        Case VALOR_OUTRASDESP
            If Len(Trim(OutrasDespesas.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(OutrasDespesas.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CONTA_DESP_ESTOQUE
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 41964
                
                objTipoFornecedor.iCodigo = objFornecedor.iTipo
                lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
                If lErro <> SUCESSO Then gError 41965
                
                If objTipoFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTipoFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 41967
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case CONTA_DESP_EST_FORN
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 133566
                
                If objFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 133567
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
         Case PIS_RETIDO
            If Len(Trim(PISRetido.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(PISRetido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case ISS_RETIDO
            If Len(Trim(ISSRetido.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ISSRetido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case COFINS_RETIDO
            If Len(Trim(COFINSRetido.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(COFINSRetido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CSLL_RETIDO
            If Len(Trim(CSLLRetido.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(CSLLRetido.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case VALOR_INSS
            If Len(Trim(ValorINSS.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorINSS.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case INSS_RETIDO
            If INSSRetido.Value = vbChecked Then
                objMnemonicoValor.colValor.Add 1
            ElseIf INSSRetido.Value = vbUnchecked Then
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case Else
            gError 39598
            
    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr
        
        Case 39593, 39594, 39595, 39596, 39597, 41964, 41965, 41966, 41967, 133566, 133567
        
        Case 39598
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157690)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_NOTAS_FISCAIS_PAGAR_ID
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal a Pagar"
    Call Form_Load
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Function Name() As String

    Name = "NFPAG"
    
End Function

Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'***** fim do trecho a ser copiado ******

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is NumNotaFiscal Then
            Call NumeroLabel_Click
        '#########################################
        'Inserido por Wagner 10/08/2006
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        '#########################################
        End If
    
    End If
    
End Sub

Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Private Sub Fornecedor_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134044

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134044

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157691)

    End Select
    
    Exit Sub

End Sub


'##################################
'Inserido por Wagner
Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long
Dim sFiltro As String
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoProxNum_Click

    If Len(Trim(Fornecedor.Text)) = 0 Then gError 132158
    If Len(Trim(Filial.Text)) = 0 Then gError 132159
        
    objFornecedor.sNomeReduzido = Fornecedor.Text

    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 132160
    If lErro = 6681 Then gError 132161
    
    sFiltro = "Fornecedor = " & objFornecedor.lCodigo & " AND Filial = " & CStr(Codigo_Extrai(Filial.Text))
    
    lErro = CF("NFiscalPag_Automatico", lCodigo, sFiltro)
    If lErro <> SUCESSO Then gError 132153
    
    NumNotaFiscal.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 132153, 132160

        Case 132158
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 132159
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 132161
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157692)
    
    End Select

    Exit Sub
    
End Sub
'##################################

'#####################################
'Inserido por Wagner 18/02/2005
'@@@@@ para ser customizada
Function NFPAG_Form_Load(ByVal objCT As Object) As Long
    NFPAG_Form_Load = SUCESSO
End Function
'@@@@@ para ser customizada
Function NFPAG_Traz_NFPag_Tela(ByVal objCT As Object, objNFsPag As ClassNFsPag) As Long
    NFPAG_Traz_NFPag_Tela = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFPAG_Move_Tela_Memoria(ByVal objCT As Object, objNFsPag As ClassNFsPag) As Long
    NFPAG_Move_Tela_Memoria = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFPAG_Limpa_Tela(ByVal objCT As Object) As Long
    NFPAG_Limpa_Tela = SUCESSO
End Function
'#####################################

'#####################################
'Inserido por Wagner 03/08/2006
Public Function BotaoProjetos_Click() As Long
        Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Function

Public Function LabelProjeto_Click() As Long
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Function

Public Function Projeto_GotFocus() As Long
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Function

Public Function Projeto_Change() As Long
    iAlterado = REGISTRO_ALTERADO
End Function

Public Function Projeto_Validate(Cancel As Boolean) As Long
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Function

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR
   
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "PROD"
    objItensPRJCR.sDescricao = "Produtos"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorProdutos.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
   Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "FR"
    objItensPRJCR.sDescricao = "Frete"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorFrete.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "SG"
    objItensPRJCR.sDescricao = "Seguro"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorSeguro.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "OD"
    objItensPRJCR.sDescricao = "Outras Despesas"

    objItensPRJCR.dValorOriginal = StrParaDbl(OutrasDespesas.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "ICMS"
    objItensPRJCR.sDescricao = "ICMS Substituto"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorICMSSubst.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "IPI"
    objItensPRJCR.sDescricao = "IPI"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorIPI.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR

    Obter_ItensPRJCR = SUCESSO

    Exit Function

Erro_Obter_ItensPRJCR:

    Obter_ItensPRJCR = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 157006)

    End Select

    Exit Function
End Function
'#####################################


Public Sub CTBGerencial_Click()
    
    Call objContabil.Contabil_Gerencial_Click

End Sub

Public Sub CTBGerencial_GotFocus()

    Call objContabil.Contabil_Gerencial_GotFocus

End Sub

Public Sub CTBGerencial_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_Gerencial_KeyPress(KeyAscii)

End Sub

Public Sub CTBGerencial_Validate(Cancel As Boolean)

    Call objContabil.Contabil_Gerencial_Validate(Cancel)

End Sub

Public Sub ISSRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    iISSRetidoAlterado = 1
    
End Sub

Public Sub ISSRetido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ISSRetido_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ISSRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ISSRetido.Text)
        If lErro <> SUCESSO Then gError 18405
    
        'Põe o valor formatado na tela
        ISSRetido.Text = Format(ISSRetido.Text, "Fixed")

    End If
    
    If iISSRetidoAlterado <> 0 Then
    
        iISSRetidoAlterado = 0
        
    End If
    
    Exit Sub

Erro_ISSRetido_Validate:

    Cancel = True

    Select Case gErr

        Case 18405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156163)

    End Select

    Exit Sub

End Sub

Public Sub INSSRetido_Click()

Dim lErro As Long

On Error GoTo Erro_INSSRetido_Click

    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_INSSRetido_Click:
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156156)
            
    End Select
    
    Exit Sub

End Sub

Public Sub PISRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    iPISRetidoAlterado = 1
    
End Sub

Public Sub PISRetido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PISRetido_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(PISRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(PISRetido.Text)
        If lErro <> SUCESSO Then gError 18405
    
        'Põe o valor formatado na tela
        PISRetido.Text = Format(PISRetido.Text, "Fixed")

    End If
    
    If iPISRetidoAlterado <> 0 Then
    
        iPISRetidoAlterado = 0
        
    End If
    
    Exit Sub

Erro_PISRetido_Validate:

    Cancel = True

    Select Case gErr

        Case 18405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156163)

    End Select

    Exit Sub

End Sub

Public Sub COFINSRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    iCOFINSRetidoAlterado = 1
    
End Sub

Public Sub COFINSRetido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_COFINSRetido_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(COFINSRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(COFINSRetido.Text)
        If lErro <> SUCESSO Then gError 18405
    
        'Põe o valor formatado na tela
        COFINSRetido.Text = Format(COFINSRetido.Text, "Fixed")

    End If
    
    If iCOFINSRetidoAlterado <> 0 Then
    
        iCOFINSRetidoAlterado = 0
        
    End If
    
    Exit Sub

Erro_COFINSRetido_Validate:

    Cancel = True

    Select Case gErr

        Case 18405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156164)

    End Select

    Exit Sub

End Sub

Public Sub CSLLRetido_Change()

    iAlterado = REGISTRO_ALTERADO
    iCSLLRetidoAlterado = 1
    
End Sub

Public Sub CSLLRetido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CSLLRetido_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(CSLLRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(CSLLRetido.Text)
        If lErro <> SUCESSO Then gError 18405
    
        'Põe o valor formatado na tela
        CSLLRetido.Text = Format(CSLLRetido.Text, "Fixed")

    End If
    
    If iCSLLRetidoAlterado <> 0 Then
    
        iCSLLRetidoAlterado = 0
        
    End If
    
    Exit Sub

Erro_CSLLRetido_Validate:

    Cancel = True

    Select Case gErr

        Case 18405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156165)

    End Select

    Exit Sub

End Sub

Public Sub ValorINSS_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorINSSAlterado = 1
    
End Sub

Public Sub ValorINSS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorINSS_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorINSS.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorINSS.Text)
    If lErro <> SUCESSO Then gError 18593

    'Põe o valor formatado na tela
    ValorINSS.Text = Format(ValorINSS.Text, "Fixed")

    If iValorINSSAlterado <> 0 Then
    
        iValorINSSAlterado = 0
        
    End If
    
    Exit Sub

Erro_ValorINSS_Validate:

    Cancel = True


    Select Case gErr

        Case 18593

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156143)

    End Select

    Exit Sub

End Sub
