VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFFATPAG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim giTeste As Integer

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo
Dim gobjAnexos As ClassAnexos

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Private iFrameAtual As Integer
Private iFornecedorAlterado As Integer
Private iEmissaoAlterada As Integer
Private iValorTotalAlterado As Integer
Private iValorIRRFAlterado As Integer
Private iValorINSSAlterado As Integer
Private iPISRetidoAlterado As Integer
Private iISSRetidoAlterado As Integer
Private iCOFINSRetidoAlterado As Integer
Private iCSLLRetidoAlterado As Integer

Public objGridParcelas As AdmGrid
Public iGrid_Parcela_Col As Integer
Public iGrid_Vencimento_Col As Integer
Public iGrid_VenctoReal_Col As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_Cobranca_Col As Integer
Public iGrid_Suspenso_Col As Integer
Public iGrid_Banco_Col As Integer
Public iGrid_Portador_Col As Integer
Public iGrid_CodigoDeBarras_Col As Integer
Public iGrid_ValorOriginal_Col As Integer
Public iGrid_MotivoDiferenca_Col As Integer

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoFilialPC As AdmEvento
Attribute objEventoFilialPC.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1

'inicio contabilidade

Public objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

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
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const TIPO1 As String = "Tipo_Documento"
Private Const VALOR_PRODUTOS As String = "Valor_Produtos"
Private Const VALOR_FRETE As String = "Valor_Frete"
Private Const VALOR_SEGURO As String = "Valor_Seguro"
Private Const VALOR_IPI As String = "Valor_IPI"
Private Const VALOR_OUTRASDESP = "Valor_OutrasDesp"
Private Const CREDITA_IPI As String = "Credita_IPI"
Private Const CONTA_DESP_ESTOQUE As String = "Conta_Desp_Estoque"
Private Const CONTA_DESP_EST_FORN As String = "Conta_Desp_Est_Forn"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Cobranca = 2
Private Const TAB_Contabilizacao = 3

'--- inicio dos properties get dos controles da tela
Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get ComboFilialPC() As Object
     Set ComboFilialPC = objUserControl.Controls("ComboFilialPC")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
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

Public Property Get NumPC() As Object
     Set NumPC = objUserControl.Controls("NumPC")
End Property

Public Property Get LabelTipo() As Object
     Set LabelTipo = objUserControl.Controls("LabelTipo")
End Property

Public Property Get LblNumPC() As Object
     Set LblNumPC = objUserControl.Controls("LblNumPC")
End Property

Public Property Get LblFilialPC() As Object
     Set LblFilialPC = objUserControl.Controls("LblFilialPC")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get INSSRetido() As Object
     Set INSSRetido = objUserControl.Controls("INSSRetido")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
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

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
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

Public Property Get ValorINSS() As Object
     Set ValorINSS = objUserControl.Controls("ValorINSS")
End Property

Public Property Get ValorIRRF() As Object
     Set ValorIRRF = objUserControl.Controls("ValorIRRF")
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

Public Property Get OutrasDespesas() As Object
     Set OutrasDespesas = objUserControl.Controls("OutrasDespesas")
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
End Property

Public Property Get ValorFrete() As Object
     Set ValorFrete = objUserControl.Controls("ValorFrete")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
End Property

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBFrame7() As Object
     Set CTBFrame7 = objUserControl.Controls("CTBFrame7")
End Property

Public Property Get CTBCclDescricao() As Object
     Set CTBCclDescricao = objUserControl.Controls("CTBCclDescricao")
End Property

Public Property Get CTBContaDescricao() As Object
     Set CTBContaDescricao = objUserControl.Controls("CTBContaDescricao")
End Property

Public Property Get CTBLabel7() As Object
     Set CTBLabel7 = objUserControl.Controls("CTBLabel7")
End Property

Public Property Get CTBCclLabel() As Object
     Set CTBCclLabel = objUserControl.Controls("CTBCclLabel")
End Property

Public Property Get CTBModelo() As Object
     Set CTBModelo = objUserControl.Controls("CTBModelo")
End Property

Public Property Get CTBBotaoLimparGrid() As Object
     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
End Property

Public Property Get CTBBotaoModeloPadrao() As Object
     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
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

Public Property Get CTBTvwContas() As Object
     Set CTBTvwContas = objUserControl.Controls("CTBTvwContas")
End Property

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
End Property

Public Property Get CTBTvwCcls() As Object
     Set CTBTvwCcls = objUserControl.Controls("CTBTvwCcls")
End Property

Public Property Get CTBLabelLote() As Object
     Set CTBLabelLote = objUserControl.Controls("CTBLabelLote")
End Property

Public Property Get CTBLabelDoc() As Object
     Set CTBLabelDoc = objUserControl.Controls("CTBLabelDoc")
End Property

Public Property Get CTBLabel8() As Object
     Set CTBLabel8 = objUserControl.Controls("CTBLabel8")
End Property

Public Property Get CTBTotalCredito() As Object
     Set CTBTotalCredito = objUserControl.Controls("CTBTotalCredito")
End Property

Public Property Get CTBTotalDebito() As Object
     Set CTBTotalDebito = objUserControl.Controls("CTBTotalDebito")
End Property

Public Property Get CTBLabelTotais() As Object
     Set CTBLabelTotais = objUserControl.Controls("CTBLabelTotais")
End Property

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
End Property

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
End Property

Public Property Get CTBLabelContas() As Object
     Set CTBLabelContas = objUserControl.Controls("CTBLabelContas")
End Property

Public Property Get CTBLabelHistoricos() As Object
     Set CTBLabelHistoricos = objUserControl.Controls("CTBLabelHistoricos")
End Property

Public Property Get CTBLabel5() As Object
     Set CTBLabel5 = objUserControl.Controls("CTBLabel5")
End Property

Public Property Get CTBLabel13() As Object
     Set CTBLabel13 = objUserControl.Controls("CTBLabel13")
End Property

Public Property Get CTBExercicio() As Object
     Set CTBExercicio = objUserControl.Controls("CTBExercicio")
End Property

Public Property Get CTBPeriodo() As Object
     Set CTBPeriodo = objUserControl.Controls("CTBPeriodo")
End Property

Public Property Get CTBLabel14() As Object
     Set CTBLabel14 = objUserControl.Controls("CTBLabel14")
End Property

Public Property Get CTBOrigem() As Object
     Set CTBOrigem = objUserControl.Controls("CTBOrigem")
End Property

Public Property Get CTBLabel21() As Object
     Set CTBLabel21 = objUserControl.Controls("CTBLabel21")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get CobrancaAutomatica() As Object
     Set CobrancaAutomatica = objUserControl.Controls("CobrancaAutomatica")
End Property

Public Property Get TipoCobranca() As Object
     Set TipoCobranca = objUserControl.Controls("TipoCobranca")
End Property

Public Property Get comboCobrador() As Object
     Set comboCobrador = objUserControl.Controls("comboCobrador")
End Property

Public Property Get CodigodeBarras() As Object
     Set CodigodeBarras = objUserControl.Controls("CodigodeBarras")
End Property

Public Property Get comboPortador() As Object
     Set comboPortador = objUserControl.Controls("comboPortador")
End Property

Public Property Get Suspenso() As Object
     Set Suspenso = objUserControl.Controls("Suspenso")
End Property

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get DataVencimentoReal() As Object
     Set DataVencimentoReal = objUserControl.Controls("DataVencimentoReal")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get CondPagtoLabel() As Object
     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
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

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get CodBarrasAvulso() As Object
     Set CodBarrasAvulso = objUserControl.Controls("CodBarrasAvulso")
End Property

Public Property Get Diferenca() As Object
     Set Diferenca = objUserControl.Controls("Diferenca")
End Property

Public Property Get MotivoDiferenca() As Object
     Set MotivoDiferenca = objUserControl.Controls("MotivoDiferenca")
End Property

Public Property Get ValorOriginal() As Object
     Set ValorOriginal = objUserControl.Controls("ValorOriginal")
End Property

Public Property Get Natureza() As Object
     Set Natureza = objUserControl.Controls("Natureza")
End Property

Public Property Get LabelNatureza() As Object
     Set LabelNatureza = objUserControl.Controls("LabelNatureza")
End Property

Public Property Get LabelNaturezaDesc() As Object
     Set LabelNaturezaDesc = objUserControl.Controls("LabelNaturezaDesc")
End Property

'--- fim dos properties get dos controles da tela

Public Sub CobrancaAutomatica_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Recalcula_Cobranca

End Sub

'Fim contabilidade

Public Sub ComboFilialPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboFilialPC_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ComboFilialPC_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim sFornecedor As String

On Error GoTo Erro_ComboFilialPC_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(ComboFilialPC.Text)) = 0 Then Exit Sub
    
    'Verifica se é uma filial selecionada
    If ComboFilialPC.Text = ComboFilialPC.List(ComboFilialPC.ListIndex) Then Exit Sub
    
    'Tenta selecionar na combo
    lErro = Combo_Item_Seleciona(ComboFilialPC)
    If lErro <> SUCESSO And lErro <> 12250 Then gError 48915
    
    'Se não encontra valor que era CÓDIGO
    If lErro = 12250 Then gError 48916
    
    Exit Sub
    
Erro_ComboFilialPC_Validate:

    Cancel = True


    Select Case gErr
    
       Case 48915
       
       Case 48916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA2", gErr, ComboFilialPC.Text)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156108)
    
    End Select
    
    Exit Sub

End Sub

Public Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CondicaoPagamento_Validate

    'Verifica se a Condicaopagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicaopagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18413

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 18414
        If lErro = 19205 Then gError 18415

        'Testa se pode ser usada em Contas a Pagar
        If objCondicaoPagto.iEmPagamento = 0 Then gError 18783

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)

        Call Recalcula_Cobranca
    
    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then gError 18416

    Exit Sub

Erro_CondicaoPagamento_Validate:
    
    Cancel = True
    
    Select Case gErr

       Case 18413, 18414 'Tratado na Rotina chamada

       Case 18415
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 18416
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)

        Case 18783
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156109)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
Dim iEmissaoAux As Integer
    
    iEmissaoAux = iEmissaoAlterada
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    iEmissaoAlterada = iEmissaoAux

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 18397

    If iEmissaoAlterada = 1 Then
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case gErr

        Case 18397

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156110)

    End Select

    Exit Sub

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
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18392

    'Se não encontra valor que era CÓDIGO
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 18393

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 18394
        
        'Se não encontrou
        If lErro = 18272 Then
        
            objFornecedor.sNomeReduzido = sFornecedor
            
            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 58616
            
            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
            
            'Sugere cadastrar nova Filial
            Error 18395
        
        End If

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        '#############################################
        'Inserido por Wagner
        lErro = CF2(Me, "NFCF_Fornecedor_Filial_Validate", objFilialFornecedor)
        If lErro <> SUCESSO Then gError 131955
        '#############################################
    
    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 18396

    Exit Sub

Erro_Filial_Validate:

    Cancel = True
    
    Select Case gErr

       Case 18392, 18394, 58616, 131955 'Inserido por Wagner

       Case 18393
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

       Case 18395
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 18396
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156111)

    End Select

    Exit Sub
End Sub

Public Sub LblFilialPc_Click()

Dim colSelecao As Collection
Dim objFilialPC As New AdmFiliais

    'Se Tipo estiver preenchida, extrai a  sigla
    If Len(Trim(ComboFilialPC.Text)) > 0 Then
        objFilialPC.iCodFilial = ComboFilialPC.Text
    End If

    'Chama a tela Tipo
    Call Chama_Tela("FilialEmpresaLista", colSelecao, objFilialPC, objEventoFilialPC)

End Sub

Public Sub NumPC_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NumPC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumPC, iAlterado)

End Sub

Public Sub NumPC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumPC_Validate

    If Len(Trim(NumPC.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(NumPC.Text)
        If lErro <> SUCESSO Then gError 57976
        
    End If
    
    Exit Sub
    
Erro_NumPC_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 57976 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156112)
    
    End Select
    
    Exit Sub

End Sub

Public Sub NumTitulo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)

End Sub

Private Sub objEventoFilialPC_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objFilialPC As AdmFiliais

On Error GoTo Erro_objEventoFilialPC_evSelecao

    Set objFilialPC = obj1
    
    'Preenche a combo de FilialPC
    ComboFilialPC.Text = objFilialPC.iCodFilial
    Me.Show
     
    Exit Sub
    
Erro_objEventoFilialPC_evSelecao:

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156113)
     
     End Select
     
     Exit Sub
     
End Sub

Private Function Carrega_FilialPC() As Long
'Carrega ComboFilialPc com as Filiais Empresas

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim colCodigo As New Collection

On Error GoTo Erro_Carrega_FilialPC
        
    'Le Código
    lErro = CF("Codigos_Le", "FiliaisEmpresa", "FilialEmpresa", TIPO_INT, colCodigo, 0)
    If lErro <> SUCESSO Then gError 48905

    For iIndice1 = 1 To colCodigo.Count
            
            iCodigo = colCodigo.Item(iIndice1)
            
            'coloca na combo
            ComboFilialPC.AddItem CStr(iCodigo)
            ComboFilialPC.ItemData(ComboFilialPC.NewIndex) = iCodigo
    Next

    'Seleciona a Filial na qual o usuário entrou no Sistema
    If giFilialEmpresa <> EMPRESA_TODA Then

       For iIndice = 0 To ComboFilialPC.ListCount - 1

            If ComboFilialPC.ItemData(iIndice) = giFilialEmpresa Then

                ComboFilialPC.ListIndex = iIndice
                Exit For

            End If

        Next

    Else

        ComboFilialPC.ListIndex = 0

    End If

    Carrega_FilialPC = SUCESSO

    Exit Function

Erro_Carrega_FilialPC:

    Carrega_FilialPC = gErr

    Select Case gErr

        'Erro já tratado
        Case 48905

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156114)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    iFrameAtual = 1
    iFornecedorAlterado = 0
    iEmissaoAlterada = 0
    iValorTotalAlterado = 0
    iValorIRRFAlterado = 0
    iValorINSSAlterado = 0
    iPISRetidoAlterado = 0
    iISSRetidoAlterado = 0
    iCOFINSRetidoAlterado = 0
    iCSLLRetidoAlterado = 0
    
    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoCondPagto = New AdmEvento
    Set objEventoFilialPC = New AdmEvento
    Set objEventoTipo = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    Set objEventoCcl = New AdmEvento
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjAnexos = New ClassAnexos
    
    lErro = CF2(Me, "NFFPAG_Form_Load")
    If lErro <> SUCESSO Then gError 132223
    
    'Carrega na combo os tipos de cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then gError 18363
    
    'Carrega combo tipos de documento
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then gError 57844
    
    'preenche a combo filialPc
    lErro = Carrega_FilialPC()
    If lErro <> SUCESSO Then gError 48914
    
    'Carrega na combo as Condições de Pagamento
        lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 18365

    'Carrega na combo de Banco de Cobradores
    lErro = Carrega_Bancos()
    If lErro <> SUCESSO Then gError 57855
    
    'Carrega na combo de Portadores
    lErro = Carrega_Portadores()
    If lErro <> SUCESSO Then gError 57857
    
    Set objGridParcelas = New AdmGrid

    'Inicializa Grid Parcelas
    lErro = CF2(Me, "NFFPAG_Inicializa_Grid_Parcelas", objGridParcelas)
    If lErro <> SUCESSO Then gError 18367
    
    'Carrega Motivos de diferença
    lErro = Carrega_MotivoDiferenca()
    If lErro <> SUCESSO Then gError 500060

    'Inicializa Máscara de Ccl
    sMascaraCcl = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then gError 18367

    Ccl.Mask = sMascaraCcl
    
    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 18362
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then gError 18362
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 18362, 18363, 18365, 18367, 48914, 57844, 57855, 57857, 132223, 500060 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156115)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colSigla As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosPag"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 18360

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", CLng(0), 0, "NumIntDoc"
    colCampoValor.Add "Fornecedor", objTituloPagar.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objTituloPagar.iFilial, 0, "Filial"
    colCampoValor.Add "NumTitulo", objTituloPagar.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objTituloPagar.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "NumParcelas", objTituloPagar.iNumParcelas, 0, "NumParcelas"
    colCampoValor.Add "ValorTotal", objTituloPagar.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "ValorSeguro", objTituloPagar.dValorSeguro, 0, "ValorSeguro"
    colCampoValor.Add "ValorFrete", objTituloPagar.dValorFrete, 0, "ValorFrete"
    colCampoValor.Add "OutrasDespesas", objTituloPagar.dOutrasDespesas, 0, "OutrasDespesas"
    colCampoValor.Add "ValorProdutos", objTituloPagar.dValorProdutos, 0, "ValorProdutos"
    colCampoValor.Add "ValorICMS", objTituloPagar.dValorICMS, 0, "ValorICMS"
    colCampoValor.Add "ValorICMSSubst", objTituloPagar.dValorICMS, 0, "ValorICMSSubst"
    colCampoValor.Add "CreditoICMS", objTituloPagar.iCreditoICMS, 0, "CreditoICMS"
    colCampoValor.Add "ValorIPI", objTituloPagar.dValorIPI, 0, "ValorIPI"
    colCampoValor.Add "CreditoIPI", objTituloPagar.iCreditoIPI, 0, "CreditoIPI"
    colCampoValor.Add "ValorIRRF", objTituloPagar.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "ValorINSS", objTituloPagar.dValorINSS, 0, "ValorINSS"
    colCampoValor.Add "INSSRetido", objTituloPagar.iINSSRetido, 0, "INSSRetido"
    colCampoValor.Add "FilialPedCompra", objTituloPagar.iFilialPedCompra, 0, "FilialPedCompra"
    colCampoValor.Add "NumPedCompra", objTituloPagar.lNumPedCompra, 0, "NumPedCompra"
    colCampoValor.Add "CondicaoPagto", objTituloPagar.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "SiglaDocumento", objTituloPagar.sSiglaDocumento, STRING_SIGLA_DOCUMENTO, "SiglaDocumento"
    colCampoValor.Add "PISRetido", objTituloPagar.dPISRetido, 0, "PISRetido"
    colCampoValor.Add "ISSRetido", objTituloPagar.dISSRetido, 0, "ISSRetido"
    colCampoValor.Add "COFINSRetido", objTituloPagar.dCOFINSRetido, 0, "COFINSRetido"
    colCampoValor.Add "CSLLRetido", objTituloPagar.dCSLLRetido, 0, "CSLLRetido"
    
    lErro = CF("TiposDocumento_Le_NaoOutrosPag", colSigla)
    If lErro <> SUCESSO Then gError 57854

    If colSigla.Count = 0 Then Exit Sub

    For iIndice = 1 To colSigla.Count
        If colSigla(iIndice) <> "NFFP" Then
            colSelecao.Add "SiglaDocumento", OP_DIFERENTE, colSigla(iIndice)
        End If
    Next
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 18360, 57854

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156116)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_Tela_Preenche

    objTituloPagar.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objTituloPagar.lNumIntDoc <> 0 Then

        'Carrega objTituloPagar com os dados passados em colCampoValor
        objTituloPagar.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objTituloPagar.iFilial = colCampoValor.Item("Filial").vValor
        objTituloPagar.sSiglaDocumento = colCampoValor.Item("SiglaDocumento").vValor
        objTituloPagar.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objTituloPagar.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objTituloPagar.iNumParcelas = colCampoValor.Item("NumParcelas").vValor
        objTituloPagar.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objTituloPagar.dValorSeguro = colCampoValor.Item("ValorSeguro").vValor
        objTituloPagar.dValorFrete = colCampoValor.Item("ValorFrete").vValor
        objTituloPagar.dOutrasDespesas = colCampoValor.Item("OutrasDespesas").vValor
        objTituloPagar.dValorProdutos = colCampoValor.Item("ValorProdutos").vValor
        objTituloPagar.dValorICMS = colCampoValor.Item("ValorICMS").vValor
        objTituloPagar.iCreditoICMS = colCampoValor.Item("CreditoICMS").vValor
        objTituloPagar.dValorICMSSubst = colCampoValor.Item("ValorICMSSubst").vValor
        objTituloPagar.dValorIPI = colCampoValor.Item("ValorIPI").vValor
        objTituloPagar.iCreditoIPI = colCampoValor.Item("CreditoIPI").vValor
        objTituloPagar.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        objTituloPagar.dValorINSS = colCampoValor.Item("ValorINSS").vValor
        objTituloPagar.iINSSRetido = colCampoValor.Item("INSSRetido").vValor
        objTituloPagar.iFilialPedCompra = colCampoValor.Item("FilialPedCompra").vValor
        objTituloPagar.lNumPedCompra = colCampoValor.Item("NumPedCompra").vValor
        objTituloPagar.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        objTituloPagar.dPISRetido = colCampoValor.Item("PISRetido").vValor
        objTituloPagar.dISSRetido = colCampoValor.Item("ISSRetido").vValor
        objTituloPagar.dCOFINSRetido = colCampoValor.Item("COFINSRetido").vValor
        objTituloPagar.dCSLLRetido = colCampoValor.Item("CSLLRetido").vValor
                
        'Traz a Nota para Tela
        lErro = Traz_NFFatPag_Tela(objTituloPagar)
        If lErro <> SUCESSO Then gError 18361

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 18361

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156117)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text

    'Chama Tela FornecedorLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1

    'Preenche campo Fornecedor
    Fornecedor.Text = objFornecedor.sNomeReduzido

    Call Fornecedor_Validate(Cancel)

    Me.Show

    Exit Sub

End Sub

Public Sub NumeroLabel_Click()

Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click

    'Se Forncedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 18389

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then gError 18390

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 57847
    
    'Move os dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 18391
    
    'Adiciona filtros: lFornecedor e iFilial e SiglaDocumento
    colSelecao.Add objTituloPagar.lFornecedor
    colSelecao.Add objTituloPagar.iFilial
    colSelecao.Add objTituloPagar.sSiglaDocumento
    
    'Chama Tela NFFatPagLista
    Call Chama_Tela("NFFatPagLista", colSelecao, objTituloPagar, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 18389
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 18390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 18391

        Case 57847
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156118)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloPagar = obj1

    'Traz os dados de objTituloPagar para Teal
    lErro = Traz_NFFatPag_Tela(objTituloPagar)
    If lErro <> SUCESSO Then gError 18359

'    'Fecha o comando das setas se estiver aberto
'    lErro = ComandoSeta_Fechar(Me.Name)
'
'    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 18359

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156119)

    End Select

    Exit Sub


End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

    'Chama a Tela CondicoesPagamentoCPLista
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Preenche campo CondicaoPagamento
    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo)
    Call CondicaoPagamento_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case 18498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156120)

     End Select

     Exit Sub

End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim objTipoDocumento As New ClassTipoDocumento

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloPagar Is Nothing) Then
    
        lErro = CF2(Me, "NFFPAG_Trata_Parametros", objTituloPagar)
        If lErro <> SUCESSO Then gError 132235

         objTipoDocumento.sSigla = objTituloPagar.sSiglaDocumento
        
        'Valida a Sigla passada
        lErro = CF("TipoDocumento_Le", objTipoDocumento)
        If lErro <> SUCESSO Then gError 57848
             
        If objTipoDocumento.iEmNFFatPag <> 1 Then gError 57849
        
        'Traz os dados para a Tela
        lErro = Traz_NFFatPag_Tela(objTituloPagar)
        If lErro <> SUCESSO Then gError 18374

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 18368, 18374, 57848, 132235

        Case 18384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_NAO_CADASTRADA", gErr, objTituloPagar.lNumIntDoc)

        Case 57849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_OUTROSPAG", gErr, objTituloPagar.sSiglaDocumento, objTituloPagar.lNumTitulo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156121)

    End Select

    Exit Function

End Function

Private Function Traz_NFFatPag_Tela(objTituloPagar As ClassTituloPagar) As Long
'Traz os dados da Nota Fiscal Fatura para a Tela

Dim lErro As Long
Dim colParcelasPag As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim iLinha As Integer
Dim iIndice As Integer, Cancel As Boolean
Dim bCancel As Boolean
Dim objNFiscal As New ClassNFiscal 'Inserido por Wagner
Dim dValorTotalParc As Double
Dim dValorTotalOrig As Double
Dim dDiferenca As Double
Dim sNaturezaEnxuta As String, sCclMascarado As String, sTela As String

On Error GoTo Erro_Traz_NFFatPag_Tela
    
    Call Limpa_Tela_NFFatPag
    
    CobrancaAutomatica.Value = vbUnchecked
'    lErro = CF("TituloPagar_Le", objTituloPagar)
'    If lErro <> SUCESSO And lErro <> 18372 Then gError 32303
'    If lErro <> SUCESSO Then gError 32304
    'Se le apenas titulos em aberto ou le todos (abertos + baixados)
    lErro = CF2(Me, "NFFPAG_Trata_Parametros", objTituloPagar)
    If lErro <> SUCESSO Then gError 32303
    
    'Coloca os dados do Título Pagar na Tela
    NumTitulo.Text = objTituloPagar.lNumTitulo

    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(Cancel)

    Filial.Text = objTituloPagar.iFilial
    Filial_Validate (bCancel)
    
    '#############################################################
    'Inserido por Wagner
    objNFiscal.lFornecedor = objTituloPagar.lFornecedor
    objNFiscal.iFilialForn = objTituloPagar.iFilial
    
    lErro = CF2(Me, "NFCF_Fornecedor_Traz_NF_Tela", objNFiscal)
    If lErro <> SUCESSO Then gError 131954
    '#############################################################
    
    'Preenche o Tipo
    Tipo.Text = objTituloPagar.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)
    
    If objTituloPagar.iFilialPedCompra <> 0 Then
        ComboFilialPC.Text = objTituloPagar.iFilialPedCompra
        Call ComboFilialPC_Validate(bSGECancelDummy)
    Else
        ComboFilialPC.Text = ""
    End If
    
    If objTituloPagar.lNumPedCompra <> 0 Then
        NumPC.PromptInclude = False
        NumPC.Text = objTituloPagar.lNumPedCompra
        NumPC.PromptInclude = True
    Else
        NumPC.PromptInclude = False
        NumPC.Text = ""
        NumPC.PromptInclude = True
    End If
    
    Call DateParaMasked(DataEmissao, objTituloPagar.dtDataEmissao)

    ValorTotal.Text = objTituloPagar.dValorTotal
    ValorICMS.Text = objTituloPagar.dValorICMS
    ValorICMSSubst.Text = objTituloPagar.dValorICMSSubst
    CreditoICMS.Value = objTituloPagar.iCreditoICMS
    ValorProdutos.Text = objTituloPagar.dValorProdutos
    ValorFrete.Text = objTituloPagar.dValorFrete
    ValorSeguro.Text = objTituloPagar.dValorSeguro
    OutrasDespesas.Text = objTituloPagar.dOutrasDespesas
    ValorIPI.Text = objTituloPagar.dValorIPI
    CreditoIPI.Value = objTituloPagar.iCreditoIPI
    ValorIRRF.Text = objTituloPagar.dValorIRRF
    ValorINSS.Text = objTituloPagar.dValorINSS
    INSSRetido.Value = objTituloPagar.iINSSRetido
    PISRetido.Text = objTituloPagar.dPISRetido
    ISSRetido.Text = objTituloPagar.dISSRetido
    COFINSRetido.Text = objTituloPagar.dCOFINSRetido
    CSLLRetido.Text = objTituloPagar.dCSLLRetido
    Historico.Text = objTituloPagar.sHistorico
    
    If Len(Trim(objTituloPagar.sCcl)) <> 0 Then
    
        sCclMascarado = String(STRING_CCL, 0)
        
        lErro = Mascara_RetornaCclEnxuta(objTituloPagar.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 18377
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
    Else
        
        Ccl.PromptInclude = False
        Ccl.Text = ""
        Ccl.PromptInclude = True
    
    End If
    
    If objTituloPagar.iCondicaoPagto <> 0 Then
    
        CondicaoPagamento.Text = CStr(objTituloPagar.iCondicaoPagto)
        Call CondicaoPagamento_Validate(bCancel)
    
    Else
    
        CondicaoPagamento.Text = ""
    
    End If
    
    'Lê as Parcelas a Pagar vinculadas ao Título
    'essa leitura vai retornar codigo de erro e nao mensagem (3o parametro = false)
    lErro = CF("ParcelasPagar_Le_Todas", objTituloPagar, colParcelasPag)
    If lErro <> SUCESSO Then gError 18377
    
    If colParcelasPag.Count > NUM_MAXIMO_PARCELAS Then gError 18674

    Call Grid_Limpa(objGridParcelas)
    
    iLinha = 0

    Set objTituloPagar.colParcelaPagar = colParcelasPag

    'Preenche as linhas do Grid Parcelas com os dados de cada Parcela
    For Each objParcelaPagar In colParcelasPag

        iLinha = iLinha + 1

        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objParcelaPagar.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_VenctoReal_Col) = Format(objParcelaPagar.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objParcelaPagar.dValor, "Standard")

        GridParcelas.TextMatrix(iLinha, iGrid_ValorOriginal_Col) = Format(objParcelaPagar.dValorOriginal, "Standard")
        
        'Se existe um Motivo de diferença
        For iIndice = 0 To MotivoDiferenca.ListCount - 1
            If MotivoDiferenca.ItemData(iIndice) = objParcelaPagar.iMotivoDiferenca Then
                GridParcelas.TextMatrix(iLinha, iGrid_MotivoDiferenca_Col) = MotivoDiferenca.List(iIndice)
                Exit For
            End If
        Next
              
        'Acumula total de Valores e Valores Originais
        dValorTotalParc = dValorTotalParc + objParcelaPagar.dValor
        dValorTotalOrig = dValorTotalOrig + objParcelaPagar.dValorOriginal

        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = objParcelaPagar.iTipoCobranca Then
                GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next

        For iIndice = 0 To comboCobrador.ListCount - 1
            If comboCobrador.ItemData(iIndice) = objParcelaPagar.iBancoCobrador Then
                GridParcelas.TextMatrix(iLinha, iGrid_Banco_Col) = comboCobrador.List(iIndice)
                Exit For
            End If
        Next

        For iIndice = 0 To comboPortador.ListCount - 1
            If comboPortador.ItemData(iIndice) = objParcelaPagar.iPortador Then
                GridParcelas.TextMatrix(iLinha, iGrid_Portador_Col) = comboPortador.List(iIndice)
                Exit For
            End If
        Next

        If objParcelaPagar.iStatus = STATUS_SUSPENSO Then
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "1"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "0"
        End If

        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = objParcelaPagar.sCodigoDeBarras
        CodigodeBarras.PromptInclude = True
        GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col) = CodigodeBarras.Text
        
    Next

    'Faz o número de linhas existentes do Grid ser igual ao número de Parcelas
    objGridParcelas.iLinhasExistentes = iLinha

    'Diferença = Soma de Valores - Soma de Valores Originais
    Diferenca.Text = Format(dValorTotalParc - dValorTotalOrig, "Standard")

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objTituloPagar.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 36364

    lErro = CF2(Me, "NFFPAG_Traz_NFFatPag_Tela", objTituloPagar)
    If lErro <> SUCESSO Then gError 32302
    
    'Faz refresh nas checkboxes
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objTituloPagar.objProjetoInfo, objTituloPagar.lNumIntDoc, PRJ_CR_TIPO_TITPAG)
    If lErro <> SUCESSO Then gError 181676
    
    Set gobjAnexos = objTituloPagar.objAnexos
    If gobjAnexos Is Nothing Then
        Set gobjAnexos = New ClassAnexos
        Set objNFiscal = New ClassNFiscal
        lErro = CF("Titulo_Le_DocumentoOriginal", objTituloPagar.lNumIntDoc, CPR_TITULO_PAGAR, objNFiscal, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError 181676
        If lErro = SUCESSO Then
            gobjAnexos.iTipoDoc = ANEXO_TIPO_NF
            gobjAnexos.lNumIntDoc = objNFiscal.lNumIntDoc
        Else
            gobjAnexos.iTipoDoc = ANEXO_TIPO_TITPAG
            gobjAnexos.lNumIntDoc = objTituloPagar.lNumIntDoc
        End If
    End If
    lErro = CF("Anexos_Le", gobjAnexos)
    If lErro <> SUCESSO Then gError 181676
    
    '??? mostrar que há (ou nao) anexos
        
    If Len(Trim(objTituloPagar.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objTituloPagar.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 181676
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
    Else
    
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
        
    End If
    
    Call Natureza_Validate(bSGECancelDummy)
     
    iAlterado = 0
    iFornecedorAlterado = 0
    iEmissaoAlterada = 0
    iValorTotalAlterado = 0
    iValorIRRFAlterado = 0
    iValorINSSAlterado = 0
    iPISRetidoAlterado = 0
    iISSRetidoAlterado = 0
    iCOFINSRetidoAlterado = 0
    iCSLLRetidoAlterado = 0
    
    Traz_NFFatPag_Tela = SUCESSO

    Exit Function

Erro_Traz_NFFatPag_Tela:

    Traz_NFFatPag_Tela = gErr

    Select Case gErr

        Case 18377, 36364, 32302, 32303, 181676 'Tratados nas Rotinas Chamadas
        
        Case 131954 'Inserido por Wagner

        Case 18674
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", gErr, colParcelasPag.Count, NUM_MAXIMO_PARCELAS)

        Case 32304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_NAO_CADASTRADA", gErr, objTituloPagar.lNumIntDoc)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156122)

    End Select

    Exit Function

End Function

Public Sub Fornecedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = 1

    Call Fornecedor_Preenche

End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean, sNaturezaEnxuta As String

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 18385

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 18386

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

            'CODIGO ESPECÍFICO
            'Verifica se chave de TipoFornecedor está preenchida
            If objFornecedor.iTipo > 0 Then

                'Coloca os dados do Tipo na tela
                lErro = TipoFornecedor_Dados(objFornecedor, objTipoFornecedor)
                If lErro <> SUCESSO Then gError 18387

            End If

            'Se Cond Pagto de Fornecedor está preenchida
            If objFornecedor.iCondicaoPagto <> 0 Then

                objCondicaoPagto.iCodigo = objFornecedor.iCondicaoPagto

                'Lê a Cond Pagto
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 18588
                If lErro = 19205 Then gError 18675

                If objCondicaoPagto.iEmPagamento = 1 Then
                    'Coloca na Tela
                    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
                    Call CondicaoPagamento_Validate(bCancel)
                End If
                
            'Se Cond Pagto de TipoFornecedor está preenchida
            ElseIf objTipoFornecedor.iCondicaoPagto <> 0 Then

                objCondicaoPagto.iCodigo = objTipoFornecedor.iCondicaoPagto

                'Lê a Cond Pagto
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 18589
                If lErro = 19205 Then gError 18676

                If objCondicaoPagto.iEmPagamento = 1 Then
                    'Coloca na Tela
                    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida
                    Call CondicaoPagamento_Validate(bCancel)
                End If

            End If
            'FIM CODIGO ESPECÍFICO

            If Len(Trim(objFornecedor.sNatureza)) <> 0 Then
                
                sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
            
                lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objFornecedor.sNatureza, sNaturezaEnxuta)
                If lErro <> SUCESSO Then gError 18589
            
                Natureza.PromptInclude = False
                Natureza.Text = sNaturezaEnxuta
                Natureza.PromptInclude = True
            
            Else

                Natureza.PromptInclude = False
                Natureza.Text = ""
                Natureza.PromptInclude = True
    
            End If

            Call Natureza_Validate(bSGECancelDummy)
            
            '########################################################
            'Inserido por Wagner
            lErro = CF2(Me, "NFCF_Fornecedor_Validate", objFornecedor.lCodigo, iCodFilial)
            If lErro <> SUCESSO Then gError 131956
            '########################################################

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            'Limpa Combo de Filial
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

        Case 18385, 18386, 18387, 18588, 18589 'Tratados nas Rotinas chamadas

        Case 18675, 18676
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case 131956 'Inserido por Wagner
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156123)

    End Select

    Exit Sub

End Sub

Public Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumTitulo.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumTitulo.ClipText)
    If lErro <> SUCESSO Then gError 18398

    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True


    Select Case gErr

        Case 18398

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156124)

    End Select

    Exit Sub

End Sub

Public Sub TipoCobranca_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub comboPortador_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub comboCobrador_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) <> 0 Then

        'Critica se é valor positivo
        lErro = Valor_Positivo_Critica(ValorTotal.Text)
        If lErro <> SUCESSO Then gError 18407
    
        'Põe o valor formatado na tela
        ValorTotal.Text = Format(ValorTotal.Text, "Standard")

    End If

    If iValorTotalAlterado = 1 Then
            
        'Recalcula a Cobrança e preenche o Grid
        Call Recalcula_Cobranca
        iValorTotalAlterado = 0
        
    End If
    
    Exit Sub

Erro_ValorTotal_Validate:

    Cancel = True


    Select Case gErr

        Case 18407

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156125)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick
    
    'Diminui "1" da data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then gError 18400

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 18400

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156126)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick
    
    'Adiciona "1" a data
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then gError 18401

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 18401

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156127)

    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagamento_Click()

Dim lErro As Long

On Error GoTo Erro_CondicaoPagamento_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Recalcula a Cobranca e Preenche o Grid
    Call Recalcula_Cobranca
    
    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156128)

      End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValorPagar As Double
Dim dtDataEmissao As Date
Dim dtDataVenctoReal As Date
Dim dValorIRRF As Double, dPISRetido As Double, dISSRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
Dim dValorINSSRetido As Double
Dim iIndice As Integer

On Error GoTo Erro_GridParcelas_Preenche
    
    Call Grid_Limpa(objGridParcelas)
            
    'Obtem Valor a Pagar
    If Len(Trim(ValorTotal.Text)) > 0 Then
    
        dValorINSSRetido = StrParaDbl(IIf(INSSRetido.Value, ValorINSS.Text, 0))
        dValorIRRF = StrParaDbl(ValorIRRF)
        If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
        If Len(Trim(ISSRetido.Text)) <> 0 And IsNumeric(ISSRetido.Text) Then dISSRetido = CDbl(ISSRetido.Text)
        If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
        If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
        
        dValorPagar = CDbl(ValorTotal) - (dValorIRRF + dPISRetido + dISSRetido + dCOFINSRetido + dCSLLRetido) - dValorINSSRetido

    End If

    'Se Valor a Pagar for positivo
    If dValorPagar > 0 Then

        objCondicaoPagto.dValorTotal = dValorPagar
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then gError 18417

        'Número de Parcelas
        objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
            GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        dtDataEmissao = CDate(DataEmissao.Text)

        objCondicaoPagto.dtDataEmissao = dtDataEmissao
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 18441

        'Número de Parcelas
        objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then gError 18443

            'Coloca Data de Vencimento Real no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        Next

    End If

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 18417, 18441, 18443, 57846

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156129)

    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer
Dim ColRateioOn As New Collection

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 36242

        If objGridInt.objGrid Is GridParcelas Then

            Select Case objGridInt.objGrid.Col

                Case iGrid_Vencimento_Col
                    lErro = Saida_Celula_Vencimento(objGridInt)
                    If lErro <> SUCESSO Then gError 18422

                Case iGrid_ValorParcela_Col
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError 18424

                Case iGrid_Cobranca_Col
                    lErro = Saida_Celula_Cobranca(objGridInt)
                    If lErro <> SUCESSO Then gError 18440

                Case iGrid_Banco_Col
                    lErro = Saida_Celula_Banco(objGridInt)
                    If lErro <> SUCESSO Then gError 57860
            
                Case iGrid_Portador_Col
                    lErro = Saida_Celula_Portador(objGridInt)
                    If lErro <> SUCESSO Then gError 57861
            
                Case iGrid_CodigoDeBarras_Col
                    lErro = Saida_Celula_CodigodeBarras(objGridInt)
                    If lErro <> SUCESSO Then gError 57861
                    
                Case iGrid_ValorOriginal_Col
                    lErro = Saida_Celula_ValorOriginal(objGridInt)
                    If lErro <> SUCESSO Then gError 500038
                
                'Motivo Diferença
                Case iGrid_MotivoDiferenca_Col
                    lErro = Saida_Celula_MotivoDiferenca(objGridInt)
                    If lErro <> SUCESSO Then gError 500039
                
            End Select

        End If
        
        lErro = CF2(Me, "NFFPAG_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 57861

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 18591

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 18422, 18424, 18440, 18512, 36242, 57860, 57861, 500038, 500039

        Case 18591
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(Trim(ValorParcela.Text)) > 0 Then
    
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 18426

        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then CobrancaAutomatica.Value = vbUnchecked
        
        If GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) = ValorParcela.Text

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            If GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = "" Then
            
                For iIndice = 0 To TipoCobranca.ListCount - 1
                    If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                        Exit For
                    End If
                Next
        
            End If
        
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 18427

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 18426, 18427
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156130)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vencimento(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim dtDataVenctoReal As Date
Dim sDataVencimento As String

On Error GoTo Erro_Saida_Celula_Vencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then gError 18420

        dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            dtDataEmissao = CDate(DataEmissao.Text)
            If dtDataVencimento < DataEmissao Then gError 18421
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
        
        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then gError 18596

        'Coloca data de Vencimento Real no Grid
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            If GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = "" Then
            
                For iIndice = 0 To TipoCobranca.ListCount - 1
                    If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                        Exit For
                    End If
                Next
        
            End If
        
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 18425

    Saida_Celula_Vencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_Vencimento:

    Saida_Celula_Vencimento = gErr

    Select Case gErr

        Case 18420, 18425, 18596
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18421
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", gErr, DataVencimento.Text, GridParcelas.Row, DataEmissao.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156131)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobranca(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim objTipoCobranca

On Error GoTo Erro_Saida_Celula_Cobranca

    Set objGridInt.objControle = TipoCobranca

    'Verifica se o Tipo de Cobrança foi preenchido
    If Len(Trim(TipoCobranca.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoCobranca.Text <> TipoCobranca.List(TipoCobranca.ListIndex) Then

            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(TipoCobranca, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18426

            If lErro = 6730 Then gError 18677
            If lErro = 6731 Then gError 18678

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 18425

    Saida_Celula_Cobranca = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobranca:

    Saida_Celula_Cobranca = gErr

    Select Case gErr

        Case 18425
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18677
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_CADASTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18678
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", gErr, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18426
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", gErr, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156132)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Banco(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Banco

    Set objGridInt.objControle = comboCobrador

    'Verifica se o Banco foi preenchido
    If Len(Trim(comboCobrador.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If comboCobrador.Text <> comboCobrador.List(comboCobrador.ListIndex) Then

            'Seleciona o Banco
            lErro = Combo_Seleciona(comboCobrador, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 57864

            If lErro = 6730 Then gError 57862
            If lErro = 6731 Then gError 57863

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 57865

    Saida_Celula_Banco = SUCESSO

    Exit Function

Erro_Saida_Celula_Banco:

    Saida_Celula_Banco = gErr

    Select Case gErr

        Case 57862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_NAO_CADASTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_INEXISTENTE1", gErr, comboCobrador.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57864, 57865
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156133)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Portador(objGridInt As AdmGrid) As Long
    
Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Portador

    Set objGridInt.objControle = comboPortador

    'Verifica se o Portador foi preenchido
    If Len(Trim(comboPortador.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If comboPortador.Text <> comboPortador.List(comboPortador.ListIndex) Then

            'Seleciona o Portador
            lErro = Combo_Seleciona(comboPortador, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 57868

            If lErro = 6730 Then gError 57866
            If lErro = 6731 Then gError 57867

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 57869

    Saida_Celula_Portador = SUCESSO

    Exit Function

Erro_Saida_Celula_Portador:

    Saida_Celula_Portador = gErr

    Select Case gErr

        Case 57866
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO1", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO2", gErr, comboPortador.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 57868, 57869
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156134)

    End Select

    Exit Function

End Function

Public Sub ValorICMS_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorICMS_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMS.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_NaoNegativo_Critica(ValorICMS.Text)
    If lErro <> SUCESSO Then gError 18403

    'Põe o valor formatado na tela
    ValorICMS.Text = Format(ValorICMS.Text, "Fixed")

    Exit Sub

Erro_ValorICMS_Validate:

    Cancel = True


    Select Case gErr

        Case 18403

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156135)

    End Select

    Exit Sub

End Sub

Public Sub ValorICMSSubst_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorICMSSubst_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorICMSSubst.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorICMSSubst.Text)
    If lErro <> SUCESSO Then gError 18408

    'Põe o valor formatado na tela
    ValorICMSSubst.Text = Format(ValorICMSSubst.Text, "Fixed")

    Exit Sub

Erro_ValorICMSSubst_Validate:

    Cancel = True


    Select Case gErr

        Case 18408

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156136)

    End Select

    Exit Sub

End Sub

Public Sub ValorProdutos_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorProdutos_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorProdutos.ClipText)) = 0 Then Exit Sub

    'Critica se é valor positivo
    lErro = Valor_Positivo_Critica(ValorProdutos.Text)
    If lErro <> SUCESSO Then gError 18406

    'Põe o valor formatado na tela
    ValorProdutos.Text = Format(ValorProdutos.Text, "Fixed")

    Exit Sub

Erro_ValorProdutos_Validate:

    Cancel = True


    Select Case gErr

        Case 18406

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156137)

    End Select

    Exit Sub

End Sub

Public Sub ValorFrete_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorFrete_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorFrete.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorFrete.Text)
    If lErro <> SUCESSO Then gError 18402

    'Põe o valor formatado na tela
    ValorFrete.Text = Format(ValorFrete.Text, "Fixed")

    Exit Sub

Erro_ValorFrete_Validate:

    Cancel = True


    Select Case gErr

        Case 18402

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156138)

    End Select

    Exit Sub

End Sub

Public Sub ValorSeguro_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valorseguro_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorSeguro.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorSeguro.Text)
    If lErro <> SUCESSO Then gError 18592

    'Põe o valor formatado na tela
    ValorSeguro.Text = Format(ValorSeguro.Text, "Fixed")

    Exit Sub

Erro_Valorseguro_Validate:

    Cancel = True


    Select Case gErr

        Case 18592

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156139)

    End Select

    Exit Sub

End Sub

Public Sub OutrasDespesas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_OutrasDespesas_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(OutrasDespesas.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(OutrasDespesas.Text)
    If lErro <> SUCESSO Then gError 18399

    'Põe o valor formatado na tela
    OutrasDespesas.Text = Format(OutrasDespesas.Text, "Fixed")

    Exit Sub

Erro_OutrasDespesas_Validate:

    Cancel = True


    Select Case gErr

        Case 18399

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156140)

    End Select

    Exit Sub

End Sub

Public Sub ValorIPI_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIPI_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIPI.ClipText)) = 0 Then Exit Sub

    'Critica se é valor não negativo
    lErro = Valor_NaoNegativo_Critica(ValorIPI.Text)
    If lErro <> SUCESSO Then gError 18404

    'Põe o valor formatado na tela
    ValorIPI.Text = Format(ValorIPI.Text, "Fixed")

    Exit Sub

Erro_ValorIPI_Validate:

    Cancel = True


    Select Case gErr

        Case 18404

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156141)

    End Select

    Exit Sub

End Sub

Public Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorIRRF_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIRRF.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
        If lErro <> SUCESSO Then gError 18405
    
        'Põe o valor formatado na tela
        ValorIRRF.Text = Format(ValorIRRF.Text, "Fixed")

    End If
    
    If iValorIRRFAlterado <> 0 Then
    
        Call Recalcula_Cobranca
        iValorIRRFAlterado = 0
        
    End If
    
    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True


    Select Case gErr

        Case 18405

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156142)

    End Select

    Exit Sub

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
    
        If INSSRetido.Value = 1 Then Call Recalcula_Cobranca
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

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 18448

    'Limpa a Tela
    Call Limpa_Tela_NFFatPag

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 18448

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156144)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'Verifica se principais campos da tela foram preenchidos para gravação

Dim lErro As Long
Dim iIndice As Integer
Dim dValorPagar As Double
Dim dValorProdutos As Double
Dim dValorICMSSubst As Double
Dim dValorFrete As Double
Dim dValorSeguro As Double
Dim dValorIPI As Double
Dim dOutrasDespesas As Double
Dim dValorIRRF As Double, dPISRetido As Double, dISSRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double
Dim dValor As Double
Dim dSoma As Double
Dim dSomaParcelas As Double
Dim dValorINSSRet As Double
Dim dtDataVencimento As Date
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim dSomaParcOriginal As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos obrigatórios estão preenchidos
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 18449
    If Len(Trim(Filial.Text)) = 0 Then gError 18450
    If Len(Trim(NumTitulo.ClipText)) = 0 Then gError 18451
    If Len(Trim(ValorProdutos.ClipText)) = 0 Then gError 18452
    If Len(Trim(ValorTotal.ClipText)) = 0 Then gError 18453
    If Len(Trim(Tipo.Text)) = 0 Then gError 57852
    
    'Conversão dos valores para Double
    If Len(Trim(ValorProdutos.Text)) > 0 Then dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(ValorICMSSubst.Text)) > 0 Then dValorICMSSubst = CDbl(ValorICMSSubst)
    If Len(Trim(ValorFrete.Text)) > 0 Then dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorSeguro.Text)) > 0 Then dValorSeguro = CDbl(ValorSeguro.Text)
    If Len(Trim(OutrasDespesas.Text)) > 0 Then dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorIPI.Text)) > 0 Then dValorIPI = CDbl(ValorIPI.Text)

    'Soma dos valores
    dSoma = dValorProdutos + dValorICMSSubst + dValorFrete + dValorSeguro + dOutrasDespesas + dValorIPI

    'Verifica se Soma dos Valores é igual ao ValorTotal
    If Format(dSoma, "0.00") <> Format(CDbl(ValorTotal.Text), "0.00") Then gError 18454

    'Verifica se GridParcelas foi preenchido
    If objGridParcelas.iLinhasExistentes = 0 Then gError 18455

    dSomaParcelas = 0
    dSomaParcOriginal = 0

    'Loop no GridParcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        'Verifica se DataVencimento foi preenchida
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 18456

        dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))

        'Verifica se DataVencimento é maior ou igual a DataEmissao
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            If dtDataVencimento < CDate(DataEmissao) Then gError 18458
        End If

        'Verifica a ordenação das Datas de Vencimento das Parcelas
        If iIndice > 1 Then
            If dtDataVencimento < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 18506
        End If

        'Verifica se Valor da Parcela foi preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 18457

        'Verifica se Valor Original da Parcela foi preenchido
        If StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col)) = 0 Then 'gError 500043
            GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col) = GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col)
        End If
                
        'Verifica se o Motivo da Diferença foi preenchido numa parcela com valor diferente Valor Original
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_MotivoDiferenca_Col))) = 0 And (StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col)) <> StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col))) Then gError 500044
        
        'Se Valor e Valor Original são iguais e Motivo Diferença foi preenchido, erro
        If StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col)) = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col)) And Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_MotivoDiferenca_Col))) > 0 Then gError 500063

        'Verifica se Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        If lErro <> SUCESSO Then gError 18679

        'Acumula Valor Parcela em dSomaParcelas
        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

        'Acumula o Valor das parcelas originais
        dSomaParcOriginal = dSomaParcOriginal + StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col))

    Next

    'Valor a Pagar
    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
    If Len(Trim(ValorINSS.Text)) > 0 And INSSRetido.Value = vbChecked Then dValorINSSRet = CDbl(ValorINSS.Text)
    If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(ISSRetido.Text)) <> 0 And IsNumeric(ISSRetido.Text) Then dISSRetido = CDbl(ISSRetido.Text)
    If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
    dValorPagar = CDbl(ValorTotal.Text) - (dValorIRRF + dPISRetido + dISSRetido + dCOFINSRetido + dCSLLRetido) - dValorINSSRet + StrParaDbl(Diferenca.Text)
    
    'Compara Soma das Parcelas com Valor a Pagar e o Valor de diferença
    If Format(dSomaParcelas, "0.00") <> Format(dValorPagar, "0.00") Then gError 18594
        
    'Compara a Soma dos Valores Originais com o Valor à Pagar
    If Format(dSomaParcOriginal, "0.00") <> Format(dValorPagar - StrParaDbl(Diferenca.Text), "0.00") Then gError 500045

    'Move dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 18459

    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    
    '??? validar anexos
    
    lErro = CF2(Me, "NFFPAG_Gravar_Registro", objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 132313

    'Grava Nota Fiscal Fatura no BD (incluindo dados contabeis (contabilidade))
    lErro = CF("NFFatPag_Grava", objTituloPagar, colParcelas, objContabil)
    If lErro <> SUCESSO Then gError 18460

    GL_objMDIForm.MousePointer = vbDefault
    
    lErro = CF2(Me, "NFFPAG_Gravar_Registro2", objTituloPagar, colParcelas, objContabil)
    If lErro <> SUCESSO Then gError 132313
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 18449
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 18450
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 18451
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)

        Case 18452
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPRODUTOS_NAO_INFORMADO", gErr)

        Case 18453
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", gErr)

        Case 18454
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_INVALIDO", gErr, ValorTotal.Text, dSoma)

        Case 18455
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", gErr)

        Case 18456
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", gErr, iIndice)

        Case 18457
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_NAO_INFORMADO", gErr, iIndice)

        Case 18458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", gErr, dtDataVencimento, iIndice, DataEmissao.Text)

        Case 18459, 18460, 18679, 20839, 181786, 132313 'Tratados nas Rotinas chamadas

        Case 18506
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_NAO_ORDENADA", gErr)

        Case 18594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_INVALIDA", gErr, dSomaParcelas, dValorPagar)

        Case 57852
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            
        Case 500043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORORIGINALPARC_NAO_INFORMADO", gErr, iIndice)
        
        Case 500044
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCAPARC_NAO_INFORMADO", gErr, iIndice)
        
        Case 500045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_VALORORIGINAL_DIFERENTE_VALORTOTAL", gErr, dSomaParcOriginal, dValorPagar - StrParaDbl(Diferenca.Text))
        
        Case 500063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCA_INFORMADO_ERRADO", gErr, iIndice)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156145)

    End Select

    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim vbMsgRes As VbMsgBoxResult
Dim colParcelas As New colParcelaPagar

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos identificadores estão preenchidos
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 18538
    If Len(Trim(Filial.Text)) = 0 Then gError 18539
    If Len(Trim(NumTitulo.Text)) = 0 Then gError 18540
    If Len(Trim(Tipo.Text)) = 0 Then gError 57853
    
    'Move dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 18575

    'Verifica se campos identificadores correspondem a Título Baixado
    lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then gError 18552

    'Se corresponder, erro
    If lErro = SUCESSO Then gError 18557

    'Lê Título Pagar a partir dos campos identificadores
    lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18551 Then gError 18547

    'Se não encontrou ---> ERRO
    If lErro <> SUCESSO Then gError 18558

    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_NFFATPAG", objTituloPagar.lNumTitulo)

    'Se não confirmar, sai
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui Nota Fiscal Fatura (incluindo dados contabeis (contabilidade))
    lErro = CF("NFFatPag_Exclui", objTituloPagar, objContabil)
    If lErro <> SUCESSO Then gError 18559

    'Limpa a Tela
    Call Limpa_Tela_NFFatPag

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 18538
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 18539
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 18540
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)

        Case 18547, 18552, 18559, 18575

        Case 18557
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_BAIXADA_EXCLUSAO", gErr, objTituloPagar.lNumTitulo)

        Case 18558
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_NAO_CADASTRADA1", gErr, objTituloPagar.lNumTitulo)

        Case 57853
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156146)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 18682

    'Limpa a Tela
    Call Limpa_Tela_NFFatPag

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 18682

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156147)

    End Select

    Exit Sub

End Sub

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim sTela As String
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoDocOriginal_Click
    
    'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número, DataEmissao)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 79381

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79382

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 79383
    
    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloPagar.lNumTitulo = CLng(NumTitulo.Text)
    objTituloPagar.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 79384

    'Se for do Tipo "OR" dá mensagem pois este Título não tem documento original
    If objTituloPagar.sSiglaDocumento = SIGLA_OUTROS_PAGAMENTOS Then gError 79385
    
    'Se for do Tipo Fatura a Pagar chama a Tela de Browser Listando todas as Notas da Fatura a Pagar que originou este Título
    If objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
        
        colSelecao.Add objTituloPagar.lNumIntDoc
        
        Call Chama_Tela("NF_FaturaReceberLista", colSelecao)
        
    Else
                
        'Le o Nome da Tela que originou este Título
        lErro = CF("Titulo_Le_DocumentoOriginal", objTituloPagar.lNumIntDoc, CPR_TITULO_PAGAR, objNFiscal, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError 79386
        
        If lErro = SUCESSO Then
'            If giTipoVersao = VERSAO_FULL Then
                'Chama a Tela
                Call Chama_Tela(sTela, objNFiscal)
'            ElseIf giTipoVersao = VERSAO_LIGHT Then
'                'Chama a Tela
'                Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
'            End If
        Else
            gError 79387
        End If
        
    End If
    
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 79381, 79382, 79384, 79386 'Tratado na Rotina chamada
        
        Case 79383
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 79385, 79387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAG_SEM_DOCORIGINAL", gErr, objTituloPagar.lNumTitulo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156148)

    End Select

    Exit Sub


End Sub

Private Function Critica_CamposPreenchidos() As Long
'Verifica se os campos chaves foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposPreenchidos

    'Se Fornecedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79377

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then gError 79378

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 79379

    'Verifica se o Titulo está Preenchido
    If Len(Trim(NumTitulo.Text)) = 0 Then gError 79380

    'verifica se a data de emissão foi preenchida
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79376

    Critica_CamposPreenchidos = SUCESSO

    Exit Function

Erro_Critica_CamposPreenchidos:

    Critica_CamposPreenchidos = gErr

    Select Case gErr

        Case 79377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 79378
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 79379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)

        Case 79380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)

        Case 79376
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156149)

    End Select

    Exit Function

End Function

Public Sub CreditoICMS_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CreditoIPI_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iEmissaoAlterada = 1
    
End Sub

Public Sub DataVencimento_Change()

    If Len(Trim(DataVencimento.ClipText)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub

Public Sub DataVencimento_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub DataVencimento_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub DataVencimento_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub DataVencimentoReal_Change()

    If Len(Trim(DataVencimentoReal.ClipText)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub

Public Sub Filial_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Call CF2(Me, "NFFPAG_Form_Unload", Cancel)
    Set gobjInfoUsu = Nothing
    
    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoFilialPC = Nothing
    Set objEventoTipo = Nothing
    
    Set objGridParcelas = Nothing

    Set gobjTelaProjetoInfo = Nothing
    
    Set gobjAnexos = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objEventoNatureza = Nothing
    Set objEventoCcl = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
     
End Sub

Public Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Public Sub GridParcelas_EnterCell()

    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Public Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

End Sub

Public Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)

End Sub

Public Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Public Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Public Sub NumTitulo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_PAGAR_ID
                
            Case TAB_Cobranca
                Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_COBRANCA
                        
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_CONTABILIZACAO
        
        End Select

    End If

End Sub

Private Function Carrega_Bancos() As Long
'Carrega a combo de Cobrador

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Bancos

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then gError 57856

   'Preenche ComboBox com código e nome dos Bancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        comboCobrador.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        comboCobrador.ItemData(comboCobrador.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_Bancos = SUCESSO

    Exit Function

Erro_Carrega_Bancos:

    Carrega_Bancos = gErr

    Select Case gErr

        Case 57856 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156150)

    End Select

    Exit Function

End Function

Private Function Carrega_Portadores() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_Portadores

    'Lê o código e a descrição de todos os Portadores Ativos
    lErro = CF("Portadores_Le_CodigosNomesRed", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 57858

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na Combo de Portadores
        comboPortador.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        comboPortador.ItemData(comboPortador.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_Portadores = SUCESSO

    Exit Function

Erro_Carrega_Portadores:

    Carrega_Portadores = gErr

    Select Case gErr

        Case 57858

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156151)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoCobranca() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 18364

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na List da Combo TipoCobranca
        TipoCobranca.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = gErr

    Select Case gErr

        Case 18364

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156152)

    End Select

    Exit Function

End Function

'Private Function Carrega_CondicaoPagamento() As Long
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 18366
'
'    For Each objCod_DescReduzida In colCod_DescReduzida
'
'        'Adiciona novo ítem na List da Combo CondicaoPagamento
'        CondicaoPagamento.AddItem CInt(objCod_DescReduzida.iCodigo) & SEPARADOR & objCod_DescReduzida.sNome
'        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCod_DescReduzida.iCodigo
'
'    Next
'
'    Carrega_CondicaoPagamento = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_CondicaoPagamento:
'
'    Carrega_CondicaoPagamento = gErr
'
'    Select Case gErr
'
'        Case 18366
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156153)
'
'    End Select
'
'    Exit Function
'
'End Function

Public Function NFFPAG_Inicializa_Grid_Parcelas(objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Suspenso")
    objGridInt.colColuna.Add ("Banco Cobrador")
    objGridInt.colColuna.Add ("Portador")
    objGridInt.colColuna.Add ("Valor Original")
    objGridInt.colColuna.Add ("Motivo da Diferença")
    objGridInt.colColuna.Add ("Código de Barras")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (TipoCobranca.Name)
    objGridInt.colCampo.Add (Suspenso.Name)
    objGridInt.colCampo.Add (comboCobrador.Name)
    objGridInt.colCampo.Add (comboPortador.Name)
    objGridInt.colCampo.Add (ValorOriginal.Name)
    objGridInt.colCampo.Add (MotivoDiferenca.Name)
    objGridInt.colCampo.Add (CodigodeBarras.Name)
        
    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_Col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_ValorParcela_Col = 3
    iGrid_Cobranca_Col = 4
    iGrid_Suspenso_Col = 5
    iGrid_Banco_Col = 6
    iGrid_Portador_Col = 7
    iGrid_ValorOriginal_Col = 8
    iGrid_MotivoDiferenca_Col = 9
    iGrid_CodigoDeBarras_Col = 10
    
    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8
    
    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 900
    
    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    NFFPAG_Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Public Sub LabelTipo_Click()

Dim colSelecao As Collection
Dim objTipoDocumento As New ClassTipoDocumento

    'Se Tipo estiver preenchida, extrai a sigla
    If Len(Trim(Tipo.Text)) > 0 Then
        objTipoDocumento.sSigla = Tipo.Text
    End If
    
    'Chama a tela Tipo
    Call Chama_Tela("TipoDocOutrosPagLista", colSelecao, objTipoDocumento, objEventoTipo)

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    Tipo.Text = objTipoDocumento.sSigla
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoTipo_evSelecao:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156154)
     
     End Select
     
     Exit Sub

End Sub

Public Sub Tipo_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tipo_Click()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub
    
    'Verifica se o Tipo foi selecionado
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Tenta Localizar o Tipo no Text da Combo
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then gError 57850

    'Se nao encontrar -> Erro
    If lErro = 60483 Then gError 57851
        
    Exit Sub
    
Erro_Tipo_Validate:

    Cancel = True


    Select Case gErr
    
        Case 57850
        
        Case 57851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", gErr, Tipo.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156155)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub INSSRetido_Click()

Dim lErro As Long

On Error GoTo Erro_INSSRetido_Click

    iAlterado = REGISTRO_ALTERADO
    
    'Recalcula as Parcelas do Grid
    Call Recalcula_Cobranca
    
    Exit Sub
    
Erro_INSSRetido_Click:
    
    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156156)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub OutrasDespesas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Suspenso_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Suspenso_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Suspenso_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Suspenso_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Suspenso
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub TipoCobranca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub comboCobrador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub comboPortador_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoCobranca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub comboCobrador_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub comboPortador_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub TipoCobranca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub comboCobrador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub comboPortador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub TipoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoCobranca
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub comboCobrador_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = comboCobrador
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub comboPortador_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = comboPortador
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorFrete_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorICMS_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorICMSSubst_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorINSS_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorINSSAlterado = 1
    
End Sub

Public Sub ValorIPI_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorIRRF_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorIRRFAlterado = 1
    
End Sub

Public Sub ValorParcela_Change()

    If Len(Trim(ValorParcela.Text)) > 0 Then
        iAlterado = REGISTRO_ALTERADO
    End If

End Sub

Public Sub ValorParcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorParcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorParcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorProdutos_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorSeguro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorTotalAlterado = 1

End Sub

Private Function TipoFornecedor_Dados(objFornecedor As ClassFornecedor, objTipoFornecedor As ClassTipoFornecedor) As Long
'Lê os dados de TipoFornecedor ligado a objFornecedor. Coloca na Tela.

Dim lErro As Long

On Error GoTo Erro_TipoFornecedor_Dados

    objTipoFornecedor.iCodigo = objFornecedor.iTipo

    'Lê o TipoFornecedor a partir do código
    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then gError 18681
    If lErro <> SUCESSO Then gError 18388

'    'Coloca os dados de TipoFornecedor na Tela
'    CreditoIPI.Value = objTipoFornecedor.iCreditoIPI
'    CreditoICMS.Value = objTipoFornecedor.iCreditoICMS

    TipoFornecedor_Dados = SUCESSO

    Exit Function

Erro_TipoFornecedor_Dados:

    TipoFornecedor_Dados = gErr

    Select Case gErr

        Case 18388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO", gErr, objTipoFornecedor.iCodigo)

        Case 18681

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156157)

    End Select

    Exit Function

End Function

Public Sub Limpa_Tela_NFFatPag()

Dim lErro As Long
Dim iIndice As Integer

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Chama função que limpa TextBoxes e MaskedEdits da Tela
    Call Limpa_Tela(Me)

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    Set gobjAnexos = New ClassAnexos

    'Limpa os campos não são limpos pela função acima
    Filial.Clear
    ComboFilialPC.ListIndex = -1
    CreditoICMS.Value = 0
    CreditoIPI.Value = 0
    CondicaoPagamento.Text = ""
    INSSRetido.Value = vbUnchecked
    
    'Coloca Tipo Default
    For iIndice = 0 To Tipo.ListCount - 1
        If SCodigo_Extrai(Tipo.List(iIndice)) = "NFFP" Then
            Tipo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Natureza.PromptInclude = False
    Natureza.Text = ""
    Natureza.PromptInclude = True
    LabelNaturezaDesc.Caption = ""
    
    Call Grid_Limpa(objGridParcelas)

    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    '#############################################
    'Inserido por Wagner
    Call CF2(Me, "NFCF_Limpa_Tela")
    '#############################################

    '#############################################
    'Inserido por Wagner 18/02/2005
    Call CF2(Me, "NFFPAG_Limpa_Tela")
    '#############################################

    iAlterado = 0

End Sub

Private Function Move_Tela_Memoria(objTituloPagar As ClassTituloPagar, colParcelas As colParcelaPagar) As Long
'Move os dados da Tela para objTituloPagar e colParcelas

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objProjetoInfo As Object
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim sCclFormatada As String, iCclPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Codigo do Fornecedor
    If Len(Trim(Fornecedor.Text)) > 0 Then
        
        objFornecedor.sNomeReduzido = Fornecedor.Text

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 18418
        If lErro <> SUCESSO Then gError 18791

        objTituloPagar.lFornecedor = objFornecedor.lCodigo
    End If

    'A filial do Fornecedor
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    End If
        
    'FilialPC
    If Len(Trim(ComboFilialPC.Text)) > 0 Then
        objTituloPagar.iFilialPedCompra = Codigo_Extrai(ComboFilialPC.Text)
    End If
    
    'NumPC
    If Len(Trim(NumPC.ClipText)) > 0 Then objTituloPagar.lNumPedCompra = CLng(NumPC.ClipText)
    
    'Numero do Titulo
    If Len(Trim(NumTitulo.ClipText)) > 0 Then objTituloPagar.lNumTitulo = CLng(NumTitulo.ClipText)

    'Data de Emissao
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objTituloPagar.dtDataEmissao = DATA_NULA
    Else
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    End If

    'Filial Empresa
    objTituloPagar.iFilialEmpresa = giFilialEmpresa
    
    'Os Valores, Condição de Pagamento
    If Len(Trim(ValorTotal.ClipText)) > 0 Then objTituloPagar.dValorTotal = CDbl(ValorTotal.Text)
    objTituloPagar.iNumParcelas = objGridParcelas.iLinhasExistentes
    If Len(Trim(ValorICMS.ClipText)) > 0 Then objTituloPagar.dValorICMS = CDbl(ValorICMS.Text)
    If Len(Trim(ValorICMSSubst.ClipText)) > 0 Then objTituloPagar.dValorICMSSubst = CDbl(ValorICMSSubst.Text)
    objTituloPagar.iCreditoICMS = CreditoICMS.Value
    If Len(Trim(ValorProdutos.ClipText)) > 0 Then objTituloPagar.dValorProdutos = CDbl(ValorProdutos.Text)
    If Len(Trim(OutrasDespesas.ClipText)) > 0 Then objTituloPagar.dOutrasDespesas = CDbl(OutrasDespesas.Text)
    If Len(Trim(ValorSeguro.ClipText)) > 0 Then objTituloPagar.dValorSeguro = CDbl(ValorSeguro)
    If Len(Trim(ValorFrete.ClipText)) > 0 Then objTituloPagar.dValorFrete = CDbl(ValorFrete.Text)
    If Len(Trim(ValorIRRF.ClipText)) > 0 Then objTituloPagar.dValorIRRF = CDbl(ValorIRRF.Text)
    If Len(Trim(ValorIPI.ClipText)) > 0 Then objTituloPagar.dValorIPI = CDbl(ValorIPI.Text)
    If Len(Trim(ValorINSS.ClipText)) > 0 Then objTituloPagar.dValorINSS = CDbl(ValorINSS.Text)
    objTituloPagar.iCreditoIPI = CreditoIPI.Value
    objTituloPagar.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
    objTituloPagar.iINSSRetido = INSSRetido.Value
    If Len(Trim(PISRetido.ClipText)) > 0 Then objTituloPagar.dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(ISSRetido.ClipText)) > 0 Then objTituloPagar.dISSRetido = CDbl(ISSRetido.Text)
    If Len(Trim(COFINSRetido.ClipText)) > 0 Then objTituloPagar.dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.ClipText)) > 0 Then objTituloPagar.dCSLLRetido = CDbl(CSLLRetido.Text)
    
    objTituloPagar.sHistorico = Historico.Text
    
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then gError 181761

    If iCclPreenchida = CCL_PREENCHIDA Then
        objTituloPagar.sCcl = sCclFormatada
    Else
        objTituloPagar.sCcl = ""
    End If
    
    'Move o Tipo
    If Len(Trim(Tipo.Text)) > 0 Then objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    
    'Move para colParcelas os dados do Grid Parcelas
    lErro = Move_GridParcelas_Memoria(colParcelas)
    If lErro <> SUCESSO Then gError 18419
    
    Set objTituloPagar.colParcelaPagar = colParcelas
    
    lErro = CF2(Me, "NFFPAG_Move_Tela_Memoria", objTituloPagar)
    If lErro <> SUCESSO Then gError 32301
    
    '##################################################
    'Inserido por Wagner 09/08/2006
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_TITPAG)
    If lErro <> SUCESSO Then gError 181761
    
    Set objTituloPagar.objProjetoInfo = objProjetoInfo
    '##################################################
    
    Set objTituloPagar.objAnexos = gobjAnexos
    
    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError 181761
    
    objTituloPagar.sNatureza = sNaturezaFormatada

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 18418, 18419, 32301, 181761

        Case 18791
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156158)

    End Select

    Exit Function

End Function

Private Function Move_GridParcelas_Memoria(colParcelas As colParcelaPagar) As Long
'Move para a memória os dados existentes no Grid

Dim lErro As Long
Dim iIndice As Integer
Dim objParcelaPag As ClassParcelaPagar

On Error GoTo Erro_Move_GridParcelas_Memoria

    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set objParcelaPag = New ClassParcelaPagar

        'Preenche objParcelaPag com a linha do GridParcelas
        objParcelaPag.iNumParcela = iIndice
        objParcelaPag.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        objParcelaPag.dtDataVencimentoReal = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))
        objParcelaPag.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))

        objParcelaPag.dValorOriginal = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col))
        objParcelaPag.iMotivoDiferenca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_MotivoDiferenca_Col))

        'Se estiver vazio usamos o Tipo Cobrança DEFAULT
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))) = 0 Then
            objParcelaPag.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
        Else
            objParcelaPag.iTipoCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
        End If

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Banco_Col))) = 0 Then
            objParcelaPag.iBancoCobrador = 0
        Else
            objParcelaPag.iBancoCobrador = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Banco_Col))
        End If

        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Portador_Col))) = 0 Then
            objParcelaPag.iPortador = 0
        Else
            objParcelaPag.iPortador = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Portador_Col))
        End If

        If GridParcelas.TextMatrix(iIndice, iGrid_Suspenso_Col) = "1" Then
            objParcelaPag.iStatus = STATUS_SUSPENSO
        Else
            objParcelaPag.iStatus = STATUS_ABERTO
        End If

        'retirar os pontos separadores para gravar sem eles
        objParcelaPag.sCodigoDeBarras = Trim(Replace(GridParcelas.TextMatrix(iIndice, iGrid_CodigoDeBarras_Col), ".", ""))
        
        'Adiciona objParcelaPag à coleção colParcelas
        With objParcelaPag
            colParcelas.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, .iPortador, .iProxSeqBaixa, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras, , .dValorOriginal, .iMotivoDiferenca
        End With
    Next

    Move_GridParcelas_Memoria = SUCESSO

    Exit Function

Erro_Move_GridParcelas_Memoria:

    Move_GridParcelas_Memoria = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156159)

    End Select

    Exit Function

End Function

'inicio contabilidade

Public Sub CTBBotaoModeloPadrao_Click()

    Call objContabil.Contabil_BotaoModeloPadrao_Click

End Sub

Public Sub CTBModelo_Click()

    Call objContabil.Contabil_Modelo_Click

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Click

'    If giTipoVersao = VERSAO_LIGHT Then
'        Call objContabil.Contabil_GridContabil_Consulta_Click
'    End If

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

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor, objTipoFornecedor As New ClassTipoFornecedor
Dim objFilial As New ClassFilialFornecedor, sContaTela As String
Dim sCclFormatada As String, iCclPreenchida As Integer, sCclMascarado As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case "Ccl"
            lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then Error 39583
        
            If iCclPreenchida = CCL_PREENCHIDA Then
                lErro = Mascara_MascararCcl(sCclFormatada, sCclMascarado)
                If lErro <> SUCESSO Then Error 39583
            
                objMnemonicoValor.colValor.Add sCclMascarado
            Else
                objMnemonicoValor.colValor.Add ""
            End If
        
        Case INSS_RETIDO
            If INSSRetido.Value = vbChecked Then
                objMnemonicoValor.colValor.Add 1
            ElseIf INSSRetido.Value = vbUnchecked Then
                objMnemonicoValor.colValor.Add 0
            End If
            
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
                If lErro <> SUCESSO Then gError 39567
                
                objMnemonicoValor.colValor.Add objFornecedor.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case FORNECEDOR_NOME
        
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 39609
            
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
                If lErro <> SUCESSO Then gError 39568
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39569
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 41600
                    
                    objMnemonicoValor.colValor.Add sContaTela
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError 39570
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case NUMERO1
            
            If Len(Trim(NumTitulo.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CLng(NumTitulo.Text)
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
        
        Case VALOR_PRODUTOS
            If Len(Trim(ValorProdutos.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorProdutos.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
        
        Case TIPO1
            If Len(Tipo.Text) > 0 Then
                objMnemonicoValor.colValor.Add Tipo.Text
            Else
                objMnemonicoValor.colValor.Add ""
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
            
        Case VALOR_OUTRASDESP
            If Len(Trim(OutrasDespesas.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(OutrasDespesas.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case CREDITA_IPI
            objMnemonicoValor.colValor.Add CreditoIPI.Value
                        
        Case CONTA_DESP_ESTOQUE
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError 39567
                
                objTipoFornecedor.iCodigo = objFornecedor.iTipo
                lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
                If lErro <> SUCESSO Then gError 41599
                
                If objTipoFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTipoFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 41968
                
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
                If lErro <> SUCESSO Then gError 133564
                
                If objFornecedor.sContaDespesa <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFornecedor.sContaDespesa, sContaTela)
                    If lErro <> SUCESSO Then gError 133565
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Else
                
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        
        
        Case Else
            gError 36229
            
    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 36229
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 39567, 39568, 39569, 39570, 39609, 41599, 41968, 133564, 133565
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156160)

    End Select

    Exit Function

End Function

Private Sub Recalcula_Cobranca()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Recalcula_Cobranca

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 Then
    
        'Passa o código da Condição para objCondicaoPagto
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
    
        'Lê Condição a partir do código
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 18411
        If lErro = 19205 Then gError 18412
    
        'Preenche o GridParcelas
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 18442

    End If
    
    Exit Sub
     
Erro_Recalcula_Cobranca:

    Select Case gErr
          
        Case 18411, 18442

        Case 18412
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156161)
     
    End Select
     
    Exit Sub

End Sub

Private Function Carrega_TipoDocumento()

Dim lErro As Long
Dim colSigla As New Collection
Dim objTipoDocumento As ClassTipoDocumento
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Outros Pagamentos
    lErro = CF("TiposDocumento_Le_TituloPag", colSigla)
    If lErro <> SUCESSO Then gError 57845
    
    'Carrega a combo com as Siglas lidas
    For Each objTipoDocumento In colSigla
        
        'Se a Sigla não for de Fatura à pagar
        If objTipoDocumento.sSigla <> "FP" Then
            Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
        End If
    
    Next

    For iIndice = 0 To Tipo.ListCount - 1
        If SCodigo_Extrai(Tipo.List(iIndice)) = "NFFP" Then
            Tipo.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Carrega_TipoDocumento = SUCESSO

    Exit Function
    
Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr
    
    Select Case gErr
    
        Case 57845
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156162)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_PAGAR_ID
    Set Form_Load_Ocx = Me
    Caption = "Nota Fiscal Fatura a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "NFFATPAG"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

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

'**** fim do trecho a ser copiado *****

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call LabelTipo_Click
        ElseIf Me.ActiveControl Is ComboFilialPC Then
            Call LblFilialPc_Click
        ElseIf Me.ActiveControl Is CondicaoPagamento Then
            Call CondPagtoLabel_Click
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        '#########################################
        'Inserido por Wagner 10/08/2006
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        '#########################################
        End If
    
    End If
    
    '##################################
    'Inserido por Wagner
    Call CF2(Me, "NFFPAG_UserControl_KeyDown", KeyCode, Shift)
    '##################################
    
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
    
        Call Recalcula_Cobranca
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
    
        Call Recalcula_Cobranca
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
    
        Call Recalcula_Cobranca
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

'@@@@@ para ser customizada
Function NFFPAG_Traz_NFFatPag_Tela(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long
    NFFPAG_Traz_NFFatPag_Tela = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFFPAG_Move_Tela_Memoria(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long
    NFFPAG_Move_Tela_Memoria = SUCESSO
End Function

'@@@@@ para ser customizada
Function NFFPAG_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    NFFPAG_Form_Unload = SUCESSO
End Function

Private Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134052

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134052

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156166)

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

    If Len(Trim(Fornecedor.Text)) = 0 Then gError 132161
    If Len(Trim(Filial.Text)) = 0 Then gError 132162
        
    objFornecedor.sNomeReduzido = Fornecedor.Text

    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 132163
    If lErro = 6681 Then gError 132164
    
    sFiltro = "Fornecedor = " & objFornecedor.lCodigo & " AND Filial = " & CStr(Codigo_Extrai(Filial.Text))
    
    lErro = CF("TitulosPag_Automatico", lCodigo, sFiltro)
    If lErro <> SUCESSO Then gError 132152
    
    NumTitulo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 132152, 132163

        Case 132161
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 132162
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 132164
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156167)
    
    End Select

    Exit Sub
    
End Sub
'##################################

'#####################################
'Inserido por Wagner 17/02/2005
Function NFFPAG_Form_Load(ByVal objCT As Object) As Long
    NFFPAG_Form_Load = SUCESSO
End Function

Function NFFPAG_Limpa_Tela(ByVal objCT As Object) As Long
    NFFPAG_Limpa_Tela = SUCESSO
End Function

Function NFFPAG_Trata_Parametros(ByVal objCT As Object, objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long

On Error GoTo Erro_NFFPAG_Trata_Parametros

    lErro = CF("TituloPagar_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 18368 'And lErro <> 18372 Then gError 18368
    
    'If lErro <> SUCESSO Then gError 18384

    NFFPAG_Trata_Parametros = SUCESSO

    Exit Function

Erro_NFFPAG_Trata_Parametros:

    NFFPAG_Trata_Parametros = gErr

    Select Case gErr
    
        Case 18368

        Case 18384
            Call Rotina_Erro(vbOKOnly, "ERRO_NFFATPAG_NAO_CADASTRADA", gErr, objTituloPagar.lNumIntDoc)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156168)

    End Select

    Exit Function
    
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

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
   Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "FR"
    objItensPRJCR.sDescricao = "Frete"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorFrete.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "SG"
    objItensPRJCR.sDescricao = "Seguro"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorSeguro.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "OD"
    objItensPRJCR.sDescricao = "Outras Despesas"

    objItensPRJCR.dValorOriginal = StrParaDbl(OutrasDespesas.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "ICMS"
    objItensPRJCR.sDescricao = "ICMS Substituto"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorICMSSubst.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "IPI"
    objItensPRJCR.sDescricao = "IPI"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorIPI.Text)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
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

'#################################################################
'Inserido por Wagner
Function NFFPAG_Saida_Celula(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
    NFFPAG_Saida_Celula = SUCESSO
End Function
Function NFFPAG_Gravar_Registro(ByVal objCT As Object, ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelaPagar As colParcelaPagar) As Long
    NFFPAG_Gravar_Registro = SUCESSO
End Function
Function NFFPAG_Gravar_Registro2(ByVal objCT As Object, ByVal objTituloPagar As ClassTituloPagar, ByVal colParcelaPagar As colParcelaPagar, ByVal objContabil As ClassContabil) As Long
    NFFPAG_Gravar_Registro2 = SUCESSO
End Function
Function NFFPAG_UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer) As Long
    NFFPAG_UserControl_KeyDown = SUCESSO
End Function
'#################################################################

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
    
        Call Recalcula_Cobranca
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

Public Sub CodigodeBarras_Change()
Dim sTexto As String

    If giTeste = 1 Then Exit Sub

    iAlterado = REGISTRO_ALTERADO
    
    If left(CodigodeBarras.Text, 1) = "8" Then
        giTeste = 1
        sTexto = CodigodeBarras.ClipText
        CodigodeBarras.Mask = ""
        CodigodeBarras.Mask = "###########.#.###########.#.###########.#.###########.#"
        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = sTexto
        CodigodeBarras.PromptInclude = True
        giTeste = 0
    Else
        giTeste = 1
        sTexto = CodigodeBarras.ClipText
        CodigodeBarras.Mask = ""
        CodigodeBarras.Mask = "#####.#####.#####.######.#####.######.#.#################"
        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = sTexto
        CodigodeBarras.PromptInclude = True
        giTeste = 0
    End If
End Sub

Public Sub CodigodeBarras_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub CodigodeBarras_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub CodigodeBarras_Validate(Cancel As Boolean)
Dim lErro As Long

    Set objGridParcelas.objControle = CodigodeBarras
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_CodigodeBarras(objGridInt As AdmGrid) As Long
'Faz a crítica da celula CodigodeBarras do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double
Dim bCancel As Boolean

On Error GoTo Erro_Saida_Celula_CodigodeBarras

    Set objGridInt.objControle = CodigodeBarras
    
    CodBarrasAvulso.Mask = CodigodeBarras.Mask
    
    CodBarrasAvulso.PromptInclude = False
    CodBarrasAvulso.Text = CodigodeBarras.Text
    CodBarrasAvulso.PromptInclude = True
    
    bCancel = False
    
    Call CodBarrasAvulso_Validate(bCancel)
    If bCancel Then gError 18426
    
    CodBarrasAvulso.PromptInclude = False
    CodBarrasAvulso.Text = ""
    CodBarrasAvulso.PromptInclude = True
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 18427

    Saida_Celula_CodigodeBarras = SUCESSO

    Exit Function

Erro_Saida_Celula_CodigodeBarras:

    Saida_Celula_CodigodeBarras = gErr

    Select Case gErr

        Case 18426, 18427
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156130)

    End Select

    Exit Function

End Function

Public Sub CodBarrasAvulso_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodBanco As Integer
Dim dNumDias As Double
Dim dtDataVenc As Date
Dim iIndice As Integer
Dim dtDataVencReal As Date
Dim dValor As Double
Dim sLinDig As String
Dim sMascara As String

On Error GoTo Erro_CodBarrasAvulso_Validate

    If Len(Trim(CodBarrasAvulso.ClipText)) <> 0 Then

        'Verifica se tem alguma linha selecionada no Grid
        If GridParcelas.Row = 0 Then gError 188320
        
        lErro = CB_Converte_LinDig(CodBarrasAvulso.ClipText, sLinDig)
        If lErro <> SUCESSO Then gError 188321
        
        CodBarrasAvulso.PromptInclude = False
        CodBarrasAvulso.Text = sLinDig
        CodBarrasAvulso.PromptInclude = True
        
        'contas de serviços publicos (concessionarias e impostos)
        If StrParaInt(left(CodBarrasAvulso.Text, 1)) = 8 Then
    
'            If Len(Trim(CodBarrasAvulso.ClipText)) = 44 Then
'
'                dValor = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 5, 11)) / 100
'
'                If dValor <> 0 Then
'
'                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = Format(dValor, "Standard")
'
'                End If
'
'            Else
            
                dValor = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 5, 7) & Mid(CodBarrasAvulso.ClipText, 13, 4)) / 100
            
                If dValor <> 0 Then
            
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = Format(dValor, "Standard")
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) = Format(dValor, "Standard")
                
                End If
            
'            End If
                
        Else
        
            'Preenche a Combo de Banco Cobrador com os 3 primeiros dígitos do Código de barras
            iCodBanco = StrParaInt(left(CodBarrasAvulso.Text, 3))
            
            For iIndice = 0 To comboCobrador.ListCount - 1
                If comboCobrador.ItemData(iIndice) = iCodBanco Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Banco_Col) = comboCobrador.List(iIndice)
                    Exit For
                End If
            Next
            

'            If Len(Trim(CodBarrasAvulso.ClipText)) = 44 Then
'
'                dNumDias = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 6, 4))
'
'                If dNumDias <> 0 Then
'
'                    dtDataVenc = DateAdd("d", dNumDias, CDate("7/10/1997"))
'                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) = Format(dtDataVenc, "dd/mm/yyyy")
'
'                    'Calcula a Data de Vencimento Real
'                    lErro = CF("DataVencto_Real", dtDataVenc, dtDataVencReal)
'                    If lErro <> SUCESSO Then gError 188321
'
'                    'Coloca data de Vencimento Real no Grid
'                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVencReal, "dd/mm/yyyy")
'
'                End If
'
'            Else
            
                dNumDias = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 34, 4))
            
                If dNumDias <> 0 Then
            
                    dtDataVenc = DateAdd("d", dNumDias, CDate("7/10/1997"))
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) = Format(dtDataVenc, "dd/mm/yyyy")
                    
                    'Calcula a Data de Vencimento Real
                    lErro = CF("DataVencto_Real", dtDataVenc, dtDataVencReal)
                    If lErro <> SUCESSO Then gError 188322
            
                    'Coloca data de Vencimento Real no Grid
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVencReal, "dd/mm/yyyy")
                    
                End If
            
            
'            End If
            
'            If Len(Trim(CodBarrasAvulso.ClipText)) = 44 Then
'
'                dValor = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 10, 10)) / 100
'
'                If dValor <> 0 Then
'
'                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = Format(dValor, "Standard")
'
'                End If
'
'            Else
            
                dValor = StrParaDbl(Mid(CodBarrasAvulso.ClipText, 38, 10)) / 100
            
                If dValor <> 0 Then
            
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) = Format(dValor, "Standard")
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) = Format(dValor, "Standard")
                
                End If
            
'            End If
                
        End If
        
        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_BANCARIA Then
                GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next
        
        giTeste = 1
        CodigodeBarras.Mask = CodBarrasAvulso.Mask
        
        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = CodBarrasAvulso.Text
        CodigodeBarras.PromptInclude = True
        
        giTeste = 0
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_CodigoDeBarras_Col) = CodBarrasAvulso.Text
    
    End If
    
    Exit Sub

Erro_CodBarrasAvulso_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 188320
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 188321, 188322
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 188327)
    
    End Select
    
    Exit Sub

End Sub

Public Sub CodBarrasAvulso_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CodBarrasAvulso, iAlterado)

End Sub

Public Sub CodBarrasAvulso_Change()

Dim sTexto As String
Static iTeste As Integer
Dim iPos As Integer

    If iTeste = 1 Then Exit Sub
     If left(CodBarrasAvulso.Text, 1) = "8" Then

        iPos = CodBarrasAvulso.SelStart
        iTeste = 1
        sTexto = CodBarrasAvulso.ClipText
        CodBarrasAvulso.Mask = "###########.#.###########.#.###########.#.###########.#"
        CodBarrasAvulso.PromptInclude = False
        CodBarrasAvulso.Text = sTexto
        CodBarrasAvulso.PromptInclude = True
        CodBarrasAvulso.SelStart = iPos
        iTeste = 0

    Else
        iPos = CodBarrasAvulso.SelStart
        iTeste = 1
        sTexto = CodBarrasAvulso.ClipText
        CodBarrasAvulso.Mask = ""
        CodBarrasAvulso.Mask = "#####.#####.#####.######.#####.######.#.#################"
        CodBarrasAvulso.PromptInclude = False
        CodBarrasAvulso.Text = sTexto
        CodBarrasAvulso.PromptInclude = True
        CodBarrasAvulso.SelStart = iPos
        iTeste = 0
    End If

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iIndice As Integer
Dim sTexto As String

On Error GoTo Erro_Rotina_Grid_Enable

    If iLocalChamada = 3 Then Exit Sub

    'Pesquisa a controle da coluna em questão
    Select Case objControl.Name
        
        Case CodigodeBarras.Name
            
            If Len(GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col)) > 0 Then
        
                If left(GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col), 1) = "8" Then
                    giTeste = 1
                    CodigodeBarras.Mask = "###########.#.###########.#.###########.#.###########.#"
                    CodigodeBarras.PromptInclude = False
                    CodigodeBarras.Text = GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col)
                    CodigodeBarras.PromptInclude = True
                    giTeste = 0
                Else
                    giTeste = 1
                    CodigodeBarras.Mask = "#####.#####.#####.######.#####.######.#.#################"
                    CodigodeBarras.PromptInclude = False
                    CodigodeBarras.Text = GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col)
                    CodigodeBarras.PromptInclude = True
                    giTeste = 0
                End If
            
            End If
        
    End Select
    
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 193460)

    End Select

    Exit Sub

End Sub

Public Sub Diferenca_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Diferenca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Diferenca_Validate

    'Se a diferença foi preenchida
    If Len(Trim(Diferenca.ClipText)) > 0 Then
    
        'Criticar o valor
        lErro = Valor_Critica(Diferenca.Text)
        If lErro <> SUCESSO Then gError 500037
        
    End If
    
    Exit Sub

Erro_Diferenca_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 500037
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)
    
    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_ValorOriginal(objGridInt As AdmGrid) As Long
'Faz a crítica da celula ValorOriginal do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_ValorOriginal

    Set objGridInt.objControle = ValorOriginal

    'Verifica se valor Original está preenchido
    If Len(Trim(ValorOriginal.Text)) > 0 Then
    
        'Critica se valor Original é positivo
        lErro = Valor_Positivo_Critica(ValorOriginal.Text)
        If lErro <> SUCESSO Then gError 500040

        ValorOriginal.Text = Format(ValorOriginal.Text, "Standard")
        If ValorOriginal.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) Then CobrancaAutomatica.Value = vbUnchecked
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 500041

    Saida_Celula_ValorOriginal = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorOriginal:

    Saida_Celula_ValorOriginal = gErr

    Select Case gErr

        Case 500040, 500041
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_MotivoDiferenca(objGridInt As AdmGrid) As Long
'Faz a saida de celula de MotivoDiferenca

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_MotivoDiferenca

    Set objGridInt.objControle = MotivoDiferenca

    'Verifica se o MotivoDiferenca está preenchido
    If Len(Trim(MotivoDiferenca.Text)) > 0 Then
    
        If MotivoDiferenca.Text <> MotivoDiferenca.List(MotivoDiferenca.ListIndex) Then

            lErro = Combo_Item_Seleciona(MotivoDiferenca)
            If lErro <> SUCESSO And lErro <> 12250 Then gError 500042
    
            'Se não encontrou o Motivo de Diferença nem com o Código nem Descrição ==> erro
            If lErro <> SUCESSO Then gError 500048
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 500049

    Saida_Celula_MotivoDiferenca = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoDiferenca:

    Saida_Celula_MotivoDiferenca = gErr

    Select Case gErr

        Case 500042, 500049
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 500048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCA_NAO_ENCONTRADO", gErr, MotivoDiferenca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Private Function Carrega_MotivoDiferenca() As Long
'Carrega a combo de Cobrador

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_MotivoDiferenca

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "MotivoDiferenca", "Codigo", "Descricao", STRING_DESCRICAO_CAMPO, colCodigoNome)
    If lErro <> SUCESSO Then gError 500047

   'Preenche ComboBox com código e Descrição dos Bancos
    For iIndice = 1 To colCodigoNome.Count
        Set objCodigoNome = colCodigoNome(iIndice)
        MotivoDiferenca.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        MotivoDiferenca.ItemData(MotivoDiferenca.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_MotivoDiferenca = SUCESSO

    Exit Function

Erro_Carrega_MotivoDiferenca:

    Carrega_MotivoDiferenca = gErr

    Select Case gErr

        Case 500047 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub ValorOriginal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorOriginal_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorOriginal_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorOriginal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorOriginal
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub MotivoDiferenca_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub MotivoDiferenca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub MotivoDiferenca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub MotivoDiferenca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = MotivoDiferenca
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Inicializa_Mascara_Natureza() As Long
'inicializa a mascara da Natureza

Dim sMascaraNatureza As String
Dim lErro As Long

On Error GoTo Erro_Inicializa_Mascara_Natureza

    'Inicializa a máscara da Natureza
    sMascaraNatureza = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Armazena em sMascaraNatureza a mascara a ser a ser exibida no campo Natureza
    lErro = MascaraItem(SEGMENTO_NATMOVCTA, sMascaraNatureza)
    If lErro <> SUCESSO Then gError 122836
    
    'coloca a mascara na tela.
    Natureza.Mask = sMascaraNatureza
    
    Inicializa_Mascara_Natureza = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Mascara_Natureza:

    Inicializa_Mascara_Natureza = gErr
    
    Select Case gErr
    
        Case 122836
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARAITEM", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Function

End Function

Private Sub objEventoNatureza_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNatMovCta As ClassNatMovCta
Dim sNaturezaEnxuta As String

On Error GoTo Erro_objEventoNatureza_evSelecao

    Set objNatMovCta = obj1

    If objNatMovCta.sCodigo = "" Then
        
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
    
    Else

        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objNatMovCta.sCodigo, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 122833

        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
    
    End If

    Call Natureza_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

Erro_objEventoNatureza_evSelecao:

    Select Case gErr

        Case 122833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub

End Sub

Public Sub Natureza_Validate(Cancel As Boolean)
     
Dim lErro As Long
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objNatMovCta As New ClassNatMovCta

On Error GoTo Erro_Natureza_Validate

    If Len(Natureza.ClipText) > 0 Then

        sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)

        'critica o formato da Natureza
        lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
        If lErro <> SUCESSO Then gError 122826
        
        'Obj recebe código
        objNatMovCta.sCodigo = sNaturezaFormatada
        
        'Verifica se a Natureza é analítica e se seu Tipo Corresponde a um pagamento
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_PAGAMENTO)
        If lErro <> SUCESSO Then gError 122843
        
        'Coloca a Descrição da Natureza na Tela
        LabelNaturezaDesc.Caption = objNatMovCta.sDescricao
        
    Else
    
        LabelNaturezaDesc.Caption = ""
    
    End If
    
    Exit Sub
    
Erro_Natureza_Validate:

    Cancel = True

    Select Case gErr
    
        Case 122826, 122843
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select

    Exit Sub
    
End Sub


Public Sub LabelNatureza_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Natureza.ClipText
    
    colSelecao.Add NATUREZA_TIPO_PAGAMENTO
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")

End Sub

Public Sub Natureza_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CclLabel_Click()

Dim objCcl As New ClassCcl
Dim colSelecao As New Collection

    Call Chama_Tela("CclLista", colSelecao, objCcl, objEventoCcl)

End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)
'Preenche Ccl

Dim objCcl As New ClassCcl
Dim sCclFormatada As String
Dim sCclMascarado As String
Dim lErro As Long

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_RetornaCclEnxuta(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError 22930

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case 22930
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACCLENXUTA", gErr, objCcl.sCcl)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175259)

    End Select

    Exit Sub

End Sub

Public Sub Ccl_Validate(Cancel As Boolean)
'verifica existência da Ccl informada

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl

On Error GoTo Erro_Ccl_Validate

    'se Ccl não estiver preenchida sai da rotina
    If Len(Trim(Ccl.Text)) = 0 Then Exit Sub

    lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
    If lErro <> SUCESSO And lErro <> 5703 Then gError 31558

    If lErro = 5703 Then gError 31559

    Exit Sub

Erro_Ccl_Validate:

    Cancel = True


    Select Case gErr

        Case 31558

        Case 31559
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CCL_NAO_CADASTRADO", gErr, Ccl.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175260)

    End Select

    Exit Sub

End Sub

Public Sub BotaoAnexos_Click()
   Call Chama_Tela_Modal("Anexos", gobjAnexos)
End Sub

