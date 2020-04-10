VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTituloPagar_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public gobjInfoUsu As Object

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Private iFrameAtual As Integer
Dim iFramePagamentoAtual As Integer

Dim objGridParcelas As AdmGrid
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_VenctoReal_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_SaldoParcela_Col As Integer
Dim iGrid_StatusParcela_Col As Integer
Dim iGrid_Cobranca_Col As Integer
Dim iGrid_Suspenso_Col As Integer
Dim iGrid_Banco_Col As Integer
Dim iGrid_Portador_Col As Integer
Dim iGrid_CodigoDeBarras_Col As Integer
Dim iGrid_ValorOriginal_Col As Integer
Dim iGrid_MotivoDiferenca_Col As Integer

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo
Dim gobjParcPagBaixa As New ClassBaixaParcPagar
Dim gobjBaixaPagar As New ClassBaixaPagar

Public iAlterado As Integer

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1

'inicio contabilidade
Dim objGrid1 As AdmGrid
Public objContabil As New ClassContabil
'Fim da Contabilidade

'Constantes públicas dos tabs
Private Const TAB_Titulos = 1
Private Const TAB_Parcelas = 2
Private Const TAB_Baixas = 3
Private Const TAB_Contabilizacao = 4

Private mvariFilialEmpresaTela As Integer

'--- inicio dos properties get dos controles da tela

Public Property Let iFilialEmpresaTela(ByVal vData As Integer)
    mvariFilialEmpresaTela = vData
End Property

Public Property Get iFilialEmpresaTela() As Integer
    If mvariFilialEmpresaTela <> 0 Then
        iFilialEmpresaTela = mvariFilialEmpresaTela
    Else
        iFilialEmpresaTela = giFilialEmpresa
    End If
End Property

Public Property Get Frame() As Object
     Set Frame = objUserControl.Controls("Frame")
End Property

Public Property Get Sequencial() As Object
     Set Sequencial = objUserControl.Controls("Sequencial")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get DataBaixa() As Object
     Set DataBaixa = objUserControl.Controls("DataBaixa")
End Property

Public Property Get ValorPago() As Object
     Set ValorPago = objUserControl.Controls("ValorPago")
End Property

Public Property Get Juros() As Object
     Set Juros = objUserControl.Controls("Juros")
End Property

Public Property Get Multa() As Object
     Set Multa = objUserControl.Controls("Multa")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get ValorBaixado() As Object
     Set ValorBaixado = objUserControl.Controls("ValorBaixado")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get Pagamento() As Object
     Set Pagamento = objUserControl.Controls("Pagamento")
End Property

Public Property Get UpDownParcela() As Object
     Set UpDownParcela = objUserControl.Controls("UpDownParcela")
End Property

Public Property Get FramePagamento() As Object
     Set FramePagamento = objUserControl.Controls("FramePagamento")
End Property

Public Property Get Frame10() As Object
     Set Frame10 = objUserControl.Controls("Frame10")
End Property

Public Property Get TipoMeioPagto() As Object
     Set TipoMeioPagto = objUserControl.Controls("TipoMeioPagto")
End Property

Public Property Get ContaCorrente() As Object
     Set ContaCorrente = objUserControl.Controls("ContaCorrente")
End Property

Public Property Get ValorPagoPagto() As Object
     Set ValorPagoPagto = objUserControl.Controls("ValorPagoPagto")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get Beneficiario() As Object
     Set Beneficiario = objUserControl.Controls("Beneficiario")
End Property

Public Property Get NumOuSequencial() As Object
     Set NumOuSequencial = objUserControl.Controls("NumOuSequencial")
End Property

Public Property Get Portador() As Object
     Set Portador = objUserControl.Controls("Portador")
End Property

Public Property Get DataEmissaoCred() As Object
     Set DataEmissaoCred = objUserControl.Controls("DataEmissaoCred")
End Property

Public Property Get SiglaDocumentoCR() As Object
     Set SiglaDocumentoCR = objUserControl.Controls("SiglaDocumentoCR")
End Property

Public Property Get ValorCredito() As Object
     Set ValorCredito = objUserControl.Controls("ValorCredito")
End Property

Public Property Get FilialEmpresaCR() As Object
     Set FilialEmpresaCR = objUserControl.Controls("FilialEmpresaCR")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
End Property

Public Property Get SaldoCredito() As Object
     Set SaldoCredito = objUserControl.Controls("SaldoCredito")
End Property

Public Property Get NumeroMP() As Object
     Set NumeroMP = objUserControl.Controls("NumeroMP")
End Property

Public Property Get CCIntNomeReduzido() As Object
     Set CCIntNomeReduzido = objUserControl.Controls("CCIntNomeReduzido")
End Property

Public Property Get FilialEmpresaPA() As Object
     Set FilialEmpresaPA = objUserControl.Controls("FilialEmpresaPA")
End Property

Public Property Get ValorPA() As Object
     Set ValorPA = objUserControl.Controls("ValorPA")
End Property

Public Property Get MeioPagtoDescricao() As Object
     Set MeioPagtoDescricao = objUserControl.Controls("MeioPagtoDescricao")
End Property

Public Property Get DataMovimento() As Object
     Set DataMovimento = objUserControl.Controls("DataMovimento")
End Property

Public Property Get DataVencParcBaixa() As Object
     Set DataVencParcBaixa = objUserControl.Controls("DataVencParcBaixa")
End Property

Public Property Get ValorParcBaixa() As Object
     Set ValorParcBaixa = objUserControl.Controls("ValorParcBaixa")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get StatusParcela() As Object
     Set StatusParcela = objUserControl.Controls("StatusParcela")
End Property

Public Property Get SaldoParcela() As Object
     Set SaldoParcela = objUserControl.Controls("SaldoParcela")
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

Public Property Get Suspenso() As Object
     Set Suspenso = objUserControl.Controls("Suspenso")
End Property

Public Property Get TipoCobranca() As Object
     Set TipoCobranca = objUserControl.Controls("TipoCobranca")
End Property

Public Property Get comboCobrador() As Object
     Set comboCobrador = objUserControl.Controls("comboCobrador")
End Property

Public Property Get comboPortador() As Object
     Set comboPortador = objUserControl.Controls("comboPortador")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get BotaoDocOriginal() As Object
     Set BotaoDocOriginal = objUserControl.Controls("BotaoDocOriginal")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get INSSRetido() As Object
     Set INSSRetido = objUserControl.Controls("INSSRetido")
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

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get CreditoICMS() As Object
     Set CreditoICMS = objUserControl.Controls("CreditoICMS")
End Property

Public Property Get ValorICMSSubst() As Object
     Set ValorICMSSubst = objUserControl.Controls("ValorICMSSubst")
End Property

Public Property Get ValorICMS() As Object
     Set ValorICMS = objUserControl.Controls("ValorICMS")
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

Public Property Get ValorProdutos() As Object
     Set ValorProdutos = objUserControl.Controls("ValorProdutos")
End Property

Public Property Get Saldo() As Object
     Set Saldo = objUserControl.Controls("Saldo")
End Property

Public Property Get NumPC() As Object
     Set NumPC = objUserControl.Controls("NumPC")
End Property

Public Property Get Status() As Object
     Set Status = objUserControl.Controls("Status")
End Property

Public Property Get LblNumPC() As Object
     Set LblNumPC = objUserControl.Controls("LblNumPC")
End Property

Public Property Get LblFilialPC() As Object
     Set LblFilialPC = objUserControl.Controls("LblFilialPC")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get ComboFilialPC() As Object
     Set ComboFilialPC = objUserControl.Controls("ComboFilialPC")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
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

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
End Property

Public Property Get CTBSeqContraPartida() As Object
     Set CTBSeqContraPartida = objUserControl.Controls("CTBSeqContraPartida")
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

Public Property Get CTBGridContabil() As Object
     Set CTBGridContabil = objUserControl.Controls("CTBGridContabil")
End Property

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Consulta_Click

End Sub

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

Public Property Get CTBDataContabil() As Object
     Set CTBDataContabil = objUserControl.Controls("CTBDataContabil")
End Property

Public Property Get CTBDocumento() As Object
     Set CTBDocumento = objUserControl.Controls("CTBDocumento")
End Property

Public Property Get CTBLote() As Object
     Set CTBLote = objUserControl.Controls("CTBLote")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoConsulta() As Object
     Set BotaoConsulta = objUserControl.Controls("BotaoConsulta")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get NumeroTitulo() As Object
     Set NumeroTitulo = objUserControl.Controls("NumeroTitulo")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
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

Public Property Get LabelTipo() As Object
     Set LabelTipo = objUserControl.Controls("LabelTipo")
End Property

Public Property Get Line1() As Object
     Set Line1 = objUserControl.Controls("Line1")
End Property

Public Property Get HistoricoTit() As Object
     Set HistoricoTit = objUserControl.Controls("HistoricoTit")
End Property

Public Property Get BancoCT() As Object
     Set BancoCT = objUserControl.Controls("BancoCT")
End Property

Public Property Get AgenciaCT() As Object
     Set AgenciaCT = objUserControl.Controls("AgenciaCT")
End Property

Public Property Get ContaCT() As Object
     Set ContaCT = objUserControl.Controls("ContaCT")
End Property

Public Property Get NumeroCT() As Object
     Set NumeroCT = objUserControl.Controls("NumeroCT")
End Property

Public Property Get ValorCT() As Object
     Set ValorCT = objUserControl.Controls("ValorCT")
End Property

Public Property Get ClienteCT() As Object
     Set ClienteCT = objUserControl.Controls("ClienteCT")
End Property

Public Property Get FilialEmpresaCT() As Object
     Set FilialEmpresaCT = objUserControl.Controls("FilialEmpresaCT")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get LabelProjeto() As Object
     Set LabelProjeto = objUserControl.Controls("LabelProjeto")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get CodigodeBarras() As Object
     Set CodigodeBarras = objUserControl.Controls("CodigodeBarras")
End Property

Public Property Get Diferenca() As Object
     Set Diferenca = objUserControl.Controls("Diferenca")
End Property

Public Property Get ValorOriginal() As Object
     Set ValorOriginal = objUserControl.Controls("ValorOriginal")
End Property

Public Property Get MotivoDiferenca() As Object
     Set MotivoDiferenca = objUserControl.Controls("MotivoDiferenca")
End Property

'--- fim dos properties get dos controles da tela

Private Function Carrega_SiglaDocumento() As Long
'Carrega as Combo de Tipo de Título

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_SiglaDocumento
    
    'Le os Tipos
    lErro = CF("TiposDocumento_Le_TituloPag", colTipoDocumento)
    If lErro <> SUCESSO Then gError 60020
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        'If objTipoDocumento.sSigla <> TIPODOC_FATURA_A_PAGAR Then
            Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
        'End If
    Next
    
    Carrega_SiglaDocumento = SUCESSO

    Exit Function

Erro_Carrega_SiglaDocumento:

    Carrega_SiglaDocumento = gErr

    Select Case gErr

        Case 60020 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158334)

    End Select

    Exit Function

End Function

Public Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsulta_Click

    'Critica se os Campos estão Preenchidos(Fornecedor, Filial, Tipo, Número)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 60021

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 60022

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 60023

    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloPagar.iFilialEmpresa = iFilialEmpresaTela

    If Len(Trim(DataEmissao.ClipText)) <> 0 Then
        objTituloPagar.dtDataEmissao = MaskedParaDate(DataEmissao)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 60024

    'Traz o Titulo para a Tela
    lErro = Traz_TituloPagar_Tela(objTituloPagar)
    If lErro <> SUCESSO Then gError 60025

    Exit Sub

Erro_BotaoConsulta_Click:

    Select Case gErr

        Case 60021, 60022, 60024, 60025 'Tratado na rotina chamada

        Case 60023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158335)

    End Select

    Exit Sub

End Sub

Private Function ParcelaPag_Le_Todas(objParcelaPagar As ClassParcelaPagar) As Long
'Lê a Parcela em Parcelas a Pagar se não encontrou procura em Parcelas Pagar Baixadas

Dim lErro As Long

On Error GoTo Erro_ParcelaPag_Le_Todas
    
    'Lê em Parcelas a Pagar
    lErro = CF("ParcelaPagar_Le_Numero", objParcelaPagar)
    If lErro <> SUCESSO And lErro <> 18987 Then gError 60029
    
    'Se não encontrou
    If lErro = 18987 Then
        
        'Procura nas Parcelas a Pagar Baixadas
        lErro = CF("ParcelaPagarBaixada_Le_Numero", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 43625 Then gError 60030
        
        'Se não encontrou ---> ERRO
        If lErro = 43625 Then gError 60031

    End If

    ParcelaPag_Le_Todas = SUCESSO

    Exit Function

Erro_ParcelaPag_Le_Todas:

    ParcelaPag_Le_Todas = gErr

    Select Case gErr
        
        Case 60029, 60030 'Tratado na rotina chamada
        
        Case 60031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158336)

    End Select

    Exit Function

End Function

Private Function Critica_CamposPreenchidos() As Long
'Verifica se os campos chaves foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposPreenchidos

    'Se Fornecedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 60032

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then gError 60033

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 60034

    'Verifica se o Titulo está Preenchido
    If Len(Trim(NumeroTitulo.Text)) = 0 Then gError 60035

    Critica_CamposPreenchidos = SUCESSO

    Exit Function

Erro_Critica_CamposPreenchidos:

    Critica_CamposPreenchidos = gErr

    Select Case gErr

        Case 60032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 60033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 60034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)

        Case 60035
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158337)

    End Select

    Exit Function

End Function

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim sTela As String
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoDocOriginal_Click
    
    lErro = TitPag_Obtem(objTituloPagar)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Se for do Tipo "OR" dá mensagem pois este Título não tem documento original
    If objTituloPagar.sSiglaDocumento = SIGLA_OUTROS_PAGAMENTOS Then gError 60041
    
    'Se for do Tipo Fatura a Pagar chama a Tela de geração de Faturas a Pagar
    If objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
        
        Call Chama_Tela("FaturasPag", objTituloPagar)
        
    Else
                
        'Le o Nome da Tela que originou este Título
        lErro = CF("Titulo_Le_DocumentoOriginal", objTituloPagar.lNumIntDoc, CPR_TITULO_PAGAR, objNFiscal, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError 60042
        
        If lErro = SUCESSO Then
'            If giTipoVersao = VERSAO_FULL Then
                'Chama a Tela
                Call Chama_Tela(sTela, objNFiscal)
'            ElseIf giTipoVersao = VERSAO_LIGHT Then
'                'Chama a Tela
'                Call Chama_Tela("NFiscalFatEntrada", objNFiscal)
'            End If
        Else
            gError 60043
        End If
        
    End If
    
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 60037, 60038, 60040, 60042, ERRO_SEM_MENSAGEM 'Tratado na Rotina chamada
        
        Case 60039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 60041, 60043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAG_SEM_DOCORIGINAL", gErr, objTituloPagar.lNumTitulo)

        Case 185050
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DIFERENTE", gErr, iFilialEmpresaTela, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158338)

    End Select

    Exit Sub


End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub DataEmissao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 60044

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case gErr

        Case 60044 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158339)

    End Select

    Exit Sub

End Sub

Public Sub LabelTipo_Click()

Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = SCodigo_Extrai(Tipo.Text)

    'Chama a tela TipoDocOutrosPagLista
    Call Chama_Tela("TiposDocBaixaPagCancLista", colSelecao, objTipoDocumento, objEventoTipo)

End Sub

Public Sub NumeroTitulo_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroTitulo)

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
    If lErro <> SUCESSO And lErro <> 60483 Then gError 60130

    'Se nao encontrar -> Erro
    If lErro = 60483 Then gError 60131

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True


    Select Case gErr

        Case 60130

        Case 60131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", gErr, Tipo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158340)

    End Select

    Exit Sub


End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui "1" da data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then gError 60045

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 60045 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158341)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Adiciona "1" a data
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then gError 60046

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 60046 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158342)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 60047

    'Se não encontra valor que era CÓDIGO
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 60048

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 60049

        'Se não encontrou
        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = sFornecedor

            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then gError 60050

            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            Error 60051

        End If

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then gError 60052

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

       Case 60047, 60049, 60050

       Case 60048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

       Case 60051
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 60052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158343)

    End Select

    Exit Sub
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFramePagamentoAtual = 0

    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento
    Set objEventoTipo = New AdmEvento
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Carregar a combo de tipos de documento com todos os tipos usados em OutrosPag.frm, mais NotaFiscalFatura (NFFP) e Fatura (FP)
    lErro = Carrega_SiglaDocumento()
    If lErro <> SUCESSO Then gError 60053

    'Carrega na combo de Banco de Cobradores
    lErro = Carrega_Bancos()
    If lErro <> SUCESSO Then gError 60878
    
    'Carrega na combo de Portadores
    lErro = Carrega_Portadores()
    If lErro <> SUCESSO Then gError 60879

    'Carrega na combo os tipos de cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then gError 60054
    
    lErro = Carrega_MotivoDiferenca()
    If lErro <> SUCESSO Then gError 500071

    Set objGridParcelas = New AdmGrid

    'Inicializa Grid Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 60055

    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade_Consulta(Me, objGrid1, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 60056
    
    Call CF2(Me, "TITPAGC_Form_Load")
    
    iAlterado = 0

    iFilialEmpresaTela = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 60053, 60054, 60055, 60056, 60878, 60879, 500071

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158344)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosPag_Todos"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 60057

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", CLng(0), 0, "NumIntDoc"
    colCampoValor.Add "SiglaDocumento", objTituloPagar.sSiglaDocumento, STRING_BUFFER_MAX_TEXTO, "SiglaDocumento"
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
    colCampoValor.Add "FilialPedCompra", objTituloPagar.iFilialPedCompra, 0, "FilialPedCompra"
    colCampoValor.Add "NumPedCompra", objTituloPagar.lNumPedCompra, 0, "NumPedCompra"
    colCampoValor.Add "INSSRetido", objTituloPagar.iINSSRetido, 0, "INSSRetido"
    colCampoValor.Add "Status", objTituloPagar.iStatus, 0, "Status"
    colCampoValor.Add "Saldo", objTituloPagar.dSaldo, 0, "Saldo"
    colCampoValor.Add "PISRetido", objTituloPagar.dPISRetido, 0, "PISRetido"
    colCampoValor.Add "ISSRetido", objTituloPagar.dISSRetido, 0, "ISSRetido"
    colCampoValor.Add "COFINSRetido", objTituloPagar.dCOFINSRetido, 0, "COFINSRetido"
    colCampoValor.Add "CSLLRetido", objTituloPagar.dCSLLRetido, 0, "CSLLRetido"

    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 60057 'Tratado nas Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158345)

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
        objTituloPagar.sSiglaDocumento = colCampoValor.Item("SiglaDocumento").vValor
        objTituloPagar.lFornecedor = colCampoValor.Item("Fornecedor").vValor
        objTituloPagar.iFilial = colCampoValor.Item("Filial").vValor
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
        objTituloPagar.iFilialPedCompra = colCampoValor.Item("FilialPedCompra").vValor
        objTituloPagar.lNumPedCompra = colCampoValor.Item("NumPedCompra").vValor
        objTituloPagar.iINSSRetido = colCampoValor.Item("INSSRetido").vValor
        objTituloPagar.iStatus = colCampoValor.Item("Status").vValor
        objTituloPagar.dSaldo = colCampoValor.Item("Saldo").vValor
        objTituloPagar.dPISRetido = colCampoValor.Item("PISRetido").vValor
        objTituloPagar.dISSRetido = colCampoValor.Item("ISSRetido").vValor
        objTituloPagar.dCOFINSRetido = colCampoValor.Item("COFINSRetido").vValor
        objTituloPagar.dCSLLRetido = colCampoValor.Item("CSLLRetido").vValor
        
        'Lê o Título
        lErro = CF("TituloPag_Le_Todos", objTituloPagar)
        If lErro <> SUCESSO Then gError 126833
        
        'Traz a Nota para Tela
        lErro = Traz_TituloPagar_Tela(objTituloPagar)
        If lErro <> SUCESSO Then gError 60058

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 60058, 126833

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158346)

    End Select

    Exit Sub

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
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 60059

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then gError 60060
    
    'Move os dados da Tela para objTituloPagar e colParcelas
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas)
    If lErro <> SUCESSO Then gError 60062

    If Len(Trim(Tipo.Text)) > 0 Then
    
        'Adiciona filtros: lFornecedor e iFilial
        colSelecao.Add objTituloPagar.lFornecedor
        colSelecao.Add objTituloPagar.iFilial
        colSelecao.Add objTituloPagar.sSiglaDocumento
        
        'Chama Tela
        Call Chama_Tela("TituloPag_TodosTipoLista", colSelecao, objTituloPagar, objEventoNumero)
    
    Else
    
        'Adiciona filtros: lFornecedor e iFilial
        colSelecao.Add objTituloPagar.lFornecedor
        colSelecao.Add objTituloPagar.iFilial
        
        'Chama Tela NFFatPagLista
        Call Chama_Tela("TituloPag_TodosLista", colSelecao, objTituloPagar, objEventoNumero)
    
    End If
        
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 60059
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 60060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 60062 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158347)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloPagar = obj1

    'Lê o Título
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 126834

    'Traz os dados de objTituloPagar para Tela
    lErro = Traz_TituloPagar_Tela(objTituloPagar)
    If lErro <> SUCESSO Then gError 60063

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 60063, 126834

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158348)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim bAchou As Boolean

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloPagar Is Nothing) Then

        'Lê o Título
        If objTituloPagar.lNumIntDoc > 0 Then
            lErro = CF("TituloPagar_Le_Todos", objTituloPagar, True)
        Else
            lErro = CF("TituloPag_Le_Todos", objTituloPagar)
        End If
        If lErro <> SUCESSO Then gError 60064
        
        If objTituloPagar.iStatus = STATUS_EXCLUIDO Then gError 213839

        'Traz os dados para a Tela
        lErro = Traz_TituloPagar_Tela(objTituloPagar)
        If lErro <> SUCESSO Then gError 60065
        
        If objTituloPagar.iParcelaTela <> 0 Then
        
            'Torna Frame atual invisível
            Frame(Opcao.SelectedItem.Index).Visible = False
            If objTituloPagar.iBaixaTela = 0 Then
                iFrameAtual = TAB_Parcelas
            Else
                iFrameAtual = TAB_Baixas
            End If
            'Torna Frame atual visível
            Frame(iFrameAtual).Visible = True
            Opcao.Tabs.Item(iFrameAtual).Selected = True
            
            Call Opcao_Click
            
            Parcela(1).Caption = objTituloPagar.iParcelaTela - 1
            
            Call UpDownParcela_UpClick(1)
            
            If objTituloPagar.iBaixaTela <> 0 Then
                
                iIndice = 0
                bAchou = False
                For iIndice = 0 To Sequencial.ListCount - 1
                    If StrParaInt(Sequencial.List(iIndice)) = objTituloPagar.iBaixaTela Then
                        Sequencial.ListIndex = iIndice
                        bAchou = True
                        Exit For
                    End If
                Next
                
                If Not bAchou Then gError 189363
            
                Call Sequencial_Click
            End If
        
        
        End If

    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 60064, 60065
        
        Case 189363 'ERRO_BAIXA_PARCELA_CANCELADA
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXA_PARCELA_CANCELADA", gErr, objTituloPagar.iParcelaTela, objTituloPagar.iBaixaTela)

        Case 213839
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_EXCLUIDO2", gErr, CStr(objTituloPagar.lNumTitulo), CStr(objTituloPagar.lFornecedor))

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158349)

    End Select

    Exit Function

End Function

Private Function Traz_TituloPagar_Tela(objTituloPagar As ClassTituloPagar) As Long
'Traz os dados da Nota Fiscal Fatura para a Tela

Dim lErro As Long
Dim colParcelasPag As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim iLinha As Integer
Dim iIndice As Integer
Dim bCancel As Boolean
Dim sTela As String
Dim sModulo As String
Dim objNFiscal As New ClassNFiscal
Dim lNumIntDoc As Long
Dim iFilialAux As Integer
Dim dSomaValor As Double
Dim dSomaValorOriginal As Double

On Error GoTo Erro_Traz_TituloPagar_Tela
    
    'Limpa a Tela
    Call Limpa_Tela_TituloPagar
    
    iFilialEmpresaTela = objTituloPagar.iFilialEmpresa
    
    iFilialAux = giFilialEmpresa
    giFilialEmpresa = iFilialEmpresaTela

    'Coloca os dados do Título Pagar na Tela
    NumeroTitulo.Text = objTituloPagar.lNumTitulo
    
    Tipo.Text = objTituloPagar.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)
    
    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(bCancel)

    Filial.Text = objTituloPagar.iFilial
    Filial_Validate (bCancel)
    
    If objTituloPagar.iStatus = STATUS_BAIXADO Then
        Status.Caption = "Baixado"
    Else
        Status.Caption = "Aberto"
    End If
    
    Saldo.Caption = Format(objTituloPagar.dSaldo, "Standard")
    
    HistoricoTit.Caption = objTituloPagar.sHistorico
    
    If objTituloPagar.iFilialPedCompra <> 0 Then
        ComboFilialPC.Caption = objTituloPagar.iFilialPedCompra
    Else
        ComboFilialPC.Caption = ""
    End If

    If objTituloPagar.lNumPedCompra <> 0 Then
        NumPC.Caption = objTituloPagar.lNumPedCompra
    Else
        NumPC.Caption = ""
    End If

    'Preenche a data de Emissao
    If objTituloPagar.dtDataEmissao <> DATA_NULA Then
        DataEmissao.PromptInclude = False
        DataEmissao.Text = Format(objTituloPagar.dtDataEmissao, "dd/mm/yy")
        DataEmissao.PromptInclude = True
    End If

    If objTituloPagar.dValorTotal > 0 Then
        ValorTotal.Caption = Format(objTituloPagar.dValorTotal, "Standard")
    Else
        ValorTotal.Caption = Format(0, "Standard")
    End If
    
    If objTituloPagar.dValorICMS > 0 Then
        ValorICMS.Caption = Format(objTituloPagar.dValorICMS, "Standard")
    Else
        ValorICMS.Caption = Format(0, "Standard")
    End If
    
    If objTituloPagar.dValorICMSSubst > 0 Then
        ValorICMSSubst.Caption = Format(objTituloPagar.dValorICMSSubst, "Standard")
    Else
        ValorICMSSubst.Caption = Format(0, "Standard")
    End If
    
    CreditoICMS.Value = objTituloPagar.iCreditoICMS
    
    If objTituloPagar.dValorProdutos > 0 Then
        ValorProdutos.Caption = Format(objTituloPagar.dValorProdutos, "Standard")
    Else
        ValorProdutos.Caption = Format(0, "Standard")
    End If
    
    If objTituloPagar.dValorFrete > 0 Then
        ValorFrete.Caption = Format(objTituloPagar.dValorFrete, "Standard")
    Else
        ValorFrete.Caption = Format(0, "Standard")
    End If
    
    If objTituloPagar.dValorSeguro > 0 Then
        ValorSeguro.Caption = Format(objTituloPagar.dValorSeguro, "Standard")
    Else
        ValorSeguro.Caption = Format(0, "Standard")
    End If
    
    If objTituloPagar.dOutrasDespesas > 0 Then
        OutrasDespesas.Caption = Format(objTituloPagar.dOutrasDespesas, "Standard")
    Else
        OutrasDespesas.Caption = Format(0, "Standard")
    End If
        
    If objTituloPagar.dValorIPI > 0 Then
        ValorIPI.Caption = Format(objTituloPagar.dValorIPI, "Standard")
    Else
        ValorIPI.Caption = Format(0, "Standard")
    End If
        
    CreditoIPI.Value = objTituloPagar.iCreditoIPI
    
    If objTituloPagar.dValorIRRF > 0 Then
        ValorIRRF.Caption = Format(objTituloPagar.dValorIRRF, "Standard")
    Else
        ValorIRRF.Caption = Format(0, "Standard")
    End If
            
    If objTituloPagar.dPISRetido > 0 Then
        PISRetido.Caption = Format(objTituloPagar.dPISRetido, "Standard")
    Else
        PISRetido.Caption = Format(0, "Standard")
    End If
            
    If objTituloPagar.dISSRetido > 0 Then
        ISSRetido.Caption = Format(objTituloPagar.dISSRetido, "Standard")
    Else
        ISSRetido.Caption = Format(0, "Standard")
    End If
            
    If objTituloPagar.dCOFINSRetido > 0 Then
        COFINSRetido.Caption = Format(objTituloPagar.dCOFINSRetido, "Standard")
    Else
        COFINSRetido.Caption = Format(0, "Standard")
    End If
            
    If objTituloPagar.dCSLLRetido > 0 Then
        CSLLRetido.Caption = Format(objTituloPagar.dCSLLRetido, "Standard")
    Else
        CSLLRetido.Caption = Format(0, "Standard")
    End If
            
    If objTituloPagar.dValorINSS > 0 Then
        ValorINSS.Caption = Format(objTituloPagar.dValorINSS, "Standard")
    Else
        ValorINSS.Caption = Format(0, "Standard")
    End If
        
    INSSRetido.Value = objTituloPagar.iINSSRetido
    
    'Lê as Parcelas a Pagar vinculadas ao Título
    lErro = CF("ParcelasPagar_Le_Todas", objTituloPagar, colParcelasPag)
    If lErro <> SUCESSO Then gError 60066

    If colParcelasPag.Count > NUM_MAXIMO_PARCELAS Then gError 60067

    Call Grid_Limpa(objGridParcelas)

    iLinha = 0

    'Preenche as linhas do Grid Parcelas com os dados de cada Parcela
    For Each objParcelaPagar In colParcelasPag

        iLinha = iLinha + 1

        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objParcelaPagar.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_VenctoReal_Col) = Format(objParcelaPagar.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objParcelaPagar.dValor, "Standard")

        GridParcelas.TextMatrix(iLinha, iGrid_ValorOriginal_Col) = Format(objParcelaPagar.dValorOriginal, "Standard")
        
        For iIndice = 0 To MotivoDiferenca.ListCount - 1
            If MotivoDiferenca.ItemData(iIndice) = objParcelaPagar.iMotivoDiferenca Then
                GridParcelas.TextMatrix(iLinha, iGrid_MotivoDiferenca_Col) = MotivoDiferenca.List(iIndice)
                Exit For
            End If
        Next
        
        'Acumula os valores e os valores originais
        dSomaValor = dSomaValor + objParcelaPagar.dValor
        dSomaValorOriginal = dSomaValorOriginal + objParcelaPagar.dValorOriginal

        GridParcelas.TextMatrix(iLinha, iGrid_SaldoParcela_Col) = Format(objParcelaPagar.dSaldo, "Standard")
        
        If objParcelaPagar.iStatus = STATUS_BAIXADO Then
            GridParcelas.TextMatrix(iLinha, iGrid_StatusParcela_Col) = "Baixada"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_StatusParcela_Col) = "Aberta"
        End If
        
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

    'Faz refresh nas checkboxes
    Call Grid_Refresh_Checkbox(objGridParcelas)

    Parcela(1).Caption = 1
    
    DataVencParcBaixa.Caption = Format(colParcelasPag(1).dtDataVencimento, "dd/mm/yyyy")
    ValorParcBaixa.Caption = Format(colParcelasPag(1).dValor, "Standard")

    lErro = Carrega_ComboSequencial(colParcelasPag(1).lNumIntDoc)
    If lErro <> SUCESSO Then gError 60068
    
    'Se for uma Parcela Baixada Traz os dados da Baixa
    If Sequencial.ListCount > 0 Then

        'Traz a Primeira Parcela da Baixa e o Primeiro Sequencial
        Call Carrega_Dados_Parcela(objTituloPagar.lNumIntDoc, 1, Sequencial.List(0))

    Else
        Call Limpa_Campos_Parcela
    End If
    
    Diferenca.Caption = Format(dSomaValor - dSomaValorOriginal, "Standard")

    
    'Se for do Tipo "Fatura" não tem contabilidade
    'If objTituloPagar.sSiglaDocumento <> TIPODOC_FATURA_A_PAGAR Then
        
        'Le a Tela e Modulo que Originaram a Contabilidade
        
        lErro = CF("Titulo_Le_DocumentoOriginal", objTituloPagar.lNumIntDoc, CPR_TITULO_PAGAR, objNFiscal, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError 60070
        
        'se o titulo foi gerado pelo CP ==> traz sua contabilidade
        If lErro = 58942 Then
            lNumIntDoc = objTituloPagar.lNumIntDoc
            
            If objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR Then
                sTela = "FaturasPag"
            Else
                sTela = "NFFATPAG"
            End If
            sModulo = MODULO_CONTASAPAGAR
            
        Else
            lNumIntDoc = objNFiscal.lNumIntDoc
            sModulo = MODULO_ESTOQUE
            
            'Alterado por cyntia
            'Recolhe o note da tela na nota fiscal pelo numintdoc passado
            lErro = CF("NFiscal_Obtem_Nome_Tela", lNumIntDoc, sTela)
            If lErro <> SUCESSO Then gError 99035
            
        End If
        
            'Traz os dados contábeis para a tela (contabilidade)
            lErro = objContabil.Contabil_Traz_Doc_Tela_Consulta(lNumIntDoc, sModulo, sTela)
            If lErro <> SUCESSO And lErro <> 58885 Then gError 60071
       
    'End If
    
    lErro = CF2(Me, "TITPAGC_Traz_TituloPagar_Tela", objTituloPagar)
    If lErro <> SUCESSO Then gError 185059
    
    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objTituloPagar.objProjetoInfo, objTituloPagar.lNumIntDoc, PRJ_CR_TIPO_TITPAG)
    If lErro <> SUCESSO Then gError 181676
    '###################################################
    
    giFilialEmpresa = iFilialAux
    
    Traz_TituloPagar_Tela = SUCESSO

    Exit Function

Erro_Traz_TituloPagar_Tela:

    giFilialEmpresa = iFilialAux

    Traz_TituloPagar_Tela = gErr

    Select Case gErr

        Case 60066, 60068, 60069, 60070, 60071, 99035, 181676, 185059 'Tratados nas Rotinas Chamadas

        Case 60067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", gErr, colParcelasPag.Count, NUM_MAXIMO_PARCELAS)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158350)

    End Select

    Exit Function

End Function

Public Sub UpDownParcela_DownClick(Index As Integer)

Dim iIndice As Integer
Dim lErro As Long
Dim iParcela As Integer
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_UpDownParcela_DownClick
    
    'Verifica se já existe alguma Parcela na Tela
    If Len(Trim(Parcela(Index).Caption)) > 0 Then
        iParcela = CInt(Parcela(Index).Caption)
    Else
        iParcela = 0
    End If

    'Verifica se existe uma Parcela inferior a Parcela da Tela
    If iParcela - 1 > 0 Then
    
        'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número)
        lErro = Critica_CamposPreenchidos()
        If lErro <> SUCESSO Then gError 60072
        
        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê o codigo através do Nome Reduzido
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 60073

        'Não achou o Fornecedor --> erro
        If lErro <> SUCESSO Then gError 60074

        objTituloPagar.lFornecedor = objFornecedor.lCodigo
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
        objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
        objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
        objTituloPagar.iFilialEmpresa = iFilialEmpresaTela
        
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
        Else
            objTituloPagar.dtDataEmissao = DATA_NULA
        End If
        
        lErro = CF("TituloPag_Le_Todos", objTituloPagar)
        If lErro <> SUCESSO Then gError 60075

        objParcelaPagar.iNumParcela = iParcela - 1
        objParcelaPagar.lNumIntTitulo = objTituloPagar.lNumIntDoc

        lErro = ParcelaPag_Le_Todas(objParcelaPagar)
        If lErro <> SUCESSO Then gError 60076
        
        Parcela(1).Caption = iParcela - 1
        
        DataVencParcBaixa.Caption = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        ValorParcBaixa.Caption = Format(objParcelaPagar.dValor, "Standard")
        
        lErro = Carrega_ComboSequencial(objParcelaPagar.lNumIntDoc)
        If lErro <> SUCESSO Then gError 60077

        'Se for uma Parcela Baixada Traz os dados da Baixa
        If Sequencial.ListCount > 0 Then

            'Carrega os Dados da Baixa (1 como Default)
            Call Carrega_Dados_Parcela(objParcelaPagar.lNumIntTitulo, objParcelaPagar.iNumParcela, Sequencial.List(0))

        Else
            Call Limpa_Campos_Parcela
        End If

    End If

    Exit Sub

Erro_UpDownParcela_DownClick:

    Select Case gErr

        Case 60072, 60073, 60075, 60076, 60077 'Tratado na Rotina chamada
        
        Case 60074
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158351)

    End Select

    Exit Sub

End Sub

Public Sub UpDownParcela_UpClick(Index As Integer)

Dim iIndice As Integer
Dim lErro As Long
Dim iParcela As Integer
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_UpDownParcela_UpClick

    'Verifica se já existe alguma Parcela na tela
    If Len(Trim(Parcela(1).Caption)) > 0 Then
        iParcela = CInt(Parcela(1).Caption)
    Else
        iParcela = 0
    End If

    'Verifica se existe a Parcela (iParcela+1)
    If iParcela + 1 <= objGridParcelas.iLinhasExistentes Then

        'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número, DataEmissao)
        lErro = Critica_CamposPreenchidos()
        If lErro <> SUCESSO Then gError 60078

        objFornecedor.sNomeReduzido = Fornecedor.Text

        'Lê o codigo através do Nome Reduzido
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 60079
    
        'Não achou o Fornecedor --> erro
        If lErro <> SUCESSO Then gError 60080
    
        objTituloPagar.lFornecedor = objFornecedor.lCodigo
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)

        objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
        objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
        objTituloPagar.iFilialEmpresa = iFilialEmpresaTela
        
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
        Else
            objTituloPagar.dtDataEmissao = DATA_NULA
        End If
        
        lErro = CF("TituloPag_Le_Todos", objTituloPagar)
        If lErro <> SUCESSO Then gError 60081

        objParcelaPagar.iNumParcela = iParcela + 1
        objParcelaPagar.lNumIntTitulo = objTituloPagar.lNumIntDoc

        lErro = ParcelaPag_Le_Todas(objParcelaPagar)
        If lErro <> SUCESSO Then gError 60082
        
        Parcela(1).Caption = iParcela + 1
        
        DataVencParcBaixa.Caption = Format(objParcelaPagar.dtDataVencimento, "dd/mm/yyyy")
        ValorParcBaixa.Caption = Format(objParcelaPagar.dValor, "Standard")
        
        lErro = Carrega_ComboSequencial(objParcelaPagar.lNumIntDoc)
        If lErro <> SUCESSO Then gError 60083

        'Se for uma Parcela Baixada Traz os dados da Baixa
        If Sequencial.ListCount > 0 Then

            'Carrega os Dados da Baixa (1 como Default)
            Call Carrega_Dados_Parcela(objParcelaPagar.lNumIntTitulo, objParcelaPagar.iNumParcela, Sequencial.List(0))

        Else
            Call Limpa_Campos_Parcela
        End If

    End If

    Exit Sub

Erro_UpDownParcela_UpClick:

    Select Case gErr

        Case 60078, 60079, 60081, 60082, 60083 'Tratado na Rotina chamada
        
        Case 60080
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158352)

    End Select

    Exit Sub

End Sub

Public Sub Sequencial_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim iParcela As Integer
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Sequencial_Click

    If Sequencial.ListIndex = -1 Then Exit Sub

    'Critica os Campos se estão Preenchidos
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 60084

    'Verifica se já existe alguma Parcela na Tela
    If Len(Trim(Parcela(1).Caption)) > 0 Then
        iParcela = CInt(Parcela(1).Caption)
    Else
        iParcela = 0
    End If
        
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 60085

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 60086

    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)

    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloPagar.iFilialEmpresa = iFilialEmpresaTela
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Lê o Numint do Título
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 60087

    'Se for uma Parcela Baixada Traz os dados da Baixa
    If Sequencial.ListCount > 0 Then
        
        'Carrega os Dados da Baixa (1 como Default)
        Call Carrega_Dados_Parcela(objTituloPagar.lNumIntDoc, iParcela, Sequencial.List(Sequencial.ListIndex))

    Else
        Call Limpa_Campos_Parcela
    End If

    Exit Sub

Erro_Sequencial_Click:

    Select Case gErr

        Case 60084, 60085, 60087 'Tratado na Rotina chamada
        
        Case 60086
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158353)

    End Select

    Exit Sub

End Sub

Function Carrega_ComboSequencial(lNumIntParcela As Long) As Long

Dim lErro As Long
Dim colSequenciais As New Collection
Dim iIndice  As Integer

On Error GoTo Erro_Carrega_ComboSequencial

    lErro = CF("BaixaParcPag_Le_Sequenciais", lNumIntParcela, colSequenciais)
    If lErro <> SUCESSO Then gError 60088
    
    Sequencial.Clear
        
    'Limpa se não encontrar
    If colSequenciais.Count = 0 Then Exit Function
    
    'Se encontrar Coloca na Combo
    For iIndice = 1 To colSequenciais.Count

        Sequencial.AddItem (CStr(colSequenciais(iIndice)))

    Next
    
    Sequencial.ListIndex = 0

    Carrega_ComboSequencial = SUCESSO

    Exit Function

Erro_Carrega_ComboSequencial:

    Carrega_ComboSequencial = gErr

    Select Case gErr

        Case 60088 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158354)

    End Select

    Exit Function

End Function

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTipoFornecedor As New ClassTipoFornecedor
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 60089

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 60090

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

            'Limpa Combo de Filial
            Filial.Clear

        End If

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True

    Select Case gErr

        Case 60089, 60090 'Tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158355)

    End Select

    Exit Sub

End Sub

Public Sub NumeroTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumeroTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumeroTitulo.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumeroTitulo.ClipText)
    If lErro <> SUCESSO Then gError 60091

    Exit Sub

Erro_NumeroTitulo_Validate:

    Cancel = True


    Select Case gErr

        Case 60091

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158356)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_Tela_TituloPagar

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158357)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Call CF2(Me, "TITPAGC_Form_Unload", Cancel)
    
    Set gobjInfoUsu = Nothing

    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoTipo = Nothing

    Set gobjParcPagBaixa = Nothing
    Set gobjBaixaPagar = Nothing
    
    Set objGridParcelas = Nothing

    Set gobjTelaProjetoInfo = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158358)

     End Select

     Exit Sub

End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame(Opcao.SelectedItem.Index).Visible = True
        Frame(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        Select Case iFrameAtual
        
            Case TAB_Titulos
                Parent.HelpContextID = IDH_TITULOPAG_CONS_TITULO
                
            Case TAB_Parcelas
                Parent.HelpContextID = IDH_TITULOPAG_CONS_PARCELAS
            
            Case TAB_Baixas
                Parent.HelpContextID = IDH_TITULOPAG_CONS_BAIXA
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_TITULOPAG_CONS_CONTABILIZACAO
                        
        End Select

    End If

End Sub

Private Function Carrega_TipoCobranca() As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 60095

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

        Case 60095 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158359)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Suspenso")
    objGridInt.colColuna.Add ("Banco Cobrador")
    objGridInt.colColuna.Add ("Portador")
    objGridInt.colColuna.Add ("Valor Original")
    objGridInt.colColuna.Add ("MotivoD iferença")
    objGridInt.colColuna.Add ("Código de Barras")
    

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (StatusParcela.Name)
    objGridInt.colCampo.Add (SaldoParcela.Name)
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
    iGrid_StatusParcela_Col = 4
    iGrid_SaldoParcela_Col = 5
    iGrid_Cobranca_Col = 6
    iGrid_Suspenso_Col = 7
    iGrid_Banco_Col = 8
    iGrid_Portador_Col = 9
    iGrid_ValorOriginal_Col = 10
    iGrid_MotivoDiferenca_Col = 11
    iGrid_CodigoDeBarras_Col = 12

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 630

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Limpa_Tela_TituloPagar()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Chama função que limpa TextBoxes e MaskedEdits da Tela
    Call Limpa_Tela(Me)

    'Limpa os campos não são limpos pela função acima
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me

    Set gobjParcPagBaixa = New ClassBaixaParcPagar
    Set gobjBaixaPagar = New ClassBaixaPagar
    
    ValorTotal.Caption = ""
    ValorICMS.Caption = ""
    ValorICMSSubst.Caption = ""
    ValorProdutos.Caption = ""
    ValorFrete.Caption = ""
    ValorSeguro.Caption = ""
    OutrasDespesas.Caption = ""
    ValorIPI.Caption = ""
    ValorIRRF.Caption = ""
    PISRetido.Caption = ""
    ISSRetido.Caption = ""
    COFINSRetido.Caption = ""
    CSLLRetido.Caption = ""
    ValorINSS.Caption = ""
    Tipo.Text = ""
    NumPC.Caption = ""
    ComboFilialPC.Caption = ""
    INSSRetido.Value = 0
    Filial.Clear
    CreditoICMS.Value = 0
    CreditoIPI.Value = 0
    HistoricoTit.Caption = ""
    
    Sequencial.Clear
    Parcela(1).Caption = ""
    Call Limpa_Campos_Parcela
    Call Grid_Limpa(objGridParcelas)
    
    iFilialEmpresaTela = 0

    Call CF2(Me, "TITPAGC_Limpa_Tela")
    
    Diferenca.Caption = ""
    
    Saldo.Caption = ""
    Status.Caption = ""
    ValorParcBaixa.Caption = ""
    DataVencParcBaixa.Caption = ""

    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade_Consulta
    
    iAlterado = 0
    
End Sub

Public Function Move_Tela_Memoria(objTituloPagar As ClassTituloPagar, colParcelas As colParcelaPagar) As Long
'Move os dados da Tela para objTituloPagar e colParcelas

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Fornecedor.Text)) > 0 Then
        'Codigo do Fornecedor
        objFornecedor.sNomeReduzido = Fornecedor.Text

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 60096
        
        If lErro <> SUCESSO Then gError 60097

        objTituloPagar.lFornecedor = objFornecedor.lCodigo
    End If

    'A filial do Fornecedor
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    End If

    'FilialPC
    If Len(Trim(ComboFilialPC.Caption)) > 0 Then
        objTituloPagar.iFilialPedCompra = Codigo_Extrai(ComboFilialPC.Caption)
    End If

    'NumPC
    If Len(Trim(NumPC.Caption)) > 0 Then objTituloPagar.lNumPedCompra = CLng(NumPC.Caption)

    'Numero do Titulo
    If Len(Trim(NumeroTitulo.ClipText)) > 0 Then objTituloPagar.lNumTitulo = CLng(NumeroTitulo.ClipText)

    'Data de Emissao
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objTituloPagar.dtDataEmissao = DATA_NULA
    Else
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    End If

    'Os Valores, sSiglaDocumento, FilialEmpresa, CondPagamento
    If Len(Trim(ValorTotal.Caption)) > 0 Then objTituloPagar.dValorTotal = CDbl(ValorTotal.Caption)
    objTituloPagar.iNumParcelas = objGridParcelas.iLinhasExistentes
    If Len(Trim(ValorICMS.Caption)) > 0 Then objTituloPagar.dValorICMS = CDbl(ValorICMS.Caption)
    If Len(Trim(ValorICMSSubst.Caption)) > 0 Then objTituloPagar.dValorICMSSubst = CDbl(ValorICMSSubst.Caption)
    objTituloPagar.iCreditoICMS = CreditoICMS.Value
    If Len(Trim(ValorProdutos.Caption)) > 0 Then objTituloPagar.dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(OutrasDespesas.Caption)) > 0 Then objTituloPagar.dOutrasDespesas = CDbl(OutrasDespesas.Caption)
    If Len(Trim(ValorSeguro.Caption)) > 0 Then objTituloPagar.dValorSeguro = CDbl(ValorSeguro.Caption)
    If Len(Trim(ValorFrete.Caption)) > 0 Then objTituloPagar.dValorFrete = CDbl(ValorFrete.Caption)
    If Len(Trim(ValorIRRF.Caption)) > 0 Then objTituloPagar.dValorIRRF = CDbl(ValorIRRF.Caption)
    If Len(Trim(ValorIPI.Caption)) > 0 Then objTituloPagar.dValorIPI = CDbl(ValorIPI.Caption)
    If Len(Trim(ValorINSS.Caption)) > 0 Then objTituloPagar.dValorINSS = CDbl(ValorINSS.Caption)
    objTituloPagar.iCreditoIPI = CreditoIPI.Value
    If Len(Trim(Tipo.Text)) > 0 Then objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloPagar.iFilialEmpresa = iFilialEmpresaTela
    If Len(Trim(PISRetido.Caption)) > 0 Then objTituloPagar.dPISRetido = CDbl(PISRetido.Caption)
    If Len(Trim(ISSRetido.Caption)) > 0 Then objTituloPagar.dISSRetido = CDbl(ISSRetido.Caption)
    If Len(Trim(COFINSRetido.Caption)) > 0 Then objTituloPagar.dCOFINSRetido = CDbl(COFINSRetido.Caption)
    If Len(Trim(CSLLRetido.Caption)) > 0 Then objTituloPagar.dCSLLRetido = CDbl(CSLLRetido.Caption)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 60096

        Case 60097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158360)

    End Select

    Exit Function

End Function

'Parte da Baixa Vinda da Tela de Cancela Baixa - com algumas adaptações

Public Sub Pagamento_Click(Index As Integer)

    FramePagamento(iFramePagamentoAtual).Visible = False
    FramePagamento(Index).Visible = True
    iFramePagamentoAtual = Index

End Sub

Private Sub Carrega_Dados_Parcela(lNumIntTitulo As Long, iNumParcela As Integer, iSequencialBaixa As Integer)
'Carrega os dados da Parcela

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim objParcPagBaixa As New ClassBaixaParcPagar
Dim objBaixaPagar As New ClassBaixaPagar

On Error GoTo Erro_Carrega_Dados_Parcela

    Call Limpa_Campos_Parcela

    'Lê a Parcela no BD
    lErro = CF("BaixaPagCancelar_Le_Parcela", lNumIntTitulo, iNumParcela, iSequencialBaixa, objParcPagBaixa)
    If lErro <> SUCESSO And lErro <> 42807 Then gError 60097

    If lErro = 42807 Then gError 60098 'Não encontrou

    objBaixaPagar.lNumIntBaixa = objParcPagBaixa.lNumIntBaixa

    'Lê  a Baixa
    lErro = CF("BaixaPagar_Le", objBaixaPagar)
    If lErro <> SUCESSO And lErro <> 42812 Then gError 60099

    If lErro <> SUCESSO Then gError 60100 'Não encontrou

    'Coloca na tela os dados da Baixa Lida
    lErro = Traz_Dados_Baixa(objParcPagBaixa, objBaixaPagar)
    If lErro <> SUCESSO Then gError 60101
    
    Set gobjParcPagBaixa = objParcPagBaixa
    Set gobjBaixaPagar = objBaixaPagar

    Exit Sub

Erro_Carrega_Dados_Parcela:

    Select Case gErr

        Case 60097, 60099, 60101 'Tratados nas rotinas chamadas

        Case 60098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAPARCPAG_INEXISTENTE", gErr)
        
        Case 60100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAPAG_INEXISTENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158361)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Campos_Parcela()
    
    DataBaixa.Caption = ""
    ValorPago.Caption = ""
    ValorBaixado.Caption = ""
    Desconto.Caption = ""
    Multa.Caption = ""
    Juros.Caption = ""

    ContaCorrente.Caption = ""
    ValorPagoPagto.Caption = ""
    Historico.Caption = ""
    Beneficiario.Caption = ""
    Portador.Caption = ""
    NumOuSequencial.Caption = ""
    DataMovimento.Caption = ""
    MeioPagtoDescricao.Caption = ""
    ValorPA.Caption = ""
    CCIntNomeReduzido.Caption = ""
    NumeroMP.Caption = ""
    FilialEmpresaPA.Caption = ""
    DataEmissaoCred.Caption = ""
    NumTitulo.Caption = ""
    ValorCredito.Caption = ""
    SaldoCredito.Caption = ""
    SiglaDocumentoCR.Caption = ""
    FilialEmpresaCR.Caption = ""
    
    BancoCT.Caption = ""
    AgenciaCT.Caption = ""
    ContaCT.Caption = ""
    NumeroCT.Caption = ""
    ValorCT.Caption = ""
    FilialEmpresaCT.Caption = ""
    ClienteCT.Caption = ""
    
    Set gobjParcPagBaixa = New ClassBaixaParcPagar
    Set gobjBaixaPagar = New ClassBaixaPagar

End Sub

Private Function Traz_Dados_Baixa(objParcPagBaixa As ClassBaixaParcPagar, objBaixaPagar As ClassBaixaPagar) As Long
'Mostra na tela os dados da baixa

Dim lErro As Long

On Error GoTo Erro_Traz_Dados_Baixa

    'Coloca os dados da Baixa na tela
    Desconto.Caption = Format(objParcPagBaixa.dValorDesconto, "Standard")
    ValorPago.Caption = Format(objParcPagBaixa.dValorBaixado - objParcPagBaixa.dValorDesconto + objParcPagBaixa.dValorMulta + objParcPagBaixa.dValorJuros, "Standard")
    Multa.Caption = Format(objParcPagBaixa.dValorMulta, "Standard")
    ValorBaixado.Caption = Format(objParcPagBaixa.dValorBaixado, "Standard")
    Juros.Caption = Format(objParcPagBaixa.dValorJuros, "Standard")
    DataBaixa.Caption = Format(objBaixaPagar.dtData, "dd/mm/yyyy")

    If objBaixaPagar.iMotivo = MOTIVO_PAGAMENTO Then
        Pagamento(0).Value = True
        Call Pagamento_Click(0)

        'Traz os dados do pagamento
        lErro = Traz_Dados_Pagamento(objBaixaPagar)
        If lErro <> SUCESSO Then gError 60102

    ElseIf objBaixaPagar.iMotivo = MOTIVO_PAGTO_ANTECIPADO Then
        Pagamento(1).Value = True
        Call Pagamento_Click(1)

        'Traz os dados do pagamento antecipado
        lErro = Traz_Dados_Pagamento_Antecipado(objBaixaPagar)
        If lErro <> SUCESSO Then gError 60103

    ElseIf objBaixaPagar.iMotivo = MOTIVO_CREDITO_FORNECEDOR Then
        Pagamento(2).Value = True
        Call Pagamento_Click(2)

        'Traz os dados do crédito
        lErro = Traz_Dados_Credito_Fornecedor(objBaixaPagar)
        If lErro <> SUCESSO Then gError 60104
        
    '#############################################################
    'Inserido por Wagner 21/06/2006
    ElseIf objBaixaPagar.iMotivo = MOTIVO_CHEQUE_DE_TERCEIROS Then
        Pagamento(3).Value = True
        Call Pagamento_Click(3)

        'Traz os dados do crédito
        lErro = Traz_Dados_ChequePre(objBaixaPagar)
        If lErro <> SUCESSO Then gError 60104
    '#############################################################

    End If

    Traz_Dados_Baixa = SUCESSO

    Exit Function

Erro_Traz_Dados_Baixa:

    Traz_Dados_Baixa = gErr

    Select Case gErr

        Case 60102, 60103, 60104

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158362)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Pagamento(objBaixaPagar As ClassBaixaPagar) As Long
'Mostra os dados de pagamento na tela

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim objBorderoPagto As New ClassBorderoPagto
Dim objPortador As New ClassPortador
Dim objChequeDet As New ClassChequeDet

On Error GoTo Erro_Traz_Dados_Pagamento

    objMovCCI.lNumMovto = objBaixaPagar.lNumMovConta

    'Lê o Movimento
    lErro = CF("MovContaCorrente_Le", objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then gError 60105

    'Se não encontrou Movimento --> erro
    If lErro = 11893 Then gError 60106

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 60107

    'Se não encontrou a conta Corrente --> erro
    If lErro <> SUCESSO Then gError 60108

    'Coloca os dados na tela
    ContaCorrente.Caption = objContaCorrente.sNomeReduzido
    ValorPagoPagto.Caption = Format(objMovCCI.dValor, "Standard")
    Historico.Caption = objMovCCI.sHistorico
    Beneficiario.Caption = ""
    NumOuSequencial.Caption = IIf(objMovCCI.lNumero <> 0, CStr(objMovCCI.lNumero), "")

    objPortador.iCodigo = objMovCCI.iPortador

    If objPortador.iCodigo <> 0 Then

        'Lê o Portador
        lErro = CF("Portador_Le", objPortador)
        If lErro <> SUCESSO And lErro <> 15971 Then gError 60109

        'Se não achou o Portador --> erro
        If lErro <> SUCESSO Then gError 60110

        Portador.Caption = objMovCCI.iPortador & SEPARADOR & objPortador.sNomeReduzido

    End If

    objBorderoPagto.dtDataEmissao = objMovCCI.dtDataMovimento
    objBorderoPagto.iCodConta = objMovCCI.iCodConta
    objBorderoPagto.lNumero = objMovCCI.lNumero

    If objMovCCI.iTipoMeioPagto = Cheque Then
        TipoMeioPagto(0).Value = True
        objChequeDet.lNumMovto = objMovCCI.lNumMovto
        lErro = CF("ChequeDet_Le", objChequeDet)
        If lErro = SUCESSO Then Beneficiario.Caption = objChequeDet.sEmNomeDeQuem

    ElseIf objMovCCI.iTipoMeioPagto = BORDERO Then
        TipoMeioPagto(1).Value = True

    ElseIf objMovCCI.iTipoMeioPagto = DINHEIRO Then
        TipoMeioPagto(2).Value = True
    End If

    Traz_Dados_Pagamento = SUCESSO

    Exit Function

Erro_Traz_Dados_Pagamento:

    Traz_Dados_Pagamento = gErr

    Select Case gErr

        Case 60105, 60107, 60109

        Case 60106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", gErr, objMovCCI.lNumMovto)

        Case 60108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContaCorrente.iCodigo)

        Case 60110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO1", gErr, objPortador.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158363)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Credito_Fornecedor(objBaixaPagar As ClassBaixaPagar) As Long

Dim lErro As Long
Dim objCreditoPag As New ClassCreditoPagar

On Error GoTo Erro_Traz_Dados_Credito_Fornecedor

    objCreditoPag.lNumIntDoc = objBaixaPagar.lNumIntDoc

    'Lê o Crédito Pagar
    lErro = CF("CreditoPagar_Le", objCreditoPag)
    If lErro <> AD_SQL_SUCESSO And 17071 Then gError 60111
    
    If lErro <> SUCESSO Then gError 60112

    'Coloca os dados na tela
    DataEmissaoCred.Caption = Format(objCreditoPag.dtDataEmissao, "dd/mm/yyyy")
    NumTitulo.Caption = objCreditoPag.lNumTitulo
    SaldoCredito.Caption = Format(objCreditoPag.dSaldo, "Standard")
    SiglaDocumentoCR.Caption = objCreditoPag.sSiglaDocumento
    ValorCredito.Caption = Format(objCreditoPag.dValorTotal, "Standard")
    FilialEmpresaCR.Caption = objCreditoPag.iFilialEmpresa

    Traz_Dados_Credito_Fornecedor = SUCESSO

    Exit Function

Erro_Traz_Dados_Credito_Fornecedor:

    Traz_Dados_Credito_Fornecedor = gErr

    Select Case gErr

        Case 60111

        Case 60112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CREDITO_PAG_FORN_INEXISTENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158364)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Pagamento_Antecipado(objBaixaPagar As ClassBaixaPagar) As Long

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objCCI As New ClassContasCorrentesInternas
Dim objAntecipPag As New ClassAntecipPag

On Error GoTo Erro_Traz_Dados_Pagamento_Antecipado

    objMovCCI.lNumMovto = objBaixaPagar.lNumMovConta

    'Lê o movimento  da Baixa
    lErro = CF("MovContaCorrente_Le", objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then gError 60113
    
    If lErro = 11893 Then gError 60114 'Não encontrou

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then gError 60115
    
    If lErro <> SUCESSO Then gError 60116 'Não encontrou

    objAntecipPag.lNumMovto = objMovCCI.lNumMovto

    lErro = CF("AntecipPag_Le_NumMovto", objAntecipPag)
    If lErro <> AD_SQL_SUCESSO And lErro <> 42845 Then gError 60117
    
    If lErro = 42845 Then gError 60118

    'Coloca os dados na tela
    DataMovimento.Caption = Format(objMovCCI.dtDataMovimento, "dd/mm/yyyy")
    ValorPA.Caption = Format(objMovCCI.dValor, "Standard")
    FilialEmpresaPA.Caption = objMovCCI.iFilialEmpresa
    CCIntNomeReduzido.Caption = objCCI.sNomeReduzido
    NumeroMP.Caption = objMovCCI.lNumero
    If objMovCCI.iTipoMeioPagto = DINHEIRO Then
        MeioPagtoDescricao.Caption = "Dinheiro"
    ElseIf objMovCCI.iTipoMeioPagto = Cheque Then
        MeioPagtoDescricao.Caption = "Cheque"
    ElseIf objMovCCI.iTipoMeioPagto = BORDERO Then
        MeioPagtoDescricao.Caption = "Borderô"
    End If

    Traz_Dados_Pagamento_Antecipado = SUCESSO

    Exit Function

Erro_Traz_Dados_Pagamento_Antecipado:

    Traz_Dados_Pagamento_Antecipado = gErr

    Select Case gErr

        Case 60113, 60115, 60117

        Case 60114
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", gErr, objMovCCI.lNumMovto)

        Case 60116
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objCCI.iCodigo)

        Case 60118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAGTO_ANTECIPADO_INEXISTENTE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158365)

    End Select

    Exit Function

End Function

'Contabilidade
Public Sub CTBBotaoImprimir_Click()

    Call objContabil.Contabil_BotaoImprimir_Consulta_Click

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TITULOPAG_CONS_TITULO
    Set Form_Load_Ocx = Me
    Caption = "Consulta de Títulos a Pagar"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TituloPagar_Consulta"

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
        ElseIf Me.ActiveControl Is NumeroTitulo Then
            Call NumeroLabel_Click
        End If

    End If

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

Private Function Carrega_Bancos() As Long
'Carrega a combo de Cobrador

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Bancos

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "Bancos", "CodBanco", "NomeReduzido", STRING_NOME_REDUZIDO, colCodigoNome)
    If lErro <> SUCESSO Then gError 60876

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

        Case 60876

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158366)

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
    If lErro <> SUCESSO Then gError 60877

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

        Case 60877

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158367)

    End Select

    Exit Function

End Function

Public Sub Fornecedor_Change()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Change
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134055

    Exit Sub

Erro_Fornecedor_Change:

    Select Case gErr

        Case 134055

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158368)

    End Select
    
    Exit Sub

End Sub

'###################################################################
'Inserido por Wagner 21/06/2006
Private Function Traz_Dados_ChequePre(objBaixaPagar As ClassBaixaPagar) As Long

Dim lErro As Long
Dim objChequepre As New ClassChequePre
Dim objCliente As New ClassCliente

On Error GoTo Erro_Traz_Dados_ChequePre

    objChequepre.lNumIntCheque = objBaixaPagar.lNumIntCheque

    'Lê o Crédito Pagar
    lErro = CF("ChequePre_Le", objChequepre)
    If lErro <> AD_SQL_SUCESSO And lErro <> 17642 Then gError 180296

    BancoCT.Caption = objChequepre.iBanco
    AgenciaCT.Caption = objChequepre.sAgencia
    ContaCT.Caption = objChequepre.sContaCorrente
    NumeroCT.Caption = objChequepre.lNumero
    ValorCT.Caption = Format(objChequepre.dValor, "STANDARD")
    FilialEmpresaCT.Caption = objChequepre.iFilialEmpresa
    
    objCliente.lCodigo = objChequepre.lCliente
    
    lErro = CF("Cliente_Le", objCliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 180297
    
    ClienteCT.Caption = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido

    Traz_Dados_ChequePre = SUCESSO

    Exit Function

Erro_Traz_Dados_ChequePre:

    Traz_Dados_ChequePre = gErr

    Select Case gErr

        Case 180296, 180297

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180298)

    End Select

    Exit Function

End Function
'#################################################################################


'#####################################
'Inserido por Wagner 21/08/2006
Public Function BotaoProjetos_Click() As Long
        Call gobjTelaProjetoInfo.BotaoProjetos_Click(True)
End Function

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR
   
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "PROD"
    objItensPRJCR.sDescricao = "Produtos"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorProdutos.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
   Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "FR"
    objItensPRJCR.sDescricao = "Frete"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorFrete.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "SG"
    objItensPRJCR.sDescricao = "Seguro"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorSeguro.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "OD"
    objItensPRJCR.sDescricao = "Outras Despesas"

    objItensPRJCR.dValorOriginal = StrParaDbl(OutrasDespesas.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "ICMS"
    objItensPRJCR.sDescricao = "ICMS Substituto"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorICMSSubst.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "IPI"
    objItensPRJCR.sDescricao = "IPI"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorIPI.Caption)

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

'############################################################
'Inserido por Wagner 08/03/2006
Public Sub BotaoConsultaNFPag_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim objNFsPag As New ClassNFsPag
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoConsultaNFPag_Click

    'Critica se os Campos estão Preenchidos(Fornecedor, Filial, Tipo, Número)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 141864

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 141865

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 141866
    
    If iFilialEmpresaTela <> giFilialEmpresa Then gError 185050

    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloPagar.iFilialEmpresa = iFilialEmpresaTela

    If Len(Trim(DataEmissao.ClipText)) <> 0 Then
        objTituloPagar.dtDataEmissao = MaskedParaDate(DataEmissao)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 141867
    
    objNFsPag.lFornecedor = objTituloPagar.lFornecedor
    objNFsPag.iFilial = objTituloPagar.iFilial

    'Adiciona filtros: lFornecedor e iFilial
    colSelecao.Add objTituloPagar.lNumIntDoc
    colSelecao.Add objNFsPag.lFornecedor
    'colSelecao.Add objNFsPag.iFilial

    'Chama Tela NFFatPagLista
    Call Chama_Tela("NFPag_TodosLista", colSelecao, objNFsPag, Nothing, "NumIntTitPag = ?")

    Exit Sub

Erro_BotaoConsultaNFPag_Click:

    Select Case gErr

        Case 141864, 141865, 141867 'Tratado na rotina chamada

        Case 141866
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case 185050
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DIFERENTE", gErr, iFilialEmpresaTela, giFilialEmpresa)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 141868)

    End Select

    Exit Sub

End Sub
'############################################################

Function TITPAGC_Form_Load(ByVal objCT As Object) As Long
    TITPAGC_Form_Load = SUCESSO
End Function

Function TITPAGC_Form_Unload(ByVal objCT As Object, Cancel As Integer) As Long
    TITPAGC_Form_Unload = SUCESSO
End Function

Function TITPAGC_Limpa_Tela(ByVal objCT As Object) As Long
    TITPAGC_Limpa_Tela = SUCESSO
End Function

Public Function TITPAGC_Saida_Celula(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
    TITPAGC_Saida_Celula = SUCESSO
End Function

Function TITPAGC_Traz_TituloPagar_Tela(ByVal objCT As Object, ByVal objTituloPagar As ClassTituloPagar) As Long
    TITPAGC_Traz_TituloPagar_Tela = SUCESSO
End Function

'########################################
'Inserido por Wagner
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iUltimaLinha As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 132253

        lErro = CF2(Me, "TITPAGC_Saida_Celula", objGridInt)
        If lErro <> SUCESSO Then gError 132254
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 132260

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 132253 To 132254

        Case 132260
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

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

Public Sub BotaoImprimirRecibo_Click()
Dim lErro As Long
Dim iParcela As Integer
Dim objParcPagBaixa As New ClassBaixaParcPagar
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim objBaixaPagar As New ClassBaixaPagar, lNumIntRel As Long
Dim objRelatorio As New AdmRelatorio

On Error GoTo Erro_BotaoImprimirRecibo_Click

    If Sequencial.ListIndex = -1 Then gError 202236

    'Critica os Campos se estão Preenchidos
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 182869

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 60022

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 60023

    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloPagar.iFilialEmpresa = iFilialEmpresaTela

    If Len(Trim(DataEmissao.ClipText)) <> 0 Then
        objTituloPagar.dtDataEmissao = MaskedParaDate(DataEmissao)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 60024
    
    'Verifica se já existe alguma Parcela na Tela
    If Len(Trim(Parcela(1).Caption)) > 0 Then
        iParcela = CInt(Parcela(1).Caption)
    Else
        iParcela = 0
    End If

    If Sequencial.ListCount = 0 Then gError 202237

    'Se for uma Parcela Baixada Traz os dados da Baixa
    If Len(Trim(Sequencial.Text)) = 0 Then gError 202236

    'Lê a Parcela no BD
    lErro = CF("BaixaPagCancelar_Le_Parcela", objTituloPagar.lNumIntDoc, iParcela, Sequencial.List(Sequencial.ListIndex), objParcPagBaixa)
    If lErro <> SUCESSO And lErro <> 42807 Then gError 60097

    If lErro = 42807 Then gError 182873 'Não encontrou

    objBaixaPagar.lNumIntBaixa = objParcPagBaixa.lNumIntBaixa

    'Lê  a Baixa
    lErro = CF("BaixaPagar_Le", objBaixaPagar)
    If lErro <> SUCESSO And lErro <> 42812 Then gError 182869

    If lErro <> SUCESSO Then gError 182873 'Não encontrou
    
    'Se encontrou a Parcela
    If lErro = SUCESSO Then

        'Se encontrou e Baixa já está cancelada --> Erro
        If objParcPagBaixa.iStatus = STATUS_EXCLUIDO Then gError 182873

        lErro = CF("RelReciboContabil_Prepara", objBaixaPagar.lNumMovConta, lNumIntRel)
        If lErro <> SUCESSO Then gError 184446
        
        lErro = objRelatorio.ExecutarDireto("Recibos Com Contabilização", "", 1, "", "NNUMINTREL", CStr(lNumIntRel))
        If lErro <> SUCESSO Then gError 184447

    End If

    Exit Sub

Erro_BotaoImprimirRecibo_Click:

    Select Case gErr
    
        Case 182869, 182870, 182872, 182874
    
        Case 182871
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", gErr)

        Case 182873
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXA_PARCELA_CANCELADA", gErr, iParcela, Sequencial.List(Sequencial.ListIndex))

        Case 202236
            Call Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_PREENCHIDO", gErr)

        Case 202237
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_SEM_BAIXA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 182875)

    End Select

    Exit Sub
    
End Sub

Public Sub BotaoCTBBaixa_Click()
    Call CF("Lancamentos_Abre_Tela", TRANSACAOCTBORIGEM_BAIXASPAG, gobjBaixaPagar.lNumIntBaixa, TRANSACAOCTBORIGEM_BAIXASPARCPAG, gobjParcPagBaixa.lNumIntDoc)
End Sub

Private Function Carrega_MotivoDiferenca() As Long
'Carrega a combo de Cobrador

Dim lErro As Long
Dim colCodigoNome As New AdmColCodigoNome
Dim objCodigoNome As New AdmCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_MotivoDiferenca

    'Leitura dos códigos e descrições dos Bancos BD
    lErro = CF("Cod_Nomes_Le", "MotivoDiferenca", "Codigo", "Descricao", STRING_DESCRICAO_CAMPO, colCodigoNome)
    If lErro <> SUCESSO Then gError 500072

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

        Case 500072 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

Public Sub BotaoAnexos_Click()
Dim objAnexos As New ClassAnexos
Dim objTituloPagar As New ClassTituloPagar
Dim lErro As Long
Dim objNFiscal As New ClassNFiscal, sTela As String

On Error GoTo Erro_BotaoAnexos_Click

    lErro = TitPag_Obtem(objTituloPagar)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'objAnexos.iTipoDoc = ANEXO_TIPO_TITPAG
    'objAnexos.lNumIntDoc = objTituloPagar.lNumIntDoc
    
    lErro = CF("Titulo_Le_DocumentoOriginal", objTituloPagar.lNumIntDoc, CPR_TITULO_PAGAR, objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 58942 Then gError ERRO_SEM_MENSAGEM
    If lErro = SUCESSO Then
        objAnexos.iTipoDoc = ANEXO_TIPO_NF
        objAnexos.lNumIntDoc = objNFiscal.lNumIntDoc
    Else
        objAnexos.iTipoDoc = ANEXO_TIPO_TITPAG
        objAnexos.lNumIntDoc = objTituloPagar.lNumIntDoc
    End If

    lErro = CF("Anexos_Le", objAnexos)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
   Call Chama_Tela_Modal("Anexos", objAnexos)
   
    Exit Sub
    
Erro_BotaoAnexos_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201570)

    End Select
    
    Exit Sub

End Sub

Private Function TitPag_Obtem(ByVal objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long
Dim sTela As String
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_TitPag_Obtem
    
    'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número, DataEmissao)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 60037

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 60038

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 60039
    
    If iFilialEmpresaTela <> giFilialEmpresa Then gError 185050
    
    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloPagar.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloPagar.iFilialEmpresa = iFilialEmpresaTela
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = CF("TituloPag_Le_Todos", objTituloPagar)
    If lErro <> SUCESSO Then gError 60040

    TitPag_Obtem = SUCESSO
    
    Exit Function
    
Erro_TitPag_Obtem:

    TitPag_Obtem = gErr

    Select Case gErr

        Case 60037, 60038, 60040, 60042 'Tratado na Rotina chamada
        
        Case 60039
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 60041, 60043
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAG_SEM_DOCORIGINAL", gErr, objTituloPagar.lNumTitulo)

        Case 185050
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DIFERENTE", gErr, iFilialEmpresaTela, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158338)

    End Select

    Exit Function

End Function

