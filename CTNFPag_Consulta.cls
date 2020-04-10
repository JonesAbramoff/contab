VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTNFPag_Consulta"
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
Dim iGrid_Cobranca_Col As Integer
Dim iGrid_Suspenso_Col As Integer

Dim glNumIntDocFat As Long

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

'inicio contabilidade
Dim objGrid1 As AdmGrid
Public objContabil As New ClassContabil
'Fim da Contabilidade

'Constantes públicas dos tabs
Private Const TAB_NF = 1
Private Const TAB_Fatura = 2
Private Const TAB_Contabilizacao = 3

Private mvariFilialEmpresaTela As Integer

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

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get BotaoProjetos() As Object
     Set BotaoProjetos = objUserControl.Controls("BotaoProjetos")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
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

Public Property Get ValorFrete() As Object
     Set ValorFrete = objUserControl.Controls("ValorFrete")
End Property

Public Property Get ValorSeguro() As Object
     Set ValorSeguro = objUserControl.Controls("ValorSeguro")
End Property

Public Property Get OutrasDespesas() As Object
     Set OutrasDespesas = objUserControl.Controls("OutrasDespesas")
End Property

Public Property Get ValorIRRF() As Object
     Set ValorIRRF = objUserControl.Controls("ValorIRRF")
End Property

Public Property Get BotaoDocOriginal() As Object
     Set BotaoDocOriginal = objUserControl.Controls("BotaoDocOriginal")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get LabelProjeto() As Object
     Set LabelProjeto = objUserControl.Controls("LabelProjeto")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get CTBGerencial() As Object
     Set CTBGerencial = objUserControl.Controls("CTBGerencial")
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

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
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

Public Property Get CTBLote() As Object
     Set CTBLote = objUserControl.Controls("CTBLote")
End Property

Public Property Get CTBDocumento() As Object
     Set CTBDocumento = objUserControl.Controls("CTBDocumento")
End Property

Public Property Get CTBDataContabil() As Object
     Set CTBDataContabil = objUserControl.Controls("CTBDataContabil")
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

Public Property Get BotaoConsultaFatura() As Object
     Set BotaoConsultaFatura = objUserControl.Controls("BotaoConsultaFatura")
End Property

Public Property Get ValorTotalFatura() As Object
     Set ValorTotalFatura = objUserControl.Controls("ValorTotalFatura")
End Property

Public Property Get SaldoFatura() As Object
     Set SaldoFatura = objUserControl.Controls("SaldoFatura")
End Property

Public Property Get StatusFatura() As Object
     Set StatusFatura = objUserControl.Controls("StatusFatura")
End Property

Public Property Get EmissaoFatura() As Object
     Set EmissaoFatura = objUserControl.Controls("EmissaoFatura")
End Property

Public Property Get NumeroFatura() As Object
     Set NumeroFatura = objUserControl.Controls("NumeroFatura")
End Property

Public Property Get TipoFatura() As Object
     Set TipoFatura = objUserControl.Controls("TipoFatura")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get BotaoConsulta() As Object
     Set BotaoConsulta = objUserControl.Controls("BotaoConsulta")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get NumeroNota() As Object
     Set NumeroNota = objUserControl.Controls("NumeroNota")
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

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get Line1() As Object
     Set Line1 = objUserControl.Controls("Line1")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsulta_Click

    'Critica se os Campos estão Preenchidos(Fornecedor, Filial, Tipo, Número)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then Error 60132

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 60133

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then Error 60134

    objNFsPag.lFornecedor = objFornecedor.lCodigo
    objNFsPag.iFilial = Codigo_Extrai(Filial.Text)
    objNFsPag.lNumNotaFiscal = CLng(NumeroNota.Text)
    objNFsPag.iFilialEmpresa = iFilialEmpresaTela 'giFilialEmpresa
    objNFsPag.dtDataEmissao = MaskedParaDate(DataEmissao)

    'Procura o Titulo (Baixados ou não)
    lErro = NFPag_Le_Todos(objNFsPag)
    If lErro <> SUCESSO Then Error 60135

    'Traz o Titulo para a Tela
    lErro = Traz_NotaPagar_Tela(objNFsPag)
    If lErro <> SUCESSO Then Error 60136

    Exit Sub

Erro_BotaoConsulta_Click:

    Select Case Err

        Case 60132, 60133, 60135, 60136 'Tratado na rotina chamada

        Case 60134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163527)

    End Select

    Exit Sub

End Sub

Private Function NFPag_Le_Todos(objNFsPag As ClassNFsPag) As Long
'Lê em NFPag se não encontrar procura em NFPag Baixados

Dim lErro As Long

On Error GoTo Erro_NFPag_Le_Todos

    If objNFsPag.lNumIntDoc = 0 Then

        'Lê em NFPags
        lErro = CF("NFPag_Le_Numero", objNFsPag)
        If lErro <> SUCESSO And lErro <> 18338 Then Error 60137
    
        'Se não encontrou
        If lErro = 18338 Then
    
            'Procura em Títulos a Receber Baixados
            lErro = CF("NFPagBaixada_Le_Numero", objNFsPag)
            If lErro <> SUCESSO And lErro <> 18344 Then Error 60138
    
            'Se não encontrou nos dois ---> ERRO
            If lErro = 18344 Then Error 60139
    
        End If
        
    Else
    
        lErro = CF("NFPag_Le_Todos2", objNFsPag)
        If lErro <> SUCESSO And lErro <> 18257 Then gError 60137
    
        'Se não encontrou nos dois ---> ERRO
        If lErro = 18257 Then Error 60139
    
    End If

    NFPag_Le_Todos = SUCESSO

    Exit Function

Erro_NFPag_Le_Todos:

    NFPag_Le_Todos = Err

    Select Case Err

        Case 60137, 60138 'Tratado na rotina chamada

        Case 60139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFPAG_NAO_CADASTRADA1", Err, objNFsPag.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163528)

    End Select

    Exit Function

End Function

Private Function Critica_CamposPreenchidos() As Long
'Verifica se os campos chaves foram preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposPreenchidos

    'Se Fornecedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 60140

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then Error 60141

    'Verifica se o Titulo está Preenchido
    If Len(Trim(NumeroNota.Text)) = 0 Then Error 60142

    Critica_CamposPreenchidos = SUCESSO

    Exit Function

Erro_Critica_CamposPreenchidos:

    Critica_CamposPreenchidos = Err

    Select Case Err

        Case 60140
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 60141
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 60142
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163529)

    End Select

    Exit Function

End Function

Public Sub BotaoConsultaFatura_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsultaFatura_Click

    If Len(Trim(TipoFatura.Caption)) > 0 And Len(Trim(NumeroFatura.Caption)) > 0 And Len(Trim(EmissaoFatura.Caption)) > 0 Then
        
        objTituloPagar.sSiglaDocumento = TipoFatura.Caption
        objTituloPagar.lNumTitulo = CLng(NumeroFatura.Caption)
        objTituloPagar.dtDataEmissao = CDate(EmissaoFatura.Caption)
        
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 60360
        
        If lErro = 6681 Then Error 60361
        
        objTituloPagar.lFornecedor = CLng(objFornecedor.lCodigo)
        objTituloPagar.iFilial = CInt(Filial.ItemData(Filial.ListIndex))
        
        objTituloPagar.lNumIntDoc = glNumIntDocFat
        
        Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    End If
    
    Exit Sub
    
Erro_BotaoConsultaFatura_Click:

    Select Case Err
    
        Case 60360
        
        Case 60361
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163530)

    End Select

    Exit Sub

End Sub

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
    If lErro <> SUCESSO Then Error 60143

    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 60144

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then Error 60145

    objNFsPag.lFornecedor = objFornecedor.lCodigo
    objNFsPag.iFilial = Codigo_Extrai(Filial.Text)
    objNFsPag.lNumNotaFiscal = CLng(NumeroNota.Text)
    objNFsPag.iFilialEmpresa = iFilialEmpresaTela ' giFilialEmpresa
    
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objNFsPag.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objNFsPag.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (Baixados ou não)
    lErro = NFPag_Le_Todos(objNFsPag)
    If lErro <> SUCESSO Then Error 60146

    'Le o Nome da Tela que originou este Título
    lErro = CF("Titulo_Le_DocumentoOriginal", objNFsPag.lNumIntDoc, CPR_NF_PAGAR, objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 58942 Then Error 60147

    If lErro = SUCESSO Then
        'Chama a Tela Estoque
        Call Chama_Tela(sTela, objNFiscal)
    Else
        Error 60148
    End If

    Exit Sub

Erro_BotaoDocOriginal_Click:

    Select Case Err

        Case 60143, 60144, 60146, 60147 'Tratado na Rotina chamada

        Case 60145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case 60148
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NFPAG_SEM_DOCORIGINAL", Err, objNFsPag.lNumNotaFiscal)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163531)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Consulta_Click

End Sub

Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao)

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 60149

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True

    Select Case Err

        Case 60149 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163532)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui "1" da data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 60150

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 60150 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163533)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Adiciona "1" a data
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 60151

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 60151 'Tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163534)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 60152

    'Se não encontra valor que era CÓDIGO
    If lErro = 6730 Then

        'Verifica de o fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 60153

        sFornecedor = Fornecedor.Text
        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 60154

        'Se não encontrou
        If lErro = 18272 Then

            objFornecedor.sNomeReduzido = sFornecedor

            'Le o Código do Fornecedor --> Para Passar para a Tela de Filiais
            lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
            If lErro <> SUCESSO And lErro <> 6681 Then Error 60155

            'Passa o Código do Fornecedor
            objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo

            'Sugere cadastrar nova Filial
            Error 60156

        End If

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não encontrou valor informado que era STRING
    If lErro = 6731 Then Error 60157

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case Err

       Case 60152, 60154, 60155

       Case 60153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

       Case 60156
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 60157
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163535)

    End Select

    Exit Sub
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objEventoNumero = New AdmEvento
    Set objEventoFornecedor = New AdmEvento

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me

    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade_Consulta(Me, objGrid1, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then Error 60158
    
    iFilialEmpresaTela = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 60158 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163536)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "NFsPag_Todas"

    'Lê os dados da Tela Notas Fiscais à Pagar
    lErro = Move_Tela_Memoria(objNFsPag)
    If lErro <> SUCESSO Then Error 60159

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

        Case 60159 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163537)

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

        'Lê o Título
        lErro = NFPag_Le_Todos(objNFsPag)
        If lErro <> SUCESSO Then gError 126827

        'Preenche a Tela
        lErro = Traz_NotaPagar_Tela(objNFsPag)
        If lErro <> SUCESSO Then gError 60160

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 60160, 126827

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163538)

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

Dim objNFsPag As New ClassNFsPag
Dim colParcelas As New colParcelaPagar
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click

    'Se Forncedor estiver vazio, erro
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 60161

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then Error 60162

    'Move os dados da Tela para objNFsPag e colParcelas
    lErro = Move_Tela_Memoria(objNFsPag)
    If lErro <> SUCESSO Then Error 60163

    'Adiciona filtros: lFornecedor e iFilial
    colSelecao.Add giFilialEmpresa
'    colSelecao.Add objNFsPag.lFornecedor
    colSelecao.Add objNFsPag.iFilial
    colSelecao.Add objNFsPag.lFornecedor

    'Chama Tela NFFatPagLista
    Call Chama_Tela("NFPag_TodosLista", colSelecao, objNFsPag, objEventoNumero, "FilialEmpresa = ? AND Filial = ?")

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 60161
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 60162
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 60163 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163539)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNFsPag As ClassNFsPag

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objNFsPag = obj1

    'Lê o Título
    lErro = NFPag_Le_Todos(objNFsPag)
    If lErro <> SUCESSO Then gError 126828

    'Traz os dados de objNFsPag para Tela
    lErro = Traz_NotaPagar_Tela(objNFsPag)
    If lErro <> SUCESSO Then gError 60164

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 60164, 126828

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163540)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objNFsPag As ClassNFsPag) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objNFsPag Is Nothing) Then

        'Lê o Título
        lErro = NFPag_Le_Todos(objNFsPag)
        If lErro <> SUCESSO Then Error 60165

        'Traz os dados para a Tela
        lErro = Traz_NotaPagar_Tela(objNFsPag)
        If lErro <> SUCESSO Then Error 60166

    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 60165, 60166 'Tratado na Rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163541)

    End Select

    Exit Function

End Function

Private Function Traz_NotaPagar_Tela(objNFsPag As ClassNFsPag) As Long
'Traz os dados da Nota Fiscal Fatura para a Tela

Dim lErro As Long
Dim bCancel As Boolean
Dim sTela As String
Dim sModulo As String
Dim objNFiscal As New ClassNFiscal
Dim lNumIntDoc As Long
Dim objTituloPagar As New ClassTituloPagar
Dim iFilialAux As Integer

On Error GoTo Erro_Traz_NotaPagar_Tela

    'Limpa a Tela
    Call Limpa_Tela_NotaPagar
    
    iFilialEmpresaTela = objNFsPag.iFilialEmpresa

    iFilialAux = giFilialEmpresa
    giFilialEmpresa = iFilialEmpresaTela

    'Coloca os dados do Título Pagar na Tela
    NumeroNota.Text = objNFsPag.lNumNotaFiscal

    Fornecedor.Text = objNFsPag.lFornecedor
    Call Fornecedor_Validate(bCancel)

    Filial.Text = objNFsPag.iFilial
    Call Filial_Validate(bCancel)

    'Preenche a data de Emissao
    If objNFsPag.dtDataEmissao <> DATA_NULA Then
        DataEmissao.PromptInclude = False
        DataEmissao.Text = Format(objNFsPag.dtDataEmissao, "dd/mm/yy")
        DataEmissao.PromptInclude = True
    End If

    'Preenche a data de Vcto
    If objNFsPag.dtDataVencimento <> DATA_NULA Then
        DataVencimento.Caption = Format(objNFsPag.dtDataVencimento, "dd/mm/yy")
    End If

    If objNFsPag.dValorTotal > 0 Then ValorTotal.Caption = Format(objNFsPag.dValorTotal, "Standard")
    If objNFsPag.dValorICMS > 0 Then ValorICMS.Caption = Format(objNFsPag.dValorICMS, "Standard")
    If objNFsPag.dValorICMSSubst > 0 Then ValorICMSSubst.Caption = Format(objNFsPag.dValorICMSSubst, "Standard")
    CreditoICMS.Value = objNFsPag.iCreditoICMS
    If objNFsPag.dValorProdutos > 0 Then ValorProdutos.Caption = Format(objNFsPag.dValorProdutos, "Standard")
    If objNFsPag.dValorFrete > 0 Then ValorFrete.Caption = Format(objNFsPag.dValorFrete, "Standard")
    If objNFsPag.dValorSeguro > 0 Then ValorSeguro.Caption = Format(objNFsPag.dValorSeguro, "Standard")
    If objNFsPag.dOutrasDespesas > 0 Then OutrasDespesas.Caption = Format(objNFsPag.dOutrasDespesas, "Standard")
    If objNFsPag.dValorIPI > 0 Then ValorIPI.Caption = Format(objNFsPag.dValorIPI, "Standard")
    CreditoIPI.Value = objNFsPag.iCreditoIPI
    If objNFsPag.dValorIRRF > 0 Then ValorIRRF.Caption = Format(objNFsPag.dValorIRRF, "Standard")
        
    If objNFsPag.dValorINSS > 0 Then ValorINSS.Caption = Format(objNFsPag.dValorINSS, "Standard")
    If objNFsPag.dPISRetido > 0 Then PISRetido.Caption = Format(objNFsPag.dPISRetido, "Standard")
    If objNFsPag.dISSRetido > 0 Then ISSRetido.Caption = Format(objNFsPag.dISSRetido, "Standard")
    If objNFsPag.dCOFINSRetido > 0 Then COFINSRetido.Caption = Format(objNFsPag.dCOFINSRetido, "Standard")
    If objNFsPag.dCSLLRetido > 0 Then CSLLRetido.Caption = Format(objNFsPag.dCSLLRetido, "Standard")
    INSSRetido.Value = objNFsPag.iINSSRetido
        
    Historico.Caption = objNFsPag.sHistorico
        
    'se a nf já está associada à fatura
    If objNFsPag.lNumIntTitPag <> 0 Then
    
        objTituloPagar.lNumIntDoc = objNFsPag.lNumIntTitPag
        
        'Le o Titulo
        lErro = TituloPag_Le_Todos_NumIntDoc(objTituloPagar)
        If lErro <> SUCESSO Then gError 60358
        
        Call Preenche_Tab_Fatura(objTituloPagar)
    
    End If
    
    lErro = CF("Titulo_Le_DocumentoOriginal", objNFsPag.lNumIntDoc, CPR_NF_PAGAR, objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 58942 Then gError 60167

    If lErro = 58942 Then
        lNumIntDoc = objNFsPag.lNumIntDoc
        sTela = "NFPAG"
        sModulo = MODULO_CONTASAPAGAR
    Else
        lNumIntDoc = objNFiscal.lNumIntDoc
        sTela = "NFiscalEntrada"
        sModulo = MODULO_ESTOQUE
    End If

    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela_Consulta(lNumIntDoc, sModulo, sTela)
    If lErro <> SUCESSO And lErro <> 58885 Then gError 60168

    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objNFsPag.objProjetoInfo, objNFsPag.lNumIntDoc, PRJ_CR_TIPO_NFPAG)
    If lErro <> SUCESSO Then gError 181676
    '###################################################

    giFilialEmpresa = iFilialAux

    Traz_NotaPagar_Tela = SUCESSO

    Exit Function

Erro_Traz_NotaPagar_Tela:

    giFilialEmpresa = iFilialAux

    Traz_NotaPagar_Tela = gErr

    Select Case gErr

        Case 60167, 60168, 60358, 181676 'Tratados nas Rotinas Chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163542)

    End Select

    Exit Function

End Function

Sub Preenche_Tab_Fatura(objTituloPagar As ClassTituloPagar)

    TipoFatura.Caption = objTituloPagar.sSiglaDocumento
    EmissaoFatura.Caption = Format(objTituloPagar.dtDataEmissao, "dd/mm/yyyy")
    SaldoFatura.Caption = Format(objTituloPagar.dSaldo, "Standard")
    NumeroFatura.Caption = objTituloPagar.lNumTitulo
    ValorTotalFatura.Caption = Format(objTituloPagar.dValorTotal, "Standard")
    
    If objTituloPagar.iStatus = STATUS_BAIXADO Then
        StatusFatura.Caption = "Baixada"
    Else
        StatusFatura.Caption = "Aberto"
    End If
    
    glNumIntDocFat = objTituloPagar.lNumIntDoc
    
End Sub
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
            If lErro <> SUCESSO Then Error 60169

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then Error 60170

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

    Select Case Err

        Case 60169, 60170 'Tratados nas Rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163543)

    End Select

    Exit Sub

End Sub

Public Sub NumeroNota_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumeroNota)

End Sub

Public Sub NumeroNota_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumeroNota_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumeroNota.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumeroNota.ClipText)
    If lErro <> SUCESSO Then Error 60171

    Exit Sub

Erro_NumeroNota_Validate:

    Cancel = True


    Select Case Err

        Case 60171

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163544)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_Tela_NotaPagar

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163545)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing

    Set gobjTelaProjetoInfo = Nothing
    
    Set objGridParcelas = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

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
        
            Case TAB_NF
                Parent.HelpContextID = IDH_NOTA_FISCAL_PAGAR_CONS_NF
                
            Case TAB_Fatura
                Parent.HelpContextID = IDH_NOTA_FISCAL_PAGAR_CONS_FATURA
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NOTA_FISCAL_PAGAR_CONS_CONTABILIZACAO
                        
        End Select

    End If

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Limpa_Tela_NotaPagar()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Chama função que limpa TextBoxes e MaskedEdits da Tela
    Call Limpa_Tela(Me)

    'Limpa os campos não são limpos pela função acima

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me

    DataVencimento.Caption = ""
    ValorTotal.Caption = ""
    ValorICMS.Caption = ""
    ValorICMSSubst.Caption = ""
    ValorProdutos.Caption = ""
    ValorFrete.Caption = ""
    ValorSeguro.Caption = ""
    OutrasDespesas.Caption = ""
    ValorIPI.Caption = ""
    ValorIRRF.Caption = ""
    
    ValorINSS.Caption = ""
    PISRetido.Caption = ""
    ISSRetido.Caption = ""
    COFINSRetido.Caption = ""
    CSLLRetido.Caption = ""
    INSSRetido.Value = 0
    
    Filial.Clear
    CreditoICMS.Value = 0
    CreditoIPI.Value = 0
    Historico.Caption = ""

    'limpa tab de fatura
    TipoFatura.Caption = ""
    EmissaoFatura.Caption = ""
    SaldoFatura.Caption = ""
    NumeroFatura.Caption = ""
    ValorTotalFatura.Caption = ""
    StatusFatura.Caption = ""
    
    glNumIntDocFat = 0
    
    iFilialEmpresaTela = 0
        
    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade_Consulta

End Sub

Public Function Move_Tela_Memoria(objNFsPag As ClassNFsPag) As Long
'Preenche objNFsPag com os dados da Tela

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Fornecedor.Text)) > 0 Then

        objFornecedor.sNomeReduzido = Fornecedor.Text

        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then Error 60172
        
        If lErro = 6681 Then Error 60173

        objNFsPag.lFornecedor = objFornecedor.lCodigo

    End If

    If Len(Trim(Filial.Text)) > 0 Then
        objNFsPag.iFilial = Codigo_Extrai(Filial.Text)
    End If

    If Len(Trim(NumeroNota.Text)) > 0 Then objNFsPag.lNumNotaFiscal = CLng(NumeroNota.Text)

    'Se as  datas nao forem digitadas completar com data_nula
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objNFsPag.dtDataEmissao = DATA_NULA
    Else
        objNFsPag.dtDataEmissao = CDate(DataEmissao.Text)
    End If

    If Len(Trim(DataVencimento.Caption)) = 0 Then
        objNFsPag.dtDataVencimento = DATA_NULA
    Else
        objNFsPag.dtDataVencimento = CDate(DataVencimento.Caption)
    End If

    If Len(Trim(ValorTotal.Caption)) > 0 Then objNFsPag.dValorTotal = CDbl(ValorTotal.Caption)
    If Len(Trim(ValorICMS.Caption)) > 0 Then objNFsPag.dValorICMS = CDbl(ValorICMS.Caption)
    If Len(Trim(ValorICMSSubst.Caption)) > 0 Then objNFsPag.dValorICMSSubst = CDbl(ValorICMSSubst.Caption)
    objNFsPag.iCreditoICMS = CreditoICMS.Value
    If Len(Trim(ValorProdutos.Caption)) > 0 Then objNFsPag.dValorProdutos = CDbl(ValorProdutos.Caption)
    If Len(Trim(ValorSeguro.Caption)) > 0 Then objNFsPag.dValorSeguro = CDbl(ValorSeguro.Caption)
    If Len(Trim(ValorFrete.Caption)) > 0 Then objNFsPag.dValorFrete = CDbl(ValorFrete.Caption)
    If Len(Trim(ValorIRRF.Caption)) > 0 Then objNFsPag.dValorIRRF = CDbl(ValorIRRF.Caption)
    If Len(Trim(OutrasDespesas.Caption)) > 0 Then objNFsPag.dOutrasDespesas = CDbl(OutrasDespesas.Caption)
    If Len(Trim(ValorIPI.Caption)) > 0 Then objNFsPag.dValorIPI = CDbl(ValorIPI.Caption)
    objNFsPag.iCreditoIPI = CreditoIPI.Value
    
    If Len(Trim(ValorINSS.Caption)) > 0 Then objNFsPag.dValorINSS = CDbl(ValorINSS.Caption)
    If Len(Trim(PISRetido.Caption)) > 0 Then objNFsPag.dPISRetido = CDbl(PISRetido.Caption)
    If Len(Trim(ISSRetido.Caption)) > 0 Then objNFsPag.dISSRetido = CDbl(ISSRetido.Caption)
    If Len(Trim(COFINSRetido.Caption)) > 0 Then objNFsPag.dCOFINSRetido = CDbl(COFINSRetido.Caption)
    If Len(Trim(CSLLRetido.Caption)) > 0 Then objNFsPag.dCSLLRetido = CDbl(CSLLRetido.Caption)
    objNFsPag.iINSSRetido = INSSRetido.Value
    
    objNFsPag.lNumIntTitPag = 0
    objNFsPag.iFilialEmpresa = iFilialEmpresaTela

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 60172 'Tratado na rotina chamada

        Case 60173
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163546)

    End Select

    Exit Function

End Function

'Contabilidade
Public Sub CTBBotaoImprimir_Click()

    Call objContabil.Contabil_BotaoImprimir_Consulta_Click

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_NOTA_FISCAL_PAGAR_CONS_NF
    Set Form_Load_Ocx = Me
    Caption = "Consulta de Notas Fiscais a Pagar"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "NFPag_Consulta"

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
        ElseIf Me.ActiveControl Is NumeroNota Then
            Call NumeroLabel_Click
        End If

    End If

End Sub

Function TituloPag_Le_Todos_NumIntDoc(objTituloPagar As ClassTituloPagar) As Long
'Lê em Titulos a Pagar se não encontrar procura em Titulos a Pagar Baixados

Dim lErro As Long

On Error GoTo Erro_TituloPag_Le_Todos_NumIntDoc
        
    'Lê em Títulos a Pagar
    lErro = CF("TituloPagar_Le", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18372 Then Error 60355
    
    'Se não encontrou
    If lErro = 18372 Then
    
        'Procura em Títulos a Receber Baixados
        lErro = CF("TituloPagarBaixado_Le", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 56661 Then Error 60356
        
        'Se não encontrou nos dois ---> ERRO
        If lErro = 56661 Then Error 60357

    End If

    TituloPag_Le_Todos_NumIntDoc = SUCESSO

    Exit Function

Erro_TituloPag_Le_Todos_NumIntDoc:

    TituloPag_Le_Todos_NumIntDoc = Err

    Select Case Err
        
        Case 60355, 60356 'Tratado na rotina chamada
        
        Case 60357
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 163547)

    End Select

    Exit Function

End Function

Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Public Sub Fornecedor_Change()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Change
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134045

    Exit Sub

Erro_Fornecedor_Change:

    Select Case gErr

        Case 134045

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163548)

    End Select
    
    Exit Sub
End Sub


'#####################################
'Inserido por Wagner 03/08/2006
Sub BotaoProjetos_Click()
    Call gobjTelaProjetoInfo.BotaoProjetos_Click(True)
End Sub

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR
   
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "PROD"
    objItensPRJCR.sDescricao = "Produtos"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorProdutos.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
   Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "FR"
    objItensPRJCR.sDescricao = "Frete"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorFrete.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "SG"
    objItensPRJCR.sDescricao = "Seguro"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorSeguro.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "OD"
    objItensPRJCR.sDescricao = "Outras Despesas"

    objItensPRJCR.dValorOriginal = StrParaDbl(OutrasDespesas.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "ICMS"
    objItensPRJCR.sDescricao = "ICMS Substituto"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorICMSSubst.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_NFPAG
    
    colItensPRJ.Add objItensPRJCR
    
    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "IPI"
    objItensPRJCR.sDescricao = "IPI"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorIPI.Caption)

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

