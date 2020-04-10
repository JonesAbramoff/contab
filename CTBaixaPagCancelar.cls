VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBaixaPagCancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'inicio contabilidade
Public iAlterado As Integer
Dim objContabil As New ClassContabil

Private gobjParcPagBaixa As New ClassBaixaParcPagar
Private gobjBaixaPagar As New ClassBaixaPagar
Private gobjTituloPag As New ClassTituloPagar
Private gobjMovCCI As New ClassMovContaCorrente
Private gobjContaCorrente As New ClassContasCorrentesInternas

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoParcela As AdmEvento
Attribute objEventoParcela.VB_VarHelpID = -1

Private Const CONTACORRENTE1 As String = "Conta_Corrente"
Private Const SEQUENCIAL1 As String = "Sequencial"
Private Const VALOR1 As String = "Valor"
Private Const FORMA1 As String = "Tipo_Meio_Pagto"
Private Const HISTORICO1 As String = "Historico"
'fim contabilidade

'mnemonicos
Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
Private Const CONTA_COD As String = "Conta_Codigo" '????
Private Const CONTA_CONTABIL_CONTA As String = "Conta_Contabil_Conta" '????
Private Const DATA_EMISSAO1 As String = "Data_Emissao"
Private Const NUMERO1 As String = "Numero"
Private Const PARCELA1 As String = "Parcela"
Private Const TIPO1 As String = "Tipo"

Private Const VALOR_BAIXADO As String = "Valor_Baixado"
Private Const VALOR_DESCONTO As String = "Valor_Desconto"
Private Const VALOR_MULTA As String = "Valor_Multa"
Private Const VALOR_JUROS As String = "Valor_Juros"
Private Const VALOR_PAGO As String = "Valor_Pago"
Private Const CTA_CONTA_CORRENTE As String = "Conta_Contabil_Conta"
Private Const TIPO_DE_BAIXA As String = "Tipo_de_Baixa"

Dim objGrid1 As AdmGrid
Dim iGrid_Conta_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Debito_Col As Integer
Dim iGrid_Credito_Col As Integer
Dim iGrid_Historico_Col As Integer

Dim iFrameAtual As Integer
Dim iFornecedorAlterado As Integer
Dim iFramePagamentoAtual As Integer
Dim gi_Baixa_Localizada As Integer

Dim gobjAntecipPag As ClassAntecipPag
Dim gobjBorderoPagto As ClassBorderoPagto
Dim gobjCreditoPagar As ClassCreditoPagar
Dim glNumIntParc As Long
Dim glNumIntBaixa As Long

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoTipo As AdmEvento
Attribute objEventoTipo.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoSequencial As AdmEvento
Attribute objEventoSequencial.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Contabilizacao = 2




'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
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

Public Property Get Portador() As Object
     Set Portador = objUserControl.Controls("Portador")
End Property

Public Property Get NumOuSequencial() As Object
     Set NumOuSequencial = objUserControl.Controls("NumOuSequencial")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get ValorPagoPagto() As Object
     Set ValorPagoPagto = objUserControl.Controls("ValorPagoPagto")
End Property

Public Property Get ContaCorrente() As Object
     Set ContaCorrente = objUserControl.Controls("ContaCorrente")
End Property

Public Property Get DataMovimento() As Object
     Set DataMovimento = objUserControl.Controls("DataMovimento")
End Property

Public Property Get MeioPagtoDescricao() As Object
     Set MeioPagtoDescricao = objUserControl.Controls("MeioPagtoDescricao")
End Property

Public Property Get ValorPA() As Object
     Set ValorPA = objUserControl.Controls("ValorPA")
End Property

Public Property Get FilialEmpresaPA() As Object
     Set FilialEmpresaPA = objUserControl.Controls("FilialEmpresaPA")
End Property

Public Property Get CCIntNomeReduzido() As Object
     Set CCIntNomeReduzido = objUserControl.Controls("CCIntNomeReduzido")
End Property

Public Property Get NumeroMP() As Object
     Set NumeroMP = objUserControl.Controls("NumeroMP")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Pagamento() As Object
     Set Pagamento = objUserControl.Controls("Pagamento")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get ValorBaixado() As Object
     Set ValorBaixado = objUserControl.Controls("ValorBaixado")
End Property

Public Property Get Desconto() As Object
     Set Desconto = objUserControl.Controls("Desconto")
End Property

Public Property Get Multa() As Object
     Set Multa = objUserControl.Controls("Multa")
End Property

Public Property Get Juros() As Object
     Set Juros = objUserControl.Controls("Juros")
End Property

Public Property Get ValorPago() As Object
     Set ValorPago = objUserControl.Controls("ValorPago")
End Property

Public Property Get DataBaixa() As Object
     Set DataBaixa = objUserControl.Controls("DataBaixa")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get BotaoTrazer() As Object
     Set BotaoTrazer = objUserControl.Controls("BotaoTrazer")
End Property

Public Property Get SiglaDocumento() As Object
     Set SiglaDocumento = objUserControl.Controls("SiglaDocumento")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Sequencial() As Object
     Set Sequencial = objUserControl.Controls("Sequencial")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get SiglaDocumentoLabel() As Object
     Set SiglaDocumentoLabel = objUserControl.Controls("SiglaDocumentoLabel")
End Property

Public Property Get SequencialLabel() As Object
     Set SequencialLabel = objUserControl.Controls("SequencialLabel")
End Property

Public Property Get NumeroTitulo() As Object
     Set NumeroTitulo = objUserControl.Controls("NumeroTitulo")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get SaldoCredito() As Object
     Set SaldoCredito = objUserControl.Controls("SaldoCredito")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
End Property

Public Property Get FilialEmpresaCR() As Object
     Set FilialEmpresaCR = objUserControl.Controls("FilialEmpresaCR")
End Property

Public Property Get ValorCredito() As Object
     Set ValorCredito = objUserControl.Controls("ValorCredito")
End Property

Public Property Get SiglaDocumentoCR() As Object
     Set SiglaDocumentoCR = objUserControl.Controls("SiglaDocumentoCR")
End Property

Public Property Get DataEmissaoCred() As Object
     Set DataEmissaoCred = objUserControl.Controls("DataEmissaoCred")
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

Public Property Get CTBBotaoModeloPadrao() As Object
     Set CTBBotaoModeloPadrao = objUserControl.Controls("CTBBotaoModeloPadrao")
End Property

Public Property Get CTBBotaoLimparGrid() As Object
     Set CTBBotaoLimparGrid = objUserControl.Controls("CTBBotaoLimparGrid")
End Property

Public Property Get CTBModelo() As Object
     Set CTBModelo = objUserControl.Controls("CTBModelo")
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

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
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

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
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

Public Property Get BotaoConsultaTitPag() As Object
     Set BotaoConsultaTitPag = objUserControl.Controls("BotaoConsultaTitPag")
End Property

'#################################################################
'Inserido por Wagner 21/06/2006
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
'################################################################

'--- fim dos properties get dos controles da tela

Public Sub BotaoTrazer_Click()

    Call Carrega_Dados_Parcela

End Sub

Public Sub DataEmissao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    If Len(Trim(DataEmissao.ClipText)) = 0 Then Exit Sub

    'Verificar se a data é valida
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 43142

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True


    Select Case Err

        Case 43142

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155404)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_Filial_Validate

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42790

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Fornecedor
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 42792

        'Lê o Fornecedor que está na tela
        sNomeRed = Trim(Fornecedor.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialFornecedor.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Fornecedor e Código da Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 42791

        'Se não existe a Filial
        If lErro = 18272 Then Error 42793

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 42794

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 42790, 42791

        Case 42792
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 42793
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIAL_FORNECEDOR")

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
                'Segura o foco
            End If

        Case 42794
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155405)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoFornecedor = Nothing
    Set objEventoTipo = Nothing
    Set objEventoNumero = Nothing
    Set objEventoSequencial = Nothing
    Set objEventoParcela = Nothing

    Set objGrid1 = Nothing

    Set gobjAntecipPag = Nothing
    Set gobjBorderoPagto = Nothing
    Set gobjCreditoPagar = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    Set objContabil = Nothing
    
    Set gobjParcPagBaixa = Nothing
    Set gobjBaixaPagar = Nothing
    Set gobjTituloPag = Nothing
    Set gobjMovCCI = Nothing
    Set gobjContaCorrente = Nothing

End Sub

Public Sub Fornecedor_Change()

    iFornecedorAlterado = REGISTRO_ALTERADO

    Call Fornecedor_Preenche

End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodFilial As Integer
Dim objFornecedor As New ClassFornecedor
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 0 Then Exit Sub

    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 42788

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then Error 42789

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se Fornecedor não está preenchido
    ElseIf Len(Trim(Fornecedor.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    iFornecedorAlterado = 0

    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True


    Select Case Err

        Case 42788, 42789

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155406)

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

Public Sub Numero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)

End Sub

Public Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    If Len(Trim(Numero.ClipText)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Numero.Text)
        If lErro <> SUCESSO Then Error 61504
    
    End If
    
    Exit Sub
    
Erro_Numero_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 61504 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155407)
    
    End Select
    
    Exit Sub

End Sub

Public Sub NumeroTitulo_Click()

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTituloPagar As New ClassTituloPagar
Dim colSelecao As New Collection
Dim objInfoBaixaParcPagSeq As New ClassInfoBaixaParcPagSeq

On Error GoTo Erro_NumeroTitulo_Click

    'Verifica se os campos chave da tela estão preenchidos
    If Len(Trim(Fornecedor.ClipText)) = 0 Then Error 43613
    If Len(Trim(Filial.Text)) = 0 Then Error 43614

    'Preenche objFornecedor
    objFornecedor.sNomeReduzido = Fornecedor.Text
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then Error 43615

    'Se não encontrou o Fornecedor --> Erro
    If lErro <> SUCESSO Then Error 43616


    'Preenche objTituloPagar
    objInfoBaixaParcPagSeq.lNumTitulo = StrParaLong(Numero.Text)
    objInfoBaixaParcPagSeq.lCodigo = objFornecedor.lCodigo
    objInfoBaixaParcPagSeq.iCodFilial = Codigo_Extrai(Filial.Text)
    objInfoBaixaParcPagSeq.iNumParcelas = StrParaLong(Parcela.Text)
    objInfoBaixaParcPagSeq.iSequencial = StrParaLong(Sequencial.Text)
    objInfoBaixaParcPagSeq.dtDataEmissao = StrParaDate(DataEmissao.Text)

    'Adiciona filtros: iStatus, iFilialEmpresa, lFornecedor, iFilial
    colSelecao.Add objInfoBaixaParcPagSeq.lCodigo
    colSelecao.Add objInfoBaixaParcPagSeq.iCodFilial
    colSelecao.Add giFilialEmpresa

    'Chama Tela TitulosPagComBaixaLista
    Call Chama_Tela("BxParcPagSeq_TitLista", colSelecao, objInfoBaixaParcPagSeq, objEventoNumero)

    Exit Sub

Erro_NumeroTitulo_Click:

    Select Case Err

        Case 43613
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 43614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 43615

        Case 43616
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", Err, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155408)

    End Select

    Exit Sub

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1

    'Preenche campo Fornecedor
    Fornecedor.Text = objFornecedor.sNomeReduzido

    'Executa o Validate
    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()
'Pega dados da tela e executa o cancelamento

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o cancelamento da Baixa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 42829

    'Limpa a tela
    Call Limpa_Tela_BaixaPagCancelar

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 42829

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155409)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

    Call Limpa_Tela_BaixaPagCancelar

End Sub

Public Sub Form_Load()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    iFrameAtual = 1
    iFramePagamentoAtual = 0

    Set objEventoFornecedor = New AdmEvento
    Set objEventoTipo = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoSequencial = New AdmEvento
    Set objEventoParcela = New AdmEvento

    gi_Baixa_Localizada = False

    'Carregar a combo de tipos de documento com todos os tipos usados em OutrosPag.frm, mais NotaFiscalFatura (NFFP) e Fatura (FP)
    lErro = Carrega_SiglaDocumento()
    If lErro <> SUCESSO Then Error 42795

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then Error 57362
   
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 42795, 57362

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155410)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 57364

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 57365

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 57364

        Case 57365
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function


Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim objInfoBaixaParcPagSeq As ClassInfoBaixaParcPagSeq

    Set objInfoBaixaParcPagSeq = obj1

    'Preenche a data de emissão
    Call DateParaMasked(DataEmissao, objInfoBaixaParcPagSeq.dtDataEmissao)
    
    'Preenche o campo Tipo
    SiglaDocumento.Text = objInfoBaixaParcPagSeq.sSiglaDocumento
    Call SiglaDocumento_Validate(bSGECancelDummy)
    
    'Preenche campo Número
    Numero.PromptInclude = False
    Numero.Text = CLng(objInfoBaixaParcPagSeq.lNumTitulo)
    Numero.PromptInclude = True
    
    'Preenche o campo Parcela
    Parcela.PromptInclude = False
    Parcela.Text = StrParaLong(objInfoBaixaParcPagSeq.iNumParcelas)
    Parcela.PromptInclude = True

    'Preenche o sequencial da baixa
    Sequencial.PromptInclude = False
    Sequencial.Text = StrParaLong(objInfoBaixaParcPagSeq.iSequencial)
    Sequencial.PromptInclude = True
    
    Me.Show
    
    Call BotaoTrazer_Click

    Exit Sub

End Sub

Private Sub objEventoSequencial_evSelecao(obj1 As Object)

Dim objBaixaParcPag As New ClassBaixaParcPagar

    Set objBaixaParcPag = obj1

    'Preenche campo Sequencial
    Sequencial.PromptInclude = False
    Sequencial.Text = CStr(objBaixaParcPag.iSequencial)
    Sequencial.PromptInclude = True

    Me.Show

    Call BotaoTrazer_Click

End Sub

Private Sub objEventoTipo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_objEventoTipo_evSelecao

    Set objTipoDocumento = obj1

    'Preenche campo Tipo
    SiglaDocumento.Text = objTipoDocumento.sSigla

    Call SiglaDocumento_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoTipo_evSelecao:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155411)

     End Select

     Exit Sub

End Sub

Private Function Carrega_SiglaDocumento() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_SiglaDocumento

    lErro = CF("TiposDocumento_Le_TituloPag", colTipoDocumento)
    If lErro <> SUCESSO Then Error 42800

    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        SiglaDocumento.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next
    
    Carrega_SiglaDocumento = SUCESSO

    Exit Function

Erro_Carrega_SiglaDocumento:

    Carrega_SiglaDocumento = Err

    Select Case Err

        Case 42800

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155412)

    End Select

    Exit Function

End Function

Private Sub Carrega_Dados_Parcela()
'Carrega os dados da Parcela

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim iNumParcela As Integer
Dim objParcPagBaixa As New ClassBaixaParcPagar
Dim objBaixaPagar As New ClassBaixaPagar
Dim iSequencialBaixa As Integer

On Error GoTo Erro_Carrega_Dados_Parcela

    gi_Baixa_Localizada = False

    If Len(Trim(Fornecedor.ClipText)) = 0 Or Len(Trim(Filial.Text)) = 0 Or Len(Trim(SiglaDocumento.Text)) = 0 Or _
       Len(Trim(Numero.ClipText)) = 0 Or Len(Trim(Parcela.ClipText)) = 0 Or Len(Trim(Sequencial.ClipText)) = 0 Then Error 57317

    Call Limpa_Campos_Parcela

    'Lê o Fornecedor
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then Error 42801

    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(SiglaDocumento.Text)
    objTituloPagar.lNumTitulo = CLng(Numero.Text)
    objTituloPagar.dtDataEmissao = MaskedParaDate(DataEmissao)

    'Lê o Titulo a Pagar Baixado com os dados da tela
    lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then Error 42802

    If lErro <> SUCESSO Then

        'Lê o Titulo a Pagar Nao Baixado com os dados da tela
        lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18551 Then Error 42802

        If lErro <> SUCESSO Then Error 42803 'Não encontrou

    End If

    iNumParcela = CInt(Parcela.Text)

    iSequencialBaixa = CInt(Sequencial)

    'Lê a Parcela no BD
    lErro = CF("BaixaPagCancelar_Le_Parcela", objTituloPagar.lNumIntDoc, iNumParcela, iSequencialBaixa, objParcPagBaixa)
    If lErro <> SUCESSO And lErro <> 42807 Then Error 42808

    If lErro = 42807 Then Error 42816 'Não encontrou

    glNumIntParc = objParcPagBaixa.lNumIntParcela
    objBaixaPagar.lNumIntBaixa = objParcPagBaixa.lNumIntBaixa

    'Lê  a Baixa
    lErro = CF("BaixaPagar_Le", objBaixaPagar)
    If lErro <> SUCESSO And lErro <> 42812 Then Error 42813

    If lErro <> SUCESSO Then Error 42814 'Não encontrou

    'Coloca na tela os dados da Baixa Lida
    lErro = Traz_Dados_Baixa(objParcPagBaixa, objBaixaPagar)
    If lErro <> SUCESSO Then Error 42815

    glNumIntBaixa = objBaixaPagar.lNumIntBaixa
    
    Set gobjParcPagBaixa = objParcPagBaixa
    Set gobjBaixaPagar = objBaixaPagar
    Set gobjTituloPag = objTituloPagar

    gi_Baixa_Localizada = True

    Exit Sub

Erro_Carrega_Dados_Parcela:

    Select Case Err

        Case 42801, 42802, 42808, 42813, 42815

        Case 42803
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)

        Case 42814
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAPAG_INEXISTENTE", Err)

        Case 42816
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAPARCPAG_INEXISTENTE", Err)

        Case 57317
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CAMPOS_CANCELARBAIXA_NAO_PREENCHIDOS", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155413)

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
                Parent.HelpContextID = IDH_CANCELAR_BAIXA_CONTAS_PAGAR_ID
                
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_CANCELAR_BAIXA_CONTAS_PAGAR_CONTABILIZACAO
                                
        End Select

    End If

End Sub

Public Sub Pagamento_Click(Index As Integer)

    FramePagamento(iFramePagamentoAtual).Visible = False
    FramePagamento(Index).Visible = True
    iFramePagamentoAtual = Index

End Sub

Private Sub Limpa_Tela_BaixaPagCancelar()

    Call Limpa_Tela(Me)

    Filial.Clear

    SiglaDocumento.Text = ""

    Call Limpa_Campos_Parcela

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0
    
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
    
    If gcolModulo.Ativo(MODULO_CONTABILIDADE) = MODULO_ATIVO Then
    
        CTBDataContabil.PromptInclude = False
        CTBDataContabil.Text = Format(objBaixaPagar.dtData, "dd/mm/yy")
        CTBDataContabil.PromptInclude = True
        Call CTBDataContabilValidate(bSGECancelDummy)
        
    End If

    If objBaixaPagar.iMotivo = MOTIVO_PAGAMENTO Then
        Pagamento(0).Value = True
        Call Pagamento_Click(0)

        'Traz os dados do pagamento
        lErro = Traz_Dados_Pagamento(objBaixaPagar)
        If lErro <> SUCESSO Then Error 42817

    ElseIf objBaixaPagar.iMotivo = MOTIVO_PAGTO_ANTECIPADO Then
        Pagamento(1).Value = True
        Call Pagamento_Click(1)

        'Traz os dados do pagamento antecipado
        lErro = Traz_Dados_Pagamento_Antecipado(objBaixaPagar)
        If lErro <> SUCESSO Then Error 42823

    ElseIf objBaixaPagar.iMotivo = MOTIVO_CREDITO_FORNECEDOR Then
        Pagamento(2).Value = True
        Call Pagamento_Click(2)

        'Traz os dados do crédito
        lErro = Traz_Dados_Credito_Fornecedor(objBaixaPagar)
        If lErro <> SUCESSO Then Error 42820
        
    '#############################################################
    'Inserido por Wagner 21/06/2006
    ElseIf objBaixaPagar.iMotivo = MOTIVO_CHEQUE_DE_TERCEIROS Then
        Pagamento(3).Value = True
        Call Pagamento_Click(3)

        'Traz os dados do crédito
        lErro = Traz_Dados_ChequePre(objBaixaPagar)
        If lErro <> SUCESSO Then Error 42820
    '#############################################################

    End If
    
    Traz_Dados_Baixa = SUCESSO

    Exit Function

Erro_Traz_Dados_Baixa:

    Traz_Dados_Baixa = Err

    Select Case Err

        Case 42817, 42820, 42823, 57363

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155414)

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

On Error GoTo Erro_Traz_Dados_Pagamento

    objMovCCI.lNumMovto = objBaixaPagar.lNumMovConta

    'Lê o Movimento
    lErro = CF("MovContaCorrente_Le", objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then Error 42818

    'Se não encontrou Movimento --> erro
    If lErro = 11893 Then Error 42825

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 42819

    'Se não encontrou a conta Corrente --> erro
    If lErro <> SUCESSO Then Error 42820
    
    Set gobjMovCCI = objMovCCI
    Set gobjContaCorrente = objContaCorrente

    'Coloca os dados na tela
    ContaCorrente.Caption = objContaCorrente.sNomeReduzido
    ValorPagoPagto.Caption = Format(objMovCCI.dValor, "Standard")
    Historico.Caption = objMovCCI.sHistorico
    NumOuSequencial.Caption = IIf(objMovCCI.lNumero <> 0, CStr(objMovCCI.lNumero), "")

    objPortador.iCodigo = objMovCCI.iPortador

    If objPortador.iCodigo <> 0 Then

        'Lê o Portador
        lErro = CF("Portador_Le", objPortador)
        If lErro <> SUCESSO And lErro <> 15971 Then Error 43143

        'Se não achou o Portador --> erro
        If lErro <> SUCESSO Then Error 43144

        Portador.Caption = objMovCCI.iPortador & SEPARADOR & objPortador.sNomeReduzido

    End If

    objBorderoPagto.dtDataEmissao = objMovCCI.dtDataMovimento
    objBorderoPagto.iCodConta = objMovCCI.iCodConta
    objBorderoPagto.lNumero = objMovCCI.lNumero

    If objMovCCI.iTipoMeioPagto = Cheque Then
        TipoMeioPagto(0).Value = True

    ElseIf objMovCCI.iTipoMeioPagto = BORDERO Then
        TipoMeioPagto(1).Value = True

    ElseIf objMovCCI.iTipoMeioPagto = DINHEIRO Then
        TipoMeioPagto(2).Value = True
    End If

    Set gobjBorderoPagto = objBorderoPagto
    
    Set gobjBorderoPagto.objTela = Me

    Traz_Dados_Pagamento = SUCESSO

    Exit Function

Erro_Traz_Dados_Pagamento:

    Traz_Dados_Pagamento = Err

    Select Case Err

        Case 42818, 42819, 43143

        Case 42825
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovCCI.lNumMovto)

        Case 42820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrente.iCodigo)

        Case 43144
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PORTADOR_NAO_CADASTRADO1", Err, objPortador.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155415)

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
    If lErro <> AD_SQL_SUCESSO And 17071 Then Error 42821
    If lErro <> SUCESSO Then Error 42822

    Set gobjCreditoPagar = objCreditoPag

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

    Traz_Dados_Credito_Fornecedor = Err

    Select Case Err

        Case 42821

        Case 42822
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CREDITO_PAG_FORN_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155416)

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
    If lErro <> SUCESSO And lErro <> 11893 Then Error 42824
    If lErro = 11893 Then Error 42826 'Não encontrou

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 42827
    If lErro <> SUCESSO Then Error 42828 'Não encontrou

    objAntecipPag.lNumMovto = objMovCCI.lNumMovto

    lErro = CF("AntecipPag_Le_NumMovto", objAntecipPag)
    If lErro <> AD_SQL_SUCESSO And lErro <> 42845 Then Error 42841
    If lErro = 42845 Then Error 42846

    Set gobjAntecipPag = objAntecipPag

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

    Traz_Dados_Pagamento_Antecipado = Err

    Select Case Err

        Case 42824, 42827, 42841

        Case 42826
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovCCI.lNumMovto)

        Case 42828
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objCCI.iCodigo)

        Case 42846
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAGTO_ANTECIPADO_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155417)

    End Select

    Exit Function

End Function

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
    Set gobjTituloPag = New ClassTituloPagar
    Set gobjMovCCI = New ClassMovContaCorrente
    Set gobjContaCorrente = New ClassContasCorrentesInternas

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iSequencial As Integer
Dim iMotivo As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se os campos chave da tela estão preenchidos
    If Len(Trim(Fornecedor.ClipText)) = 0 Then Error 42830
    If Len(Trim(Filial.Text)) = 0 Then Error 42831
    If Len(Trim(SiglaDocumento)) = 0 Then Error 42832
    If Len(Trim(Numero.ClipText)) = 0 Then Error 42833
    If Len(Trim(Parcela.ClipText)) = 0 Then Error 42835
    If Len(Trim(Sequencial.ClipText)) = 0 Then Error 42836

    If gi_Baixa_Localizada = False Then Error 42834

    iSequencial = CInt(Sequencial.Text)

    If Pagamento(0).Value = True Then
        iMotivo = MOTIVO_PAGAMENTO
    ElseIf Pagamento(1).Value = True Then
        iMotivo = MOTIVO_PAGTO_ANTECIPADO
    ElseIf Pagamento(2).Value = True Then
        iMotivo = MOTIVO_CREDITO_FORNECEDOR
    ElseIf Pagamento(3).Value = True Then
        iMotivo = MOTIVO_CHEQUE_DE_TERCEIROS
    End If

    'verifica se a data contábil é igual a data corrente ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(gdtDataAtual)
    If lErro <> SUCESSO Then Error 20838
    
    If gobjBorderoPagto Is Nothing Then
        Set gobjBorderoPagto = New ClassBorderoPagto
        Set gobjBorderoPagto.objTela = Me
    End If

    lErro = CF("BaixaPagCancelar_Grava", gobjBorderoPagto, gobjAntecipPag, gobjCreditoPagar, glNumIntParc, iSequencial, iMotivo, objContabil)
    If lErro <> SUCESSO Then Error 42847

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = Err
    
    Select Case Err

        Case 20838, 42847

        Case 42830
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 42831
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 42832
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)

        Case 42833
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 42835
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_PREENCHIDA", Err)

        Case 42836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SEQUENCIAL_NAO_INFORMADO", Err)

        Case 42834
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAPAG_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155418)

    End Select

    Exit Function

End Function

Public Sub Parcela_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Parcela, iAlterado)

End Sub

Public Sub Sequencial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Sequencial, iAlterado)

End Sub

Public Sub Sequencial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Sequencial_Validate

    If Len(Trim(Sequencial.ClipText)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Sequencial.Text)
        If lErro <> SUCESSO Then Error 61503
    
    End If
    
    Exit Sub
    
Erro_Sequencial_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 61503 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155419)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Parcela_Validate

    If Len(Trim(Parcela.ClipText)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Parcela.Text)
        If lErro <> SUCESSO Then Error 61502
    
    End If
    
    Exit Sub
    
Erro_Parcela_Validate:

    Cancel = True
    
    Select Case Err
    
        Case 61502 'Erro tratado na rotina chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155420)
    
    End Select
    
    Exit Sub

End Sub

Public Sub SequencialLabel_Click()

Dim lErro As Long
Dim iCodFilial As Integer
Dim iNumParcela As Integer
Dim objFornecedor As New ClassFornecedor
Dim objTituloPagar As New ClassTituloPagar
Dim objBaixaParcPag As New ClassBaixaParcPagar
Dim objParcelaPagar As New ClassParcelaPagar
Dim colSelecao As New Collection

On Error GoTo Erro_SequencialLabel_Click

    'Verifica se os campos chave da tela estão preenchidos
    If Len(Trim(Fornecedor.ClipText)) = 0 Then Error 43136
    If Len(Trim(Filial.Text)) = 0 Then Error 43137
    If Len(Trim(Numero.ClipText)) = 0 Then Error 43138
    If Len(Trim(Parcela.ClipText)) = 0 Then Error 43139
    If Len(Trim(SiglaDocumento)) = 0 Then Error 43145
    If Len(Trim(DataEmissao.Text)) = 0 Then Error 43628

    'Lê o Fornecedor
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
    If lErro <> SUCESSO Then Error 43617

    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    objTituloPagar.sSiglaDocumento = SCodigo_Extrai(SiglaDocumento.Text)
    objTituloPagar.lNumTitulo = CLng(Numero.Text)
    objTituloPagar.dtDataEmissao = MaskedParaDate(DataEmissao)

    objParcelaPagar.iNumParcela = CInt(Parcela.Text)

    'Lê o Titulo a Pagar Baixado com os dados da tela
    lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then Error 43618

    If lErro <> SUCESSO Then

        'Lê o Titulo a Pagar não Baixado com os dados da tela
        lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18551 Then Error 43619

        If lErro <> SUCESSO Then Error 43620

        'Preenche objParcelaPagar
        objParcelaPagar.lNumIntTitulo = objTituloPagar.lNumIntDoc
        'Lê a Parcela no BD
        lErro = CF("ParcelaPagar_Le_Numero", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 18987 Then Error 43621

        If lErro <> SUCESSO Then Error 43627

    Else

        'Preenche objParcelaPagar
        objParcelaPagar.lNumIntTitulo = objTituloPagar.lNumIntDoc
        'Lê a Parcela no BD
        lErro = CF("ParcelaPagarBaixada_Le_Numero", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 43625 Then Error 43626

        If lErro <> SUCESSO Then Error 43625

    End If

    'Preenche objBaixaParcPag
    objBaixaParcPag.iStatus = STATUS_EXCLUIDO

    'Adiciona filtros
    colSelecao.Add objParcelaPagar.lNumIntDoc
    colSelecao.Add objBaixaParcPag.iStatus

    Call Chama_Tela("BaixasParcPagSeqLista", colSelecao, objBaixaParcPag, objEventoSequencial)

    Exit Sub

Erro_SequencialLabel_Click:

    Select Case Err

        Case 43136
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 43137
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 43138
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case 43139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_NAO_PREENCHIDA", Err)

        Case 43145
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)

        Case 43617, 43618, 43619, 43621, 43626

        Case 43620
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)

        Case 43625, 43627
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_BAIXADA_INEXISTENTE", Err)

        Case 43628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_NAO_PREENCHIDA", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155421)

    End Select

    Exit Sub

End Sub

Public Sub SiglaDocumento_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_SiglaDocumento_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(SiglaDocumento.Text)) = 0 Then Exit Sub

    'Verifica se o Tipo foi selecionado
    If SiglaDocumento.Text = SiglaDocumento.List(SiglaDocumento.ListIndex) Then Exit Sub

    'Tenta Localizar o Tipo no Text da Combo
    lErro = CF("SCombo_Seleciona", SiglaDocumento)
    If lErro <> SUCESSO And lErro <> 60483 Then Error 43134

    'Se não encontrar -> Erro
    If lErro = 60483 Then Error 43135

    Exit Sub

Erro_SiglaDocumento_Validate:

    Cancel = True


    Select Case Err

        Case 43134

        Case 43135
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, SiglaDocumento.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155422)

    End Select

    Exit Sub

End Sub

Public Sub SiglaDocumentoLabel_Click()

Dim objTipoDocumento As New ClassTipoDocumento
Dim colSelecao As Collection

    objTipoDocumento.sSigla = SCodigo_Extrai(SiglaDocumento.Text)

    'Chama a tela TipoDocOutrosPagLista
    Call Chama_Tela("TiposDocBaixaPagCancLista", colSelecao, objTipoDocumento, objEventoTipo)

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 43140

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        'Erro já tratado
        Case 43140

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155423)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a data
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 43141

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        'Erro já tratado
        Case 43141

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155424)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

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

Private Sub CTBDataContabilValidate(Cancel As Boolean)

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
Dim objFornecedor As New ClassFornecedor, sContaTela As String
Dim objFilial As New ClassFilialFornecedor ', objConta As New ClassContasCorrentesInternas

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case FORNECEDOR_COD

            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then

                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                objMnemonicoValor.colValor.Add objFornecedor.lCodigo

            Else

                objMnemonicoValor.colValor.Add 0

            End If

        Case FORNECEDOR_NOME

            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then

                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                objMnemonicoValor.colValor.Add objFilial.sNome

            Else

                objMnemonicoValor.colValor.Add ""

            End If

        Case FILIAL_CONTA

            If Len(Filial.Text) > 0 Then

                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

                If objFilial.sContaContabil <> "" Then

                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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
                If lErro <> SUCESSO Then Error 57370

                objMnemonicoValor.colValor.Add objFilial.sCgc

            Else

                objMnemonicoValor.colValor.Add ""

            End If

        Case NUMERO1
                objMnemonicoValor.colValor.Add StrParaLong(Numero.Text)
                
        Case PARCELA1
                objMnemonicoValor.colValor.Add StrParaInt(Parcela.Text)

        Case SEQUENCIAL1
                objMnemonicoValor.colValor.Add StrParaInt(Sequencial.Text)

        Case DATA_EMISSAO1
                objMnemonicoValor.colValor.Add StrParaDate(DataEmissao.Text)

        Case TIPO1
                objMnemonicoValor.colValor.Add SiglaDocumento.Text
                
        Case CTA_CONTA_CORRENTE
                
            If gobjContaCorrente.sContaContabil <> "" Then
                lErro = Mascara_RetornaContaTela(gobjContaCorrente.sContaContabil, sContaTela)
                If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            Else
                sContaTela = ""
            End If
            
            objMnemonicoValor.colValor.Add sContaTela
            
        Case TIPO_DE_BAIXA
            objMnemonicoValor.colValor.Add gobjBaixaPagar.iMotivo

        Case VALOR_BAIXADO
            objMnemonicoValor.colValor.Add StrParaDbl(ValorBaixado.Caption)
            
        Case VALOR_DESCONTO
            objMnemonicoValor.colValor.Add StrParaDbl(Desconto.Caption)
            
        Case VALOR_MULTA
            objMnemonicoValor.colValor.Add StrParaDbl(Multa.Caption)
            
        Case VALOR_JUROS
            objMnemonicoValor.colValor.Add StrParaDbl(Juros.Caption)
            
        Case VALOR_PAGO
            objMnemonicoValor.colValor.Add StrParaDbl(ValorPago.Caption)
            
        Case Else
            Error 57373

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case ERRO_SEM_MENSAGEM

        Case 57373
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155425)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_CANCELAR_BAIXA_CONTAS_PAGAR_ID
    Set Form_Load_Ocx = Me
    Caption = "Cancelar Baixa em Contas a Pagar"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BaixaPagCancelar"

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

'***** fim do trecho a ser copiado ******

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fornecedor Then
            Call FornecedorLabel_Click
        ElseIf Me.ActiveControl Is Numero Then
            Call NumeroTitulo_Click
        ElseIf Me.ActiveControl Is Sequencial Then
            Call SequencialLabel_Click
        ElseIf Me.ActiveControl Is SiglaDocumento Then
            Call SiglaDocumentoLabel_Click
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

Public Sub BotaoConsultaTitPag_Click()
'Abre a tela de consulta de títulos a pagar

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objTituloPagar As New ClassTituloPagar

On Error GoTo Erro_ConsultaTitPag_Click

    'Verifica se os campos necessários para consultar o título foram preenchidos
    lErro = Critica_CamposPreenchidos_Consulta_TitPag()
    If lErro <> SUCESSO Then gError 79857
    
    'Guarda em objFornecedor o nome reduzido do Fornecedor, que servirá como parâmetro para obter o código do mesmo
    objFornecedor.sNomeReduzido = Fornecedor.Text
    
    'Lê os dados do Fornecedor e obtém o código do mesmo
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79858
    
    'Se não encontrou o fornecedor => erro
    If lErro = 6681 Then gError 79859
    
    'Guarda no obj os dados que serão utilizados na tela de consulta para exibir o título desejado
    With objTituloPagar
        
        .lFornecedor = objFornecedor.lCodigo
        .iFilial = Codigo_Extrai(Filial.Text)
        .dtDataEmissao = DataEmissao.Text
        .lNumTitulo = Numero.Text
        .sSiglaDocumento = SCodigo_Extrai(SiglaDocumento.Text)
        
    End With
    
    'Chama a tela de consulta de títulos a pagar
    Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    Exit Sub

Erro_ConsultaTitPag_Click:

    Select Case gErr
    
        Case 79857, 79858
        
        Case 79859
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO", gErr, objTituloPagar.lFornecedor)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155426)
    
    End Select
    
    Exit Sub

End Sub

Private Function Critica_CamposPreenchidos_Consulta_TitPag() As Long
'Verifica se os campos necessários para chamar a tela de consulta de título a pagar estão preenchidos

Dim lErro As Long

On Error GoTo Erro_Critica_CamposPreenchidos_Consulta_TitPag

    'Verifica se o campo Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79853
    
    'Verifica se o campo Filial foi preenchido
    If Len(Trim(Filial.Text)) = 0 Then gError 79854
    
    'Verifica se o campo Emissão foi preenchido
    If Len(Trim(DataEmissao.ClipText)) = 0 Then gError 79855
    
    'Verifica se o campo Número foi preenchido
    If Len(Trim(Numero.Text)) = 0 Then gError 79856
    
    Critica_CamposPreenchidos_Consulta_TitPag = SUCESSO
    
    Exit Function
    
Erro_Critica_CamposPreenchidos_Consulta_TitPag:

    Critica_CamposPreenchidos_Consulta_TitPag = gErr
    
    Select Case gErr
    
        Case 79853
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 79854
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 79855
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_OBRIGATORIA_DOC_ORIGINAL", gErr)
        
        Case 79856
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155427)
    
    End Select
    
    Exit Function
    
End Function


Public Sub LabelParcela_Click()

Dim lErro As Long
Dim objInfoBaixaParcPagSeq As New ClassInfoBaixaParcPagSeq
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_LabelParcela_Click

    'Verifica se os campos chave da tela estão preenchidos
    
    If Len(Trim(Filial.Text)) = 0 Then gError 80472
    If Len(Trim(SiglaDocumento.Text)) = 0 Then gError 80473
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 80471
    If Len(Trim(Numero.Text)) = 0 Then gError 87003
    
    objFornecedor.sNomeReduzido = Fornecedor.Text
    
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 12729 Then gError 87001
    If lErro = 12729 Then gError 87002
        
    'Adiciona dados a coleção
    colSelecao.Add objFornecedor.lCodigo
    colSelecao.Add Codigo_Extrai(Filial.Text)
    colSelecao.Add SCodigo_Extrai(SiglaDocumento.Text)
    colSelecao.Add StrParaDate(DataEmissao.Text)
    colSelecao.Add StrParaLong(Numero.Text)
    
    Set objInfoBaixaParcPagSeq = New ClassInfoBaixaParcPagSeq
    
    'Preenche o obj com informações da tela
    objInfoBaixaParcPagSeq.iCodFilial = Codigo_Extrai(Filial.Text)
    objInfoBaixaParcPagSeq.sSiglaDocumento = SCodigo_Extrai(SiglaDocumento.Text)
    objInfoBaixaParcPagSeq.dtDataEmissao = StrParaDate(DataEmissao.Text)
    objInfoBaixaParcPagSeq.lNumTitulo = StrParaLong(Numero.Text)
    
    'Chama a tela
    Call Chama_Tela("bxParcelaTitSelLista", colSelecao, objInfoBaixaParcPagSeq, objEventoParcela)
    
    Exit Sub
    
Erro_LabelParcela_Click:

    Select Case gErr
    
        Case 80471
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
    
        Case 80472
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 80473
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SIGLA_DOCUMENTO_NAO_PREENCHIDO", gErr)
                    
        Case 87001
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_INEXISTENTE", gErr)
                                        
        Case 87002
        
        Case 87003
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155428)
            
    End Select
    
    Exit Sub

End Sub
Public Property Get LabelParcela() As Object
     Set LabelParcela = objUserControl.Controls("LabelParcela")
End Property

Public Sub objEventoParcela_evSelecao(obj1 As Object)

Dim lErro As Long, bCancela As Boolean
Dim objInfoBaixaParcPagSeq As ClassInfoBaixaParcPagSeq

On Error GoTo Erro_objEventoParcela_evSelecao

    Set objInfoBaixaParcPagSeq = obj1

    If Not (objInfoBaixaParcPagSeq Is Nothing) Then
        Parcela.PromptInclude = False
        Parcela.Text = CStr(objInfoBaixaParcPagSeq.iNumParcelas)
        Parcela.PromptInclude = True
        Call Parcela_Validate(bCancela)
    End If

    Me.Show

    Exit Sub

Erro_objEventoParcela_evSelecao:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155429)

    End Select

    Exit Sub

End Sub

Private Sub Fornecedor_Preenche()
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134054

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134054

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155430)

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

        Case 180296

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180298)

    End Select

    Exit Function

End Function
'#################################################################################



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


