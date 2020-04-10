VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBaixaPagChequeTerc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Alteracao Daniel em 07/05/02
Dim gColNumIntPag As New Collection

Private gcolTiposDeCobranca As New AdmColCodigoNome

'Property Variables:
Dim m_Caption As String
Event Unload()

'Indicação de seleção no grid
Const S_MARCADO As String = "1"
Const S_DESMARCADO As String = "0"

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'Mnemônicos
Private Const FORNECEDOR_COD As String = "Fornecedor_Codigo"
Private Const FORNECEDOR_NOME As String = "Fornecedor_Nome"
Private Const FILIAL_COD As String = "FilialForn_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialForn_Nome"
Private Const FILIAL_CONTA As String = "FilialForn_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialForn_CGC_CPF"
'###########Parcelas#############
Private Const DATA_BAIXA As String = "Data_Baixa"
Private Const TIPO1 As String = "Tipo"
Private Const VALOR_BAIXADO As String = "Valor_Baixado"
Private Const VALOR_DESCONTO As String = "Valor_Desconto"
Private Const VALOR_MULTA As String = "Valor_Multa"
Private Const VALOR_JUROS As String = "Valor_Juros"
Private Const SELECIONADA1 As String = "Selecionada"
Private Const COBRANCA1 As String = "Cobranca"
Private Const VALOR_PAGAR As String = "Valor_Pagar"
Private Const FILIAL_EMPRESA As String = "Filial_Empresa"
Private Const COD_CONTA_CORRENTE As String = "CodContaCorrente"
Private Const NOMERED_CONTA_CORRENTE As String = "NomeRedContaCorrente"
Private Const CTA_CONTA_CORRENTE As String = "Conta_Contabil_Conta"
Private Const VALOR_PAGO As String = "Valor_Pago"
Private Const HISTORICO1 As String = "Historico"
Private Const DATA_MOVIMENTO As String = "Data_Adiantamento"
Private Const CCINT_NOME_REDUZIDO As String = "CCInt_Nome_Reduzido"
Private Const MEIO_PAGTO_DESC As String = "Meio_Pagto_Descricao"
Private Const NUMERO_MP As String = "Numero_MP"
Private Const VALOR_BAIXADOPA As String = "Valor_BaixadoPA"
Private Const SELECIONADOPA1 As String = "SelecionadoPA"
Private Const FILIAL_EMPRESAPA As String = "Filial_EmpresaPA"
Private Const DATA_EMISSAO As String = "Data_EmissaoCR"
Private Const SIGLA_DOCUMENTO As String = "Sigla_DocumentoCR"
Private Const NUM_TITULO As String = "Num_TituloCR"
Private Const VALOR_BAIXADOCR As String = "Valor_BaixadoCR"
Private Const SELECIONADO1 As String = "SelecionadoCR"
Private Const FILIAL_EMPRESACR As String = "Filial_EmpresaCR"
Private Const PAGTO_DINHEIRO As String = "Pagto_Dinheiro"
Private Const PAGTO_ANTECIPADO As String = "Pagto_Antecipado"
Private Const CRED_DEV As String = "Cred_Dev"
Private Const VALOR_BAIXADO_CT As String = "Valor_Baixado_Cheque"
Private Const VALOR_DESCONTO_CT As String = "Valor_Desc_Cheque"
Private Const VALOR_MULTA_CT As String = "Valor_Multa_Cheque"
Private Const VALOR_JUROS_CT As String = "Valor_Juros_Cheque"
Private Const SELECIONADOCT1 As String = "Selecionado_Cheque"
Private Const NUM_TITULO1 As String = "Num_Titulo"
Private Const NUM_PARCELA As String = "Parcela"

'Fim da contabilidade

'Grid Parcelas:
Dim objGridParcelas As AdmGrid
Dim iFornecedorAlterado As Integer
Dim iGrid_Selecionar_Col As Integer
Dim iGrid_DataVencimento_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_DataEmissaoTitulo_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Saldo_Col As Integer
Dim iGrid_ValorBaixar_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_Multa_Col As Integer
Dim iGrid_Juros_Col As Integer
Dim iGrid_Cobranca_Col As Integer
Dim iGrid_Portador_Col As Integer
Dim iGrid_ValorPagar_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_FilialEmpresa_Col As Integer
Dim iFramePrincipalAlterado As Integer

'Grid Cheques Pré
Dim objGridChequePre As AdmGrid
Dim iGrid_SelecionadoCT_Col As Integer
Dim iGrid_ClienteCT_Col As Integer
Dim iGrid_FilialCT_Col As Integer
Dim iGrid_BancoCT_Col As Integer
Dim iGrid_AgenciaCT_Col As Integer
Dim iGrid_ContaCorrenteCT_Col As Integer
Dim iGrid_NumeroCT_Col As Integer
Dim iGrid_DataDepositoCT_Col As Integer
Dim iGrid_ValorCT_Col As Integer
Dim iGrid_FilialEmpresaCT_Col As Integer
Dim iGrid_ValorBaixadoCT_Col As Integer
Dim iGrid_ValorDescontoCT_Col As Integer
Dim iGrid_ValorJurosCT_Col As Integer
Dim iGrid_ValorMultaCT_Col As Integer

Public iAlterado As Integer
Dim iGridAtual As Integer
Dim iFrameAtual As Integer
Dim iFramePagamentoAtual As Integer
Dim gcolInfoParcPag As Collection
Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1

Dim gcolCheques As Collection

'Indicação de número máximo de linhas dos Grids
Private Const NUM_MAX_PARCELAS_BAIXA = 300
Private Const NUM_MAX_CREDITOS = 200
Private Const NUM_MAX_PAGTOS_ANTECIPADOS = 200

'Constantes públicas dos tabs
Private Const TAB_Titulos = 1
Private Const TAB_Parcelas = 2
Private Const TAB_Contabilizacao = 3

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get FramePagamento() As Object
     Set FramePagamento = objUserControl.Controls("FramePagamento")
End Property

Public Property Get Portador() As Object
     Set Portador = objUserControl.Controls("Portador")
End Property

Public Property Get ContaCorrente() As Object
     Set ContaCorrente = objUserControl.Controls("ContaCorrente")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get ValorPago() As Object
     Set ValorPago = objUserControl.Controls("ValorPago")
End Property

Public Property Get Selecionado() As Object
     Set Selecionado = objUserControl.Controls("Selecionado")
End Property

Public Property Get SaldoCredito() As Object
     Set SaldoCredito = objUserControl.Controls("SaldoCredito")
End Property

Public Property Get ValorCredito() As Object
     Set ValorCredito = objUserControl.Controls("ValorCredito")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
End Property

Public Property Get SiglaDocumento() As Object
     Set SiglaDocumento = objUserControl.Controls("SiglaDocumento")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get CCIntNomeReduzido() As Object
     Set CCIntNomeReduzido = objUserControl.Controls("CCIntNomeReduzido")
End Property

Public Property Get MeioPagtoDescricao() As Object
     Set MeioPagtoDescricao = objUserControl.Controls("MeioPagtoDescricao")
End Property

Public Property Get NumeroMP() As Object
     Set NumeroMP = objUserControl.Controls("NumeroMP")
End Property

Public Property Get DataMovimento() As Object
     Set DataMovimento = objUserControl.Controls("DataMovimento")
End Property

Public Property Get FrameParcelas() As Object
     Set FrameParcelas = objUserControl.Controls("FrameParcelas")
End Property

Public Property Get Selecionada() As Object
     Set Selecionada = objUserControl.Controls("Selecionada")
End Property

Public Property Get ValorAPagar() As Object
     Set ValorAPagar = objUserControl.Controls("ValorAPagar")
End Property

Public Property Get ValorJuros() As Object
     Set ValorJuros = objUserControl.Controls("ValorJuros")
End Property

Public Property Get ValorBaixado() As Object
     Set ValorBaixado = objUserControl.Controls("ValorBaixado")
End Property

Public Property Get ValorMulta() As Object
     Set ValorMulta = objUserControl.Controls("ValorMulta")
End Property

Public Property Get ValorDesconto() As Object
     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get Saldo() As Object
     Set Saldo = objUserControl.Controls("Saldo")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Pagamento() As Object
     Set Pagamento = objUserControl.Controls("Pagamento")
End Property

Public Property Get NomePortador() As Object
     Set NomePortador = objUserControl.Controls("NomePortador")
End Property

Public Property Get Cobranca() As Object
     Set Cobranca = objUserControl.Controls("Cobranca")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get UpDownDataBaixa() As Object
     Set UpDownDataBaixa = objUserControl.Controls("UpDownDataBaixa")
End Property

Public Property Get DataBaixa() As Object
     Set DataBaixa = objUserControl.Controls("DataBaixa")
End Property

Public Property Get FilialEmpresa() As Object
     Set FilialEmpresa = objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get TotalBaixar() As Object
     Set TotalBaixar = objUserControl.Controls("TotalBaixar")
End Property

Public Property Get Frame8() As Object
     Set Frame8 = objUserControl.Controls("Frame8")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Fornecedor() As Object
     Set Fornecedor = objUserControl.Controls("Fornecedor")
End Property

Public Property Get FornecLabel() As Object
     Set FornecLabel = objUserControl.Controls("FornecLabel")
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
End Property

Public Property Get UpDownEmissaoInic() As Object
     Set UpDownEmissaoInic = objUserControl.Controls("UpDownEmissaoInic")
End Property

Public Property Get EmissaoInic() As Object
     Set EmissaoInic = objUserControl.Controls("EmissaoInic")
End Property

Public Property Get UpDownEmissaoFim() As Object
     Set UpDownEmissaoFim = objUserControl.Controls("UpDownEmissaoFim")
End Property

Public Property Get EmissaoFim() As Object
     Set EmissaoFim = objUserControl.Controls("EmissaoFim")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get UpDownVencInic() As Object
     Set UpDownVencInic = objUserControl.Controls("UpDownVencInic")
End Property

Public Property Get VencInic() As Object
     Set VencInic = objUserControl.Controls("VencInic")
End Property

Public Property Get UpDownVencFim() As Object
     Set UpDownVencFim = objUserControl.Controls("UpDownVencFim")
End Property

Public Property Get VencFim() As Object
     Set VencFim = objUserControl.Controls("VencFim")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get TituloInic() As Object
     Set TituloInic = objUserControl.Controls("TituloInic")
End Property

Public Property Get TituloFim() As Object
     Set TituloFim = objUserControl.Controls("TituloFim")
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

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get DataEmissaoTitulo() As Object
     Set DataEmissaoTitulo = objUserControl.Controls("DataEmissaoTitulo")
End Property

Public Property Get BotaoConsultaDocOriginal() As Object
     Set BotaoConsultaDocOriginal = objUserControl.Controls("BotaoConsultaDocOriginal")
End Property

'#######################################################################
'Inserido por Wagner 06/06/2006
Public Property Get GridChequePre() As Object
     Set GridChequePre = objUserControl.Controls("GridChequePre")
End Property
Public Property Get SelecionadoCT() As Object
     Set SelecionadoCT = objUserControl.Controls("SelecionadoCT")
End Property
Public Property Get ClienteCT() As Object
     Set ClienteCT = objUserControl.Controls("ClienteCT")
End Property
Public Property Get FilialCT() As Object
     Set FilialCT = objUserControl.Controls("FilialCT")
End Property
Public Property Get BancoCT() As Object
     Set BancoCT = objUserControl.Controls("BancoCT")
End Property
Public Property Get AgenciaCT() As Object
     Set AgenciaCT = objUserControl.Controls("AgenciaCT")
End Property
Public Property Get ContaCorrenteCT() As Object
     Set ContaCorrenteCT = objUserControl.Controls("ContaCorrenteCT")
End Property
Public Property Get NumeroCT() As Object
     Set NumeroCT = objUserControl.Controls("NumeroCT")
End Property
Public Property Get DataDepositoCT() As Object
     Set DataDepositoCT = objUserControl.Controls("DataDepositoCT")
End Property
Public Property Get ValorCT() As Object
     Set ValorCT = objUserControl.Controls("ValorCT")
End Property
Public Property Get ValorBaixadoCT() As Object
     Set ValorBaixadoCT = objUserControl.Controls("ValorBaixadoCT")
End Property
Public Property Get ValorJurosCT() As Object
     Set ValorJurosCT = objUserControl.Controls("ValorJurosCT")
End Property
Public Property Get ValorMultaCT() As Object
     Set ValorMultaCT = objUserControl.Controls("ValorMultaCT")
End Property
Public Property Get ValorDescontoCT() As Object
     Set ValorDescontoCT = objUserControl.Controls("ValorDescontoCT")
End Property
Public Property Get FilialEmpresaCT() As Object
     Set FilialEmpresaCT = objUserControl.Controls("FilialEmpresaCT")
End Property
Public Property Get TotalCheque() As Object
     Set TotalCheque = objUserControl.Controls("TotalCheque")
End Property
'#######################################################################

'--- fim dos properties get dos controles da tela


Public Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava a Baixa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 42769

    'Limpa a tela
    Call Limpa_Tela_BaixaPagar
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 42769

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155357)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alterações e confirma se deseja salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 42749

    'Limpa a tela
    Call Limpa_Tela_BaixaPagar
    
    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 42749

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155358)

    End Select

    Exit Sub

End Sub

Public Sub Cobranca_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Cobranca_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Cobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Cobranca
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ContaCorrente_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaCorrente_Validate(Cancel As Boolean)


End Sub

Public Sub DataBaixa_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataBaixa_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataBaixa, iAlterado)

End Sub

Public Sub DataBaixa_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixa_Validate

    'Se a DataBaixa está preenchida
    If Len(DataBaixa.ClipText) > 0 Then

        'Verifica se a DataBaixa é válida
        lErro = Data_Critica(DataBaixa.Text)
        If lErro <> SUCESSO Then Error 42754

    Else
    
        Error 42755
    
    End If

    Exit Sub

Erro_DataBaixa_Validate:

    Cancel = True


    Select Case Err

        Case 42754

        Case 42755
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_SEM_PREENCHIMENTO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155360)

    End Select

    Exit Sub

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

Public Sub EmissaoFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub EmissaoFim_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(EmissaoFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub EmissaoFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoFim_Validate

    'Se a data EmissaoFim está preenchida
    If Len(EmissaoFim.ClipText) > 0 Then

        'Verifica se a data EmissaoFim é válida
        lErro = Data_Critica(EmissaoFim.Text)
        If lErro <> SUCESSO Then Error 42768

        'Se a data EmissaoInic Está preenchida
        If Len(EmissaoInic.ClipText) > 0 Then

            'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
            If CDate(EmissaoFim.Text) < CDate(EmissaoInic.Text) Then Error 15898

        End If

    End If

    Exit Sub

Erro_EmissaoFim_Validate:

    Cancel = True


    Select Case Err

        Case 42768

        Case 15898
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155361)

    End Select

    Exit Sub

End Sub

Public Sub EmissaoInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub EmissaoInic_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(EmissaoInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub EmissaoInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoInic_Validate

    'Se a data EmissaoInic está preenchida
    If Len(EmissaoInic.ClipText) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(EmissaoInic.Text)
        If lErro <> SUCESSO Then Error 42770

    End If

    Exit Sub

Erro_EmissaoInic_Validate:

    Cancel = True


    Select Case Err

        Case 42770

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155362)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42771

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Fornecedor
        If Len(Trim(Fornecedor.Text)) = 0 Then Error 42772

        'Lê o Fornecedor que está na tela
        sNomeRed = Trim(Fornecedor.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialFornecedor.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Fornecedor e Código da Filial
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sNomeRed, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then Error 42773

        'Se não existe a Filial
        If lErro = 18272 Then Error 42774

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialFornecedor.iCodFilial) & SEPARADOR & objFilialFornecedor.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 42775

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 42771, 42773

        Case 42772
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)

        Case 42774
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIAL_FORNECEDOR")

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            Else
                'Segura o foco
            End If

        Case 42775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155363)

    End Select

    Exit Sub

End Sub

Public Sub FilialEmpresa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub FilialEmpresa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub FilialEmpresa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = FilialEmpresa
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub


Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
        
    Set gcolInfoParcPag = Nothing
    Set gcolTiposDeCobranca = Nothing
    Set gcolCheques = Nothing
    
    'Eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objEventoFornecedor = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    Set objGridParcelas = Nothing
    Set objGridChequePre = Nothing
    
End Sub

Public Sub Fornecedor_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFornecedorAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
    Call Fornecedor_Preenche
    
End Sub

Public Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 0 Then Exit Sub
    
    'Se Fornecedor está preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
        If lErro <> SUCESSO Then Error 15888

        'Lê coleção de códigos, nomes de Filiais do Fornecedor
        lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
        If lErro <> SUCESSO Then Error 15889

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

        Case 15888, 15889

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155364)

    End Select

    Exit Sub

End Sub

Public Sub FornecLabel_Click()
'Chamada do Browse de Fornecedores

Dim colSelecao As Collection
Dim objFornecedor As New ClassFornecedor

    'Passa o Fornecedor que está na tela para o Obj
    objFornecedor.sNomeReduzido = Trim(Fornecedor.Text)

    'Chama a tela com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

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

Public Sub Historico_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NomePortador_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub NomePortador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub NomePortador_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = NomePortador
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Numero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Numero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Numero
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Me.Show

    'Preenche Fornecedor na tela com NomeReduzido
    Set objFornecedor = obj1
    
    Fornecedor.Text = CStr(objFornecedor.sNomeReduzido)

    'Chama Validate de Fornecedor
    Call Fornecedor_Validate(bSGECancelDummy)
    
    Exit Sub

End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame selecionado visível
        Frame1(Opcao.SelectedItem.Index).Visible = True
        
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
    
        'Se Frame selecionado foi o de Títulos
        If Opcao.SelectedItem.Index = TAB_Titulos Then
            
            Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_TITULO
            iFramePrincipalAlterado = 0

        'Se Frame selecionado foi o de Parcelas
        ElseIf Opcao.SelectedItem.Index = TAB_Parcelas Then
            
            Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_PARCELAS
            
            If iFramePrincipalAlterado <> 0 Then
                
                
                lErro = Carrega_Tab_Parcelas()
                If lErro <> SUCESSO Then Error 42767
                
                iFramePrincipalAlterado = 0

            End If
            
        'Se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        ElseIf Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then
        
            Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_CONTABILIZACAO
            Call objContabil.Contabil_Carga_Modelo_Padrao
            iFramePrincipalAlterado = 0
    
        End If

    End If
            
    Exit Sub

Erro_Opcao_Click:

    Select Case Err

        Case 42767
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155365)
            
    End Select

    Exit Sub

End Sub

Public Sub Parcela_GotFocus()


    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Parcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Parcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Parcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Portador_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Portador_Validate(Cancel As Boolean)


End Sub

Public Sub Saldo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Saldo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Saldo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Saldo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Selecionada_Click()

Dim lErro As Long
Dim objInfoParcPag As ClassInfoParcPag

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

    Call Calcula_Total_Pagar

End Sub

Public Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Tipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub TituloFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TituloFim_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(TituloFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub TituloFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TituloFim_Validate

    'Se TituloFim e TituloInic estão preenchidos
    If Len(Trim(TituloFim.Text)) And Len(Trim(TituloInic.Text)) > 0 Then

        'Verifica se TituloFim é maior ou igual que TituloInic
        If CLng(TituloFim.Text) < CLng(TituloInic.Text) Then Error 15909

    End If

    Exit Sub

Erro_TituloFim_Validate:

    Cancel = True


    Select Case Err

        Case 15909
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155367)

    End Select

    Exit Sub

End Sub

Public Sub TituloInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TituloInic_GotFocus()
    
Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(TituloInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub UpDownDataBaixa_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_DownClick

    'Diminui a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15938

    Exit Sub

Erro_UpDownDataBaixa_DownClick:

    Select Case Err

        Case 15938

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155368)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataBaixa_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_UpClick

    'Aumenta a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15939

    Exit Sub

Erro_UpDownDataBaixa_UpClick:

    Select Case Err

        Case 15939

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155369)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15903

    Exit Sub

Erro_UpDownEmissaoFim_DownClick:

    Select Case Err

        Case 15903

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155370)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15908

    Exit Sub

Erro_UpDownEmissaoFim_UpClick:

    Select Case Err

        Case 15908

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155371)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_DownClick

    'Diminui a data EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15901

    Exit Sub

Erro_UpDownEmissaoInic_DownClick:

    Select Case Err

        Case 15901

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155372)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_UpClick

    'Aumenta a data EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15907

    Exit Sub

Erro_UpDownEmissaoInic_UpClick:

    Select Case Err

        Case 15907

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155373)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_DownClick

    'Diminui a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15902

    Exit Sub

Erro_UpDownVencFim_DownClick:

    Select Case Err

        Case 15902

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155374)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_UpClick

    'Aumenta a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15906

    Exit Sub

Erro_UpDownVencFim_UpClick:

    Select Case Err

        Case 15906

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155375)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_DownClick

    'Diminui a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 15904

    Exit Sub

Erro_UpDownVencInic_DownClick:

    Select Case Err

        Case 15904

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155376)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_UpClick

    'Aumenta a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 15905

    Exit Sub

Erro_UpDownVencInic_UpClick:

    Select Case Err

        Case 15905

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155377)

    End Select

    Exit Sub

End Sub

Public Sub ValorAPagar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorAPagar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorAPagar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorAPagar
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorBaixado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorBaixado_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorBaixado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorBaixado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorBaixado
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorDesconto_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorDesconto
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ValorJuros_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorMulta_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorMulta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorMulta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorMulta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorMulta
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub
Public Sub ValorJuros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorJuros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorJuros_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorJuros
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Selecionada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Selecionada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Selecionada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Selecionada
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39677

        If objGridInt Is objGridParcelas Then

            Select Case GridParcelas.Col

                Case iGrid_ValorBaixar_Col

                    lErro = Saida_Celula_ValorBaixado(objGridInt)
                    If lErro <> SUCESSO Then Error 15880

                Case iGrid_Desconto_Col

                    lErro = Saida_Celula_ValorDesconto(objGridInt)
                    If lErro <> SUCESSO Then Error 15881

                Case iGrid_Multa_Col

                    lErro = Saida_Celula_ValorMulta(objGridInt)
                    If lErro <> SUCESSO Then Error 15882

                Case iGrid_Juros_Col

                    lErro = Saida_Celula_ValorJuros(objGridInt)
                    If lErro <> SUCESSO Then Error 15883
                    
            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 15884

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 15880, 15881, 15882, 15883, 15952, 15958

        Case 5898, 6666, 7777, 8888, 9999

        Case 15884
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39677
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155378)

    End Select

    Exit Function

End Function

Private Function ContaCorrente_Carrega(objComboBox As ComboBox) As Long


End Function

Private Function Portador_Carrega(objComboBox As ComboBox) As Long

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iSubTipo As Integer

On Error GoTo Erro_Form_Load

    'Visibilidade para versão LIGHT
'    If giTipoVersao = VERSAO_LIGHT Then
'        FilialEmpresa.left = POSICAO_FORA_TELA
'        FilialEmpresa.TabStop = False
'    End If
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    iFrameAtual = 1
    iFramePagamentoAtual = 0

    Set objGridParcelas = New AdmGrid
    Set objEventoFornecedor = New AdmEvento

    Set gcolInfoParcPag = New Collection
    Set gcolCheques = New Collection

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, gcolTiposDeCobranca)
    If lErro <> SUCESSO Then Error 57340

    'Executa a Inicialização do grid Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 15879
    
    '#######################################################
    'Inserido por Wagner 06/06/2006
    'Executa a Inicialização do grid Pagamentos Antecipados
    lErro = Inicializa_Grid_ChequePre(objGridChequePre)
    If lErro <> SUCESSO Then Error 15886
    '#######################################################
       
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR, 0)
    If lErro <> SUCESSO Then Error 39676
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 6760, 6761, 15878, 15879, 15885, 15886, 39676, 57340, 79673

        Case 79674
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155381)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Parcelas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Baixar")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Valor Baixar")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Cobrança")
    objGridInt.colColuna.Add ("Portador")
    objGridInt.colColuna.Add ("Valor a Pagar")
    objGridInt.colColuna.Add ("Valor Parcela")
    
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
'    End If
            
    'campos de edição do grid
    objGridInt.colCampo.Add (Selecionada.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (DataEmissaoTitulo.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Saldo.Name)
    objGridInt.colCampo.Add (ValorBaixado.Name)
    objGridInt.colCampo.Add (ValorDesconto.Name)
    objGridInt.colCampo.Add (ValorMulta.Name)
    objGridInt.colCampo.Add (ValorJuros.Name)
    objGridInt.colCampo.Add (Cobranca.Name)
    objGridInt.colCampo.Add (NomePortador.Name)
    objGridInt.colCampo.Add (ValorAPagar.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (FilialEmpresa.Name)
'    End If

    iGrid_Selecionar_Col = 1
    iGrid_DataVencimento_Col = 2
    iGrid_Tipo_Col = 3
    iGrid_DataEmissaoTitulo_Col = 4
    iGrid_Numero_Col = 5
    iGrid_Parcela_Col = 6
    iGrid_Saldo_Col = 7
    iGrid_ValorBaixar_Col = 8
    iGrid_Desconto_Col = 9
    iGrid_Multa_Col = 10
    iGrid_Juros_Col = 11
    iGrid_Cobranca_Col = 12
    iGrid_Portador_Col = 13
    iGrid_ValorPagar_Col = 14
    iGrid_ValorParcela_Col = 15
    iGrid_FilialEmpresa_Col = 16

    objGridInt.objGrid = GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    'largura da primeira coluna
    GridParcelas.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO
        
    Exit Function

End Function



Function Trata_Parametros(Optional objBaixaPag As ClassBaixaPagar) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0
    
    Trata_Parametros = SUCESSO
    
    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155382)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub ValorPago_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

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

Public Sub VencFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VencFim_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(VencFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub VencFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_VencFim_Validate

    'Se a data VencFim está preenchida
    If Len(VencFim.ClipText) > 0 Then

        'Verifica se a data VencFim é válida
        lErro = Data_Critica(VencFim.Text)
        If lErro <> SUCESSO Then Error 15899

        'Se a data vencInica está preenchida
        If Len(VencInic.ClipText) > 0 Then

            'Verifica se a VencFim é maior ou igual a VencInic
            If CDate(VencFim.Text) < CDate(VencInic.Text) Then Error 15900

        End If

    End If

    Exit Sub

Erro_VencFim_Validate:

    Cancel = True


    Select Case Err

        Case 15899

        Case 15900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155383)

    End Select

    Exit Sub

End Sub

Public Sub VencInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub VencInic_GotFocus()
    
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(VencInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub VencInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_VencInic_Validate

    'Se a data VencInic está preenchida
    If Len(VencInic.ClipText) > 0 Then

        'Verifica se a data VencInic é válida
        lErro = Data_Critica(VencInic.Text)
        If lErro <> SUCESSO Then Error 15896

    End If

    Exit Sub

Erro_VencInic_Validate:

    Cancel = True


    Select Case Err

        Case 15896

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155384)

    End Select

    Exit Sub

End Sub

Private Sub Grid_Parcelas_Preenche(colInfoParcPag As Collection)
'Preenche o Grid Parcelas com os dados de gcolInfoParcPag

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim objFilialEmpresa As New AdmFiliais
Dim objCodDescricao As New AdmCodigoNome

On Error GoTo Erro_Grid_Parcelas_Preenche

    'Se o número de parcelas for maior que o número de linhas do Grid
    If colInfoParcPag.Count + 1 > GridParcelas.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de parcelas
        GridParcelas.Rows = colInfoParcPag.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridParcelas)

    End If

    iLinha = 0

    'Percorre todas as Parcelas da Coleção
    For Each objInfoParcPag In colInfoParcPag

        iLinha = iLinha + 1

        'Passa para a tela os dados da Parcela em questão
        GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col) = S_DESMARCADO
        GridParcelas.TextMatrix(iLinha, iGrid_DataVencimento_Col) = Format(objInfoParcPag.dtDataVencimento, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcPag.sSiglaDocumento
        If objInfoParcPag.dtDataEmissao <> DATA_NULA Then GridParcelas.TextMatrix(iLinha, iGrid_DataEmissaoTitulo_Col) = Format(objInfoParcPag.dtDataEmissao, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col) = objInfoParcPag.lNumTitulo
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcPag.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Saldo_Col) = Format(objInfoParcPag.dValor, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixar_Col) = Format(objInfoParcPag.dValor, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorPagar_Col) = Format(objInfoParcPag.dValor, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_Portador_Col) = objInfoParcPag.sNomeRedPortador
        
        If objInfoParcPag.iTipoCobranca <> 0 Then
        
            For Each objCodDescricao In gcolTiposDeCobranca
                If objCodDescricao.iCodigo = objInfoParcPag.iTipoCobranca Then
                
                    GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = objInfoParcPag.iTipoCobranca & SEPARADOR & objCodDescricao.sNome
                    Exit For
                    
                End If
            Next
            
        End If
        
        If objFilialEmpresa.iCodFilial <> objInfoParcPag.iFilialEmpresa Then
        
            objFilialEmpresa.iCodFilial = objInfoParcPag.iFilialEmpresa
            'Lê a Filial Empresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then Error 43367
        
        End If
        
'        If giTipoVersao = VERSAO_FULL Then
            GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objInfoParcPag.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
'        End If
        
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objInfoParcPag.dValorOriginal, "Standard")
        
    Next
    
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridParcelas.iLinhasExistentes = colInfoParcPag.Count

    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    Exit Sub

Erro_Grid_Parcelas_Preenche:

    Select Case Err

        Case 43367
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155385)

    End Select

    Exit Sub

End Sub

Private Function Frame_Parcelas_Limpa()
'Limpa os comandos do Frame Parcelas

    'Limpa os Grids
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridChequePre)
    
    'Limpa os Labels
    TotalBaixar.Caption = Format(0, "Fixed")

    'Limpa os MaskEdBox
    DataBaixa = Format(gdtDataAtual, "dd/mm/yy")
    
End Function

Private Function Saida_Celula_ValorBaixado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorBaixado As Double
Dim dValorDesconto As Double
Dim dValorSaldo As Double

On Error GoTo Erro_Saida_Celula_ValorBaixado

    Set objGridInt.objControle = ValorBaixado
    
    'Se ValorBaixado está preenchido
    If Len(Trim(ValorBaixado.Text)) <> 0 Then

        'Verifica se ValorBaixado é válido
        lErro = Valor_Positivo_Critica(ValorBaixado.Text)
        If lErro <> SUCESSO Then Error 15940

        'Lê os dados da tela
        dValorBaixado = CDbl(Trim(ValorBaixado.Text))
        dValorSaldo = CDbl(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col)))
        
        'Verifica se Saldo é maior que o ValorBaixado
        If dValorBaixado > dValorSaldo Then Error 15941

        'Verifica se Desconto é maior ou igual ao ValorBaixado
        If dValorBaixado <= StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col)) Then Error 15942

        lErro = Parcela_RecalculaValorPagar(ValorBaixado.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col))
        If lErro <> SUCESSO Then Error 56760
                
        'Formata o Valor na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col) = Format(dValorBaixado, "Standard")
        
    Else
    
        lErro = Parcela_RecalculaValorPagar(ValorBaixado.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col))
        If lErro <> SUCESSO Then Error 56761
                
    End If

    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Selecionar_Col)) = 1 Then Call Calcula_Total_Pagar
    
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15943

    Saida_Celula_ValorBaixado = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBaixado:

    Saida_Celula_ValorBaixado = Err

    Select Case Err

        Case 15940, 56760, 56761
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_MAIOR_SALDO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_MENOR_IGUAL_DESCONTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15943
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155388)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorDesconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorBaixado As Double
Dim dValorDesconto As Double

On Error GoTo Erro_Saida_Celula_ValorDesconto

    Set objGridParcelas.objControle = ValorDesconto
    
    'Se Desconto está preenchido
    If Len(Trim(ValorDesconto.Text)) <> 0 Then

        'Verifica se o Desconto é válido
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then Error 15944

        'Verifica se o Desconto é maior ou igual ao ValorBaixado
        If dValorDesconto >= StrParaDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col)) Then Error 15945

        lErro = Parcela_RecalculaValorPagar(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), ValorDesconto.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col))
        If lErro <> SUCESSO Then Error 56758
        
        'Formata o Desconto na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col) = Format(dValorDesconto, "Standard")

    Else
    
        lErro = Parcela_RecalculaValorPagar(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), ValorDesconto.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col))
        If lErro <> SUCESSO Then Error 56759
        
    End If
    
    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Selecionar_Col)) = 1 Then Call Calcula_Total_Pagar

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15946

    Saida_Celula_ValorDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorDesconto:

    Saida_Celula_ValorDesconto = Err

    Select Case Err

        Case 15944, 56758, 56759
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_MENOR_IGUAL_DESCONTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15946
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155389)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorMulta(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorMulta As Double
Dim dtDataBaixa As Date
Dim dtDataVencimento As Date

On Error GoTo Erro_Saida_Celula_ValorMulta

    Set objGridParcelas.objControle = ValorMulta

    'Se Multa está preenchida
    If Len(Trim(ValorMulta.Text)) <> 0 Then

        'Verifica se a Multa é válida
        lErro = Valor_NaoNegativo_Critica(ValorMulta.Text)
        If lErro <> SUCESSO Then Error 15947

        If Len(Trim(DataBaixa.Text)) = 0 Then Error 56766
        
        'Lê os dados da tela
        dValorMulta = CDbl(Trim(ValorMulta.Text))
        
        dtDataBaixa = CDate(DataBaixa.Text)
        dtDataVencimento = CDate(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col)))

        If gobjCRFAT.iPermiteMultaSemAtraso <> MARCADO Then

            'Verifica se a Multa é maior que zero quando a DataBaixa é menor ou igual à Data Vencimento
            If dtDataBaixa <= dtDataVencimento And dValorMulta > 0 Then Error 15948
        
        End If

        lErro = Parcela_RecalculaValorPagar(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col), ValorMulta.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col))
        If lErro <> SUCESSO Then Error 56764
                
        'Formata a Multa na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col) = Format(dValorMulta, "Standard")

    Else
    
        lErro = Parcela_RecalculaValorPagar(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col), ValorMulta.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col))
        If lErro <> SUCESSO Then Error 56765
                    
    End If

    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Selecionar_Col)) = 1 Then Call Calcula_Total_Pagar
    
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15949

    Saida_Celula_ValorMulta = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorMulta:

    Saida_Celula_ValorMulta = Err

    Select Case Err

        Case 15947, 56764, 56765
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56766
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_JUROS_DATABAIXA_VAZIA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 15948
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_INCOMPATIVEL_DATABAIXA_DATAVENC", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15949
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155390)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorJuros(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorJuros As Double
Dim dtDataBaixa As Date
Dim dtDataVencimento As Date

On Error GoTo Erro_Saida_Celula_ValorJuros

    Set objGridParcelas.objControle = ValorJuros
    
    'Se ValorJuros está preenchido
    If Len(Trim(ValorJuros.Text)) <> 0 Then

        'Verifica se ValorJuros é válido
        lErro = Valor_NaoNegativo_Critica(ValorJuros.Text)
        If lErro <> SUCESSO Then Error 15950

        If Len(Trim(DataBaixa.Text)) = 0 Then Error 56767
        
        'Lê os dados da tela
        dValorJuros = CDbl(Trim(ValorJuros.Text))
        
        dtDataBaixa = CDate(DataBaixa.Text)
        dtDataVencimento = CDate(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col)))

        If gobjCRFAT.iPermiteMultaSemAtraso <> MARCADO Then

            'Verifica se juros é maior que zero quando a DataBaixa é menor ou igual à Data Vencimento
            If dtDataBaixa <= dtDataVencimento And dValorJuros > 0 Then Error 56768

        End If

        lErro = Parcela_RecalculaValorPagar(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col), ValorJuros.Text)
        If lErro <> SUCESSO Then Error 56762
                
        'Formata os Juros na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Juros_Col) = Format(dValorJuros, "Standard")

    Else
    
        lErro = Parcela_RecalculaValorPagar(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Desconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Multa_Col), ValorJuros.Text)
        If lErro <> SUCESSO Then Error 56763
                
    End If

    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Selecionar_Col)) = 1 Then Call Calcula_Total_Pagar
        
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 15951

    Saida_Celula_ValorJuros = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorJuros:

    Saida_Celula_ValorJuros = Err

    Select Case Err

        Case 15950, 56762, 56763
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56767
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_JUROS_DATABAIXA_VAZIA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 56768
            lErro = Rotina_Erro(vbOKOnly, "ERRO_JUROS_INCOMPATIVEL_DATABAIXA_DATAVENC", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 15951
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155391)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Verifica se os principais dados para gravação foram preenchidos

Dim lErro As Long
Dim iIndice As Integer
Dim iParcelaSelecionada As Integer
Dim dSomaPagar As Double
Dim dValorParcela As Double, dValorParcelaBaixa As Double
Dim colBaixaParcPagar As New colBaixaParcPagar
Dim objInfoParcPag As ClassInfoParcPag
Dim dtDataBaixa As Date, vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se DataBaixa está preenchida
    If Len(Trim(DataBaixa.ClipText)) = 0 Then Error 15985

    'Lê a DataBaixa na tela
    dtDataBaixa = MaskedParaDate(DataBaixa)

    iParcelaSelecionada = 0

    'Percorre todas as linhas do Grid Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gcolInfoParcPag.Item(iIndice)
        
        'Se a checkbox Selecionada está marcada
        If GridParcelas.TextMatrix(iIndice, iGrid_Selecionar_Col) = S_MARCADO Then

            'Conta quantas Parcelas estão selecionadas
            iParcelaSelecionada = iParcelaSelecionada + 1
            
            'Verifica se Valor a Baixar está preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorBaixar_Col))) = 0 Then Error 15986
            
            dValorParcelaBaixa = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorBaixar_Col))
            dValorParcela = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorPagar_Col))
            
            'Faz o somatório do Valor total a pagar
            dSomaPagar = Round(dSomaPagar + dValorParcela, 2)

            'Passa os dados da tela para o Obj
            objInfoParcPag.dValorDesconto = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Desconto_Col))
            objInfoParcPag.dValorMulta = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Multa_Col))
            objInfoParcPag.dValorJuros = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Juros_Col))
        
            'Titulos com pagto após o vcto nao deveriam ter desconto
            If dtDataBaixa > objInfoParcPag.dtDataVencimento And objInfoParcPag.dValorDesconto > 0 Then
            
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PAGTO_ATRASO_DESC", iIndice)
                If vbMsgRes = vbYes Then Error 59087
            
            End If
            
            'Titulos com pagto até o vcto nao deveriam ter juros ou multa
            If dtDataBaixa <= objInfoParcPag.dtDataVencimento And (objInfoParcPag.dValorJuros > 0 Or objInfoParcPag.dValorMulta > 0) Then
                
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_PAGTO_EM_DIA_MULTA", iIndice)
                If vbMsgRes = vbYes Then Error 59088
            
            End If
                                    
            Call colBaixaParcPagar.Add(0, 0, objInfoParcPag.lNumIntParc, 0, objInfoParcPag.dValorMulta, objInfoParcPag.dValorJuros, objInfoParcPag.dValorDesconto, dValorParcelaBaixa, STATUS_LANCADO)
            
        End If
    Next

    'Verifica se alguma célula Selecionada do GridParcelas está marcada
    If iParcelaSelecionada < 1 Then Error 15987

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataBaixa.Text))
    If lErro <> SUCESSO Then Error 20837
       
    lErro = Grava_Pagamento_ChequePre(dSomaPagar, dtDataBaixa, colBaixaParcPagar)
    If lErro <> SUCESSO Then Error 15895
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 59087, 59088 'desistiu por que descobriu que tinha digitado valor errado
        
        Case 15895, 15897, 15975, 20837
        
        Case 15985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_SEM_PREENCHIMENTO", Err)

        Case 15986
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_PARCELAS_NAO_INFORMADO", Err)

        Case 15987
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_MARCADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155392)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_BaixaPagar()
'Limpa a tela

Dim lErro As Long

    'Limpa a coleção de Parcelas
    Set gcolInfoParcPag = Nothing
    Set gcolInfoParcPag = New Collection
    Set gcolCheques = New Collection
    
    'Limpa os comandos da tela
    Call Limpa_Tela(Me)
    
    'Limpa os Labels
    TotalBaixar.Caption = Format(0, "Fixed")
    
    'Limpa os Grid
    Call Grid_Limpa(objGridParcelas)
    
    '#################################################
    'Inserido por Wagner 20/06/2006
    Call Grid_Limpa(objGridChequePre)
    
    'Torna Frame atual invisível
    Frame1(Opcao.SelectedItem.Index).Visible = False
    iFrameAtual = TAB_Titulos
    'Torna Frame atual visível
    Frame1(iFrameAtual).Visible = True
    Opcao.Tabs.Item(iFrameAtual).Selected = True
    
    Call Opcao_Click
    
    TotalCheque.Caption = ""
    '#################################################
    
    'Deixa as ComboBOx sem ítem selecionado
    Filial.Clear
    
    'Passa a data do sistema para a tela
    DataBaixa = Format(gdtDataAtual, "dd/mm/yy")
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    'Zera iAlterado
    iAlterado = 0

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
Dim iLinha As Integer
Dim iCodigo As Integer
Dim sNomeRed As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objFornecedor As New ClassFornecedor, objTipoFornecedor As New ClassTipoFornecedor
Dim objFilial As New ClassFilialFornecedor, sContaTela As String
Dim objFilialCliente  As New ClassFilialCliente

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case FORNECEDOR_COD
            
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 41975
                
                objMnemonicoValor.colValor.Add objFornecedor.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case FORNECEDOR_NOME
        
            'Preenche NomeReduzido com o fornecedor da tela
            If Len(Trim(Fornecedor.Text)) > 0 Then
                
                objFornecedor.sNomeReduzido = Fornecedor.Text
                lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
                If lErro <> SUCESSO Then Error 41976
            
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
                If lErro <> SUCESSO Then Error 41977
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 41978
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41979
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CGC_CPF
            
            If Len(Trim(Filial.Text)) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", Fornecedor.Text, objFilial)
                If lErro <> SUCESSO Then Error 41980
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case DATA_BAIXA
            If Len(DataBaixa.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataBaixa.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case TIPO1
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                    
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
                    
            Next
            
        Case VALOR_BAIXADO
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixar_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixar_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
                            
            Next
            
        Case VALOR_DESCONTO
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Desconto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_Desconto_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case VALOR_MULTA
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Multa_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_Multa_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
            
        Case VALOR_JUROS
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
            
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Juros_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_Juros_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
            
'######################################################################
'Inserido por Wagner 19/07/2006
        Case VALOR_BAIXADO_CT
            For iLinha = 1 To objGridChequePre.iLinhasExistentes
                
                If Len(GridChequePre.TextMatrix(iLinha, iGrid_ValorBaixadoCT_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridChequePre.TextMatrix(iLinha, iGrid_ValorBaixadoCT_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
                            
            Next
            
        Case VALOR_DESCONTO_CT
            For iLinha = 1 To objGridChequePre.iLinhasExistentes
                
                If Len(GridChequePre.TextMatrix(iLinha, iGrid_ValorDescontoCT_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridChequePre.TextMatrix(iLinha, iGrid_ValorDescontoCT_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case VALOR_MULTA_CT
            For iLinha = 1 To objGridChequePre.iLinhasExistentes
                
                If Len(GridChequePre.TextMatrix(iLinha, iGrid_ValorMultaCT_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridChequePre.TextMatrix(iLinha, iGrid_ValorMultaCT_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
            
        Case VALOR_JUROS_CT
            For iLinha = 1 To objGridChequePre.iLinhasExistentes
            
                If Len(GridChequePre.TextMatrix(iLinha, iGrid_ValorJurosCT_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridChequePre.TextMatrix(iLinha, iGrid_ValorJurosCT_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
            
        Case SELECIONADOCT1
            For iLinha = 1 To objGridChequePre.iLinhasExistentes
                objMnemonicoValor.colValor.Add CInt(GridChequePre.TextMatrix(iLinha, iGrid_SelecionadoCT_Col))
            Next
'#############################################################################
        
        Case "FilialCli_Conta_Ctb"
        
            For iLinha = 1 To objGridChequePre.iLinhasExistentes
                Set objFilialCliente = New ClassFilialCliente
                objFilialCliente.iCodFilial = Codigo_Extrai(GridChequePre.TextMatrix(iLinha, iGrid_FilialCT_Col))
                If objFilialCliente.iCodFilial <> 0 Then
                
                    lErro = CF("FilialCliente_Le_NomeRed_CodFilial", GridChequePre.TextMatrix(iLinha, iGrid_ClienteCT_Col), objFilialCliente)
                    If lErro <> SUCESSO Then Error 56522
                    
                    If objFilialCliente.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objFilialCliente.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then Error 56523
                    
                    Else
                    
                        sContaTela = ""
                        
                    End If
                
                Else
                
                    sContaTela = ""
                
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
            Next
        
        Case SELECIONADA1
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                    
                objMnemonicoValor.colValor.Add CInt(GridParcelas.TextMatrix(iLinha, iGrid_Selecionar_Col))
                
            Next
            
        Case COBRANCA1
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case VALOR_PAGAR
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
            
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_ValorPagar_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorPagar_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case FILIAL_EMPRESA
            
'            If giTipoVersao = VERSAO_FULL Then
            
                For iLinha = 1 To objGridParcelas.iLinhasExistentes
                    
                    If Len(GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col)) > 0 Then
                        objMnemonicoValor.colValor.Add Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col))
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
    
                Next
            
'            ElseIf giTipoVersao = VERSAO_LIGHT Then
'                objMnemonicoValor.colValor.Add giFilialEmpresa
'            End If

        Case NUM_TITULO1
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add StrParaLong(GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case NUM_PARCELA
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add StrParaLong(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
            
        Case Else
            Error 39678
            
    End Select

    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err
    
    Select Case Err
        
        Case 39678
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 39696, 41975 To 41980
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155393)
    
    End Select
    
    Exit Function

End Function

Private Function Carrega_Tab_Parcelas() As Long
'Carrega os dados das parcelas, créditos e pagamentos antecipados para tela

Dim lErro As Long
Dim iCodFilialFornecedor As Integer
Dim iFilialForn As Integer
Dim lCodForn As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim colCreditos As New Collection
Dim colPagtoAntecipado As New Collection
Dim colChequePre As New Collection
Dim dtEmissaoInic As Date
Dim dtEmissaoFim As Date
Dim dtVencInic As Date
Dim dtVencFim As Date
Dim lTituloInic As Long
Dim lTituloFim As Long
    
On Error GoTo Erro_Carrega_Tab_Parcelas
    
    'Verifica se Fornecedor e a filial estão preenchidos
    If Len(Trim(Fornecedor.Text)) = 0 Then Error 15917
    If Len(Trim(Filial.Text)) = 0 Then Error 15918

    If iFramePrincipalAlterado <> 0 Then

        'Limpa o frame Parcelas
        Call Frame_Parcelas_Limpa

        'Lê os dados do Fornecedor
        lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilialFornecedor)
        If lErro <> SUCESSO Then Error 29009
        
        'PassaFornecedor que está na tela para o Obj
        lCodForn = objFornecedor.lCodigo
        objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    
        'Passa a Filial que está na tela para o Obj
        iFilialForn = Codigo_Extrai(Filial.Text)
        objFilialFornecedor.iCodFilial = iFilialForn
    
        dtEmissaoInic = MaskedParaDate(EmissaoInic)
        dtEmissaoFim = MaskedParaDate(EmissaoFim)
        dtVencInic = MaskedParaDate(VencInic)
        dtVencFim = MaskedParaDate(VencFim)
            
        'Se as datas EmissaoFim e EmissaoInic estão preenchidas
        If dtEmissaoInic <> DATA_NULA And dtEmissaoFim <> DATA_NULA Then
            'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
            If dtEmissaoFim < dtEmissaoInic Then Error 15919
        End If
    
        'Se as datas VencFim e VencInic estão preenchidas
        If dtVencInic <> DATA_NULA And dtVencFim <> DATA_NULA Then
            'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
            If dtVencFim < dtVencInic Then Error 15920
        End If
    
        'Lê TituloInic e TituloFim que estão na tela
        lTituloInic = StrParaLong(TituloInic.Text)
        lTituloFim = StrParaLong(TituloFim.Text)
            
        'Verifica se TituloFim é maior ou igual que TituloInic
        If lTituloFim < lTituloInic Then Error 15921
        
        'Limpa gcolInfoParcPag antes de carregar as novas parcelas
        Set gcolInfoParcPag = New Collection

        'Limpa gcolInfoParcPag antes de carregar as novas parcelas
        Set gColNumIntPag = New Collection

        'Preenche a Coleção de Parcelas
        lErro = CF("ParcelasPagar_Le_BaixaManual", lCodForn, iFilialForn, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, gcolInfoParcPag)
        If lErro <> SUCESSO Then Error 15922
    
        'Verifica o número máximo de parcelas
        If gcolInfoParcPag.Count > NUM_MAX_PARCELAS_BAIXA Then Error 43092
                       
        '#########################################################
        'Inserido por Wagner 06/06/2006
        'Preenche a Coleção de ChequePre
        lErro = CF("ChequePre_Le_Todos", colChequePre, giFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 15935 Then Error 15924
        '#########################################################
        
        'Passa a data do sistema para a tela
        DataBaixa = Format(gdtDataAtual, "dd/mm/yy")
    
        'Preenche o GridParcelas
        Call Grid_Parcelas_Preenche(gcolInfoParcPag)
    
        '#########################################################
        'Inserido por Wagner 06/06/2006
        Call Grid_ChequePre_Preenche(colChequePre)
        '#########################################################
    
    End If
        
    Carrega_Tab_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Carrega_Tab_Parcelas:

    Carrega_Tab_Parcelas = Err
    
    Select Case Err
    
        Case 15917
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", Err)
            
        Case 15918
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
            
        Case 15919
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_INICIAL_MAIOR", Err)
            
        Case 15920
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_INICIAL_MAIOR", Err)
            
        Case 15921
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", Err)
            
        Case 15922, 15923, 15924, 29009
            
        Case 43092
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAS_SUPERIOR_NUM_MAX_PARCELAS_BAIXA", Err)
            
        Case 43093
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CREDITOS_SUPERIOR_NUM_MAX_CREDITOS", Err)
            
        Case 43094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PAGTOSANTECIPADOS_SUPERIOR_NUM_MAX_PAGTOS_ANTECIPADOS", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155394)
            
    End Select
    
    Exit Function
        
End Function

Private Sub Calcula_Total_Pagar()

Dim iIndice As Integer
Dim dTotalSelecionado As Double

    dTotalSelecionado = 0
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        If StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_Selecionar_Col)) <> 0 Then
            
            dTotalSelecionado = dTotalSelecionado + StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorPagar_Col))
        End If
    Next
    
    If dTotalSelecionado > 0 Then
        TotalBaixar.Caption = Format(dTotalSelecionado, "Standard")
    Else
        TotalBaixar.Caption = ""
    End If
    
    Exit Sub
    
End Sub

Private Function Parcela_RecalculaValorPagar(sValorBaixar As String, sValorDesconto As String, sValorMulta As String, sValorJuros As String) As Long
'obs: chamada dentro de saida celula

Dim lErro As Long, dValorPagar As Double

On Error GoTo Erro_Parcela_RecalculaValorPagar

    'Lê os dados da tela
    dValorPagar = StrParaDbl(sValorBaixar) - StrParaDbl(sValorDesconto) + StrParaDbl(sValorMulta) + StrParaDbl(sValorJuros)
    
    If dValorPagar < 0 Then Error 56757
    
    'Preenche ValorPagar
    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorPagar_Col) = Format(dValorPagar, "Standard")

    Parcela_RecalculaValorPagar = SUCESSO
     
    Exit Function
    
Erro_Parcela_RecalculaValorPagar:

    Parcela_RecalculaValorPagar = Err
     
    Select Case Err
          
        Case 56757
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PAGAR_NEGATIVO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155398)
     
    End Select
     
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BAIXA_PARCELAS_PAGAR_TITULO
    Set Form_Load_Ocx = Me
    Caption = "Baixa de Parcelas a Pagar com cheque de terceiros"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BaixaPagChequeTerc"
    
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
            Call FornecLabel_Click
        End If
    
    End If
    
End Sub

Public Sub BotaoConsultaDocOriginal_Click()
'Exibe na tela o documento que originou a parcela, o adiantamento ou o crédito selecionado
'Esse botão não está mais visível na tela, mas esse código ainda é chamado a partir de mnuGridConsultaDocOriginal_Click

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor

On Error GoTo Erro_BotaoConsultaDocOriginal_Click
    
    'Se nenhum grid foi selecionado => erro
    If iGridAtual = NENHUM_GRID_SELECIONADO Then gError 79883
    
    'Se o Fornecedor não foi preenchido => erro
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 79884
    
    'Se a filial não foi preenchida => erro
    If Len(Trim(Filial.Text)) = 0 Then gError 79891
    
    'Guarda no obj o parâmetro que será utilizado na obtenção do código do fornecedor
    objFornecedor.sNomeReduzido = Fornecedor.Text
        
    'Obtém o código do fornecedor que será passado como parâmetro
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79876
    
    'Se não encontrou o fornecedor = > erro
    If lErro = 6681 Then gError 79882
    
    'Guarda no objFilialFornecedor o código do fornecedor e da filial
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    objFilialFornecedor.iCodFilial = Codigo_Extrai(Filial.Text)
   
    'Verifica em qual grid está o documento que será consultado
    Select Case iGridAtual
    
        'Se for o grid parcelas
        Case GRID_PARCELAS
        
            'Abre a tela de consulta de títulos a pagar
            lErro = Consulta_TituloPagar(objFilialFornecedor)
            If lErro <> SUCESSO Then gError 79877
          
    End Select

    Exit Sub
    
Erro_BotaoConsultaDocOriginal_Click:

    Select Case gErr
    
        Case 79876 To 79879
        
        Case 79882
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case 79883
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 79884
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 79891
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155399)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub mnuGridMarcarTodos_Click()
    Select Case iGridAtual
        'Se for o grid parcelas
        Case GRID_PARCELAS
            Call Grid_Marca_Desmarca(objGridParcelas, iGrid_Selecionar_Col, MARCADO)
    End Select
    Call Calcula_Total_Pagar
End Sub

Public Sub mnuGridDesmarcarTodos_Click()
    Select Case iGridAtual
        'Se for o grid parcelas
        Case GRID_PARCELAS
            Call Grid_Marca_Desmarca(objGridParcelas, iGrid_Selecionar_Col, DESMARCADO)
    End Select
    TotalBaixar.Caption = ""
End Sub

Private Function Consulta_TituloPagar(objFilialFornecedor As ClassFilialFornecedor) As Long
'Exibe o título referente à parcela selecionada no grid
            
Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
            
On Error GoTo Erro_Consulta_TituloPagar

    'Se não há linha selecionada no grid parcelas => erro
    If GridParcelas.Row <= 0 Then gError 79874
    
    'Se a linha selecionada não contém dados => erro
    If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col))) = 0 Then gError 79875
    
    'Preenche o obj com os dados que serão utilizados para leitura do título que será consultado
    With objTituloPagar
        
        .lFornecedor = objFilialFornecedor.lCodFornecedor
        .iFilial = objFilialFornecedor.iCodFilial
        .sSiglaDocumento = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col)
        .lNumTitulo = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col)
        .dtDataEmissao = StrParaDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataEmissaoTitulo_Col))
    
    End With
    
    Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    Consulta_TituloPagar = SUCESSO
    
    Exit Function

Erro_Consulta_TituloPagar:

    Consulta_TituloPagar = gErr
    
    Select Case gErr
    
        Case 79874
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 79875
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155400)
    
    End Select
    
    Exit Function
    
End Function

Public Sub GridParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Seta o Grid Parcelas como grid atual
    iGridAtual = GRID_PARCELAS

    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then

        'Seta objTela como a Tela de Baixas a Receber
        Set PopUpMenuGridMD.objTela = Me

        'Chama o Menu PopUp
        PopUpMenuGridMD.PopupMenu PopUpMenuGridMD.mnuGrid, vbPopupMenuRightButton

        'Limpa o objTela
        Set PopUpMenuGridMD.objTela = Nothing

    End If

End Sub

Public Sub mnuGridConsultaDocOriginal_Click()
    Call BotaoConsultaDocOriginal_Click
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
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134053

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134053

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155403)

    End Select
    
    Exit Sub

End Sub

'#############################################################
'Inserido por Wagner 06/0/2006
Private Function Inicializa_Grid_ChequePre(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIncremento As Integer

On Error GoTo Erro_Inicializa_Grid_ChequePre

    Set objGridInt = New AdmGrid

    'Tela em questão
    Set objGridInt.objForm = Me

    'Títulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Data de Deposito")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Banco")
    objGridInt.colColuna.Add ("Agência")
    objGridInt.colColuna.Add ("Conta Corrente")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Vlr Baixar Rec")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Filial Empresa")
    
    objGridInt.colCampo.Add (SelecionadoCT.Name)
    objGridInt.colCampo.Add (DataDepositoCT.Name)
    objGridInt.colCampo.Add (ValorCT.Name)
    objGridInt.colCampo.Add (ClienteCT.Name)
    objGridInt.colCampo.Add (FilialCT.Name)
    objGridInt.colCampo.Add (BancoCT.Name)
    objGridInt.colCampo.Add (AgenciaCT.Name)
    objGridInt.colCampo.Add (ContaCorrenteCT.Name)
    objGridInt.colCampo.Add (NumeroCT.Name)
    objGridInt.colCampo.Add (ValorBaixadoCT.Name)
    objGridInt.colCampo.Add (ValorDescontoCT.Name)
    objGridInt.colCampo.Add (ValorMultaCT.Name)
    objGridInt.colCampo.Add (ValorJurosCT.Name)
    objGridInt.colCampo.Add (FilialEmpresaCT.Name)
    
    iGrid_SelecionadoCT_Col = 1
    iGrid_DataDepositoCT_Col = 2
    iGrid_ValorCT_Col = 3
    iGrid_ClienteCT_Col = 4
    iGrid_FilialCT_Col = 5
    iGrid_BancoCT_Col = 6
    iGrid_AgenciaCT_Col = 7
    iGrid_ContaCorrenteCT_Col = 8
    iGrid_NumeroCT_Col = 9
    iGrid_ValorBaixadoCT_Col = 10
    iGrid_ValorDescontoCT_Col = 11
    iGrid_ValorMultaCT_Col = 12
    iGrid_ValorJurosCT_Col = 13
    iGrid_FilialEmpresaCT_Col = 14

    objGridInt.objGrid = GridChequePre

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    GridChequePre.ColWidth(0) = 400

    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR

    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ChequePre = SUCESSO

    Exit Function

Erro_Inicializa_Grid_ChequePre:

    Inicializa_Grid_ChequePre = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143637)

    End Select

    Exit Function

End Function

Public Sub GridChequePre_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridChequePre, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequePre, iAlterado)
    End If

End Sub

Public Sub GridChequePre_GotFocus()

    Call Grid_Recebe_Foco(objGridChequePre)

End Sub

Public Sub GridChequePre_EnterCell()

    Call Grid_Entrada_Celula(objGridChequePre, iAlterado)

End Sub

Public Sub GridChequePre_LeaveCell()

    Call Saida_Celula(objGridChequePre)

End Sub

Public Sub GridChequePre_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridChequePre)

End Sub

Public Sub GridChequePre_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridChequePre, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridChequePre, iAlterado)
    End If

End Sub

Public Sub GridChequePre_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridChequePre)

End Sub

Public Sub GridChequePre_RowColChange()

    Call Grid_RowColChange(objGridChequePre)

End Sub

Public Sub GridChequePre_Scroll()

    Call Grid_Scroll(objGridChequePre)

End Sub

Public Sub SelecionadoCT_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridChequePre)

End Sub

Public Sub SelecionadoCT_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridChequePre)

End Sub

Public Sub SelecionadoCT_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridChequePre.objControle = SelecionadoCT
    lErro = Grid_Campo_Libera_Foco(objGridChequePre)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub SelecionadoCT_Click()
  
    Call Calcula_Total_Cheque

End Sub

Private Sub Grid_ChequePre_Preenche(colChequePre As Collection)
'Preenche o Grid ChequePre com os dados de colChequePre

Dim lErro As Long
Dim iLinha As Integer
Dim objChequepre As ClassChequePre
Dim objFilialCliente As ClassFilialCliente
Dim sNomeFilial As String
Dim objCliente As ClassCliente
Dim objFilialEmpresa As New AdmFiliais
Dim objInfoParcRec As New ClassInfoParcRec
Dim colInfoParcRec As Collection
Dim dValorBaixado As Double
Dim dValorDesconto As Double
Dim dValorJuros As Double
Dim dValorMulta As Double

On Error GoTo Erro_Grid_ChequePre_Preenche

    Set gcolCheques = colChequePre

    For Each objChequepre In colChequePre
    
        Set objFilialCliente = New ClassFilialCliente
        
        iLinha = iLinha + 1

        Set objCliente = New ClassCliente
        Set objFilialCliente = New ClassFilialCliente
        
        objCliente.lCodigo = objChequepre.lCliente

        If objCliente.lCodigo <> 0 Then
        
            lErro = CF("Cliente_Le", objCliente)
            If lErro <> SUCESSO Then gError 180240
                        
            objFilialCliente.iCodFilial = objChequepre.iFilial
            objFilialCliente.lCodCliente = objCliente.lCodigo
            
            lErro = CF("FilialCliente_Le", objFilialCliente)
            If lErro <> SUCESSO Then gError 180241
            
            sNomeFilial = CStr(objChequepre.iFilial) & SEPARADOR & CStr(objFilialCliente.sNome)
        Else
        
            objCliente.sNomeReduzido = "Não Especificado"
            sNomeFilial = "Não Especificado"
        End If
        
        GridChequePre.TextMatrix(iLinha, iGrid_ClienteCT_Col) = objCliente.sNomeReduzido
        GridChequePre.TextMatrix(iLinha, iGrid_FilialCT_Col) = sNomeFilial
        GridChequePre.TextMatrix(iLinha, iGrid_BancoCT_Col) = objChequepre.iBanco
        GridChequePre.TextMatrix(iLinha, iGrid_AgenciaCT_Col) = objChequepre.sAgencia
        GridChequePre.TextMatrix(iLinha, iGrid_ContaCorrenteCT_Col) = objChequepre.sContaCorrente
        GridChequePre.TextMatrix(iLinha, iGrid_NumeroCT_Col) = objChequepre.lNumero
        GridChequePre.TextMatrix(iLinha, iGrid_DataDepositoCT_Col) = Format(objChequepre.dtDataDeposito, "dd/mm/yyyy")
        GridChequePre.TextMatrix(iLinha, iGrid_ValorCT_Col) = Format(objChequepre.dValor, "Standard")

        Set colInfoParcRec = New Collection

        'Le as Parcelas a Receber associadas ao cheque
        lErro = CF("ParcelasReceber_Le_ChequePre", objChequepre, colInfoParcRec)
        If lErro <> SUCESSO Then gError 181229
        
        dValorBaixado = 0
        dValorDesconto = 0
        dValorJuros = 0
        dValorMulta = 0
        For Each objInfoParcRec In colInfoParcRec

            dValorBaixado = dValorBaixado + objInfoParcRec.dValorReceber
            dValorDesconto = dValorDesconto + objInfoParcRec.dValorDesconto
            dValorJuros = dValorJuros + objInfoParcRec.dValorJuros
            dValorMulta = dValorMulta + objInfoParcRec.dValorMulta

        Next

        GridChequePre.TextMatrix(iLinha, iGrid_ValorBaixadoCT_Col) = Format(dValorBaixado, "Standard")
        GridChequePre.TextMatrix(iLinha, iGrid_ValorDescontoCT_Col) = Format(dValorDesconto, "Standard")
        GridChequePre.TextMatrix(iLinha, iGrid_ValorJurosCT_Col) = Format(dValorJuros, "Standard")
        GridChequePre.TextMatrix(iLinha, iGrid_ValorMultaCT_Col) = Format(dValorMulta, "Standard")

        GridChequePre.TextMatrix(iLinha, iGrid_ValorCT_Col) = Format(objChequepre.dValor, "Standard")

        objFilialEmpresa.iCodFilial = objChequepre.iFilialEmpresa
        
        'Lê a Filial Empresa
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO Then gError 180242
            
        GridChequePre.TextMatrix(iLinha, iGrid_FilialEmpresaCT_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome

    Next
    
    'Passa para o Obj o número de Ceques passados pela Coleção
    objGridChequePre.iLinhasExistentes = colChequePre.Count

    Call Grid_Refresh_Checkbox(objGridChequePre)
    
    Exit Sub

Erro_Grid_ChequePre_Preenche:

    Select Case gErr

        Case 180240 To 180242, 181229
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180243)

    End Select

    Exit Sub
    
End Sub

Private Function Grava_Pagamento_ChequePre(dSomaPagar As Double, dtDataBaixa As Date, colBaixaParcPagar As colBaixaParcPagar) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim iChequePreSelecionado As Integer
Dim dValor As Double
Dim objChequepre As ClassChequePre
Dim colCheques As New Collection
Dim objFornecedor As New ClassFornecedor
Dim iFilial As Integer

On Error GoTo Erro_Grava_Pagamento_ChequePre

    iChequePreSelecionado = 0

    'Percorre todas as linhas do Grid PagtosChequePres
    For iIndice = 1 To objGridChequePre.iLinhasExistentes

        'Se a checkbox Selecionada está marcada
        If GridChequePre.TextMatrix(iIndice, iGrid_SelecionadoCT_Col) = 1 Then

            'Soma quantos Pagamentos ChequePres estão selecionados
            iChequePreSelecionado = iChequePreSelecionado + 1

            colCheques.Add gcolCheques.Item(iIndice)
        
        End If
    
    Next

    'Verifica se alguma célula Selecionada do Grid PagtosChequePres está marcada
    If iChequePreSelecionado < 1 Then gError 180244
    
    dValor = StrParaDbl(TotalCheque.Caption)
    
    If dSomaPagar - dValor < -DELTA_VALORMONETARIO Then gError 180245
    
    'Lê os dados do Fornecedor
    lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iFilial)
    If lErro <> SUCESSO Then gError 180248

    'Tem um parâmetro a mais para a contabilidade
    lErro = CF("BaixaPagar_Grava_ChequePre", colBaixaParcPagar, colCheques, dtDataBaixa, objFornecedor, iFilial, objContabil, StrParaDate(CTBDataContabil.Text))
    If lErro <> SUCESSO Then gError 180246

    Grava_Pagamento_ChequePre = SUCESSO
    
    Exit Function
    
Erro_Grava_Pagamento_ChequePre:

    Grava_Pagamento_ChequePre = gErr
    
    Select Case gErr
    
        Case 180244
            Call Rotina_Erro(vbOKOnly, "ERRO_CHEQUESPRE_NAO_MARCADO", gErr)
        
        Case 180245
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCHEQUES_MAIOR_VALORTITULOS", gErr, dValor, dSomaPagar)
            
        Case 180246, 180248

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 180247)

    End Select

    Exit Function
    
End Function

Private Sub Calcula_Total_Cheque()

Dim iIndice As Integer
Dim dTotalSelecionado As Double

    dTotalSelecionado = 0
    
    For iIndice = 1 To objGridChequePre.iLinhasExistentes
        
        If StrParaInt(GridChequePre.TextMatrix(iIndice, iGrid_SelecionadoCT_Col)) <> 0 Then
            
            dTotalSelecionado = dTotalSelecionado + StrParaDbl(GridChequePre.TextMatrix(iIndice, iGrid_ValorCT_Col))
        End If
    Next
    
    If dTotalSelecionado > 0 Then
        TotalCheque.Caption = Format(dTotalSelecionado, "Standard")
    Else
        TotalCheque.Caption = ""
    End If
    
    Exit Sub
    
End Sub
'#############################################################


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


