VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBaixaRecCancelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const NUM_LINHAS_GRID_PARC = 20
Private Const NUM_LINHAS_GRID_MOVCCI = 20
Private Const NUM_LINHAS_GRID_BAIXASMOV = 20

'inicio contabilidade

Public objGrid1 As AdmGrid
Public objContabil As New ClassContabil

Public WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Public WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const FILIAL_COD As String = "FilialCli_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialCli_Nome"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialCli_CGC_CPF"
Private Const DATA_CANCELAMENTO As String = "Data_Cancelamento"
Private Const NUM_TITULO As String = "Numero_Titulo"
Private Const PARCELA1 As String = "Parcela"
Private Const VALOR_CANCELAR As String = "Valor_Cancelar"
Private Const VALOR_DESCONTO As String = "Valor_Desconto"
Private Const VALOR_MULTA As String = "Valor_Multa"
Private Const VALOR_JUROS As String = "Valor_Juros"
Private Const PAGTO_DINHEIRO As String = "Pagto_Dinheiro"
Private Const PAGTO_ANTECIPADO As String = "Pagto_Antecipado"
Private Const CRED_DEV As String = "Credito_Devolução"
Private Const PERDA As String = "Perda"
Private Const CTA_CONTA_CORRENTE As String = "Conta_Contabil_Conta"
Private Const CTACARTEIRACOBRADOR As String = "Cta_CarteiraCobrador"
Private Const DATA_CREDITO As String = "Data_Credito"
Private Const BAIXAR As String = "Baixar"

'fim contabilidade

Private gcolCodNomeTiposMovCCI As New AdmColCodigoNome

'Property Variables:
Dim m_Caption As String
Event Unload()

'Declaração do obj global a tela (funciona como um espelho da tela,
'pois todas as informações contidas na tela devem constar nesse obj também)
'Ele será utilizado tanto para leitura de dados como para gravação também
Dim gobjBaixaRecCancelar As New ClassBaixaRecCancelar

'Declaração de obj para utilização de Browser
Private WithEvents objEventoClientes As AdmEvento
Attribute objEventoClientes.VB_VarHelpID = -1

'Variáveis para controle da tela
Public iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFrameAtual As Integer
Dim iFramePrincipalAlterado As Integer
Dim iFrameRecebimentoAtual As Integer
Dim iFrameDetalhesAtual As Integer

'GridBaixasMovimento
Dim objGridBaixasMovimento As AdmGrid
Dim iGrid_FilialEmpresaBaixada_Col As Integer
Dim iGrid_TipoBaixada_Col As Integer
Dim iGrid_NumeroBaixada_Col As Integer
Dim iGrid_ParcelaBaixada_Col As Integer
Dim iGrid_ValorBaixada_Col As Integer
Dim iGrid_ValorParcelaBaixada_Col As Integer
Dim iGrid_SequencialBaixada_Col As Integer

'GridMovimentosCCI
Dim objGridMovimentosCCI As AdmGrid
Dim iGrid_CancelarMov_Col As Integer
Dim iGrid_FilialEmpresaMov_Col As Integer
Dim iGrid_HistoricoMovCCI_Col As Integer
Dim iGrid_DataMovCCI_Col As Integer
Dim iGrid_CtaCorrenteMovCCI_Col As Integer
Dim iGrid_ValorMovCCI_Col As Integer
Dim iGrid_TipoMovCCI_Col As Integer

'Grid Parcelas:
Dim objGridParcelas As AdmGrid
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_DataBaixa_Col As Integer
Dim iGrid_DataCredito_Col As Integer
Dim iGrid_ValorBaixa_Col As Integer
Dim iGrid_ValorParcela_Col As Integer
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_ParcelaSelecionada_Col As Integer
Dim iGrid_MultaCancelar_Col As Integer
Dim iGrid_JurosCancelar_Col As Integer
Dim iGrid_DescontoCancelar_Col As Integer
Dim iGrid_ValorCancelar_Col As Integer
Dim iGrid_ValorTotalParcela_Col As Integer

'Tab Principal (Opcao)
Private Const TAB_Titulos = 1
Private Const TAB_Baixas = 2

'TabDetalhes
Private Const TAB_Cancelamento = 1
Private Const TAB_Detalhamento = 2

Private Const TELA_TODA = 1
Private Const CTACORRENTE_TODAS = 1
Private Const CTACORRENTE_APENAS = 2

'Número máximo de linhas nos grids
Private Const NUM_MAX_BAIXAS_GRID = 300
Private Const NUM_MAX_MOVCCI_GRID = 300

Dim m_objUserControl As Object

'--- inicio dos properties get dos controles da tela

Public Property Get TipoBaixada() As Object
     Set TipoBaixada = objUserControl.Controls("TipoBaixada")
End Property

Public Property Get ValorParcelaBaixada() As Object
     Set ValorParcelaBaixada = objUserControl.Controls("ValorParcelaBaixada")
End Property

Public Property Get ValorBaixada() As Object
     Set ValorBaixada = objUserControl.Controls("ValorBaixada")
End Property

Public Property Get ParcelaBaixada() As Object
     Set ParcelaBaixada = objUserControl.Controls("ParcelaBaixada")
End Property

Public Property Get NumeroBaixada() As Object
     Set NumeroBaixada = objUserControl.Controls("NumeroBaixada")
End Property

Public Property Get SequencialBaixada() As Object
     Set SequencialBaixada = objUserControl.Controls("SequencialBaixada")
End Property

Public Property Get GridBaixasMovimento() As Object
     Set GridBaixasMovimento = objUserControl.Controls("GridBaixasMovimento")
End Property

Public Property Get GridMovimentosCCI() As Object
     Set GridMovimentosCCI = objUserControl.Controls("GridMovimentosCCI")
End Property

Public Property Get CancelarMov() As Object
     Set CancelarMov = objUserControl.Controls("CancelarMov")
End Property

Public Property Get FilialEmpresaMov() As Object
     Set FilialEmpresaMov = objUserControl.Controls("FilialEmpresaMov")
End Property

Public Property Get HistoricoMovCCI() As Object
     Set HistoricoMovCCI = objUserControl.Controls("HistoricoMovCCI")
End Property

Public Property Get DataMovCCI() As Object
     Set DataMovCCI = objUserControl.Controls("DataMovCCI")
End Property

Public Property Get ValorMovCCI() As Object
     Set ValorMovCCI = objUserControl.Controls("ValorMovCCI")
End Property

Public Property Get TipoMovCCI() As Object
     Set TipoMovCCI = objUserControl.Controls("TipoMovCCI")
End Property

Public Property Get ValorTotalParcela() As Object
     Set ValorTotalParcela = objUserControl.Controls("ValorTotalParcela")
End Property

Public Property Get ValorCancelar() As Object
     Set ValorCancelar = objUserControl.Controls("ValorCancelar")
End Property

Public Property Get FilialEmpresaBaixada() As Object
     Set FilialEmpresaBaixada = objUserControl.Controls("FilialEmpresaBaixada")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get FrameDetalhamentoBaixa() As Object
     Set FrameDetalhamentoBaixa = objUserControl.Controls("FrameDetalhamentoBaixa")
End Property

Public Property Get FrameDadosBaixa() As Object
     Set FrameDadosBaixa = objUserControl.Controls("FrameDadosBaixa")
End Property

Public Property Get FrameBaixasMovCCI() As Object
     Set FrameBaixasMovCCI = objUserControl.Controls("FrameBaixasMovCCI")
End Property

Public Property Get CtaCorrenteMovCCI() As Object
     Set CtaCorrenteMovCCI = objUserControl.Controls("CtaCorrenteMovCCI")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get FilialEmpresa() As Object
     Set FilialEmpresa = objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get LabelFil() As Object
     Set LabelFil = objUserControl.Controls("LabelFil")
End Property

Public Property Get LabelCliente() As Object
     Set LabelCliente = objUserControl.Controls("LabelCliente")
End Property

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
End Property

Public Property Get Frame7() As Object
     Set Frame7 = objUserControl.Controls("Frame7")
End Property

Public Property Get UpDownBaixaInic() As Object
     Set UpDownBaixaInic = objUserControl.Controls("UpDownBaixaInic")
End Property

Public Property Get BaixaInic() As Object
     Set BaixaInic = objUserControl.Controls("BaixaInic")
End Property

Public Property Get UpDownBaixaFim() As Object
     Set UpDownBaixaFim = objUserControl.Controls("UpDownBaixaFim")
End Property

Public Property Get BaixaFim() As Object
     Set BaixaFim = objUserControl.Controls("BaixaFim")
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

Public Property Get FrameBaixaBaixa() As Object
     Set FrameBaixaBaixa = objUserControl.Controls("FrameBaixaBaixa")
End Property

Public Property Get FrameMovCCI() As Object
     Set FrameMovCCI = objUserControl.Controls("FrameMovCCI")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get Numero() As Object
     Set Numero = objUserControl.Controls("Numero")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get Sequencial() As Object
     Set Sequencial = objUserControl.Controls("Sequencial")
End Property

Public Property Get DataBaixaParc() As Object
     Set DataBaixaParc = objUserControl.Controls("DataBaixaParc")
End Property

Public Property Get DataCreditoParc() As Object
     Set DataCreditoParc = objUserControl.Controls("DataCreditoParc")
End Property

Public Property Get ValorPagoBaixa() As Object
     Set ValorPagoBaixa = objUserControl.Controls("ValorPagoBaixa")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
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

Public Property Get DataCancelamento() As Object
     Set DataCancelamento = objUserControl.Controls("DataCancelamento")
End Property

Public Property Get TotalCancelar() As Object
     Set TotalCancelar = objUserControl.Controls("TotalCancelar")
End Property

Public Property Get ItensSelecionados() As Object
     Set ItensSelecionados = objUserControl.Controls("ItensSelecionados")
End Property

Public Property Get UpDownDataCancelamento() As Object
     Set UpDownDataCancelamento = objUserControl.Controls("UpDownDataCancelamento")
End Property

Public Property Get HistoricoCancelamento() As Object
     Set HistoricoCancelamento = objUserControl.Controls("HistoricoCancelamento")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Recebimento() As Object
     Set Recebimento = objUserControl.Controls("Recebimento")
End Property

Public Property Get FrameRecebimento() As Object
     Set FrameRecebimento = objUserControl.Controls("FrameRecebimento")
End Property

Public Property Get SaldoDebito() As Object
     Set SaldoDebito = objUserControl.Controls("SaldoDebito")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
End Property

Public Property Get FilialEmpresaCR() As Object
     Set FilialEmpresaCR = objUserControl.Controls("FilialEmpresaCR")
End Property

Public Property Get ValorDebito() As Object
     Set ValorDebito = objUserControl.Controls("ValorDebito")
End Property

Public Property Get SiglaDocumentoCR() As Object
     Set SiglaDocumentoCR = objUserControl.Controls("SiglaDocumentoCR")
End Property

Public Property Get DataEmissaoCred() As Object
     Set DataEmissaoCred = objUserControl.Controls("DataEmissaoCred")
End Property

Public Property Get CCIntNomeReduzido() As Object
     Set CCIntNomeReduzido = objUserControl.Controls("CCIntNomeReduzido")
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

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get ContaCorrente() As Object
     Set ContaCorrente = objUserControl.Controls("ContaCorrente")
End Property

Public Property Get HistoricoPerda() As Object
     Set HistoricoPerda = objUserControl.Controls("HistoricoPerda")
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

Public Property Get TabDetalhes() As Object
     Set TabDetalhes = objUserControl.Controls("TabDetalhes")
End Property

Public Property Get ClienteTitCartao() As Object
     Set ClienteTitCartao = objUserControl.Controls("ClienteTitCartao")
End Property

Public Property Get FilialTitCartao() As Object
     Set FilialTitCartao = objUserControl.Controls("FilialTitCartao")
End Property

Public Property Get TipoTitCartao() As Object
     Set TipoTitCartao = objUserControl.Controls("TipoTitCartao")
End Property

Public Property Get NumTitCartao() As Object
     Set NumTitCartao = objUserControl.Controls("NumTitCartao")
End Property

Public Property Get DataEmissaoTitCartao() As Object
     Set DataEmissaoTitCartao = objUserControl.Controls("DataEmissaoTitCartao")
End Property



'--- fim dos properties get dos controles da tela

'Tratamento do GridParcelas
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

Public Sub GridParcelas_LeaveCell()
    Call Saida_Celula(objGridParcelas)
End Sub
'Fim do tratamento do GridParcelas

'Tratamento do GridMovimentosCCI
Public Sub GridMovimentosCCI_Click()
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGridMovimentosCCI, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMovimentosCCI, iAlterado)
    Else
        
'        If objGridMovimentosCCI.iLinhaAntiga <> GridMovimentosCCI.Row Then
'
'            Call Traz_Parcelas_MovCCI_Tela
'
'        End If
    
    End If
    
End Sub

Public Sub GridMovimentosCCI_GotFocus()
    Call Grid_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub GridMovimentosCCI_EnterCell()
    Call Grid_Entrada_Celula(objGridMovimentosCCI, iAlterado)
End Sub

Public Sub GridMovimentosCCI_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridMovimentosCCI)
End Sub

Public Sub GridMovimentosCCI_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer
    Call Grid_Trata_Tecla(KeyAscii, objGridMovimentosCCI, iExecutaEntradaCelula)
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridMovimentosCCI, iAlterado)
    End If
End Sub

Public Sub GridMovimentosCCI_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridMovimentosCCI)
End Sub

Public Sub GridMovimentosCCI_RowColChange()
    Call Grid_RowColChange(objGridMovimentosCCI)
End Sub

Public Sub GridMovimentosCCI_Scroll()
    Call Grid_Scroll(objGridMovimentosCCI)
End Sub

Public Sub GridMovimentosCCI_LeaveCell()
    Call Saida_Celula(objGridMovimentosCCI)
End Sub
'Fim do tratamento do GridMovimentosCCI

'Tratamento do GridBaixasMovimento
Public Sub GridBaixasMovimento_Click()
Dim iExecutaEntradaCelula As Integer
    Call Grid_Click(objGridBaixasMovimento, iExecutaEntradaCelula)
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBaixasMovimento, iAlterado)
    End If
End Sub

Public Sub GridBaixasMovimento_GotFocus()
    Call Grid_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub GridBaixasMovimento_EnterCell()
    Call Grid_Entrada_Celula(objGridBaixasMovimento, iAlterado)
End Sub

Public Sub GridBaixasMovimento_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridBaixasMovimento)
End Sub

Public Sub GridBaixasMovimento_KeyPress(KeyAscii As Integer)
Dim iExecutaEntradaCelula As Integer
    Call Grid_Trata_Tecla(KeyAscii, objGridBaixasMovimento, iExecutaEntradaCelula)
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBaixasMovimento, iAlterado)
    End If
End Sub

Public Sub GridBaixasMovimento_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridBaixasMovimento)
End Sub

Public Sub GridBaixasMovimento_RowColChange()
    Call Grid_RowColChange(objGridBaixasMovimento)
End Sub

Public Sub GridBaixasMovimento_Scroll()
    Call Grid_Scroll(objGridBaixasMovimento)
End Sub

Public Sub GridBaixasMovimento_LeaveCell()
    Call Saida_Celula(objGridBaixasMovimento)
End Sub
'Fim do tratamento do GridBaixasMovimento

'Tratamento dos campos do GridParcelas
Public Sub FilialEmpresa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub FilialEmpresa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub FilialEmpresa_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = FilialEmpresa
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
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

Public Sub DataBaixaParc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub DataBaixaParc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub DataBaixaParc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = DataBaixaParc
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorPagoBaixa_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub ValorPagoBaixa_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub ValorPagoBaixa_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = ValorPagoBaixa
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
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

Public Sub Sequencial_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub Sequencial_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub Sequencial_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = Sequencial
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorCancelar_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorCancelar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub ValorCancelar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub ValorCancelar_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = ValorCancelar
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub MultaCancelar_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub MultaCancelar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub MultaCancelar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub MultaCancelar_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = MultaCancelar
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Property Get MultaCancelar() As Object
     Set MultaCancelar = objUserControl.Controls("MultaCancelar")
End Property

Public Sub JurosCancelar_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub JurosCancelar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub JurosCancelar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub JurosCancelar_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = JurosCancelar
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Property Get JurosCancelar() As Object
     Set JurosCancelar = objUserControl.Controls("JurosCancelar")
End Property

Public Sub DescontoCancelar_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DescontoCancelar_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub DescontoCancelar_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub DescontoCancelar_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = DescontoCancelar
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Property Get DescontoCancelar() As Object
     Set DescontoCancelar = objUserControl.Controls("DescontoCancelar")
End Property

Public Sub ValorTotalParcela_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorTotalParcela_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub ValorTotalParcela_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub ValorTotalParcela_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridParcelas.objControle = ValorTotalParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
End Sub
'Fim do tratamento dos campos do GridParcelas

'Tratamento dos campos do GridMovimentosCCI
Public Sub CancelarMov_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CancelarMov_Click()
'Alterado por Wagner

Dim iIndice As Integer

On Error GoTo Erro_CancelarMov_Click

    For iIndice = 1 To objGridMovimentosCCI.iLinhasExistentes
        
        gobjBaixaRecCancelar.colInfoBaixaRecCancelar(iIndice).iItemSelecionado = StrParaInt(GridMovimentosCCI.TextMatrix(iIndice, iGrid_CancelarMov_Col))
    
    Next
    
    If GridMovimentosCCI.TextMatrix(GridMovimentosCCI.Row, iGrid_CancelarMov_Col) = MARCADO Then Call Traz_Parcelas_MovCCI_Tela
    
'    'Se o usuário selecionou a linha
'    If GridMovimentosCCI.TextMatrix(GridMovimentosCCI.Row, iGrid_CancelarMov_Col) = True Then
'
'        'Guarda no obj a informação de que a linha está selecionada
'        gobjBaixaRecCancelar.colInfoBaixaRecCancelar(GridMovimentosCCI.Row).iItemSelecionado = MARCADO
'
'        Call Traz_Parcelas_MovCCI_Tela
'
'    'Senão
'    Else
'
'        'Guarda no obj global a informação de que a linha não está selecionada
'        gobjBaixaRecCancelar.colInfoBaixaRecCancelar(GridMovimentosCCI.Row).iItemSelecionado = DESMARCADO
'
'    End If
    
    'Atualiza o valor total do cancelamento
    Call Calcula_Total_Cancelar(TIPOCANCELAMENTO_MOVCCI)
    
    'Atualiza os optionbutton do grid
    Call Grid_Refresh_Checkbox(objGridMovimentosCCI)
    
    Call Trata_Data_Cancelamento

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_CancelarMov_Click:

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155431)
        
    End Select
    
    Exit Sub
    
End Sub

Public Sub CancelarMov_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub CancelarMov_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMovimentosCCI)
End Sub

Public Sub CancelarMov_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMovimentosCCI.objControle = CancelarMov
    lErro = Grid_Campo_Libera_Foco(objGridMovimentosCCI)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub FilialEmpresaMov_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub FilialEmpresaMov_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMovimentosCCI)
End Sub

Public Sub FilialEmpresaMov_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMovimentosCCI.objControle = FilialEmpresaMov
    lErro = Grid_Campo_Libera_Foco(objGridMovimentosCCI)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub DataMovCCI_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub DataMovCCI_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMovimentosCCI)
End Sub

Public Sub DataMovCCI_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMovimentosCCI.objControle = DataMovCCI
    lErro = Grid_Campo_Libera_Foco(objGridMovimentosCCI)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub CtaCorrenteMovCCI_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub CtaCorrenteMovCCI_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMovimentosCCI)
End Sub

Public Sub CtaCorrenteMovCCI_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMovimentosCCI.objControle = CtaCorrenteMovCCI
    lErro = Grid_Campo_Libera_Foco(objGridMovimentosCCI)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub TipoMovCCI_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub TipoMovCCI_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMovimentosCCI)
End Sub

Public Sub TipoMovCCI_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMovimentosCCI.objControle = TipoMovCCI
    lErro = Grid_Campo_Libera_Foco(objGridMovimentosCCI)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorMovCCI_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub ValorMovCCI_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMovimentosCCI)
End Sub

Public Sub ValorMovCCI_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMovimentosCCI.objControle = ValorMovCCI
    lErro = Grid_Campo_Libera_Foco(objGridMovimentosCCI)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub HistoricoMovCCI_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridMovimentosCCI)
End Sub

Public Sub HistoricoMovCCI_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridMovimentosCCI)
End Sub

Public Sub HistoricoMovCCI_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridMovimentosCCI.objControle = HistoricoMovCCI
    lErro = Grid_Campo_Libera_Foco(objGridMovimentosCCI)
    If lErro <> SUCESSO Then Cancel = True
End Sub
'Fim do tratamento dos campos do GridMovCCI

'Início do tratamento dos campos do GridBaixasMovimento
Public Sub NumeroBaixada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumeroBaixada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub NumeroBaixada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub NumeroBaixada_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = NumeroBaixada
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub FilialEmpresaBaixada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FilialEmpresaBaixada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub FilialEmpresaBaixada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub FilialEmpresaBaixada_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = FilialEmpresaBaixada
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub TipoBaixada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoBaixada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub TipoBaixada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub TipoBaixada_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = TipoBaixada
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ParcelaBaixada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ParcelaBaixada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub ParcelaBaixada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub ParcelaBaixada_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = ParcelaBaixada
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorParcelaBaixada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorParcelaBaixada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub ValorParcelaBaixada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub ValorParcelaBaixada_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = ValorParcelaBaixada
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub SequencialBaixada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub SequencialBaixada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub SequencialBaixada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub SequencialBaixada_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = SequencialBaixada
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub ValorBaixada_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorBaixada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBaixasMovimento)
End Sub

Public Sub ValorBaixada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBaixasMovimento)
End Sub

Public Sub ValorBaixada_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBaixasMovimento.objControle = ValorBaixada
    lErro = Grid_Campo_Libera_Foco(objGridBaixasMovimento)
    If lErro <> SUCESSO Then Cancel = True
End Sub
'Fim do tratamento dos campos do GridBaixasMovimento

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Parcelas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Filial Empresa")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Data Baixa")
    objGridInt.colColuna.Add ("Data Crédito")
    objGridInt.colColuna.Add ("Valor Baixado")
    objGridInt.colColuna.Add ("Valor Parcela")
    objGridInt.colColuna.Add ("Sequencial")
    objGridInt.colColuna.Add ("Valor Baixado Cancelar")
    objGridInt.colColuna.Add ("Multa Cancelar")
    objGridInt.colColuna.Add ("Juros Cancelar")
    objGridInt.colColuna.Add ("Desconto Cancelar")
    objGridInt.colColuna.Add ("Valor Total Cancelar")

    'campos de edição do grid
    objGridInt.colCampo.Add (ParcelaSelecionada.Name)
    objGridInt.colCampo.Add (FilialEmpresa.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (DataBaixaParc.Name)
    objGridInt.colCampo.Add (DataCreditoParc.Name)
    objGridInt.colCampo.Add (ValorPagoBaixa.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (Sequencial.Name)
    objGridInt.colCampo.Add (ValorCancelar.Name)
    objGridInt.colCampo.Add (MultaCancelar.Name)
    objGridInt.colCampo.Add (JurosCancelar.Name)
    objGridInt.colCampo.Add (DescontoCancelar.Name)
    objGridInt.colCampo.Add (ValorTotalParcela.Name)


    iGrid_ParcelaSelecionada_Col = 1
    iGrid_FilialEmpresa_Col = 2
    iGrid_Tipo_Col = 3
    iGrid_Numero_Col = 4
    iGrid_Parcela_Col = 5
    iGrid_DataBaixa_Col = 6
    iGrid_DataCredito_Col = 6 + 1
    iGrid_ValorBaixa_Col = 7 + 1
    iGrid_ValorParcela_Col = 8 + 1
    iGrid_Sequencial_Col = 9 + 1

    iGrid_ValorCancelar_Col = 10 + 1
    iGrid_MultaCancelar_Col = 11 + 1
    iGrid_JurosCancelar_Col = 12 + 1
    iGrid_DescontoCancelar_Col = 13 + 1
    iGrid_ValorTotalParcela_Col = 14 + 1

    objGridInt.objGrid = GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_LINHAS_GRID_PARC + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 5

    'largura da primeira coluna
    GridParcelas.ColWidth(0) = 0

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    
    'Indica para executar a Rotina Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Public Sub LabelCliente_Click()
'Chamada do Browse de Clientes

Dim colSelecao As Collection
Dim objCliente As New ClassCliente

    'Passa o Cliente que está na tela para o Obj
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    'Chama a tela com a lista de Clientes
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClientes)

    Exit Sub

End Sub

Private Sub objEventoClientes_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Me.Show

    'Preenche Cliente na tela com NomeReduzido
    Set objCliente = obj1

    Cliente.Text = CStr(objCliente.sNomeReduzido)

    'Chama Validate de Cliente
    Call Cliente_Validate(bCancel)

    Exit Sub

End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna o frame clicado visível
        Frame1(Opcao.SelectedItem.Index).Visible = True

        'Torna o frame atual invisível
        Frame1(iFrameAtual).Visible = False

        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        'Se Frame selecionado foi o de Títulos
        If Opcao.SelectedItem.Index = TAB_Titulos Then

'            iFramePrincipalAlterado = 0 '??? deve ser desnecessario

        'Se Frame selecionado foi o de Baixas
        ElseIf Opcao.SelectedItem.Index = TAB_Baixas And iFramePrincipalAlterado <> 0 Then

            lErro = Carrega_Tab_Baixas()
            If lErro <> SUCESSO Then gError 94810

            iFramePrincipalAlterado = 0

        End If

'        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual

            Case TAB_Titulos
                '??? criar idh Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_TITULOS

            Case TAB_Baixas
                '??? criar idh Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_PARCELAS

'            Case TAB_Contabilizacao
'                '??? criar idh Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_CONTABILIZACAO

        End Select

    End If
    
    Exit Sub
    
Erro_Opcao_Click:

    Select Case gErr
    
        Case 94810
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155432)
        
    End Select
    
    Exit Sub

End Sub

Public Sub Opcao_BeforeClick(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Opcao_BeforeClick

    Call TabStrip_TrataBeforeClick(Cancel, Opcao) '??? conferir se isto nao está em desuso
    
    'Se estava no tab de filtros e está passando para outro tab
    If iFrameAtual = TAB_Titulos Then
    
        'Valida os filtros
        lErro = Valida_Filtros()
        If lErro <> SUCESSO Then gError 94803
    
    End If

    Exit Sub

Erro_Opcao_BeforeClick:

    Cancel = True

    Select Case gErr

        Case 94803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155433)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Tab_Baixas() As Long
'Função que carrega o Frame de Baixa a Baixa

Dim lErro As Long

On Error GoTo Erro_Carrega_Tab_Baixas

    DoEvents '??? para que ?
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    Set gobjBaixaRecCancelar.colInfoBaixaRecCancelar = Nothing
    Set gobjBaixaRecCancelar.colParcelasMovCCI = Nothing
    
    Call Limpa_Tela_Tab_Baixas
    
    Select Case gobjBaixaRecCancelar.iTipoCancelamento

        Case TIPOCANCELAMENTO_BAIXABAIXA
        
            'Guarda no obj global a informação do tipo de baixa selecionado
            If TipoBaixaRecebimentos.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_RECEBIMENTO
            ElseIf TipoBaixaAdiantamentos.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_RECEBTO_ANTECIPADO
            ElseIf TipoBaixaDebitos.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_DEBITO_CLIENTE
            ElseIf TipoBaixaPerdas.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_PERDA
            ElseIf TipoBaixaCartao.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_CARTAO_DEBITO_CREDITO
            End If

            'preenche gobjBaixaRecCancelar.colInfoBaixaRecCancelar com baixas
            lErro = CF("BaixaRecCancelar_Le_BaixaBaixa", gobjBaixaRecCancelar)
            If lErro <> SUCESSO And lErro <> 94813 Then gError 94811
            
            If lErro = 94813 Then gError 94814
            
            'Verifica o número máximo de linhas do grid
            If gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count > NUM_MAX_BAIXAS_GRID Then gError 94812

            Call Traz_Baixas_Tela
            
        Case TIPOCANCELAMENTO_MOVCCI

            'preenche gobjBaixaRecCancelar.colInfoBaixaRecCancelar com movimentos associados a baixas
            lErro = CF("BaixaRecCancelar_Le_MovCCI", gobjBaixaRecCancelar)
            If lErro <> SUCESSO And lErro <> 94778 Then gError 94815
            
            If lErro = 94778 Then gError 94816
            
            'Verifica o número máximo de linhas do grid
            If gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count > NUM_MAX_MOVCCI_GRID Then gError 94817

            Call Traz_MovimentosCCI_Tela

    End Select
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Carrega_Tab_Baixas = SUCESSO

    Exit Function

Erro_Carrega_Tab_Baixas:

    Carrega_Tab_Baixas = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 94811, 94815
        
        Case 94812
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAS_SUPERIOR_NUM_MAX_BAIXA_GRID", gErr)

        Case 94814, 94816
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXARECCANCELAR_SEM_BAIXAS", gErr)
            
        Case 94817
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVCCI_SUPERIOR_NUM_MAX_MOVCCI_GRID", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155434)

    End Select

    Exit Function

End Function

Private Sub Traz_Baixas_Tela()
'Preenche o grid de parcelas p/cancelamento baixa a baixa

Dim iLinha As Integer, lErro As Long
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_Baixas_Tela

    'Recoloca o Numero de Linhas caso seja maior que o Numero de Linhas Visiveis na Tela
    If (gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count + 1) > NUM_LINHAS_GRID_PARC Then
        GridParcelas.Rows = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count + 1
    Else
        GridParcelas.Rows = NUM_LINHAS_GRID_PARC + 1
    End If
       
    'Renicializa
    Call Grid_Inicializa(objGridParcelas)
    
    For Each objInfoBaixaRecCancelar In gobjBaixaRecCancelar.colInfoBaixaRecCancelar
    
        'Incrementa a variável de controle da linha atual
        iLinha = iLinha + 1
        
        'Passa para a tela os dados da Parcela em questão
        If objFilialEmpresa.iCodFilial <> objInfoBaixaRecCancelar.iFilialEmpresa Then
        
            objFilialEmpresa.iCodFilial = objInfoBaixaRecCancelar.iFilialEmpresa
            'Lê a Filial Empresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError 81825
        
        End If
        
        GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objInfoBaixaRecCancelar.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome

        GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoBaixaRecCancelar.sSiglaDocumento
        GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col) = objInfoBaixaRecCancelar.lNumDocumento
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoBaixaRecCancelar.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_DataBaixa_Col) = Format(objInfoBaixaRecCancelar.dtDataBaixa, "dd/mm/yy")
        GridParcelas.TextMatrix(iLinha, iGrid_DataCredito_Col) = Format(objInfoBaixaRecCancelar.dtDataCredito, "dd/mm/yy")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixa_Col) = Format(objInfoBaixaRecCancelar.dValorBaixado, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objInfoBaixaRecCancelar.dValorParcela, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_Sequencial_Col) = objInfoBaixaRecCancelar.iSequencial
        GridParcelas.TextMatrix(iLinha, iGrid_MultaCancelar_Col) = Format(objInfoBaixaRecCancelar.dValorMulta, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_JurosCancelar_Col) = Format(objInfoBaixaRecCancelar.dValorJuros, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_DescontoCancelar_Col) = Format(objInfoBaixaRecCancelar.dValorDesconto, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorCancelar_Col) = Format(objInfoBaixaRecCancelar.dValorBaixado, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorTotalParcela_Col) = Format(objInfoBaixaRecCancelar.dValorPago, "Standard")

    Next
        
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridParcelas.iLinhasExistentes = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridParcelas)


    Exit Sub

Erro_Traz_Baixas_Tela:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155435)

    End Select

    Exit Sub

End Sub

Public Sub Traz_MovimentosCCI_Tela()
'Preenche o grid de movtos

Dim iLinha As Integer, lErro As Long, iIndice As Integer
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar
Dim colParcelasMovCCI As Collection
Dim objFilialEmpresa As New AdmFiliais, objCodNome As AdmCodigoNome

On Error GoTo Erro_Traz_MovimentosCCI_Tela

    'Recoloca o Numero de Linhas caso seja maior que o Numero de Linhas Visiveis na Tela
    If (gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count + 1) > NUM_LINHAS_GRID_MOVCCI Then
        GridMovimentosCCI.Rows = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count + 1
    Else
        GridMovimentosCCI.Rows = NUM_LINHAS_GRID_MOVCCI + 1
    End If
       
    'Renicializa
    Call Grid_Inicializa(objGridMovimentosCCI)
    
    For Each objInfoBaixaRecCancelar In gobjBaixaRecCancelar.colInfoBaixaRecCancelar
    
        Set colParcelasMovCCI = New Collection
        
        'Incrementa a variável de controle da linha atual
        iLinha = iLinha + 1
        
        'Passa para a tela os dados da Parcela em questão
        
        If objFilialEmpresa.iCodFilial <> objInfoBaixaRecCancelar.iFilialEmpresa Then
        
            objFilialEmpresa.iCodFilial = objInfoBaixaRecCancelar.iFilialEmpresa
            'Lê a Filial Empresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO Then gError 81825
        
        End If
        
        GridMovimentosCCI.TextMatrix(iLinha, iGrid_FilialEmpresaMov_Col) = objInfoBaixaRecCancelar.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
        
        GridMovimentosCCI.TextMatrix(iLinha, iGrid_DataMovCCI_Col) = objInfoBaixaRecCancelar.dtDataCredito
        
        For iIndice = 0 To ContaCorrenteFiltro.ListCount
        
            If ContaCorrenteFiltro.ItemData(iIndice) = objInfoBaixaRecCancelar.iCodConta Then
                GridMovimentosCCI.TextMatrix(iLinha, iGrid_CtaCorrenteMovCCI_Col) = ContaCorrenteFiltro.List(iIndice)
                Exit For
            End If
        
        Next
        
        For Each objCodNome In gcolCodNomeTiposMovCCI
        
            If objCodNome.iCodigo = objInfoBaixaRecCancelar.iTipoMovCCI Then Exit For
            
        Next
        
        GridMovimentosCCI.TextMatrix(iLinha, iGrid_TipoMovCCI_Col) = objInfoBaixaRecCancelar.iTipoMovCCI & SEPARADOR & objCodNome.sNome
        
        GridMovimentosCCI.TextMatrix(iLinha, iGrid_ValorMovCCI_Col) = Format(objInfoBaixaRecCancelar.dValorMovCCI, "Standard")
        GridMovimentosCCI.TextMatrix(iLinha, iGrid_HistoricoMovCCI_Col) = objInfoBaixaRecCancelar.sHistoricoMovCCI
        
        gobjBaixaRecCancelar.colParcelasMovCCI.Add colParcelasMovCCI
    
    Next
        
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridMovimentosCCI.iLinhasExistentes = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Count

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridMovimentosCCI)

    Exit Sub

Erro_Traz_MovimentosCCI_Tela:

    Select Case gErr

        Case 81825
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155436)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alterações e confirma se deseja salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 57883

    'Limpa a tela
    Call Limpa_Tela_BaixaRecCancelar

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 57883 'Erro tratado na rotina chamda

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155437)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long, iSubTipo As Integer

On Error GoTo Erro_Form_Load

    iFrameAtual = 1
    iFramePrincipalAlterado = 0
    iFrameRecebimentoAtual = MOTIVO_RECEBIMENTO
    iFrameDetalhesAtual = 1
    
    'Instancia obj's referentes a evento
    Set objEventoClientes = New AdmEvento
    
    'Instancia os obj's referentes a grid
    Set objGridParcelas = New AdmGrid
    Set objGridBaixasMovimento = New AdmGrid
    Set objGridMovimentosCCI = New AdmGrid

    lErro = Carrega_ContaCorrentes
    If lErro <> SUCESSO Then gError 81828
    
    lErro = Carrega_TiposMovCCI()
    If lErro <> SUCESSO Then gError 81827
    
    'Seta a data do dia como a data default para cancelamento
    DataCancelamento.PromptInclude = False
    DataCancelamento.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataCancelamento.PromptInclude = True
    
    TotalCancelar.Caption = Format(0, "Fixed")
    ItensSelecionados.Caption = 0
    
    'Executa a Inicialização do grid Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 57884

    'Executa a Inicialização do GridBaixasMovimento
    lErro = Inicializa_GridBaixasMovimento(objGridBaixasMovimento)
    If lErro <> SUCESSO Then gError 94240

    'Executa a Inicialização do objGridMovimentosCCI
    lErro = Inicializa_GridMovimentosCCI(objGridMovimentosCCI)
    If lErro <> SUCESSO Then gError 94241

    'Obtem o código do subtipo que terá seus modelos carregados
    lErro = CF("Obtem_Subtipo_Contabil", "MotivosBaixa", MOTIVO_PAGAMENTO, iSubTipo)
    If lErro <> SUCESSO And lErro <> 79671 Then gError 79673
    
    'Se não encontrou o subtipo => erro
    If lErro = 79671 Then gError 79674
    
    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASARECEBER, iSubTipo)
    If lErro <> SUCESSO Then gError 92851
    
    Call TipoCancBaixaBaixa_Click
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 81827, 81828, 92851, 94240, 94241, 79673

        Case 79674
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155438)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub Inicializa_gobjBaixaRecCancelar()
'Inicializa o obj que serve de "espelho" à tela

Dim lErro As Long, iCodigo As Integer, iSubTipo As Integer

On Error GoTo Erro_Inicializa_gobjBaixaRecCancelar

    Set gobjBaixaRecCancelar = New ClassBaixaRecCancelar

    'Guarda no obj global a informação de que não existem cliente e filial selecionados
    gobjBaixaRecCancelar.lCliente = 0
    gobjBaixaRecCancelar.iFilialCliente = COMBO_INDICE

    'Guarda no obj global a informação do tipo de cancelamento selecionado
    If TipoCancBaixaBaixa.Value = True Then
        gobjBaixaRecCancelar.iTipoCancelamento = TIPOCANCELAMENTO_BAIXABAIXA
    Else
        gobjBaixaRecCancelar.iTipoCancelamento = TIPOCANCELAMENTO_MOVCCI
    End If

    Select Case gobjBaixaRecCancelar.iTipoCancelamento

        Case TIPOCANCELAMENTO_BAIXABAIXA
        
            'Guarda no obj global a informação do tipo de baixa selecionado
            If TipoBaixaRecebimentos.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_RECEBIMENTO
                
            ElseIf TipoBaixaAdiantamentos.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_RECEBTO_ANTECIPADO
                
            ElseIf TipoBaixaDebitos.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_DEBITO_CLIENTE
                
            ElseIf TipoBaixaPerdas.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_PERDA
                
            ElseIf TipoBaixaCartao.Value = True Then
                gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_CARTAO_DEBITO_CREDITO
            End If

            iCodigo = gobjBaixaRecCancelar.iTipoBaixas
    
            lErro = CF("Obtem_Subtipo_Contabil", "MotivosBaixa", iCodigo, iSubTipo)
            If lErro <> SUCESSO And lErro <> 79671 Then gError 79664
            
            'Se não encontrou o subtipo => erro
            If lErro = 79671 Then gError 79678

        Case TIPOCANCELAMENTO_MOVCCI
        
            iSubTipo = 9999
        
    End Select
    
    'Processa alteração do Subtipo para que os novos modelos sejam carregados
    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(iSubTipo)
    If lErro <> SUCESSO Then gError 79662
        
    'Guarda no obj global a informação que não existe uma conta específica selecionada
    gobjBaixaRecCancelar.iCtaCorrenteTodas = CTACORRENTE_TODAS
    gobjBaixaRecCancelar.iCtaCorrenteFiltro = COMBO_INDICE

    'Guarda no obj global a informação de que não foram informadas datas de baixa nem de vencimento
    gobjBaixaRecCancelar.dtDataBaixaInicial = StrParaDate(BaixaInic.Text) 'DATA_NULA
    gobjBaixaRecCancelar.dtDataBaixaFinal = StrParaDate(BaixaFim.Text) 'DATA_NULA
    gobjBaixaRecCancelar.dtDataVenctoInicial = DATA_NULA
    gobjBaixaRecCancelar.dtDataVenctoFinal = DATA_NULA

    'Guarda no obj global a data da baixa
    gobjBaixaRecCancelar.dtDataCancelamento = gdtDataHoje
    
    Exit Sub

Erro_Inicializa_gobjBaixaRecCancelar:

    Select Case gErr

        Case 79662, 79664

        Case 79678
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155439)

    End Select

    Exit Sub

End Sub

Public Sub TabDetalhes_Click()

Dim iIndice As Integer

    'Se FrameDetalhes atual não corresponde ao Tab clicado
    If TabDetalhes.SelectedItem.Index <> iFrameDetalhesAtual Then

        If TabStrip_PodeTrocarTab(iFrameDetalhesAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Torna Frame selecionado visível
        FrameDetalhamentoBaixa(TabDetalhes.SelectedItem.Index).Visible = True

        'Torna Frame atual invisível
        FrameDetalhamentoBaixa(iFrameDetalhesAtual).Visible = False

        'Armazena novo valor de iFrameDetalhesAtual
        iFrameDetalhesAtual = TabDetalhes.SelectedItem.Index

        'Se Frame selecionado foi o de Títulos
        Select Case iFrameDetalhesAtual
        
            Case TAB_Cancelamento
    
                '??? Criar IDH Parent.HelpContextID = IDH_CANCELAR_BAIXA_CONTAS_RECEBER_TITULOS
    
            Case TAB_Detalhamento
    
                '??? Criar IDH Parent.HelpContextID = IDH_CANCELAR_BAIXA_CONTAS_RECEBER_TITULOS
        
        End Select

    End If

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera campos globais
    Set objEventoClientes = Nothing
    Set objGridParcelas = Nothing

    Set objGridBaixasMovimento = Nothing
    Set objGridMovimentosCCI = Nothing
    Set gobjBaixaRecCancelar = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing

    Set gcolCodNomeTiposMovCCI = Nothing

End Sub


Public Sub Cliente_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iClienteAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

    Call Cliente_Preenche

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 0 Then Exit Sub

    'Se Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código ou CPF ou CGC)
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 57885

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 57886

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

        gobjBaixaRecCancelar.lCliente = objCliente.lCodigo
        gobjBaixaRecCancelar.iFilialCliente = iCodFilial

    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

        gobjBaixaRecCancelar.lCliente = 0
        gobjBaixaRecCancelar.iFilialCliente = 0

    End If

    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 57885, 57886 'Erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155440)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objFilialCliente As New ClassFilialCliente
Dim iCodigo As Integer
Dim sNomeRed As String

On Error GoTo Erro_Filial_Validate

    gobjBaixaRecCancelar.iFilialCliente = COMBO_INDICE
    
    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then
        gobjBaixaRecCancelar.iFilialCliente = Filial.ItemData(Filial.ListIndex)
        Exit Sub
    End If

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 57887

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Cliente
        If Len(Trim(Cliente.Text)) = 0 Then gError 57888

        'Lê o Cliente que está na tela
        sNomeRed = Trim(Cliente.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Cliente e Código da Filial
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 57889

        'Se não existe a Filial
        If lErro = 17660 Then gError 57890

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 57891

    gobjBaixaRecCancelar.iFilialCliente = Codigo_Extrai(Filial.Text)
    
    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 57887, 57889

        Case 57888
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 57890
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE1", Filial.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
                'Segura o foco
            End If

        Case 57891
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155441)

    End Select

    Exit Sub

End Sub

Public Sub BaixaInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BaixaInic_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(BaixaInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub BaixaInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_BaixaInic_Validate

    'Se a data BaixaInic está preenchida
    If Len(BaixaInic.ClipText) > 0 Then

        'Verifica se a data BaixaInic é válida
        lErro = Data_Critica(BaixaInic.Text)
        If lErro <> SUCESSO Then gError 94818
        
        gobjBaixaRecCancelar.dtDataBaixaInicial = StrParaDate(BaixaInic.Text)
    
    Else
    
        gobjBaixaRecCancelar.dtDataBaixaInicial = DATA_NULA

    End If

    Exit Sub

Erro_BaixaInic_Validate:

    Cancel = True

    Select Case gErr

        Case 94818

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155442)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaInic_DownClick

    'Diminui a data Inic em 1 dia
    lErro = Data_Up_Down_Click(BaixaInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 57893

    Exit Sub

Erro_UpDownBaixaInic_DownClick:

    Select Case Err

        Case 57893

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155443)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaInic_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(BaixaInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 57894

    Exit Sub

Erro_UpDownBaixaInic_UpClick:

    Select Case Err

        Case 57894

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155444)

    End Select

    Exit Sub

End Sub

Public Sub BaixaFim_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub BaixaFim_GotFocus()

Dim iFrameAux As Integer

    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(BaixaFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub BaixaFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_BaixaFim_Validate

    'Se a data EmissaoFim está preenchida
    If Len(BaixaFim.ClipText) > 0 Then

        'Verifica se a data EmissaoFim é válida
        lErro = Data_Critica(BaixaFim.Text)
        If lErro <> SUCESSO Then gError 94819
        
        gobjBaixaRecCancelar.dtDataBaixaFinal = StrParaDate(BaixaFim.Text)
    
    Else
    
        gobjBaixaRecCancelar.dtDataBaixaFinal = DATA_NULA
    
    End If

    Exit Sub

Erro_BaixaFim_Validate:

    Cancel = True

    Select Case gErr

        Case 94819

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155445)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaFim_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(BaixaFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 57897

    Exit Sub

Erro_UpDownBaixaFim_DownClick:

    Select Case Err

        Case 57897 'Erro tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155446)

    End Select

    Exit Sub

End Sub

Public Sub UpDownBaixaFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownBaixaFim_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(BaixaFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 57898

    Exit Sub

Erro_UpDownBaixaFim_UpClick:

    Select Case Err

        Case 57898 'Erro tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155447)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(VencInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 57899

    Exit Sub

Erro_UpDownVencInic_DownClick:

    Select Case Err

        Case 57899 'Erro tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155448)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(VencInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 57900

    Exit Sub

Erro_UpDownVencInic_UpClick:

    Select Case Err

        Case 57900 'Erro tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155449)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 57901

    Exit Sub

Erro_UpDownVencFim_DownClick:

    Select Case Err

        Case 57901 'Erro tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155450)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 57902

    Exit Sub

Erro_UpDownVencFim_UpClick:

    Select Case Err

        Case 57902 'Erro tratado na rotina chamada

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155451)

    End Select

    Exit Sub

End Sub

Public Sub TituloInic_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TituloInic_Validate

    'Se o campo TituloInic está preenchido
    If Len(Trim(TituloInic.Text)) > 0 Then

        'Critica o valor passado
        lErro = Valor_Positivo_Critica(TituloInic.Text)
        If lErro <> SUCESSO Then gError 94821
        
        gobjBaixaRecCancelar.lTituloInicial = StrParaLong(TituloInic.Text)
    
    Else
    
        gobjBaixaRecCancelar.lTituloInicial = 0
        
    End If

    Exit Sub

Erro_TituloInic_Validate:

    Cancel = True


    Select Case gErr

        Case 94821

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155452)

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
        If lErro <> SUCESSO Then gError 94820

        gobjBaixaRecCancelar.dtDataVenctoInicial = StrParaDate(VencInic.Text)
    Else
        gobjBaixaRecCancelar.dtDataVenctoInicial = DATA_NULA
    End If

    Exit Sub

Erro_VencInic_Validate:

    Cancel = True

    Select Case gErr

        Case 94820

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155453)

    End Select

    Exit Sub

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
        If lErro <> SUCESSO Then Error 57905

        'Se a data vencInic está preenchida
        If Len(VencInic.ClipText) > 0 Then

            'Verifica se a VencFim é maior ou igual a VencInic
            If CDate(VencFim.Text) < CDate(VencInic.Text) Then Error 57906

        End If

    End If

    Exit Sub

Erro_VencFim_Validate:

    Cancel = True


    Select Case Err

        Case 57905 'Erro tratado na rotina chamada

        Case 57906
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155454)

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

    'Se TituloFim está preenchido
    If Len(Trim(TituloFim.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(TituloFim.Text)
        If lErro <> SUCESSO Then gError 94822
        
        gobjBaixaRecCancelar.lTituloFinal = StrParaLong(TituloFim.Text)
    
    Else
    
        gobjBaixaRecCancelar.lTituloFinal = StrParaLong(TituloFim.Text)
        
    End If

    Exit Sub

Erro_TituloFim_Validate:

    Cancel = True

    Select Case gErr

        Case 94822

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155455)

    End Select

    Exit Sub

End Sub

Public Sub Recebimento_Click(Index As Integer)

Dim lErro As Long, iCodigo As Integer, iSubTipo As Integer

On Error GoTo Erro_Recebimento_Click

    'Se o Recebimento clicado é diferente do atual
    If iFrameRecebimentoAtual <> Index Then

        FrameRecebimento(iFrameRecebimentoAtual).Visible = False
        FrameRecebimento(Index).Visible = True
        iFrameRecebimentoAtual = Index

        'o indice é o codigo do motivo de baixa
        iCodigo = Index

        'Lê o código do subtipo
        lErro = CF("Obtem_Subtipo_Contabil", "MotivosBaixa", iCodigo, iSubTipo)
        If lErro <> SUCESSO And lErro <> 79671 Then gError 79664
        
        'Se não encontrou o subtipo => erro
        If lErro = 79671 Then gError 79678
    
        'Processa alteração do Subtipo para que os novos modelos sejam carregados
        lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(iSubTipo)
        If lErro <> SUCESSO Then gError 79662

    End If

    Exit Sub

Erro_Recebimento_Click:

    Select Case gErr
    
        Case 79662, 79664
        
        Case 79666, 79678
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO", gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155456)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub BotaoGravar_Click()
'Pega dados da tela e executa o cancelamento da Baixa

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Grava o cancelamento da Baixa
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 57934

    'Limpa a tela
    Call Limpa_Tela_BaixaRecCancelar

    iFramePrincipalAlterado = REGISTRO_ALTERADO 'para garantir que seja pesquisado o bd apos uma gravacao, qdo trocar de tab
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 57934

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 155457)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iSequencial As Integer
Dim iMotivo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If gobjBaixaRecCancelar.dtDataCancelamento = DATA_NULA Then gError 81856
       
    If StrParaInt(ItensSelecionados) = 0 Then gError 81857
    
    Select Case gobjBaixaRecCancelar.iTipoCancelamento

        Case TIPOCANCELAMENTO_BAIXABAIXA
        
            lErro = Valida_Cancelamento_BaixaBaixa
            If lErro <> SUCESSO Then gError 81849
    
            lErro = Move_Tela_Memoria_BaixaBaixa()
            If lErro <> SUCESSO Then gError 81852

        Case TIPOCANCELAMENTO_MOVCCI
        
            lErro = Valida_Cancelamento_MocvCCI
            If lErro <> SUCESSO Then gError 81850
    
            lErro = Move_Tela_Memoria_MocvCCI()
            If lErro <> SUCESSO Then gError 81851

    End Select
    
    lErro = Critica_Data_Cancelamento
    If lErro <> SUCESSO Then gError 196998
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataCancelamento.Text))
    If lErro <> SUCESSO Then gError 92053
    
    lErro = CF("BaixaRecCancelar_Grava", gobjBaixaRecCancelar, objContabil)
    If lErro <> SUCESSO Then gError 57937

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case 81856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CANC_VAZIA", gErr)
        
        Case 81857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 57935
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAREC_INEXISTENTE", gErr)

        Case 57936
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TECLAR_BOTAO_TRAZER", gErr)

        Case 57937, 81850, 81851, 81852, 92053, 81849, 196998 'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155458)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_BaixaRecCancelar()

    'Limpa a tela
    Call Limpa_Tela(Me)
    Filial.ListIndex = -1

    Call Limpa_Tela_Tab_Baixas

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    CtaCorrenteTodas.Value = True
    
    Call Inicializa_gobjBaixaRecCancelar

    iAlterado = 0

End Sub

Private Sub Limpa_Tela_Tab_Baixas()
'Limpa o frame de Baixas

    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridMovimentosCCI)
    Call Grid_Limpa(objGridBaixasMovimento)
    
    'Frame Dados da Baixa
    DataBaixa.Caption = ""
    Desconto.Caption = ""
    ValorPago.Caption = ""
    Multa.Caption = ""
    ValorBaixado.Caption = ""
    Juros.Caption = ""

    'Débito
    DataEmissaoCred.Caption = ""
    NumTitulo.Caption = ""
    SaldoDebito.Caption = ""
    SiglaDocumentoCR.Caption = ""
    ValorDebito.Caption = ""

    'Adiantamento de Cliente
    DataMovimento.Caption = ""
    MeioPagtoDescricao.Caption = ""
    CCIntNomeReduzido.Caption = ""
    ValorPA.Caption = ""

    'Dados de Recebimento
    ContaCorrente.Caption = ""
    Valor.Caption = ""
    Historico.Caption = ""

    'Perda
    Historico.Caption = ""
    
    'Cartao
    ClienteTitCartao.Caption = ""
    FilialTitCartao.Caption = ""
    TipoTitCartao.Caption = ""
    NumTitCartao.Caption = ""
    DataEmissaoTitCartao.Caption = ""
    
    'Cancelamento
    DataCancelamento.PromptInclude = False
    DataCancelamento.Text = Format(gdtDataHoje, "dd/mm/yy")
    DataCancelamento.PromptInclude = True

    HistoricoCancelamento.Text = ""
    ItensSelecionados.Caption = 0
    TotalCancelar.Caption = Format(0, "Fixed")

End Sub

Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 92852

        Select Case objGridInt.objGrid.Name

            Case GridParcelas.Name
                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then gError 94269

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 57983

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 57983, 94269
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 92852

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155459)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)
'Observacao: esta implementacao atípica de Rotina_Grid_Enable se deve à necessidade de alterar o conteudo da tela em funcao da linha corrente no grid

On Error GoTo Erro_Rotina_Grid_Enable
    
    If objControl.Container.Name = FrameBaixaBaixa.Name Then

        Call Traz_DetalhesBaixa_Tela

    ElseIf objControl.Container.Name = FrameMovCCI.Name Then

        'Call Traz_Parcelas_MovCCI_Tela
            
    End If

    Exit Sub
    
Erro_Rotina_Grid_Enable:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155460)
            
    End Select
    
    Exit Sub
    
End Sub

Private Function Traz_Parcelas_MovCCI_Tela() As Long
'Preenche o grid de baixas associadas ao movimento da linha corrente do grid de movtos
'Se necessario faz a leitura no bd
'Obs.: esta funcao pode se chamar recursivamente

Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar
Dim colParcelasMovCCI As Collection
Dim lErro As Long
Dim iLinha As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_Parcelas_MovCCI_Tela

    Set colParcelasMovCCI = gobjBaixaRecCancelar.colParcelasMovCCI(GridMovimentosCCI.Row)
    
    If colParcelasMovCCI.Count > 0 Then

        'Recoloca o Numero de Linhas caso seja maior que o Numero de Linhas Visiveis na Tela
        If (colParcelasMovCCI.Count + 1) > NUM_LINHAS_GRID_BAIXASMOV Then
            GridBaixasMovimento.Rows = colParcelasMovCCI.Count + 1
        Else
            GridBaixasMovimento.Rows = NUM_LINHAS_GRID_BAIXASMOV + 1
        End If
           
        'Renicializa
        Call Grid_Inicializa(objGridBaixasMovimento)
    
        Call Grid_Limpa(objGridBaixasMovimento)
    
        For Each objInfoBaixaRecCancelar In colParcelasMovCCI
        
            iLinha = iLinha + 1
            
            If objFilialEmpresa.iCodFilial <> objInfoBaixaRecCancelar.iFilialEmpresa Then
            
                objFilialEmpresa.iCodFilial = objInfoBaixaRecCancelar.iFilialEmpresa
                'Lê a Filial Empresa
                lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                If lErro <> SUCESSO Then gError 81826
            
            End If
        
            GridBaixasMovimento.TextMatrix(iLinha, iGrid_FilialEmpresaBaixada_Col) = objInfoBaixaRecCancelar.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
            GridBaixasMovimento.TextMatrix(iLinha, iGrid_TipoBaixada_Col) = objInfoBaixaRecCancelar.sSiglaDocumento
            GridBaixasMovimento.TextMatrix(iLinha, iGrid_NumeroBaixada_Col) = objInfoBaixaRecCancelar.lNumDocumento
            GridBaixasMovimento.TextMatrix(iLinha, iGrid_ParcelaBaixada_Col) = objInfoBaixaRecCancelar.iNumParcela
            GridBaixasMovimento.TextMatrix(iLinha, iGrid_ValorParcelaBaixada_Col) = Format(objInfoBaixaRecCancelar.dValorParcela, "Standard")
            GridBaixasMovimento.TextMatrix(iLinha, iGrid_ValorBaixada_Col) = Format(objInfoBaixaRecCancelar.dValorBaixado, "Standard")
            GridBaixasMovimento.TextMatrix(iLinha, iGrid_SequencialBaixada_Col) = objInfoBaixaRecCancelar.iSequencial
        
        Next
        
        objGridBaixasMovimento.iLinhasExistentes = iLinha

    Else
        
        lErro = Movcci_ObtemBaixas(colParcelasMovCCI, GridMovimentosCCI.Row)
        If lErro <> SUCESSO Then gError 94836
        
        lErro = Traz_Parcelas_MovCCI_Tela()
        If lErro <> SUCESSO Then gError 94837
    
    End If

    Traz_Parcelas_MovCCI_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Parcelas_MovCCI_Tela:

    Select Case gErr
        
        Case 94836, 94837, 81826
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155461)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Traz_DetalhesBaixa_Tela()

Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar
Dim objBaixaReceber As New ClassBaixaReceber
Dim lErro As Long

On Error GoTo Erro_Traz_DetalhesBaixa_Tela
    
    'If objGridParcelas.iLinhasExistentes > 0 And GridParcelas.Row <= objGridParcelas.iLinhasExistentes Then
    
        Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar(GridParcelas.Row)
        
        DataBaixa.Caption = Format(objInfoBaixaRecCancelar.dtDataBaixa, "dd/mm/yy")
        ValorPago.Caption = Format(objInfoBaixaRecCancelar.dValorPago, "Standard")
        ValorBaixado.Caption = Format(objInfoBaixaRecCancelar.dValorBaixado, "Standard")
        Desconto.Caption = Format(objInfoBaixaRecCancelar.dValorDesconto, "Standard")
        Multa.Caption = Format(objInfoBaixaRecCancelar.dValorMulta, "Standard")
        Juros.Caption = Format(objInfoBaixaRecCancelar.dValorJuros, "Standard")
        
        Select Case objInfoBaixaRecCancelar.iMotivoBaixa
        
            Case MOTIVO_RECEBIMENTO
            
                ContaCorrente.Caption = objInfoBaixaRecCancelar.iCodConta & "-" & objInfoBaixaRecCancelar.sContaCorrente
                Valor.Caption = Format(objInfoBaixaRecCancelar.dValorMovCCI, "Standard")
                Historico.Caption = objInfoBaixaRecCancelar.sHistoricoBaixa
            
            Case MOTIVO_RECEBTO_ANTECIPADO
            
                DataMovimento.Caption = Format(objInfoBaixaRecCancelar.dtDataBaixa, "dd/mm/yy")
                CCIntNomeReduzido.Caption = objInfoBaixaRecCancelar.iCodConta & "-" & objInfoBaixaRecCancelar.sContaCorrente
                
                If objInfoBaixaRecCancelar.iTipoMeioPagto = DINHEIRO Then
                    MeioPagtoDescricao.Caption = "Dinheiro"
                ElseIf objInfoBaixaRecCancelar.iTipoMeioPagto = Cheque Then
                    MeioPagtoDescricao.Caption = "Cheque"
                ElseIf objInfoBaixaRecCancelar.iTipoMeioPagto = BORDERO Then
                    MeioPagtoDescricao.Caption = "Borderô"
                End If
                
                ValorPA.Caption = Format(objInfoBaixaRecCancelar.dValorMovCCI, "Standard")
            
            Case MOTIVO_DEBITO_CLIENTE
            
                DataEmissaoCred.Caption = Format(objInfoBaixaRecCancelar.dtDataEmissaoDebito, "dd/mm/yy")
                NumTitulo.Caption = objInfoBaixaRecCancelar.lNumDocumento
                SiglaDocumentoCR.Caption = objInfoBaixaRecCancelar.sSiglaDocumento
                ValorDebito.Caption = Format(objInfoBaixaRecCancelar.dValorDebito, "Standard")
                SaldoDebito.Caption = Format(objInfoBaixaRecCancelar.dSaldoDebito, "Standard")
            
            Case MOTIVO_PERDA
            
                HistoricoPerda.Caption = objInfoBaixaRecCancelar.sHistoricoBaixa
        
            Case MOTIVO_CARTAO_DEBITO_CREDITO
            
                objBaixaReceber.lNumIntDoc = objInfoBaixaRecCancelar.lNumIntDebRecCli
            
                lErro = Traz_Dados_Titulo_Cartao(objBaixaReceber)
                If lErro <> SUCESSO Then gError 183056
        
        
        End Select
    
    'End If
    
    Exit Sub
    
Erro_Traz_DetalhesBaixa_Tela:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155462)
            
    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CANCELAR_BAIXA_CONTAS_RECEBER_TITULOS
    Set Form_Load_Ocx = Me
    Caption = "Cancelar Baixa em Contas a Receber"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BaixaRecCancelar"

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

        If Me.ActiveControl Is Cliente Then
            Call LabelCliente_Click
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

Private Function Inicializa_GridBaixasMovimento(objGridInt As AdmGrid) As Long
'Executa a Inicialização do GridBaixasMovimento

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Numero")
    objGridInt.colColuna.Add ("Filial Emp")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Sequencial")
    objGridInt.colColuna.Add ("Valor Baixado")

    'campos de edição do grid
    objGridInt.colCampo.Add (NumeroBaixada.Name)
    objGridInt.colCampo.Add (FilialEmpresaBaixada.Name)
    objGridInt.colCampo.Add (TipoBaixada.Name)
    objGridInt.colCampo.Add (ParcelaBaixada.Name)
    objGridInt.colCampo.Add (ValorParcelaBaixada.Name)
    objGridInt.colCampo.Add (SequencialBaixada.Name)
    objGridInt.colCampo.Add (ValorBaixada.Name)

    iGrid_NumeroBaixada_Col = 1
    iGrid_FilialEmpresaBaixada_Col = 2
    iGrid_TipoBaixada_Col = 3
    iGrid_ParcelaBaixada_Col = 4
    iGrid_ValorParcelaBaixada_Col = 5
    iGrid_SequencialBaixada_Col = 6
    iGrid_ValorBaixada_Col = 7

    objGridInt.objGrid = GridBaixasMovimento

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_LINHAS_GRID_BAIXASMOV + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6

    'largura da primeira coluna
    objGridInt.objGrid.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridBaixasMovimento = SUCESSO

    Exit Function

End Function

Private Function Inicializa_GridMovimentosCCI(objGridInt As AdmGrid) As Long
'Executa a Inicialização do GridMovimentosCCI

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Cancelar")
    objGridInt.colColuna.Add ("Filial Emp")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("C/C.")
    objGridInt.colColuna.Add ("Origem")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Histórico")

    'campos de edição do grid
    objGridInt.colCampo.Add (CancelarMov.Name)
    objGridInt.colCampo.Add (FilialEmpresaMov.Name)
    objGridInt.colCampo.Add (DataMovCCI.Name)
    objGridInt.colCampo.Add (CtaCorrenteMovCCI.Name)
    objGridInt.colCampo.Add (TipoMovCCI.Name)
    objGridInt.colCampo.Add (ValorMovCCI.Name)
    objGridInt.colCampo.Add (HistoricoMovCCI.Name)
    
    iGrid_CancelarMov_Col = 1
    iGrid_FilialEmpresaMov_Col = 2
    iGrid_DataMovCCI_Col = 3
    iGrid_CtaCorrenteMovCCI_Col = 4
    iGrid_TipoMovCCI_Col = 5
    iGrid_ValorMovCCI_Col = 6
    iGrid_HistoricoMovCCI_Col = 7
    
    objGridInt.objGrid = GridMovimentosCCI

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_LINHAS_GRID_MOVCCI + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 6

    'largura da primeira coluna
    GridMovimentosCCI.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR

    'Indica para executar a Rotina Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridMovimentosCCI = SUCESSO

    Exit Function

End Function

Public Sub ParcelaSelecionada_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ParcelaSelecionada_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ParcelaSelecionada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ParcelaSelecionada
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ParcelaSelecionada_Click()

    'Se o usuário selecionou a parcela
    If GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ParcelaSelecionada_Col) = True Then
        
        'Guarda no obj a informação de que a parcela está selecionada
        gobjBaixaRecCancelar.colInfoBaixaRecCancelar(GridParcelas.Row).iItemSelecionado = MARCADO
    
    'Senão
    Else
        
        'Guarda no obj global a informação de que a parcela não está selecionada
        gobjBaixaRecCancelar.colInfoBaixaRecCancelar(GridParcelas.Row).iItemSelecionado = DESMARCADO
    
    End If
    
    'Calcula o valor total de adiantamentos a serem baixados
    Call Calcula_Total_Cancelar(TIPOCANCELAMENTO_BAIXABAIXA)
    
    'Atualiza os CheckBox na tela
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    Call Trata_Data_Cancelamento

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Calcula_Total_Cancelar(iTipoCancelamento As Integer)
'Atualiza na tela o valor total cancelado e o numero de itens selecionados
'Atualiza no objeto global o valor total cancelado

Dim iIndice As Integer, iItensSelecionados As Integer
Dim dTotalSelecionado As Double
Dim objGrid As AdmGrid
Dim ControleGrid As Object
Dim iColunaSeleciona As Integer
Dim iColunaValor As Integer

On Error GoTo Erro_Calcula_Total_Cancelar

    dTotalSelecionado = 0
    
    'Verifica o tipo de baixa passado como parâmetro
    Select Case iTipoCancelamento
    
        'Se for cancelamento baixa a baixa => o grid utilizado para o cálculo deve ser o de baixas (parcelas)
        Case TIPOCANCELAMENTO_BAIXABAIXA
        
            'Instancia o obj e controle com os seus correspondentes para o grid Adiantamentos
            Set objGrid = objGridParcelas
            Set ControleGrid = GridParcelas
            
            'Indica a coluna que deverá ser testada para verificar se a linha está selecionada
            iColunaSeleciona = iGrid_ParcelaSelecionada_Col
        
            'Indica a coluna que contém o valor que será acumulado
            iColunaValor = iGrid_ValorTotalParcela_Col
        
        'Se for baixa de movtos => o grid utilizado para o cálculo deve ser o de movtos
        Case TIPOCANCELAMENTO_MOVCCI
            
            '??? trata como se fosse possivel selecionar mais de uma linha
            
            'Instancia o obj e controle com os seus correspondentes para o grid Débitos
            Set objGrid = objGridMovimentosCCI
            Set ControleGrid = GridMovimentosCCI
        
            'Indica a coluna que deverá ser testada para verificar se a linha está selecionada
            iColunaSeleciona = iGrid_CancelarMov_Col
        
            'Indica a coluna que contém o valor que será acumulado
            iColunaValor = iGrid_ValorMovCCI_Col
        
    End Select
    
    'Para cada linha do grid
    For iIndice = 1 To objGrid.iLinhasExistentes
        
        'Se a linha está selecionada
        If StrParaInt(ControleGrid.TextMatrix(iIndice, iColunaSeleciona)) <> 0 Then
            
            'Adiciona o valor da linha ao valor total das linhas selecionadas
            dTotalSelecionado = dTotalSelecionado + StrParaDbl(ControleGrid.TextMatrix(iIndice, iColunaValor))
            
            iItensSelecionados = iItensSelecionados + 1
            
        End If
        
    Next
    
    'Exibe na tela o valor encontrado acima
    TotalCancelar.Caption = Format(dTotalSelecionado, "Standard")
    ItensSelecionados.Caption = CStr(iItensSelecionados)
    
    'Guarda no obj global o valor encontrado
    gobjBaixaRecCancelar.dValorBaixasCancelar = dTotalSelecionado
    
    Exit Sub
    
Erro_Calcula_Total_Cancelar:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155463)

    End Select
    
    Exit Sub
    
End Sub

Public Property Get ParcelaSelecionada() As Object
     Set ParcelaSelecionada = objUserControl.Controls("ParcelaSelecionada")
End Property

Public Sub TipoCancBaixaBaixa_Click()

    FrameDataBaixa.top = 2280
    FrameFiltrosBaixaMovInt.Caption = "Filtros - Baixa a Baixa"
    FrameFiltrosBaixaBaixa.Visible = True
    FrameFiltrosMovInt.Visible = False
    FrameCliente.Visible = True
    FrameTipoBaixas.Visible = True
    
    'Torna o Frame com o grid de Baixas visível
    FrameBaixaBaixa.Visible = True
    
    'Torna o Frame com o grid de movimentos invisível
    FrameMovCCI.Visible = False
            
    FrameDadosBaixa.Visible = True
    FrameBaixasMovCCI.Visible = False
            
    Call Limpa_Tela_BaixaRecCancelar
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoCancMovimentoIntegral_Click()

    FrameDataBaixa.top = 1080
    FrameFiltrosBaixaMovInt.Caption = "Filtros - Movimento Integral"
    FrameFiltrosBaixaBaixa.Visible = False
    FrameFiltrosMovInt.Visible = True
    FrameCliente.Visible = False
    FrameTipoBaixas.Visible = False
    
    'Torna o Frame com o grid de movimentos visível
    FrameMovCCI.Visible = True
    
    'Torna o Frame com o grid de Baixas invisível
    FrameBaixaBaixa.Visible = False
    
    FrameDadosBaixa.Visible = False
    FrameBaixasMovCCI.Visible = True
            
    Call Inicializa_gobjBaixaRecCancelar

    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Property Get TipoCancBaixaBaixa() As Object
     Set TipoCancBaixaBaixa = objUserControl.Controls("TipoCancBaixaBaixa")
End Property

Public Property Get TipoCancMovimentoIntegral() As Object
     Set TipoCancMovimentoIntegral = objUserControl.Controls("TipoCancMovimentoIntegral")
End Property

Public Property Get FrameFiltrosBaixaBaixa() As Object
     Set FrameFiltrosBaixaBaixa = objUserControl.Controls("FrameFiltrosBaixaBaixa")
End Property

Public Property Get FrameFiltrosMovInt() As Object
     Set FrameFiltrosMovInt = objUserControl.Controls("FrameFiltrosMovInt")
End Property

Public Property Get FrameFiltrosBaixaMovInt() As Object
     Set FrameFiltrosBaixaMovInt = objUserControl.Controls("FrameFiltrosBaixaMovInt")
End Property

Public Property Get FrameCliente() As Object
     Set FrameCliente = objUserControl.Controls("FrameCliente")
End Property

Public Property Get FrameTipoBaixas() As Object
     Set FrameTipoBaixas = objUserControl.Controls("FrameTipoBaixas")
End Property

Public Property Get FrameDataBaixa() As Object
     Set FrameDataBaixa = objUserControl.Controls("FrameDataBaixa")
End Property

Public Sub TipoBaixaRecebimentos_Click()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_RECEBIMENTO
    Call Recebimento_Click(MOTIVO_RECEBIMENTO)
End Sub

Public Sub TipoBaixaAdiantamentos_Click()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_RECEBTO_ANTECIPADO
    Call Recebimento_Click(MOTIVO_RECEBTO_ANTECIPADO)
End Sub

Public Sub TipoBaixaDebitos_Click()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_DEBITO_CLIENTE
    Call Recebimento_Click(MOTIVO_DEBITO_CLIENTE)
End Sub

Public Sub TipoBaixaPerdas_Click()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_PERDA
    Call Recebimento_Click(MOTIVO_PERDA)
End Sub

Public Sub CtaCorrenteTodas_Click()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    gobjBaixaRecCancelar.iCtaCorrenteTodas = CTACORRENTE_TODAS
    ContaCorrenteFiltro.ListIndex = -1
    ContaCorrenteFiltro.Enabled = False
    
End Sub

Public Sub CtaCorrenteApenas_Click()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    gobjBaixaRecCancelar.iCtaCorrenteFiltro = CTACORRENTE_APENAS
    ContaCorrenteFiltro.Enabled = True
    
End Sub

Public Sub ContaCorrenteFiltro_Change()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ContaCorrenteFiltro_Click()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    If ContaCorrenteFiltro.ListIndex <> -1 Then
        gobjBaixaRecCancelar.iCtaCorrenteFiltro = ContaCorrenteFiltro.ItemData(ContaCorrenteFiltro.ListIndex)
    End If

End Sub

Public Sub ContaCorrenteFiltro_Validate(Cancel As Boolean)

Dim sContaCorrenteFiltro As String
Dim iCodigo As Integer
Dim objContaCorrente As New ClassContasCorrentesInternas
Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_ContaCorrenteFiltro_Validate

    'Indica que não existe conta corrente selecionada (só existirá uma conta selecionada depois da execução do validate)
    gobjBaixaRecCancelar.iCtaCorrenteFiltro = COMBO_INDICE

    'Se a ComboBox ContaCorrenteFiltro não está preenchida => sai da função
    If Len(Trim(ContaCorrenteFiltro.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox ContaCorrenteFiltro
    If ContaCorrenteFiltro.Text = ContaCorrenteFiltro.List(ContaCorrenteFiltro.ListIndex) Then
        gobjBaixaRecCancelar.iCtaCorrenteFiltro = ContaCorrenteFiltro.ItemData(ContaCorrenteFiltro.ListIndex)
        Exit Sub
    End If

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(ContaCorrenteFiltro, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 94823

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Lê os dados da conta corrente a partir do código passado como parâmetro
        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrente)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 94824

        'Se não encontrou a conta => o número de erro recebe tratamento para perguntar se deseja criar a conta
        If lErro = 11807 Then gError 94825

        'Coloca os dados da conta na combo
        ContaCorrenteFiltro.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then

        'Lê os dados da conta corrente a partir do nome red. passado como parâmetro
        objContaCorrente.sNomeReduzido = ContaCorrente.Text
        lErro = CF("ContaCorrenteInt_Le_NomeReduzido", objContaCorrente)
        If lErro <> SUCESSO And lErro <> 15981 Then gError 94826

        'Se não encontrou a conta => Erro
        If lErro = 15981 Then gError 94827
        
        'Coloca os dados da conta na combo
        ContaCorrenteFiltro.Text = CStr(objContaCorrente.iCodigo) & SEPARADOR & objContaCorrente.sNomeReduzido
    
    End If

    'Guarda no obj global a tela o código da conta corrente selecionada
    gobjBaixaRecCancelar.iCtaCorrenteFiltro = Codigo_Extrai(ContaCorrenteFiltro.Text)

    Exit Sub

Erro_ContaCorrenteFiltro_Validate:

    Cancel = True

    Select Case gErr

        Case 94823, 94824, 94826

        Case 94825
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONTACORRENTE", ContaCorrenteFiltro.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("CtaCorrenteInt", objContaCorrente)
            Else
                'Segura o foco
            End If

        Case 94827
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_CADASTRADA", gErr, objContaCorrente.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155464)

    End Select

    Exit Sub

End Sub

Public Property Get ContaCorrenteFiltro() As Object
     Set ContaCorrenteFiltro = objUserControl.Controls("ContaCorrenteFiltro")
End Property

Public Property Get CtaCorrenteApenas() As Object
     Set CtaCorrenteApenas = objUserControl.Controls("CtaCorrenteApenas")
End Property

Public Property Get CtaCorrenteTodas() As Object
     Set CtaCorrenteTodas = objUserControl.Controls("CtaCorrenteTodas")
End Property

Public Property Get TipoBaixaRecebimentos() As Object
     Set TipoBaixaRecebimentos = objUserControl.Controls("TipoBaixaRecebimentos")
End Property

Public Property Get TipoBaixaAdiantamentos() As Object
     Set TipoBaixaAdiantamentos = objUserControl.Controls("TipoBaixaAdiantamentos")
End Property

Public Property Get TipoBaixaDebitos() As Object
     Set TipoBaixaDebitos = objUserControl.Controls("TipoBaixaDebitos")
End Property

Public Property Get TipoBaixaPerdas() As Object
     Set TipoBaixaPerdas = objUserControl.Controls("TipoBaixaPerdas")
End Property

Public Property Get TipoBaixaCartao() As Object
     Set TipoBaixaCartao = objUserControl.Controls("TipoBaixaCartao")
End Property

Private Function Carrega_ContaCorrentes() As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As AdmCodigoNome

On Error GoTo Erro_ContaCorrente_Carrega

    'Lê Codigos, NomesReduzidos de ContasCorrentes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then gError 94239

    'Preeche list de ComboBox
    For Each objCodigoNomeRed In colCodigoNomeRed

        ContaCorrenteFiltro.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        ContaCorrenteFiltro.ItemData(ContaCorrenteFiltro.NewIndex) = objCodigoNomeRed.iCodigo

    Next

    Carrega_ContaCorrentes = SUCESSO

    Exit Function

Erro_ContaCorrente_Carrega:

    Carrega_ContaCorrentes = gErr
    
    Select Case gErr

        Case 94239

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155465)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid)
'Executa a saída de célula do GridParcelas

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    Select Case objGridInt.objGrid.Col

        Case iGrid_ValorTotalParcela_Col
            lErro = Saida_Celula_ValorTotalParcela(objGridInt)
            If lErro <> SUCESSO Then gError 94271

        Case iGrid_MultaCancelar_Col
            lErro = Saida_Celula_MultaCancelar(objGridInt)
            If lErro <> SUCESSO Then gError 94272

        Case iGrid_JurosCancelar_Col
            lErro = Saida_Celula_JurosCancelar(objGridInt)
            If lErro <> SUCESSO Then gError 94273

        Case iGrid_DescontoCancelar_Col
            lErro = Saida_Celula_DescontoCancelar(objGridInt)
            If lErro <> SUCESSO Then gError 94274

        Case iGrid_ValorCancelar_Col
            lErro = Saida_Celula_ValorCancelar(objGridInt)
            If lErro <> SUCESSO Then gError 94275

    End Select

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        Case 94271 To 94275
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155466)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorTotalParcela(objGridInt As AdmGrid)

Dim lErro As Long
Dim dValor As Double
Dim iValorAlterado As Boolean
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar

On Error GoTo Erro_Saida_Celula_ValorTotalParcela

    Set objGridInt.objControle = ValorTotalParcela

    If Len(Trim(ValorTotalParcela.Text)) <> 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorTotalParcela.Text)
        If lErro <> SUCESSO Then gError 94278

        dValor = StrParaDbl(ValorTotalParcela.Text)

        Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Item(GridParcelas.Row)
        
        If dValor > (objInfoBaixaRecCancelar.dValorPago + DELTA_VALORMONETARIO) Then gError 81844
        
        'Coloca Valor no formato Standard da tela
        ValorTotalParcela.Text = Format(dValor, "STANDARD")

    End If

    If ValorTotalParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorTotalParcela_Col) Then
        iValorAlterado = True
    Else
        iValorAlterado = False
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 94277

    If iValorAlterado Then
    
        'acerta total geral a cancelar
        Call Calcula_Total_Cancelar(TIPOCANCELAMENTO_BAIXABAIXA)
    
    End If
    
    Saida_Celula_ValorTotalParcela = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorTotalParcela:

    Saida_Celula_ValorTotalParcela = gErr

    Select Case gErr
'??? 94278
        Case 94276, 94277, 94278
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 81844
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCANC_MAIOR_ORIGINAL", gErr, Format(objInfoBaixaRecCancelar.dValorPago, "standard"))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155467)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_MultaCancelar(objGridInt As AdmGrid)

Dim lErro As Long
Dim dValor As Double
Dim iValorAlterado As Boolean
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar

On Error GoTo Erro_Saida_Celula_MultaCancelar

    Set objGridInt.objControle = MultaCancelar

    If Len(Trim(MultaCancelar.Text)) <> 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(MultaCancelar.Text)
        If lErro <> SUCESSO Then gError 81842
    
        dValor = StrParaDbl(MultaCancelar.Text)
    
        Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Item(GridParcelas.Row)
        
        If dValor > (objInfoBaixaRecCancelar.dValorMulta + DELTA_VALORMONETARIO) Then gError 81848
        
        'Coloca Valor no formato Standard da tela
        MultaCancelar.Text = Format(dValor, "STANDARD")

    End If
    
    If MultaCancelar.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_MultaCancelar_Col) Then
        iValorAlterado = True
    Else
        iValorAlterado = False
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 82203

    If iValorAlterado Then
    
        'recalcula total a cancelar da linha
        Call ValorTotalCancelar_Calcula(GridParcelas.Row)
        
        'acerta total geral a cancelar
        Call Calcula_Total_Cancelar(TIPOCANCELAMENTO_BAIXABAIXA)

    End If
    
    Saida_Celula_MultaCancelar = SUCESSO

    Exit Function

Erro_Saida_Celula_MultaCancelar:

    Saida_Celula_MultaCancelar = gErr

    Select Case gErr

        Case 82202, 82203, 81842
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 81848
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCANC_MAIOR_ORIGINAL", gErr, Format(objInfoBaixaRecCancelar.dValorMulta, "standard"))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155468)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_JurosCancelar(objGridInt As AdmGrid)

Dim lErro As Long
Dim dValor As Double
Dim iValorAlterado As Boolean
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar

On Error GoTo Erro_Saida_Celula_JurosCancelar

    Set objGridInt.objControle = JurosCancelar
    
    If Len(Trim(JurosCancelar.Text)) <> 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(JurosCancelar.Text)
        If lErro <> SUCESSO Then gError 81843
    
        dValor = StrParaDbl(JurosCancelar.Text)
    
        Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Item(GridParcelas.Row)
        
        If dValor > (objInfoBaixaRecCancelar.dValorJuros + DELTA_VALORMONETARIO) Then gError 81847
        
        'Coloca Valor no formato Standard da tela
        JurosCancelar.Text = Format(dValor, "STANDARD")

    End If
    
    If JurosCancelar.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_JurosCancelar_Col) Then
        iValorAlterado = True
    Else
        iValorAlterado = False
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 82203

    If iValorAlterado Then
    
        'recalcula total a cancelar da linha
        Call ValorTotalCancelar_Calcula(GridParcelas.Row)
        
        'acerta total geral a cancelar
        Call Calcula_Total_Cancelar(TIPOCANCELAMENTO_BAIXABAIXA)
    
    End If
    
    Saida_Celula_JurosCancelar = SUCESSO

    Exit Function

Erro_Saida_Celula_JurosCancelar:

    Saida_Celula_JurosCancelar = gErr

    Select Case gErr

        Case 82202, 82203, 81843
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 81847
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCANC_MAIOR_ORIGINAL", gErr, Format(objInfoBaixaRecCancelar.dValorJuros, "standard"))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155469)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DescontoCancelar(objGridInt As AdmGrid)

Dim lErro As Long
Dim dValor As Double
Dim iValorAlterado As Boolean
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar

On Error GoTo Erro_Saida_Celula_DescontoCancelar

    Set objGridInt.objControle = DescontoCancelar
    
    If Len(Trim(DescontoCancelar.Text)) <> 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(DescontoCancelar.Text)
        If lErro <> SUCESSO Then gError 94278
    
        dValor = StrParaDbl(DescontoCancelar.Text)
    
        Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Item(GridParcelas.Row)
        
        If dValor > (objInfoBaixaRecCancelar.dValorDesconto + DELTA_VALORMONETARIO) Then gError 81846
        
        'Coloca Valor no formato Standard da tela
        DescontoCancelar.Text = Format(dValor, "STANDARD")

    End If
    
    If DescontoCancelar.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DescontoCancelar_Col) Then
        iValorAlterado = True
    Else
        iValorAlterado = False
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 82203

    If iValorAlterado Then
    
        'recalcula total a cancelar da linha
        Call ValorTotalCancelar_Calcula(GridParcelas.Row)
        
        'acerta total geral a cancelar
        Call Calcula_Total_Cancelar(TIPOCANCELAMENTO_BAIXABAIXA)
    
    End If
    
    Saida_Celula_DescontoCancelar = SUCESSO

    Exit Function

Erro_Saida_Celula_DescontoCancelar:

    Saida_Celula_DescontoCancelar = gErr

    Select Case gErr

        Case 82202, 82203, 94278
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 81846
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCANC_MAIOR_ORIGINAL", gErr, Format(objInfoBaixaRecCancelar.dValorDesconto, "standard"))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155470)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorCancelar(objGridInt As AdmGrid)

Dim lErro As Long
Dim dValor As Double
Dim iValorAlterado As Boolean
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar

On Error GoTo Erro_Saida_Celula_ValorCancelar

    Set objGridInt.objControle = ValorCancelar

    If Len(Trim(ValorCancelar.Text)) <> 0 Then
    
        'Critica o valor informado
        lErro = Valor_NaoNegativo_Critica(ValorCancelar.Text)
        If lErro <> SUCESSO Then gError 94278

        dValor = StrParaDbl(ValorCancelar.Text)

        Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Item(GridParcelas.Row)
        
        If dValor > (objInfoBaixaRecCancelar.dValorBaixado + DELTA_VALORMONETARIO) Then gError 81845
        
        'Coloca Valor no formato Standard da tela
        ValorCancelar.Text = Format(dValor, "STANDARD")

    End If

    If ValorCancelar.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorCancelar_Col) Then
        iValorAlterado = True
    Else
        iValorAlterado = False
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 94279

    If iValorAlterado Then
    
        'recalcula total a cancelar da linha
        Call ValorTotalCancelar_Calcula(GridParcelas.Row)
        
        'acerta total geral a cancelar
        Call Calcula_Total_Cancelar(TIPOCANCELAMENTO_BAIXABAIXA)
    
    End If
    
    Saida_Celula_ValorCancelar = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorCancelar:

    Saida_Celula_ValorCancelar = gErr

    Select Case gErr

        Case 94278, 94279
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 81845
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORCANC_MAIOR_ORIGINAL", gErr, Format(objInfoBaixaRecCancelar.dValorBaixado, "standard"))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155471)

    End Select

    Exit Function

End Function

Public Sub DataCancelamento_Change()
    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataCancelamento_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataCancelamento_Validate

    'Se a data DataCancelamento está preenchida
    If Len(Trim(DataCancelamento.ClipText)) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(DataCancelamento.Text)
        If lErro <> SUCESSO Then gError 94800

        'Guarda no obj a data De que deve ser passada como parâmetro
        gobjBaixaRecCancelar.dtDataCancelamento = StrParaDate(DataCancelamento.Text)

    Else

        'Guarda no obj a data De que deve ser passada como parâmetro
        gobjBaixaRecCancelar.dtDataCancelamento = DATA_NULA

    End If

    Exit Sub

Erro_DataCancelamento_Validate:

    Cancel = True

    Select Case gErr

        Case 94800

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155472)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataCancelamento_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCancelamento_DownClick

    'Diminui a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataCancelamento, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 94801

    Exit Sub

Erro_UpDownDataCancelamento_DownClick:

    Select Case gErr

        Case 94801

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155473)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataCancelamento_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCancelamento_UpClick

    'Aumenta a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataCancelamento, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 94802

    Exit Sub

Erro_UpDownDataCancelamento_UpClick:

    Select Case gErr

        Case 94802

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155474)

    End Select

    Exit Sub

End Sub

Public Sub HistoricoCancelamento_Change()
    gobjBaixaRecCancelar.sHistoricoCancelamento = HistoricoCancelamento.Text
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Valida_Filtros() As Long

On Error GoTo Erro_Valida_Filtros

    'Se as datas de emissão (De e Até) foram preenchidas => verifica se data Até não é menor do que a data De
    If gobjBaixaRecCancelar.dtDataBaixaInicial <> DATA_NULA And gobjBaixaRecCancelar.dtDataBaixaFinal <> DATA_NULA Then

        'Se a data Até for menor do que a Data De => erro
        If gobjBaixaRecCancelar.dtDataBaixaFinal < gobjBaixaRecCancelar.dtDataBaixaInicial Then gError 94804

    End If

    Select Case gobjBaixaRecCancelar.iTipoCancelamento

        Case TIPOCANCELAMENTO_BAIXABAIXA

            'Se o cliente não foi preenchido => erro
            If gobjBaixaRecCancelar.lCliente = 0 Then gError 94805
        
            'Se a filial do cliente não foi preenchida = > erro
            If gobjBaixaRecCancelar.iFilialCliente = 0 Then gError 94806
            
            'Se os números de títulos(De e Até) foram preenchidos => verifica se o número Até não é menor do que o número De
            If gobjBaixaRecCancelar.lTituloInicial <> 0 And gobjBaixaRecCancelar.lTituloFinal <> 0 Then
    
                'Se a data Até for menor do que a Data De => erro
                If gobjBaixaRecCancelar.lTituloFinal < gobjBaixaRecCancelar.lTituloInicial Then gError 94807
    
            End If
            
            'Se as datas de vencimento (De e Até) foram preenchidas => verifica se data Até não é menor do que a data De
            If gobjBaixaRecCancelar.dtDataVenctoInicial <> DATA_NULA And gobjBaixaRecCancelar.dtDataVenctoFinal <> DATA_NULA Then
        
                'Se a data Até for menor do que a Data De => erro
                If gobjBaixaRecCancelar.dtDataVenctoFinal < gobjBaixaRecCancelar.dtDataVenctoInicial Then gError 94808
        
            End If
            
        Case TIPOCANCELAMENTO_MOVCCI

            'Se o filtro de seleção por apenas uma conta corrente foi selecionado => Verifica se foi selecionada uma conta
            If gobjBaixaRecCancelar.iCtaCorrenteTodas = CTACORRENTE_APENAS Then

                'Se a Conta Corrente que será usada para filtro não foi informado => erro
                If gobjBaixaRecCancelar.iCtaCorrenteFiltro = COMBO_INDICE Then gError 94809

            End If
        
    End Select

    Valida_Filtros = SUCESSO

    Exit Function

Erro_Valida_Filtros:

    Valida_Filtros = gErr

    Select Case gErr

        Case 94804
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_INICIAL_MAIOR", gErr)
        
        Case 94805
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 94806
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_INFORMADA", gErr)

        Case 94807
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_ATE_MENOR_NUMERO_DE", gErr)
        
        Case 94808
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_VENCTO_INICIAL_MAIOR", gErr)

        Case 94809
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTA_NAO_INFORMADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155475)

    End Select

    Exit Function

End Function

Private Function Carrega_TiposMovCCI() As Long

Dim lErro As Long

On Error GoTo Erro_Carrega_TiposMovCCI

    lErro = CF("Cod_Nomes_Le", "TiposMovtoCtaCorrente", "Codigo", "NomeReduzido", STRING_TIPOMOVCCI_NOMERED, gcolCodNomeTiposMovCCI)
    
    Carrega_TiposMovCCI = SUCESSO
     
    Exit Function
    
Erro_Carrega_TiposMovCCI:

    Carrega_TiposMovCCI = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155476)
     
    End Select
     
    Exit Function

End Function

Private Function Valida_Cancelamento_BaixaBaixa() As Long

Dim lErro As Long, iCodConta As Integer, iIndice As Integer
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar

On Error GoTo Erro_Valida_Cancelamento_BaixaBaixa

    If gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_RECEBIMENTO Then

        iCodConta = -1
        
        For Each objInfoBaixaRecCancelar In gobjBaixaRecCancelar.colInfoBaixaRecCancelar

            If objInfoBaixaRecCancelar.iItemSelecionado = MARCADO Then
                        
                If iCodConta = -1 Then
                
                    iCodConta = objInfoBaixaRecCancelar.iCodConta
                    
                Else
                
                    If iCodConta <> objInfoBaixaRecCancelar.iCodConta Then gError 81841
                    
                End If
                
            End If
            
        Next
        
    End If

    'Verifica se alguma linha do GridParcelas foi selecionada
    For iIndice = 1 To objGridParcelas.iLinhasExistentes - 1 'Alterado por Leo em 08/12/01 iGrid_Parcela_Selecionada

        If GridParcelas.TextMatrix(iIndice, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO Then
        
            '??? verificar se algum valor p/zerar é dif de zero, colocar contador de linhas validas,...
            Exit For
            
        End If
        
    Next

    'Se não foi -> Erro 'Alterado por Leo em 08/12/01 iGrid_Parcela_Selecionada
    If iIndice = objGridParcelas.iLinhasExistentes And GridParcelas.TextMatrix(iIndice, iGrid_ParcelaSelecionada_Col) <> SELECIONAR_CHECADO Then gError 57938

    Valida_Cancelamento_BaixaBaixa = SUCESSO
     
    Exit Function
    
Erro_Valida_Cancelamento_BaixaBaixa:

    Valida_Cancelamento_BaixaBaixa = gErr
     
    Select Case gErr
          
        Case 57938
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 81841
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXAS_CONTAS_DIF", gErr)
                   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155477)
     
    End Select
     
    Exit Function

End Function

Private Function Move_Tela_Memoria_BaixaBaixa() As Long

Dim lErro As Long
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar, iLinha As Integer

On Error GoTo Erro_Move_Tela_Memoria_BaixaBaixa

    For Each objInfoBaixaRecCancelar In gobjBaixaRecCancelar.colInfoBaixaRecCancelar
    
        'Incrementa a variável de controle da linha atual
        iLinha = iLinha + 1
        
        If objInfoBaixaRecCancelar.iItemSelecionado = MARCADO Then
        
            objInfoBaixaRecCancelar.dValorBaixadoCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorCancelar_Col)), 2)
            objInfoBaixaRecCancelar.dValorMultaCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_MultaCancelar_Col)), 2)
            objInfoBaixaRecCancelar.dValorJurosCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_JurosCancelar_Col)), 2)
            objInfoBaixaRecCancelar.dValorDescontoCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_DescontoCancelar_Col)), 2)
            objInfoBaixaRecCancelar.dValorTotalCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorTotalParcela_Col)), 2)
        
            If (objInfoBaixaRecCancelar.dValorTotalCanc + DELTA_VALORMONETARIO) < (objInfoBaixaRecCancelar.dValorBaixadoCanc + objInfoBaixaRecCancelar.dValorMultaCanc + objInfoBaixaRecCancelar.dValorJurosCanc - objInfoBaixaRecCancelar.dValorDescontoCanc) Then gError 81849
            
            objInfoBaixaRecCancelar.dtDataCancelamento = gobjBaixaRecCancelar.dtDataCancelamento
        
        End If
        
    Next

    Move_Tela_Memoria_BaixaBaixa = SUCESSO
     
    Exit Function
    
Erro_Move_Tela_Memoria_BaixaBaixa:

    Move_Tela_Memoria_BaixaBaixa = gErr
     
    Select Case gErr
          
        Case 81849
            Call Rotina_Erro(vbOKOnly, "ERRO_TOTALCANC_MENOR", gErr, CStr(iLinha))
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155478)
     
    End Select
     
    Exit Function

End Function

Private Sub ValorTotalCancelar_Calcula(iLinha As Integer)
'Gera valor default p/total a cancelar para baixa da linha corrente

Dim dValorBaixadoCanc As Double, dValorMultaCanc As Double, dValorJurosCanc As Double, dValorDescontoCanc As Double, dValorTotalCanc As Double

    dValorBaixadoCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorCancelar_Col)), 2)
    dValorMultaCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_MultaCancelar_Col)), 2)
    dValorJurosCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_JurosCancelar_Col)), 2)
    dValorDescontoCanc = Round(StrParaDbl(GridParcelas.TextMatrix(iLinha, iGrid_DescontoCancelar_Col)), 2)

    dValorTotalCanc = dValorBaixadoCanc - dValorDescontoCanc + dValorMultaCanc + dValorJurosCanc
    If dValorTotalCanc < 0 Then dValorTotalCanc = 0
    
    GridParcelas.TextMatrix(iLinha, iGrid_ValorTotalParcela_Col) = Format(dValorTotalCanc, "STANDARD")

End Sub

Private Function Valida_Cancelamento_MocvCCI() As Long

Dim lErro As Long, iIndice As Integer, bAchou As Boolean

On Error GoTo Erro_Valida_Cancelamento_MocvCCI

    'Verifica se alguma linha do GridMovimentosCCI foi selecionada
    bAchou = False
    For iIndice = 1 To objGridMovimentosCCI.iLinhasExistentes

        If GridMovimentosCCI.TextMatrix(iIndice, iGrid_CancelarMov_Col) = True Then
            
            bAchou = True
            Exit For
            
        End If
        
    Next

    If bAchou = False Then gError 81853
    
    Valida_Cancelamento_MocvCCI = SUCESSO
     
    Exit Function
    
Erro_Valida_Cancelamento_MocvCCI:

    Valida_Cancelamento_MocvCCI = gErr
     
    Select Case gErr
          
        Case 81853
            Call Rotina_Erro(vbOKOnly, "ERRO_SEL_LINHA_GRIDMOV", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155479)
     
    End Select
     
    Exit Function

End Function

Private Function Move_Tela_Memoria_MocvCCI() As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria_MocvCCI

    Move_Tela_Memoria_MocvCCI = SUCESSO
     
    Exit Function
    
Erro_Move_Tela_Memoria_MocvCCI:

    Move_Tela_Memoria_MocvCCI = gErr
     
    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155480)
     
    End Select
     
    Exit Function

End Function

'inicio contabilidade

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

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
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

Public Property Get CTBLabelCcl() As Object
     Set CTBLabelCcl = objUserControl.Controls("CTBLabelCcl")
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

'Início contabilidade
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

'****
Public Sub CTBSeqContraPartida_GotFocus()

    Call objContabil.Contabil_SeqContraPartida_GotFocus

End Sub

'****
Public Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)

    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)

End Sub

'****
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
'Traz o lote selecionado para a tela

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
Dim iLinha As Integer, sCampoGlobal As String
Dim iCodigo As Integer, sContaTela As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objCliente As New ClassCliente, objTipoCliente As New ClassTipoCliente
Dim objFilial As New ClassFilialCliente, objMnemonico As New ClassMnemonicoCTBValor
Dim colParcelasMovCCI As Collection

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case "Cta_Contabil_Cta_Mov"
        
            For iLinha = 1 To objGridMovimentosCCI.iLinhasExistentes
            
                If GridMovimentosCCI.TextMatrix(iLinha, iGrid_CancelarMov_Col) = True Then
                
                    iCodigo = Codigo_Extrai(GridMovimentosCCI.TextMatrix(iLinha, iGrid_CtaCorrenteMovCCI_Col))
                    
                    lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
                    If lErro <> SUCESSO Then gError 49725
                    
                    If objContaCorrenteInt.sContaContabil <> "" Then
                    
                        lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then gError 56525
                    
                    End If
                                
                    Exit For
                
                End If
                                    
            Next
            
            objMnemonicoValor.colValor.Add sContaTela
        
        Case "Cta_CartCobr_Mov"

            lErro = Movcci_ObtemBaixas(colParcelasMovCCI)
            If lErro <> SUCESSO Then gError 81859
            
            'Percorre todas as linhas do Grid
            For iLinha = 1 To objGridBaixasMovimento.iLinhasExistentes

                'Passa a linha do Grid para o Obj
                Set objInfoBaixaRecCancelar = colParcelasMovCCI.Item(iLinha)

                objCarteiraCobrador.iCobrador = objInfoBaixaRecCancelar.iCobrador
                objCarteiraCobrador.iCodCarteiraCobranca = objInfoBaixaRecCancelar.iCarteiraCobrador

                lErro = CartCobr_ObtemCtaTela(objCarteiraCobrador, sContaTela)
                If lErro <> SUCESSO Then gError 81860
                
                objMnemonicoValor.colValor.Add sContaTela
                    
            Next
        
        Case "Valor_Multa_Mov"
            lErro = Movcci_ObtemBaixas(colParcelasMovCCI)
            If lErro <> SUCESSO Then gError 81861
            
            For iLinha = 1 To objGridBaixasMovimento.iLinhasExistentes
                
                Set objInfoBaixaRecCancelar = colParcelasMovCCI.Item(iLinha)
                objMnemonicoValor.colValor.Add Round(objInfoBaixaRecCancelar.dValorMultaCanc, 2)
                
            Next
                    
        Case "Valor_Juros_Mov"
        
            lErro = Movcci_ObtemBaixas(colParcelasMovCCI)
            If lErro <> SUCESSO Then gError 81862
            
            For iLinha = 1 To objGridBaixasMovimento.iLinhasExistentes
                
                Set objInfoBaixaRecCancelar = colParcelasMovCCI.Item(iLinha)
                objMnemonicoValor.colValor.Add Round(objInfoBaixaRecCancelar.dValorJurosCanc, 2)
                
            Next
        
        Case "Valor_Desconto_Mov"
        
            lErro = Movcci_ObtemBaixas(colParcelasMovCCI)
            If lErro <> SUCESSO Then gError 81863
            
            For iLinha = 1 To objGridBaixasMovimento.iLinhasExistentes
                
                Set objInfoBaixaRecCancelar = colParcelasMovCCI.Item(iLinha)
                objMnemonicoValor.colValor.Add Round(objInfoBaixaRecCancelar.dValorDescontoCanc, 2)
                
            Next
        
        Case "Valor_Cancelar_Mov"
        
            lErro = Movcci_ObtemBaixas(colParcelasMovCCI)
            If lErro <> SUCESSO Then gError 81864
            
            For iLinha = 1 To objGridBaixasMovimento.iLinhasExistentes
                
                Set objInfoBaixaRecCancelar = colParcelasMovCCI.Item(iLinha)
                objMnemonicoValor.colValor.Add Round(objInfoBaixaRecCancelar.dValorBaixadoCanc, 2)
                
            Next
        
        Case "Parcela_Mov"
        
            For iLinha = 1 To objGridBaixasMovimento.iLinhasExistentes
                If Len(GridBaixasMovimento.TextMatrix(iLinha, iGrid_ParcelaBaixada_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CInt(GridBaixasMovimento.TextMatrix(iLinha, iGrid_ParcelaBaixada_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case "Numero_Titulo_Mov"
            
            For iLinha = 1 To objGridBaixasMovimento.iLinhasExistentes
                If Len(GridBaixasMovimento.TextMatrix(iLinha, iGrid_NumeroBaixada_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CLng(GridBaixasMovimento.TextMatrix(iLinha, iGrid_NumeroBaixada_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
        
        Case CLIENTE_COD
            
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 92855
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case CLIENTE_NOME
        
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then gError 56520
            
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case BAIXAR
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                    
                objMnemonicoValor.colValor.Add CInt(GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col))
                
            Next
        
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
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 56521
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 56522
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then gError 56523
                
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
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then gError 56524
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
        
        Case DATA_CANCELAMENTO
            If Len(DataCancelamento.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataCancelamento.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case NUM_TITULO
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO And Len(GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CLng(GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case PARCELA1
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO And Len(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case VALOR_CANCELAR
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorCancelar_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorCancelar_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case VALOR_DESCONTO
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_DescontoCancelar_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_DescontoCancelar_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case VALOR_MULTA
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_MultaCancelar_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_MultaCancelar_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case VALOR_JUROS
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_JurosCancelar_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_JurosCancelar_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case PAGTO_DINHEIRO
            
            objMnemonicoValor.colValor.Add Recebimento(MOTIVO_RECEBIMENTO).Value
            
        Case PAGTO_ANTECIPADO
            
            objMnemonicoValor.colValor.Add Recebimento(MOTIVO_RECEBTO_ANTECIPADO).Value
            
        Case CRED_DEV
            
            objMnemonicoValor.colValor.Add Recebimento(MOTIVO_DEBITO_CLIENTE).Value
                
        Case PERDA
            objMnemonicoValor.colValor.Add Recebimento(MOTIVO_PERDA).Value
            
        Case CTA_CONTA_CORRENTE
            
            sContaTela = ""
    
            'Percorre todas as linhas do Grid Parcelas
            For iLinha = 1 To objGridParcelas.iLinhasExistentes

                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO Then
                
                    'Passa a linha do Grid para o Obj
                    Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Item(iLinha)
    
                    If objInfoBaixaRecCancelar.iCodConta <> 0 Then
                    
                        iCodigo = objInfoBaixaRecCancelar.iCodConta
                        
                        lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
                        If lErro <> SUCESSO Then gError 49725
                        
                        If objContaCorrenteInt.sContaContabil <> "" Then
                        
                            lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                            If lErro <> SUCESSO Then gError 56525
                        
                        End If
                                    
                    End If
                                    
                    Exit For
                                    
                End If
                
            Next
            
            objMnemonicoValor.colValor.Add sContaTela
        
            
        Case CTACARTEIRACOBRADOR

            'Percorre todas as linhas do Grid Parcelas
            For iLinha = 1 To objGridParcelas.iLinhasExistentes

                sContaTela = ""
    
                If GridParcelas.TextMatrix(iLinha, iGrid_ParcelaSelecionada_Col) = SELECIONAR_CHECADO Then
                
                    'Passa a linha do Grid para o Obj
                    Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar.Item(iLinha)
    
                    objCarteiraCobrador.iCobrador = objInfoBaixaRecCancelar.iCobrador
                    objCarteiraCobrador.iCodCarteiraCobranca = objInfoBaixaRecCancelar.iCarteiraCobrador
    
                    lErro = CartCobr_ObtemCtaTela(objCarteiraCobrador, sContaTela)
                    If lErro <> SUCESSO Then gError 81865
    
                End If
            
                objMnemonicoValor.colValor.Add sContaTela
                    
            Next
            
        Case Else
            Error 39695
    
    End Select
    
    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
    
    Select Case gErr
        
        Case 49724, 49725, 49726, 56519 To 56526, 56800, 92855, 81859 To 81865
        
        Case 56797, 56799
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", gErr, objCarteiraCobrador.iCobrador)
        
        Case 56801
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", gErr, objMnemonico.sMnemonico)
        
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155481)
            
    End Select
    
    Exit Function
    
End Function

Private Function Movcci_ObtemBaixas(colParcelasMovCCI As Collection, Optional iLinha As Integer = 0) As Long

Dim lErro As Long, iIndice As Integer
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar

On Error GoTo Erro_Movcci_ObtemBaixas

    If iLinha = 0 Then
    
        For iIndice = 1 To objGridMovimentosCCI.iLinhasExistentes
    
            If GridMovimentosCCI.TextMatrix(iIndice, iGrid_CancelarMov_Col) = True Then
                
                iLinha = iIndice
                Exit For
                
            End If
            
        Next
    
    End If
    
    If iLinha <> 0 Then
    
        Set colParcelasMovCCI = gobjBaixaRecCancelar.colParcelasMovCCI(iLinha)
    
        If colParcelasMovCCI.Count = 0 Then
            
            Set objInfoBaixaRecCancelar = gobjBaixaRecCancelar.colInfoBaixaRecCancelar(iLinha)
            
            lErro = CF("BaixaRecCancelar_Le_ParcelasMovCCI", objInfoBaixaRecCancelar.lNumMovCta, colParcelasMovCCI)
            If lErro <> SUCESSO And lErro <> 94834 Then gError 81858
            
            If lErro = 94834 Then gError 94838
    
        End If
            
    End If
    
    Movcci_ObtemBaixas = SUCESSO
     
    Exit Function
    
Erro_Movcci_ObtemBaixas:

    Movcci_ObtemBaixas = gErr
     
    Select Case gErr
          
        Case 81858
        
        Case 94838
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXARECCANCELAR_SEM_BAIXAS", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155482)
     
    End Select
     
    Exit Function

End Function

Private Function CartCobr_ObtemCtaTela(objCarteiraCobrador As ClassCarteiraCobrador, sContaTela As String) As Long

Dim lErro As Long, sCampoGlobal As String, objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_CartCobr_ObtemCtaTela

    If objCarteiraCobrador.iCobrador = COBRADOR_PROPRIA_EMPRESA Then

        Select Case objCarteiraCobrador.iCodCarteiraCobranca

            Case CARTEIRA_CARTEIRA
                sCampoGlobal = "CtaReceberCarteira"

            Case CARTEIRA_CHEQUEPRE
                sCampoGlobal = "CtaRecChequePre"

            Case CARTEIRA_JURIDICO
                sCampoGlobal = "CtaJuridico"

            Case Else
                Error 56799

        End Select

        objMnemonico.sMnemonico = sCampoGlobal
        lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
        If lErro <> SUCESSO And lErro <> 39690 Then Error 56800
        If lErro <> SUCESSO Then Error 56801

        sContaTela = objMnemonico.sValor

    Else

        lErro = CF("CarteiraCobrador_Le", objCarteiraCobrador)
        If lErro <> SUCESSO And lErro <> 23551 Then Error 49726
        If lErro <> SUCESSO Then Error 56797

        If objCarteiraCobrador.sContaContabil <> "" Then

            lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaContabil, sContaTela)
            If lErro <> SUCESSO Then Error 56526

        End If

    End If

    CartCobr_ObtemCtaTela = SUCESSO
     
    Exit Function
    
Erro_CartCobr_ObtemCtaTela:

    CartCobr_ObtemCtaTela = gErr
     
    Select Case gErr
          
        Case 49726, 56526, 56800
        
        Case 56797, 56799
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56801
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155483)
     
    End Select
     
    Exit Function

End Function

Private Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134020

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134020

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155484)

    End Select
    
    Exit Sub
    
End Sub

Public Sub TipoBaixaCartao_Click()
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    gobjBaixaRecCancelar.iTipoBaixas = MOTIVO_CARTAO_DEBITO_CREDITO
    Call Recebimento_Click(MOTIVO_CARTAO_DEBITO_CREDITO)
End Sub

Private Function Traz_Dados_Titulo_Cartao(objBaixaReceber As ClassBaixaReceber) As Long
'Mostra os dados do recebimento utilizado para a baixa

Dim lErro As Long
Dim objTitRec As New ClassTituloReceber
Dim objCliente As New ClassCliente
Dim sCliente As String
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_Traz_Dados_Titulo_Cartao

    
    objTitRec.lNumIntDoc = objBaixaReceber.lNumIntDoc

    lErro = CF("TituloReceber_Le", objTitRec)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 183165

    If lErro <> SUCESSO Then
    
        lErro = CF("TituloReceberBaixado_Le", objTitRec)
        If lErro <> SUCESSO And lErro <> 56570 Then gError 183166
    
    End If

    If lErro = SUCESSO Then
    
        objCliente.lCodigo = objTitRec.lCliente
    
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 183167
        
        If lErro = SUCESSO Then
            
            ClienteTitCartao.Caption = objCliente.sNomeReduzido
            
            sCliente = objCliente.sNomeReduzido
            objFilialCliente.iCodFilial = objTitRec.iFilial
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 183168
    
            If lErro = SUCESSO Then
    
                'Coloca na tela a Filial lida
                FilialTitCartao.Caption = objTitRec.iFilial & SEPARADOR & objFilialCliente.sNome
            
            End If
            
        End If
        
        TipoTitCartao.Caption = objTitRec.sSiglaDocumento
        NumTitCartao.Caption = objTitRec.lNumTitulo
        DataEmissaoTitCartao.Caption = Format(objTitRec.dtDataEmissao, "dd/mm/yyyy")

    End If

    Traz_Dados_Titulo_Cartao = SUCESSO

    Exit Function

Erro_Traz_Dados_Titulo_Cartao:

    Traz_Dados_Titulo_Cartao = gErr

    Select Case gErr

        Case 183165 To 183168

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158398)

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

Private Function Trata_Data_Cancelamento() As Long

Dim lErro As Long
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar
Dim dtMaiorDataBaixa As Date
Dim iIndice As Integer

On Error GoTo Erro_Trata_Data_Cancelamento

    dtMaiorDataBaixa = DATA_NULA

    Select Case gobjBaixaRecCancelar.iTipoCancelamento

        Case TIPOCANCELAMENTO_BAIXABAIXA

            For Each objInfoBaixaRecCancelar In gobjBaixaRecCancelar.colInfoBaixaRecCancelar
                
                If objInfoBaixaRecCancelar.iItemSelecionado = MARCADO Then
                    'If dtMaiorDataBaixa < objInfoBaixaRecCancelar.dtDataBaixa Then dtMaiorDataBaixa = objInfoBaixaRecCancelar.dtDataBaixa
                    If dtMaiorDataBaixa < objInfoBaixaRecCancelar.dtDataCredito Then dtMaiorDataBaixa = objInfoBaixaRecCancelar.dtDataCredito
                End If
                
            Next
    
        Case TIPOCANCELAMENTO_MOVCCI
        
            For iIndice = 1 To objGridMovimentosCCI.iLinhasExistentes
                If GridMovimentosCCI.TextMatrix(iIndice, iGrid_CancelarMov_Col) = True Then
                    dtMaiorDataBaixa = StrParaDate(GridMovimentosCCI.TextMatrix(iIndice, iGrid_DataMovCCI_Col))
                    Exit For
                End If
            Next

    End Select
    
    If dtMaiorDataBaixa <> DATA_NULA Then
        DataCancelamento.PromptInclude = False
        DataCancelamento.Text = Format(dtMaiorDataBaixa, "dd/mm/yy")
        DataCancelamento.PromptInclude = True
        Call DataCancelamento_Validate(bSGECancelDummy)
    
'        CTBDataContabil.PromptInclude = False
'        CTBDataContabil.Text = Format(dtMaiorDataBaixa, "dd/mm/yy")
'        CTBDataContabil.PromptInclude = True
'        Call CTBDataContabil_Validate(bSGECancelDummy)
    End If

    Trata_Data_Cancelamento = SUCESSO
     
    Exit Function
    
Erro_Trata_Data_Cancelamento:

    Trata_Data_Cancelamento = gErr
     
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196994)
     
    End Select
     
    Exit Function

End Function

Private Function Critica_Data_Cancelamento() As Long

Dim lErro As Long
Dim objInfoBaixaRecCancelar As ClassInfoBaixaRecCancelar
Dim iLinha As Integer
Dim dtDataCanc As Date
Dim vbMsgResult As VbMsgBoxResult

On Error GoTo Erro_Critica_Data_Cancelamento

    dtDataCanc = StrParaDate(DataCancelamento.Text)

    Select Case gobjBaixaRecCancelar.iTipoCancelamento

        Case TIPOCANCELAMENTO_BAIXABAIXA
        
            For Each objInfoBaixaRecCancelar In gobjBaixaRecCancelar.colInfoBaixaRecCancelar
                
                iLinha = iLinha + 1
                If objInfoBaixaRecCancelar.iItemSelecionado = MARCADO Then
                
                    If dtDataCanc <> objInfoBaixaRecCancelar.dtDataCredito Then
                        vbMsgResult = Rotina_Aviso(vbYesNoCancel, "AVISO_DATA_BAIXA_DIF_DATA_CANC_BAIXA")
                        If vbMsgResult = vbNo Then gError 196995
                        
                        If vbMsgResult = vbCancel Then Exit For
                    End If
                    
                End If
                
            Next
            
        Case TIPOCANCELAMENTO_MOVCCI
        
            For iLinha = 1 To objGridMovimentosCCI.iLinhasExistentes
                If GridMovimentosCCI.TextMatrix(iLinha, iGrid_CancelarMov_Col) = True Then
                    If dtDataCanc <> StrParaDate(GridMovimentosCCI.TextMatrix(iLinha, iGrid_DataMovCCI_Col)) Then
                        vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_DATA_BAIXA_DIF_DATA_CANC_BAIXA2")
                        If vbMsgResult = vbNo Then gError 196996
                    End If
                    
                    Exit For
                End If
            Next

            
    End Select

    Critica_Data_Cancelamento = SUCESSO
     
    Exit Function
    
Erro_Critica_Data_Cancelamento:

    Critica_Data_Cancelamento = gErr
     
    Select Case gErr
    
        Case 196995, 196996

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196997)
     
    End Select
     
    Exit Function

End Function
