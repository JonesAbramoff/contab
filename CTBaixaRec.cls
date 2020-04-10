VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTBaixaRec"
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

'Início da contabilidade
Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoCodigoViagem As AdmEvento
Attribute objEventoCodigoViagem.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1
Private WithEvents objEventoProdutoDe As AdmEvento
Attribute objEventoProdutoDe.VB_VarHelpID = -1
Private WithEvents objEventoProdutoAte As AdmEvento
Attribute objEventoProdutoAte.VB_VarHelpID = -1

Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const FILIAL_COD As String = "FilialCli_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialCli_Nome"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialCli_CGC_CPF"
'###########Parcelas#############
Private Const CLIENTE_COD_GRID As String = "Cliente_Cod_Parc"
Private Const CLIENTE_NOME_GRID As String = "Cliente_Nome_Parc"
Private Const FILIAL_COD_GRID As String = "FilialCli_Cod_Parc"
Private Const FILIAL_NOME_RED_GRID As String = "FilialCli_Nome_Parc"
Private Const FILIAL_CONTA_GRID As String = "FilCli_Cta_Ctb_Parc"
Private Const FILIAL_CGC_CPF_GRID As String = "FilCli_CGC_CPF_Parc"

Private Const DATA_BAIXA As String = "Data_Baixa"
Private Const NUM_TITULO As String = "Numero_Titulo"
Private Const PARCELA1 As String = "Parcela"
Private Const VALOR_BAIXAR As String = "Valor_Baixar"
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
Private Const NOME_CONTA_CORRENTE As String = "Nome_Conta_Corente"

'Fim da contabilidade

''COLUNAS DOS GRIDS
''Grid Parcelas:
'Const VALOR_BAIXAR_COL = 7
'Const VALOR_DESCONTO_COL = 8
'Const VALOR_MULTA_COL = 9
'Const VALOR_JUROS_COL = 10

'Indicação de número máximo de linhas dos Grids
Const NUM_MAX_PARCELAS_BAIXA = 300
Const NUM_MAX_RECEB_ANTECIPADOS = 600
Const NUM_MAX_DEBITOS = 600

'Grid Parcelas:

Public objGridParcelas As AdmGrid
Public iGrid_Baixar_Col As Integer
Public iGrid_DataVencimento_Col As Integer
Public iGrid_Tipo_Col As Integer
Public iGrid_Numero_Col As Integer
Public iGrid_NossoNumero_Col As Integer
Public iGrid_Emissao_Col As Integer
Public iGrid_Parcela_Col As Integer
Public iGrid_Saldo_Col As Integer
Public iGrid_ValorBaixar_Col As Integer
Public iGrid_ValorDesconto_Col As Integer
Public iGrid_ValorMulta_Col As Integer
Public iGrid_ValorJuros_Col As Integer
Public iGrid_ValorAReceber_Col As Integer
Public iGrid_ValorParcela_Col As Integer
Public iGrid_FilialEmpresa_Col As Integer
Public iGrid_Cobrador_Col As Integer
Public iGrid_Cliente_Col As Integer
Public iGrid_Filial_Col As Integer

'Grid RecebAntecipados:
Public objGridRecebAntecipados As AdmGrid
Public iGrid_SelecionarRA_Col As Integer
Public iGrid_DataMovimentoRA_Col As Integer
Public iGrid_ContaCorrenteRA_Col As Integer
Public iGrid_MeioPagtoRA_Col As Integer
Public iGrid_ValorRA_Col As Integer
Public iGrid_SaldoRA_Col As Integer
Public iGrid_Doc_Col As Integer
Public iGrid_Seq_Col As Integer
Public iGrid_Hist_Col As Integer
Public iGrid_FilialRA_Col As Integer

'Grid Débitos:
Dim objGridDebitos As AdmGrid
Dim iGrid_SelecionarDB_Col As Integer
Dim iGrid_DataEmissaoDB_Col As Integer
Dim iGrid_TipoDocDB_Col As Integer
Dim iGrid_NumTituloDB_Col As Integer
Dim iGrid_ValorDB_Col As Integer
Dim iGrid_SaldoDB_Col As Integer
Dim iGrid_FilialDB_Col As Integer
Dim iGrid_OBSDB_Col As Integer

Public iAlterado As Integer
Public iFramePrincipalAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFrameAtual As Integer
Dim iGridAtual As Integer
Dim iFrameRecebimentoAtual As Integer
Dim gcolInfoParcRec As Collection
Dim gcolRecebAntecipado As New Collection
Dim gcolDebitos As New Collection
Private dtDataMultaJuros As Date

Private WithEvents objEventoClientes As AdmEvento
Attribute objEventoClientes.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Titulos = 1
Private Const TAB_Parcelas = 2
Private Const TAB_Contabilizacao = 3

'--- inicio dos properties get dos controles da tela
Public Property Get CobradorFiltro() As Object
     Set CobradorFiltro = objUserControl.Controls("CobradorFiltro")
End Property

Public Property Get NossoNumero() As Object
     Set NossoNumero = objUserControl.Controls("NossoNumero")
End Property

Public Property Get Emissao() As Object
     Set Emissao = objUserControl.Controls("Emissao")
End Property

Public Property Get Vendedor() As Object
     Set Vendedor = objUserControl.Controls("Vendedor")
End Property

Public Property Get LabelViagem() As Object
     Set LabelViagem = objUserControl.Controls("LabelViagem")
End Property

Public Property Get CodigoViagem() As Object
     Set CodigoViagem = objUserControl.Controls("CodigoViagem")
End Property

Public Property Get FormaPagto() As Object
     Set FormaPagto = objUserControl.Controls("FormaPagto")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get FrameRecebimento() As Object
     Set FrameRecebimento = objUserControl.Controls("FrameRecebimento")
End Property

Public Property Get FrameTipoMeioPagto() As Object
     Set FrameTipoMeioPagto = objUserControl.Controls("FrameTipoMeioPagto")
End Property

Public Property Get ContaCorrente() As Object
     Set ContaCorrente = objUserControl.Controls("ContaCorrente")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get ValorReceber() As Object
     Set ValorReceber = objUserControl.Controls("ValorReceber")
End Property

Public Property Get HistoricoPerda() As Object
     Set HistoricoPerda = objUserControl.Controls("HistoricoPerda")
End Property

Public Property Get SelecionarDB() As Object
     Set SelecionarDB = objUserControl.Controls("SelecionarDB")
End Property

Public Property Get DataEmissaoDB() As Object
     Set DataEmissaoDB = objUserControl.Controls("DataEmissaoDB")
End Property

Public Property Get SaldoDB() As Object
     Set SaldoDB = objUserControl.Controls("SaldoDB")
End Property

Public Property Get ValorDB() As Object
     Set ValorDB = objUserControl.Controls("ValorDB")
End Property

Public Property Get NumTituloDB() As Object
     Set NumTituloDB = objUserControl.Controls("NumTituloDB")
End Property

Public Property Get TipoDocDB() As Object
     Set TipoDocDB = objUserControl.Controls("TipoDocDB")
End Property

Public Property Get FilialDB() As Object
     Set FilialDB = objUserControl.Controls("FilialDB")
End Property

Public Property Get GridDebitos() As Object
     Set GridDebitos = objUserControl.Controls("GridDebitos")
End Property

Public Property Get SelecionarRA() As Object
     Set SelecionarRA = objUserControl.Controls("SelecionarRA")
End Property

Public Property Get ContaCorrenteRA() As Object
     Set ContaCorrenteRA = objUserControl.Controls("ContaCorrenteRA")
End Property

Public Property Get MeioPagtoRA() As Object
     Set MeioPagtoRA = objUserControl.Controls("MeioPagtoRA")
End Property

Public Property Get DataMovimentoRA() As Object
     Set DataMovimentoRA = objUserControl.Controls("DataMovimentoRA")
End Property

Public Property Get ValorRA() As Object
     Set ValorRA = objUserControl.Controls("ValorRA")
End Property

Public Property Get SaldoRA() As Object
     Set SaldoRA = objUserControl.Controls("SaldoRA")
End Property

Public Property Get FilialRA() As Object
     Set FilialRA = objUserControl.Controls("FilialRA")
End Property

Public Property Get GridRecebAntecipados() As Object
     Set GridRecebAntecipados = objUserControl.Controls("GridRecebAntecipados")
End Property

Public Property Get Cobranca() As Object
     Set Cobranca = objUserControl.Controls("Cobranca")
End Property

Public Property Get Selecionar() As Object
     Set Selecionar = objUserControl.Controls("Selecionar")
End Property

Public Property Get DataVencimento() As Object
     Set DataVencimento = objUserControl.Controls("DataVencimento")
End Property

Public Property Get Saldo() As Object
     Set Saldo = objUserControl.Controls("Saldo")
End Property

Public Property Get ValorBaixar() As Object
     Set ValorBaixar = objUserControl.Controls("ValorBaixar")
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

Public Property Get ValorDesconto() As Object
     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
End Property

Public Property Get ValorMulta() As Object
     Set ValorMulta = objUserControl.Controls("ValorMulta")
End Property

Public Property Get ValorJuros() As Object
     Set ValorJuros = objUserControl.Controls("ValorJuros")
End Property

Public Property Get Cobrador() As Object
     Set Cobrador = objUserControl.Controls("Cobrador")
End Property

Public Property Get GridParcelas() As Object
     Set GridParcelas = objUserControl.Controls("GridParcelas")
End Property

Public Property Get ValorParcela() As Object
     Set ValorParcela = objUserControl.Controls("ValorParcela")
End Property

Public Property Get FilialEmpresa() As Object
     Set FilialEmpresa = objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get ValorAReceber() As Object
     Set ValorAReceber = objUserControl.Controls("ValorAReceber")
End Property

Public Property Get Frame10() As Object
     Set Frame10 = objUserControl.Controls("Frame10")
End Property

Public Property Get Recebimento() As Object
     Set Recebimento = objUserControl.Controls("Recebimento")
End Property

Public Property Get UpDownDataBaixa() As Object
     Set UpDownDataBaixa = objUserControl.Controls("UpDownDataBaixa")
End Property

Public Property Get DataBaixa() As Object
     Set DataBaixa = objUserControl.Controls("DataBaixa")
End Property

Public Property Get UpDownDataCredito() As Object
     Set UpDownDataCredito = objUserControl.Controls("UpDownDataCredito")
End Property

Public Property Get DataCredito() As Object
     Set DataCredito = objUserControl.Controls("DataCredito")
End Property

Public Property Get TotalBaixar() As Object
     Set TotalBaixar = objUserControl.Controls("TotalBaixar")
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

Public Property Get Frame9() As Object
     Set Frame9 = objUserControl.Controls("Frame9")
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

Public Property Get Frame8() As Object
     Set Frame8 = objUserControl.Controls("Frame8")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get LabelCli() As Object
     Set LabelCli = objUserControl.Controls("LabelCli")
End Property

Public Property Get LabelFilial() As Object
     Set LabelFilial = objUserControl.Controls("LabelFilial")
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

Public Property Get ClienteItem() As Object
     Set ClienteItem = objUserControl.Controls("ClienteItem")
End Property

Public Property Get FilialClienteItem() As Object
     Set FilialClienteItem = objUserControl.Controls("FilialClienteItem")
End Property

Public Property Get ImprimirRecibo() As Object
     Set ImprimirRecibo = objUserControl.Controls("ImprimirRecibo")
End Property

Public Property Get Doc() As Object
     Set Doc = objUserControl.Controls("Doc")
End Property

Public Property Get Seq() As Object
     Set Seq = objUserControl.Controls("Seq")
End Property

Public Property Get Hist() As Object
     Set Hist = objUserControl.Controls("Hist")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get TipoDocSeleciona() As Object
     Set TipoDocSeleciona = objUserControl.Controls("TipoDocSeleciona")
End Property

Public Property Get TipoDocTodos() As Object
     Set TipoDocTodos = objUserControl.Controls("TipoDocTodos")
End Property

Public Property Get TipoDocApenas() As Object
     Set TipoDocApenas = objUserControl.Controls("TipoDocApenas")
End Property

Public Property Get ProdutoInicial() As Object
     Set ProdutoInicial = objUserControl.Controls("ProdutoInicial")
End Property

Public Property Get DescProdInic() As Object
     Set DescProdInic = objUserControl.Controls("DescProdInic")
End Property

Public Property Get ProdutoFinal() As Object
     Set ProdutoFinal = objUserControl.Controls("ProdutoFinal")
End Property

Public Property Get DescProdFim() As Object
     Set DescProdFim = objUserControl.Controls("DescProdFim")
End Property

Public Property Get OBSDB() As Object
     Set OBSDB = objUserControl.Controls("OBSDB")
End Property

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
    If lErro <> SUCESSO Then Error 42756

    Call Limpa_Tela_BaixaRec
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 42756

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143394)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Verifica se houve alterações e confirma se deseja salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 42757

    'Limpa a tela
    Call Limpa_Tela_BaixaRec

    'Passa a data do sistema para a tela
    DataBaixa = Format(gdtDataAtual, "dd/mm/yy")
    DataCredito = Format(gdtDataAtual, "dd/mm/yy")

    'Zera iAlterado
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 42757

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143395)

    End Select

    Exit Sub

End Sub

Public Sub DataBaixa_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataBaixa, iAlterado)

End Sub

Public Sub DataCredito_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataCredito, iAlterado)

End Sub

Public Sub DataEmissaoDB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDebitos)

End Sub

Public Sub DataEmissaoDB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)

End Sub

Public Sub DataEmissaoDB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDebitos.objControle = DataEmissaoDB
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Public Sub EmissaoFim_GotFocus()
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(EmissaoFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub EmissaoInic_GotFocus()
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(EmissaoInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux
    
End Sub

Public Sub HistoricoPerda_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LabelCli_Click()
'Chamada do Browse de Clientes

Dim colSelecao As Collection
Dim objCliente As New ClassCliente

    'Passa o Cliente que está na tela para o Obj
    objCliente.sNomeReduzido = Trim(Cliente.Text)

    'Chama a tela com a lista de Clientes
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClientes)

    Exit Sub

End Sub

Public Sub TipoDocDB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDebitos)

End Sub

Public Sub TipoDocDB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)

End Sub

Public Sub TipoDocDB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDebitos.objControle = TipoDocDB
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub NumTituloDB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDebitos)

End Sub

Public Sub NumTituloDB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)

End Sub

Public Sub NumTituloDB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDebitos.objControle = NumTituloDB
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub TituloFim_GotFocus()
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(TituloFim, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub TituloInic_GotFocus()
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(TituloInic, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub ValorDB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDebitos)

End Sub

Public Sub ValorDB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)

End Sub

Public Sub ValorDB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDebitos.objControle = ValorDB
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub SaldoDB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDebitos)

End Sub

Public Sub SaldoDB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)

End Sub

Public Sub SaldoDB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDebitos.objControle = SaldoDB
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub SelecionarDB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDebitos)

End Sub

Public Sub SelecionarDB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)

End Sub

Public Sub SelecionarDB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDebitos.objControle = SelecionarDB
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub FilialDB_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDebitos)

End Sub

Public Sub FilialDB_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDebitos)

End Sub

Public Sub FilialDB_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridDebitos.objControle = FilialDB
    lErro = Grid_Campo_Libera_Foco(objGridDebitos)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ContaCorrenteRA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRecebAntecipados)

End Sub

Public Sub ContaCorrenteRA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRecebAntecipados)

End Sub

Public Sub ContaCorrenteRA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRecebAntecipados.objControle = ContaCorrenteRA
    lErro = Grid_Campo_Libera_Foco(objGridRecebAntecipados)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ContaCorrente_Click()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaCorrente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim vbMsgRes As VbMsgBoxResult
Dim iCodigo As Integer

On Error GoTo Erro_ContaCorrente_Validate

    'Verifica se a Conta está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Exit Sub

    'Verifica se esta preenchida com o ítem selecionado na ComboBox CodConta
    If ContaCorrente.Text = ContaCorrente.List(ContaCorrente.ListIndex) Then Exit Sub

    'Verifica se o a Conta existe na Combo, e , se existir, seleciona
    lErro = Combo_Seleciona(ContaCorrente, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 42770

    'Se a Conta(CODIGO) não existe na Combo
    If lErro = 6730 Then

        'Passa o Código da Conta para o Obj
        objContaCorrenteInt.iCodigo = iCodigo

        'Lê os dados da Conta
        lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then Error 42771

        'Se a Conta não estiver cadastrada
        If lErro = 11807 Then Error 42772
        
        If giFilialEmpresa <> EMPRESA_TODA Then

            'Verifica se a Conta é Filial Empresa corrente
            If objContaCorrenteInt.iFilialEmpresa <> giFilialEmpresa Then Error 43100
        
        End If
        
        'Passa o código da Conta para a tela
        ContaCorrente.Text = CStr(objContaCorrenteInt.iCodigo) & SEPARADOR & objContaCorrenteInt.sNomeReduzido

    End If

    'Se a Conta(STRING) não existe na Combo
    If lErro = 6731 Then Error 42773
    
    Exit Sub

Erro_ContaCorrente_Validate:

    Cancel = True


    Select Case Err

        Case 42770, 42771

        Case 42772
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CODCONTACORRENTE_INEXISTENTE", objContaCorrenteInt.iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CtaCorrenteInt", objContaCorrenteInt)
            Else
            End If

        Case 42773
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, ContaCorrente.Text)
            
        Case 43100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_FILIAL_DIFERENTE", Err, objContaCorrenteInt.iCodigo, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143396)

    End Select

    Exit Sub

End Sub

Public Sub DataBaixa_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataBaixa_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataBaixa_Validate

    'Se a DataBaixa está preenchida
    If Len(DataBaixa.ClipText) > 0 Then

        'Verifica se a DataBaixa é válida
        lErro = Data_Critica(DataBaixa.Text)
        If lErro <> SUCESSO Then Error 46374

        Call Calcula_Multa_Juros_Desc_Parcelas
        
    'Se a DataBaixa não está preenchida
    Else
    
        Error 46375
        
    End If
    
    DataCredito.PromptInclude = False
    DataCredito.Text = Format(DataBaixa.Text, "dd/mm/yy")
    DataCredito.PromptInclude = True

    Exit Sub

Erro_DataBaixa_Validate:

    Cancel = True


    Select Case Err

        Case 46374

        Case 46375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_SEM_PREENCHIMENTO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143397)

    End Select

    Exit Sub

End Sub

Public Sub DataMovimentoRA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRecebAntecipados)

End Sub

Public Sub DataMovimentoRA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRecebAntecipados)

End Sub

Public Sub DataMovimentoRA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRecebAntecipados.objControle = DataMovimentoRA
    lErro = Grid_Campo_Libera_Foco(objGridRecebAntecipados)
    If lErro <> SUCESSO Then Cancel = True
    
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

Public Sub EmissaoFim_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoFim_Validate

    'Se a data EmissaoFim está preenchida
    If Len(EmissaoFim.ClipText) > 0 Then

        'Verifica se a data EmissaoFim é válida
        lErro = Data_Critica(EmissaoFim.Text)
        If lErro <> SUCESSO Then Error 42774

        'Se a data EmissaoInic Está preenchida
        If Len(EmissaoInic.ClipText) > 0 Then

            'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
            If CDate(EmissaoFim.Text) < CDate(EmissaoInic.Text) Then Error 42775

        End If

    End If

    Exit Sub

Erro_EmissaoFim_Validate:

    Cancel = True


    Select Case Err

        Case 42774

        Case 42775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143398)

    End Select

    Exit Sub

End Sub

Public Sub EmissaoInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub EmissaoInic_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_EmissaoInic_Validate

    'Se a data EmissaoInic está preenchida
    If Len(EmissaoInic.ClipText) > 0 Then

        'Verifica se a data EmissaoInic é válida
        lErro = Data_Critica(EmissaoInic.Text)
        If lErro <> SUCESSO Then Error 42776

    End If

    Exit Sub

Erro_EmissaoInic_Validate:

    Cancel = True


    Select Case Err

        Case 42776

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143399)

    End Select

    Exit Sub

End Sub

Public Sub Filial_Click()

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

    'Verifica se foi preenchida a ComboBox Filial
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox Filial
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 29022

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        'Verifica se foi preenchido o Cliente
        If Len(Trim(Cliente.Text)) = 0 Then Error 29023

        'Lê o Cliente que está na tela
        sNomeRed = Trim(Cliente.Text)

        'Passa o Código da Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = iCodigo

        'Lê Filial no BD a partir do NomeReduzido do Cliente e Código da Filial
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sNomeRed, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 29024

        'Se não existe a Filial
        If lErro = 17660 Then Error 29025

        'Encontrou Filial no BD, coloca no Text da Combo
        Filial.Text = CStr(objFilialCliente.iCodFilial) & SEPARADOR & objFilialCliente.sNome

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then Error 29026

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 29022, 29024

        Case 29023
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 29025
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE1", Filial.Text)

            If vbMsgRes = vbYes Then
                'Chama a tela de Filiais
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
                'Segura o foco
            End If

        Case 29026
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143400)

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

Public Sub FilialRA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRecebAntecipados)

End Sub

Public Sub FilialRA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRecebAntecipados)

End Sub

Public Sub FilialRA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRecebAntecipados.objControle = FilialRA
    lErro = Grid_Campo_Libera_Foco(objGridRecebAntecipados)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

Dim lErro As Long

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gcolInfoParcRec = Nothing
    Set objEventoClientes = Nothing
    Set objGridParcelas = Nothing
    Set objGridRecebAntecipados = Nothing
    Set objGridDebitos = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    Set objEventoCodigoViagem = Nothing
    Set objEventoVendedor = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing

    Set gcolRecebAntecipado = Nothing
    Set gcolDebitos = Nothing

    Set objEventoProdutoDe = Nothing
    Set objEventoProdutoAte = Nothing

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
        If lErro <> SUCESSO Then Error 29027

        'Lê coleção de códigos, nomes de Filiais do Cliente
        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then Error 29028

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se Cliente não está preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If
    
    iClienteAlterado = 0

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True
    
    Select Case Err

        Case 29027
            
        Case 29028

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143401)

    End Select

    Exit Sub

End Sub

Public Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcol As New Collection

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If
    
    colcol.Add gcolInfoParcRec
    Call Ordenacao_ClickGrid(objGridParcelas, Nothing, colcol)

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

Public Sub GridRecebAntecipados_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridRecebAntecipados, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRecebAntecipados, iAlterado)
    End If

End Sub

Public Sub GridRecebAntecipados_GotFocus()
    Call Grid_Recebe_Foco(objGridRecebAntecipados)
End Sub

Public Sub GridRecebAntecipados_EnterCell()
    Call Grid_Entrada_Celula(objGridRecebAntecipados, iAlterado)
End Sub

Public Sub GridRecebAntecipados_LeaveCell()
    Call Saida_Celula(objGridRecebAntecipados)
End Sub

Public Sub GridRecebAntecipados_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridRecebAntecipados)
End Sub

Public Sub GridRecebAntecipados_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridRecebAntecipados, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridRecebAntecipados, iAlterado)
    End If

End Sub

Public Sub GridRecebAntecipados_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridRecebAntecipados)
End Sub

Public Sub GridRecebAntecipados_RowColChange()
    Call Grid_RowColChange(objGridRecebAntecipados)
End Sub

Public Sub GridRecebAntecipados_Scroll()
    Call Grid_Scroll(objGridRecebAntecipados)
End Sub

Public Sub GridDebitos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDebitos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDebitos, iAlterado)
    End If

End Sub

Public Sub GridDebitos_GotFocus()
    Call Grid_Recebe_Foco(objGridDebitos)
End Sub

Public Sub GridDebitos_EnterCell()
    Call Grid_Entrada_Celula(objGridDebitos, iAlterado)
End Sub

Public Sub GridDebitos_LeaveCell()
    Call Saida_Celula(objGridDebitos)
End Sub

Public Sub GridDebitos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridDebitos)
End Sub

Public Sub GridDebitos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDebitos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDebitos, iAlterado)
    End If

End Sub

Public Sub GridDebitos_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridDebitos)
End Sub

Public Sub GridDebitos_RowColChange()
    Call Grid_RowColChange(objGridDebitos)
End Sub

Public Sub GridDebitos_Scroll()
    Call Grid_Scroll(objGridDebitos)
End Sub

Public Sub Historico_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub MeioPagtoRA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRecebAntecipados)

End Sub

Public Sub MeioPagtoRA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRecebAntecipados)

End Sub

Public Sub MeioPagtoRA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRecebAntecipados.objControle = MeioPagtoRA
    lErro = Grid_Campo_Libera_Foco(objGridRecebAntecipados)
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
        
        'Torna Frame de Títulos visível
        Frame1(Opcao.SelectedItem.Index).Visible = True
    
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
    
        'Se Frame selecionado foi o de Títulos
        If Opcao.SelectedItem.Index = TAB_Titulos Then
        
            'iFramePrincipalAlterado = REGISTRO_ALTERADO
                    
        'Se Frame selecionado foi o de Parcelas
        ElseIf Opcao.SelectedItem.Index = TAB_Parcelas And iFramePrincipalAlterado <> 0 Then
        
            lErro = Carrega_Tab_Parcelas()
            If lErro <> SUCESSO Then Error 29015
            
            iFramePrincipalAlterado = 0

        End If

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual
        
            Case TAB_Titulos
                Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_TITULOS
                
            Case TAB_Parcelas
                Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_PARCELAS
            
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_CONTABILIZACAO
                        
        End Select
    
    End If

    Exit Sub

Erro_Opcao_Click:

    iFramePrincipalAlterado = REGISTRO_ALTERADO

    Select Case Err

        Case 29015
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143402)

    End Select

    Exit Sub

End Sub

Public Sub Recebimento_Click(Index As Integer)

Dim iCodigo As Integer
Dim iSubTipo As Integer
Dim lErro As Long

On Error GoTo Erro_Recebimento_Click

    'Se o Recebimento clicado é diferente do atual
    If iFrameRecebimentoAtual <> Index Then
    
        '#########################################
        'Inserido por Wagner
        lErro = Valida_Clientes_Diferentes()
        If lErro <> SUCESSO And lErro <> 140295 Then gError 140299
        
        'Se existe clientes diferentes e a baixa não é por recebimento, nem por perda
        If lErro <> SUCESSO And Index <> 0 And Index <> 3 Then
            'Não sai do Frame que está e gera erro
            Recebimento(iFrameRecebimentoAtual).Value = True
            gError 140300
        End If
        '#########################################

        'Exibe o grid correspondente ao Recebimento clicado
        FrameRecebimento(Index).Visible = True
        FrameRecebimento(iFrameRecebimentoAtual).Visible = False
        iFrameRecebimentoAtual = Index
        
        'Define o parâmetro que será passado para Contabil_Processa_Alteracao_Subtipo
        Select Case Index

            Case 0 'Recebimento
                iCodigo = MOTIVO_RECEBIMENTO

            Case 1 'Adiantamento
                iCodigo = MOTIVO_RECEBTO_ANTECIPADO

            Case 2 'Devoluções / Débitos
                iCodigo = MOTIVO_DEBITO_CLIENTE

            Case 3 'Perda
                iCodigo = MOTIVO_PERDA

            Case Else
                gError 79666

        End Select

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
    
        Case 79662, 79664, 140299
        
        Case 79666, 79678
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO", gErr)
    
        Case 140300 'Inserido por Wagner
            Call Rotina_Erro(vbOKOnly, "ERRO_PARC_CLIENTES_DIFERENTES", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143403)
    
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

Public Sub SaldoRA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRecebAntecipados)

End Sub

Public Sub SaldoRA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRecebAntecipados)

End Sub

Public Sub SaldoRA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRecebAntecipados.objControle = SaldoRA
    lErro = Grid_Campo_Libera_Foco(objGridRecebAntecipados)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub


Public Sub Selecionar_Click()
'Função alterado por Wagner

Dim lErro As Long

On Error GoTo Erro_Selecionar_Click

    lErro = Valida_Clientes_Diferentes()
    If lErro <> SUCESSO And lErro <> 140295 Then gError 140297
    
    If lErro <> SUCESSO Then
    
        'Se a baixa não é por recebimento nem por perda
        If Recebimento(0).Value = False And Recebimento(3).Value = False Then
            GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Baixar_Col) = "0"
            
            Call Grid_Refresh_Checkbox(objGridParcelas)
            
            gError 140298
        End If

    End If

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    
    Call Calcula_Total_Receber

    Exit Sub

Erro_Selecionar_Click:

    Select Case gErr
    
        Case 140297
        
        Case 140298
             Call Rotina_Erro(vbOKOnly, "ERRO_PARC_CLIENTES_DIFERENTES", gErr)
    
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Sub
    
End Sub

Public Sub SelecionarRA_Click()

Dim iLinha As Integer

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

    'Passa a linha atual para iLinha
    iLinha = GridRecebAntecipados.Row

    'Se SelecionarRA foi marcado
    If GridRecebAntecipados.TextMatrix(iLinha, iGrid_SelecionarRA_Col) = GRID_CHECKBOX_ATIVO Then

        'Percorre todas as Parcelas da Coleção
        For iLinha = 1 To objGridRecebAntecipados.iLinhasExistentes

            'Desmarca as outras Parcelas da Coleção
             If iLinha <> GridRecebAntecipados.Row Then GridRecebAntecipados.TextMatrix(iLinha, iGrid_SelecionarRA_Col) = GRID_CHECKBOX_INATIVO

        Next

    End If

    'Atualiza os CheckBox na tela
    Call Grid_Refresh_Checkbox(objGridRecebAntecipados)

End Sub

Public Sub SelecionarDB_Click()

Dim iLinha As Integer

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

    'Passa a linha atual para iLinha
    iLinha = GridDebitos.Row

    'Se SelecionarDB foi marcado
    If GridDebitos.TextMatrix(iLinha, iGrid_SelecionarDB_Col) = GRID_CHECKBOX_ATIVO Then

        'Percorre todas as Parcelas da Coleção
        For iLinha = 1 To objGridDebitos.iLinhasExistentes

            'Desmarca as outras Parcelas da Coleção
             If iLinha <> GridDebitos.Row Then GridDebitos.TextMatrix(iLinha, iGrid_SelecionarDB_Col) = GRID_CHECKBOX_INATIVO

        Next

    End If

    'Atualiza os CheckBox na tela
    Call Grid_Refresh_Checkbox(objGridDebitos)

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

Public Sub TituloFim_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_TituloFim_Validate

    'Se TituloFim e TituloInic estão preenchidos
    If Len(Trim(TituloFim.Text)) > 0 And Len(Trim(TituloInic.Text)) > 0 Then

        'Verifica se TituloFim é maior ou igual que TituloInic
        If CLng(Trim(TituloFim.Text)) < CLng(Trim(TituloInic.Text)) Then Error 29041

    End If

    Exit Sub

Erro_TituloFim_Validate:

    Cancel = True


    Select Case Err

        Case 29041
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143404)

    End Select

    Exit Sub

End Sub

Public Sub TituloInic_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub UpDownDataBaixa_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_DownClick

    'Diminui a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 29042

    Call Calcula_Multa_Juros_Desc_Parcelas

    Exit Sub

Erro_UpDownDataBaixa_DownClick:

    Select Case Err

        Case 29042

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143405)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataBaixa_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataBaixa_UpClick

    'Aumenta a DataBaixa em 1 dia
    lErro = Data_Up_Down_Click(DataBaixa, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 29043

    Call Calcula_Multa_Juros_Desc_Parcelas
    
    Exit Sub

Erro_UpDownDataBaixa_UpClick:

    Select Case Err

        Case 29043

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143406)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_DownClick

    'Diminui a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 29044

    Exit Sub

Erro_UpDownEmissaoFim_DownClick:

    Select Case Err

        Case 29044

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143407)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoFim_UpClick

    'Aumenta a data EmissaoFim em 1 dia
    lErro = Data_Up_Down_Click(EmissaoFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 29045

    Exit Sub

Erro_UpDownEmissaoFim_UpClick:

    Select Case Err

        Case 29045

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143408)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_DownClick

    'Diminui a data EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 29046

    Exit Sub

Erro_UpDownEmissaoInic_DownClick:

    Select Case Err

        Case 29046

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143409)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissaoInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoInic_UpClick

    'Aumenta a data EmissaoInic em 1 dia
    lErro = Data_Up_Down_Click(EmissaoInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 29047

    Exit Sub

Erro_UpDownEmissaoInic_UpClick:

    Select Case Err

        Case 29047

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143410)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_DownClick

    'Diminui a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 29048

    Exit Sub

Erro_UpDownVencFim_DownClick:

    Select Case Err

        Case 29048

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143411)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencFim_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencFim_UpClick

    'Aumenta a data VencFim em 1 dia
    lErro = Data_Up_Down_Click(VencFim, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 29049

    Exit Sub

Erro_UpDownVencFim_UpClick:

    Select Case Err

        Case 29049

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143412)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_DownClick

    'Diminui a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 29050

    Exit Sub

Erro_UpDownVencInic_DownClick:

    Select Case Err

        Case 29050

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143413)

    End Select

    Exit Sub

End Sub

Public Sub UpDownVencInic_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownVencInic_UpClick

    'Aumenta a data VencInic em 1 dia
    lErro = Data_Up_Down_Click(VencInic, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 29051

    Exit Sub

Erro_UpDownVencInic_UpClick:

    Select Case Err

        Case 29051

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143414)

    End Select

    Exit Sub

End Sub

Public Sub ValorAReceber_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorAReceber_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorAReceber_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorAReceber
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorBaixar_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorBaixar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorBaixar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorBaixar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorBaixar
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

Public Sub Selecionar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Selecionar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Selecionar_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Selecionar
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub SelecionarRA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRecebAntecipados)

End Sub

Public Sub SelecionarRA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRecebAntecipados)

End Sub

Public Sub SelecionarRA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRecebAntecipados.objControle = SelecionarRA
    lErro = Grid_Campo_Libera_Foco(objGridRecebAntecipados)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39694

        If objGridInt Is objGridParcelas Then

            Select Case GridParcelas.Col

                Case iGrid_ValorBaixar_Col

                    lErro = Saida_Celula_ValorBaixar(objGridInt)
                    If lErro <> SUCESSO Then Error 29052

                Case iGrid_ValorDesconto_Col

                    lErro = Saida_Celula_ValorDesconto(objGridInt)
                    If lErro <> SUCESSO Then Error 29053

                Case iGrid_ValorMulta_Col

                    lErro = Saida_Celula_ValorMulta(objGridInt)
                    If lErro <> SUCESSO Then Error 29054

                Case iGrid_ValorJuros_Col

                    lErro = Saida_Celula_ValorJuros(objGridInt)
                    If lErro <> SUCESSO Then Error 29055

            End Select

        ElseIf objGridInt Is objGridRecebAntecipados Then

            Select Case GridRecebAntecipados.Col
                    
            End Select
    
        ElseIf objGridInt Is objGridDebitos Then

            Select Case GridDebitos.Col

            End Select
        
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 29058

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 29052, 29053, 29054, 29055, 29056, 29057

        Case 29058
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 39694
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143415)

    End Select

    Exit Function

End Function

Private Function ContaCorrente_Carrega(objComboBox As ComboBox) As Long

Dim lErro As Long
Dim colCodigoNomeRed As New AdmColCodigoNome
Dim objCodigoNomeRed As AdmCodigoNome

On Error GoTo Erro_ContaCorrente_Carrega

    'Lê Codigos, NomesReduzidos de ContasCorrentes
    lErro = CF("ContasCorrentesInternas_Le_CodigosNomesRed", colCodigoNomeRed)
    If lErro <> SUCESSO Then Error 6758

    'Preeche list de ComboBox
    For Each objCodigoNomeRed In colCodigoNomeRed
        objComboBox.AddItem CStr(objCodigoNomeRed.iCodigo) & SEPARADOR & objCodigoNomeRed.sNome
        objComboBox.ItemData(objComboBox.NewIndex) = objCodigoNomeRed.iCodigo
    Next

    ContaCorrente_Carrega = SUCESSO

    Exit Function

Erro_ContaCorrente_Carrega:

    ContaCorrente_Carrega = Err

    Select Case Err

        Case 6758  'já tratado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143416)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iSubTipo As Integer
Dim colCodigoDescricao As AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    'Visibilidade para versão LIGHT
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        LabelFilial.left = POSICAO_FORA_TELA
'        Filial.TabStop = False
'        Filial.left = POSICAO_FORA_TELA
'        FilialEmpresa.left = POSICAO_FORA_TELA
'        FilialEmpresa.TabStop = False
'        Cobrador.left = POSICAO_FORA_TELA
'        Cobrador.TabStop = False
'        FilialDB.left = POSICAO_FORA_TELA
'        FilialDB.TabStop = False
'        Recebimento.Item(1).left = POSICAO_FORA_TELA
'        Recebimento.Item(1).TabStop = False
'
'        Recebimento.Item(2).left = 2660
'        Recebimento.Item(3).left = 5200
'
'    End If
                    
    iFrameAtual = 1
    iFrameRecebimentoAtual = 0
    iFramePrincipalAlterado = 0
    dtDataMultaJuros = DATA_NULA
    
    Set objGridParcelas = New AdmGrid
    Set objGridRecebAntecipados = New AdmGrid
    Set objGridDebitos = New AdmGrid
    Set objEventoClientes = New AdmEvento
    Set objEventoVendedor = New AdmEvento

    Set objEventoCodigoViagem = New AdmEvento
    
    Set gcolInfoParcRec = New Collection
        
    'Carrega list de ComboBox ContaCorrente
    lErro = ContaCorrente_Carrega(ContaCorrente)
    If lErro <> SUCESSO Then gError 29133
    
    lErro = Carrega_TipoDocumento(TipoDocSeleciona)
    If lErro <> SUCESSO Then gError 185035

    FormaPagto.AddItem ""
    FormaPagto.ItemData(FormaPagto.NewIndex) = 0
    
    Set colCodigoDescricao = New AdmColCodigoNome
    lErro = CF("FormasPagamento_Le_CodNome", colCodigoDescricao, True)
    If lErro <> SUCESSO Then gError 29133
    
    'Preenche ListBox Condições com DescReduzidas de CondicoesPagto
    For Each objCodigoDescricao In colCodigoDescricao
        
        FormaPagto.AddItem objCodigoDescricao.sNome
        FormaPagto.ItemData(FormaPagto.NewIndex) = objCodigoDescricao.iCodigo
        
    Next
    
    'Executa a Inicialização do grid Parcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 29135

    'Executa a Inicialização do grid Recebimentos Antecipados
    lErro = Inicializa_Grid_RecebAntecipados(objGridRecebAntecipados)
    If lErro <> SUCESSO Then gError 29136

    'Executa a Inicialização do grid Debitos
    lErro = Inicializa_Grid_Debitos(objGridDebitos)
    If lErro <> SUCESSO Then gError 29137
    
    'Obtem o código do subtipo que terá seus modelos carregados
    lErro = CF("Obtem_Subtipo_Contabil", "MotivosBaixa", MOTIVO_RECEBIMENTO, iSubTipo)
    If lErro <> SUCESSO And lErro <> 79671 Then gError 79665
    
    'Se não encontrou o subtipo => erro
    If lErro = 79671 Then gError 79672
    
    'Inicializa a contabilidade carregando os modelos do subtipo encontrado
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASARECEBER, iSubTipo)
    If lErro <> SUCESSO Then gError 79667
    
    lErro = CF2(Me, "BR_Form_Load")
    If lErro <> SUCESSO Then gError 182904
    
    lErro = Carrega_Cobrador()
    If lErro <> SUCESSO Then gError 182904
    
    Set objEventoProdutoDe = New AdmEvento
    Set objEventoProdutoAte = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoInicial)
    If lErro <> SUCESSO Then gError 185035

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoFinal)
    If lErro <> SUCESSO Then gError 185035
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 29133, 29135, 29136, 29137, 39693, 79665, 79667, 182904, 185035
        
        Case 79672
            Call Rotina_Erro(vbOKOnly, "ERRO_SUBTIPOCONTABIL_TIPOBAIXA_NAO_ENCONTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143417)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
    Inicializa_Grid_Parcelas = CF2(Me, "BR_Inicializa_Grid_Parcelas", objGridInt)
End Function

Public Function BR_Inicializa_Grid_Parcelas(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Parcelas

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Baixar")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Nosso Número")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Valor Baixar")
    objGridInt.colColuna.Add ("Desconto")
    objGridInt.colColuna.Add ("Multa")
    objGridInt.colColuna.Add ("Juros")
    objGridInt.colColuna.Add ("Valor Recebido")
    
    '########################################
    'Inserido por Wagner
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Filial")
    '########################################
    
    'Visibilidade para versão LIGHT
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Cobrador")
'    End If
    objGridInt.colColuna.Add ("Valor Parcela")
    
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
'    End If
    
    'campos de edição do grid
    objGridInt.colCampo.Add (Selecionar.Name)
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (NossoNumero.Name)
    objGridInt.colCampo.Add (Emissao.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Saldo.Name)
    objGridInt.colCampo.Add (ValorBaixar.Name)
    objGridInt.colCampo.Add (ValorDesconto.Name)
    objGridInt.colCampo.Add (ValorMulta.Name)
    objGridInt.colCampo.Add (ValorJuros.Name)
    objGridInt.colCampo.Add (ValorAReceber.Name)
    
    '######################################
    'Inserido por Wagner
    objGridInt.colCampo.Add (ClienteItem.Name)
    objGridInt.colCampo.Add (FilialClienteItem.Name)
    '######################################
    
    'Visibilidade para versão LIGHT
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (Cobrador.Name)
'    End If
    objGridInt.colCampo.Add (ValorParcela.Name)
    
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (FilialEmpresa.Name)
'    End If
    
    iGrid_Baixar_Col = 1
    iGrid_DataVencimento_Col = 2
    iGrid_Tipo_Col = 3
    iGrid_Numero_Col = 4
    iGrid_NossoNumero_Col = 5
    iGrid_Emissao_Col = 6
    iGrid_Parcela_Col = 7
    iGrid_Saldo_Col = 8
    iGrid_ValorBaixar_Col = 9
    iGrid_ValorDesconto_Col = 10
    iGrid_ValorMulta_Col = 11
    iGrid_ValorJuros_Col = 12
    iGrid_ValorAReceber_Col = 13
    iGrid_Cliente_Col = 14
    iGrid_Filial_Col = 15
    
    'Visibilidade para versão LIGHT
'    If giTipoVersao = VERSAO_FULL Then
        iGrid_Cobrador_Col = 16
        iGrid_ValorParcela_Col = 17
        iGrid_FilialEmpresa_Col = 18
'    ElseIf giTipoVersao = VERSAO_LIGHT Then
'        iGrid_ValorParcela_Col = 16
'    End If
    
    objGridInt.objGrid = GridParcelas

    'todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_PARCELAS_BAIXA + 1

    'linhas visiveis do grid
    If m_objUserControl.Height > 9000 Then
        objGridInt.iLinhasVisiveis = 20
    Else
        objGridInt.iLinhasVisiveis = 7
    End If
    
    'largura da primeira coluna
    GridParcelas.ColWidth(0) = 600

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'incluir barra de rolagem horizontal
    objGridInt.iIncluirHScroll = GRID_INCLUIR_HSCROLL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    BR_Inicializa_Grid_Parcelas = SUCESSO
    
    Exit Function

End Function

Private Function Inicializa_Grid_RecebAntecipados(objGridInt As AdmGrid) As Long
    Inicializa_Grid_RecebAntecipados = CF2(Me, "BR_Inicializa_Grid_RecebAntecipados", objGridInt)
End Function

Public Function BR_Inicializa_Grid_RecebAntecipados(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Recebimentos Antecipados

    'tela em questão
    Set objGridRecebAntecipados.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Data Movimento")
    objGridInt.colColuna.Add ("Conta Corrente")
    objGridInt.colColuna.Add ("Meio Pagamento")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")
    objGridInt.colColuna.Add ("Doc")
    objGridInt.colColuna.Add ("Seq")
    objGridInt.colColuna.Add ("Histórico")
    objGridInt.colColuna.Add ("Filial Empresa")

   'campos de edição do grid
    objGridInt.colCampo.Add (SelecionarRA.Name)
    objGridInt.colCampo.Add (DataMovimentoRA.Name)
    objGridInt.colCampo.Add (ContaCorrenteRA.Name)
    objGridInt.colCampo.Add (MeioPagtoRA.Name)
    objGridInt.colCampo.Add (ValorRA.Name)
    objGridInt.colCampo.Add (SaldoRA.Name)
    objGridInt.colCampo.Add (Doc.Name)
    objGridInt.colCampo.Add (Seq.Name)
    objGridInt.colCampo.Add (Hist.Name)
    objGridInt.colCampo.Add (FilialRA.Name)

    iGrid_SelecionarRA_Col = 1
    iGrid_DataMovimentoRA_Col = 2
    iGrid_ContaCorrenteRA_Col = 3
    iGrid_MeioPagtoRA_Col = 4
    iGrid_ValorRA_Col = 5
    iGrid_SaldoRA_Col = 6
    iGrid_Doc_Col = 7
    iGrid_Seq_Col = 8
    iGrid_Hist_Col = 9
    iGrid_FilialRA_Col = 10

    objGridInt.objGrid = GridRecebAntecipados

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 5 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 3

    'largura da primeira coluna
    GridRecebAntecipados.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    BR_Inicializa_Grid_RecebAntecipados = SUCESSO
    
    Exit Function

End Function

Private Function Inicializa_Grid_Debitos(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Debitos

    'tela em questão
    Set objGridDebitos.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Data Emissão")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Saldo")
        
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
'    End If
    objGridInt.colColuna.Add ("Observação")
    
    'campos de edição do grid
    objGridInt.colCampo.Add (SelecionarDB.Name)
    objGridInt.colCampo.Add (DataEmissaoDB.Name)
    objGridInt.colCampo.Add (TipoDocDB.Name)
    objGridInt.colCampo.Add (NumTituloDB.Name)
    objGridInt.colCampo.Add (ValorDB.Name)
    objGridInt.colCampo.Add (SaldoDB.Name)
        
'    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (FilialDB.Name)
'    End If
    objGridInt.colCampo.Add (OBSDB.Name)
    
    iGrid_SelecionarDB_Col = 1
    iGrid_DataEmissaoDB_Col = 2
    iGrid_TipoDocDB_Col = 3
    iGrid_NumTituloDB_Col = 4
    iGrid_ValorDB_Col = 5
    iGrid_SaldoDB_Col = 6
        
'    If giTipoVersao = VERSAO_FULL Then
        iGrid_FilialDB_Col = 7
'    End If
    iGrid_OBSDB_Col = 8
    
    objGridInt.objGrid = GridDebitos

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 5 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 4

    'largura da primeira coluna
    GridDebitos.ColWidth(0) = 300

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas nem excluir as existentes
    objGridInt.iProibidoIncluir = 1
    objGridInt.iProibidoExcluir = 1

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Debitos = SUCESSO
    
    Exit Function

End Function

Function Trata_Parametros(Optional objBaixaReceber As ClassBaixaReceber) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143418)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub ValorRA_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridRecebAntecipados)

End Sub

Public Sub ValorRA_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridRecebAntecipados)

End Sub

Public Sub ValorRA_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridRecebAntecipados.objControle = ValorRA
    lErro = Grid_Campo_Libera_Foco(objGridRecebAntecipados)
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
        If lErro <> SUCESSO Then Error 29080

        'Se a data vencInic está preenchida
        If Len(VencInic.ClipText) > 0 Then

            'Verifica se a VencFim é maior ou igual a VencInic
            If CDate(VencFim.Text) < CDate(VencInic.Text) Then Error 29081

        End If

    End If

    Exit Sub

Erro_VencFim_Validate:

    Cancel = True


    Select Case Err

        Case 29080

        Case 29081
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143419)

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
        If lErro <> SUCESSO Then Error 29082

    End If

    Exit Sub

Erro_VencInic_Validate:

    Cancel = True


    Select Case Err

        Case 29082

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143420)

    End Select

    Exit Sub

End Sub

Private Function Grid_Parcelas_Preenche(colInfoParcRec As Collection) As Long
'Preenche o Grid Parcelas com os dados de gcolInfoParcRec

Dim iLinha As Integer
Dim objInfoParcRec As ClassInfoParcRec
Dim objFilialEmpresa As New AdmFiliais
Dim objCobrador As New ClassCobrador
Dim lErro As Long
'#########################
'Inserido por Wagner
Dim objCliente As ClassCliente
Dim objFilialCliente As ClassFilialCliente
'#########################

On Error GoTo Erro_Grid_Parcelas_Preenche

    'Se o número de parcelas for maior que o número de linhas do Grid
    If colInfoParcRec.Count + 1 > GridParcelas.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de parcelas
        GridParcelas.Rows = colInfoParcRec.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridParcelas)

    End If

    iLinha = 0
    
    'Percorre todas as Parcelas da Coleção
    For Each objInfoParcRec In colInfoParcRec

        iLinha = iLinha + 1

        'Passa para a tela os dados da Parcela em questão
        GridParcelas.TextMatrix(iLinha, iGrid_DataVencimento_Col) = objInfoParcRec.dtVencimento
        GridParcelas.TextMatrix(iLinha, iGrid_Tipo_Col) = objInfoParcRec.sSiglaDocumento
        GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col) = objInfoParcRec.lNumTitulo
        If iGrid_NossoNumero_Col <> 0 Then GridParcelas.TextMatrix(iLinha, iGrid_NossoNumero_Col) = right(objInfoParcRec.sNossoNumero, 10)
        If iGrid_Emissao_Col <> 0 Then GridParcelas.TextMatrix(iLinha, iGrid_Emissao_Col) = Format(objInfoParcRec.dtDataEmissao, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col) = objInfoParcRec.iNumParcela
        GridParcelas.TextMatrix(iLinha, iGrid_Saldo_Col) = Format(objInfoParcRec.dValor, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixar_Col) = Format(objInfoParcRec.dValor, "Standard")
        GridParcelas.TextMatrix(iLinha, iGrid_ValorAReceber_Col) = Format(objInfoParcRec.dValor, "Standard")
        
        '#################################################
        'Inserido por Wagner
        Set objCliente = New ClassCliente
        Set objFilialCliente = New ClassFilialCliente
        
        objCliente.lCodigo = objInfoParcRec.lCliente
        
        objFilialCliente.lCodCliente = objCliente.lCodigo
        objFilialCliente.iCodFilial = objInfoParcRec.iFilialCliente
            
        'le o nome reduzido do cliente
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 140296
            
        'le o nome reduzido da filial  cliente
        lErro = CF("FilialCliente_Le", objFilialCliente)
        If lErro <> SUCESSO And lErro <> 12567 Then gError 140297
            
        GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col) = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
        GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col) = objInfoParcRec.iFilialCliente & SEPARADOR & objFilialCliente.sNome
        '#################################################
        
        'preenche o objFilialEmpresa
        objFilialEmpresa.iCodFilial = objInfoParcRec.iFilialEmpresa
        
        'le o Nome da Filial
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 49722
                
'        If giTipoVersao = VERSAO_FULL Then
            GridParcelas.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
'        End If
        
        GridParcelas.TextMatrix(iLinha, iGrid_Baixar_Col) = objInfoParcRec.iMarcada
        GridParcelas.TextMatrix(iLinha, iGrid_ValorParcela_Col) = Format(objInfoParcRec.dValorOriginal, "Standard")
        
        'Preenche o objCobrador
        objCobrador.iCodigo = objInfoParcRec.iCobrador
        
        'le o nomeReduzido do cobrador
        lErro = CF("Cobrador_Le", objCobrador)
        If lErro <> SUCESSO And lErro <> 19294 Then gError 49723
        
        If lErro = 19294 Then gError 49728
        
'        If giTipoVersao = VERSAO_FULL Then
            GridParcelas.TextMatrix(iLinha, iGrid_Cobrador_Col) = objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
'        End If
        
        lErro = CF2(Me, "BR_Traz_Parcela_Tela", objInfoParcRec.lNumIntParc, iLinha)
        If lErro <> SUCESSO Then gError 182832
        
    Next

    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGridParcelas.iLinhasExistentes = gcolInfoParcRec.Count

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridParcelas)
    
    Grid_Parcelas_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_Parcelas_Preenche:

    Grid_Parcelas_Preenche = gErr
    
    Select Case gErr
    
        Case 49722, 49723, 140296, 140297, 182832 'Inserido por Wagner
        
        Case 49728
            lErro = Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_NAO_ENCONTRADO", gErr, CStr(objCobrador.iCodigo))
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143421)

    End Select

    Exit Function

End Function

Private Function Grid_RecebAntecipados_Preenche(colRecebAntecipado As Collection) As Long
'Preenche o Grid Recebimentos Antecipados com os dados de colRecebAntecipado

Dim iLinha As Integer
Dim objRecebAntecipados As ClassAntecipReceb
Dim objFilialEmpresa As New AdmFiliais
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim lErro As Long

On Error GoTo Erro_Grid_RecebAntecipados_Preenche

    'Se o número de antecipacoes for maior que o número de linhas do Grid
    If colRecebAntecipado.Count + 1 > GridRecebAntecipados.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de antecipacoes
        GridRecebAntecipados.Rows = colRecebAntecipado.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridRecebAntecipados)

    End If

    iLinha = 0
    
    'Percorre todos os Recebimentos antecipados da Coleção
    For Each objRecebAntecipados In colRecebAntecipado

        iLinha = iLinha + 1

        'Passa para a tela os dados do Recebimento antecipado em questão
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_DataMovimentoRA_Col) = objRecebAntecipados.dtData
        
        'Lê a Conta Corrente Interna para pegar o nome reduzido
        lErro = CF("ContaCorrenteInt_Le", objRecebAntecipados.iCodConta, objContaCorrenteInt)
        If lErro <> SUCESSO And lErro <> 11807 Then gError 49730
        
        If lErro = 11807 Then gError 49736
        
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_ContaCorrenteRA_Col) = objRecebAntecipados.iCodConta & SEPARADOR & objContaCorrenteInt.sNomeReduzido
        
        'preenche o objTipoMeioPagto
        objTipoMeioPagto.iTipo = objRecebAntecipados.iTipoMeioPagto
        
        'Lê o Tipo Meio de Pagamento
        lErro = CF("TipoMeioPagto_Le", objTipoMeioPagto)
        If lErro <> SUCESSO And lErro <> 11909 Then gError 49731
        
        'Se não encontrou o Tipo Meio Pagto --> Erro
        If lErro <> SUCESSO Then gError 49732
        
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_MeioPagtoRA_Col) = objRecebAntecipados.iTipoMeioPagto & SEPARADOR & objTipoMeioPagto.sDescricao
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_ValorRA_Col) = Format(objRecebAntecipados.dValor, "Standard")
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_SaldoRA_Col) = Format(objRecebAntecipados.dSaldoNaoApropriado, "Standard")
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_SelecionarRA_Col) = GRID_CHECKBOX_INATIVO
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_Doc_Col) = objRecebAntecipados.sNumRefExterna
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_Seq_Col) = objRecebAntecipados.lSequencial
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_Hist_Col) = objRecebAntecipados.sHistorico
        
        'preenche o objFilialEmpresa
        objFilialEmpresa.iCodFilial = objRecebAntecipados.iFilialEmpresa
        
        'Le o nome da filial
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 49733
                
        If lErro = 27378 Then gError 49734
        
        GridRecebAntecipados.TextMatrix(iLinha, iGrid_FilialRA_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
        
        lErro = CF2(Me, "BR_Traz_RA_Tela", objRecebAntecipados, iLinha)
        If lErro <> SUCESSO Then gError 182874

    Next

    'Passa para o Obj o número de Recebimentos antecipados passados pela Coleção
    objGridRecebAntecipados.iLinhasExistentes = colRecebAntecipado.Count

    
    'Modifica a cor da Label adiantamento para verde
    If objGridRecebAntecipados.iLinhasExistentes <> 0 Then Recebimento(1).ForeColor = vbGreen

    Call Grid_Refresh_Checkbox(objGridRecebAntecipados)
    
    Grid_RecebAntecipados_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_RecebAntecipados_Preenche:

    Grid_RecebAntecipados_Preenche = gErr
    
    Select Case gErr
    
        Case 49730, 49731, 182874
        
        Case 49732
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", gErr, objTipoMeioPagto.iTipo)

        Case 49736
            Call Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", gErr, objContaCorrenteInt.iCodigo)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143422)

    End Select

    Exit Function

End Function

Private Function Grid_Debitos_Preenche(colDebitos As Collection) As Long
'Preenche o Grid Débitos com os dados de colDebitos

Dim iLinha As Integer
Dim objDebitos As ClassDebitoRecCli
Dim objFilialEmpresa As New AdmFiliais
Dim lErro As Long

On Error GoTo Erro_Grid_Debitos_Preenche

    'Se o número de debitos for maior que o número de linhas do Grid
    If colDebitos.Count + 1 > GridDebitos.Rows Then
    
        'Altera o número de linhas do Grid de acordo com o número de debitos
        GridDebitos.Rows = colDebitos.Count + 1
        
        'Chama rotina de Inicialização do Grid
        Call Grid_Inicializa(objGridDebitos)

    End If

    iLinha = 0

    'Percorre todos os Débitos da Coleção
    For Each objDebitos In colDebitos

        iLinha = iLinha + 1

        'Passa para a tela os dados do Débito em questão
        GridDebitos.TextMatrix(iLinha, iGrid_DataEmissaoDB_Col) = IIf(objDebitos.dtDataEmissao <> DATA_NULA, objDebitos.dtDataEmissao, "")
        GridDebitos.TextMatrix(iLinha, iGrid_TipoDocDB_Col) = objDebitos.sSiglaDocumento
        GridDebitos.TextMatrix(iLinha, iGrid_NumTituloDB_Col) = objDebitos.lNumTitulo
        GridDebitos.TextMatrix(iLinha, iGrid_ValorDB_Col) = Format(objDebitos.dValorTotal, "Standard")
        GridDebitos.TextMatrix(iLinha, iGrid_SaldoDB_Col) = Format(objDebitos.dSaldo, "Standard")
        GridDebitos.TextMatrix(iLinha, iGrid_OBSDB_Col) = objDebitos.sObservacao
        GridDebitos.TextMatrix(iLinha, iGrid_SelecionarDB_Col) = GRID_CHECKBOX_INATIVO
        
        'preenche o objFilialEmpresa
        objFilialEmpresa.iCodFilial = objDebitos.iFilialEmpresa
        
        'Le o nome da filial
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then Error 49735
        
'        If giTipoVersao = VERSAO_FULL Then
            GridDebitos.TextMatrix(iLinha, iGrid_FilialDB_Col) = objDebitos.iFilialEmpresa & SEPARADOR & objFilialEmpresa.sNome
'        End If
        
    Next

    'Passa para o Obj o número de Débitos passados pela Coleção
    objGridDebitos.iLinhasExistentes = colDebitos.Count
    
    'Modifica a cor da Label Débitos / Devoluções para verde
    If objGridDebitos.iLinhasExistentes <> 0 Then Recebimento(2).ForeColor = vbGreen

    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGridDebitos)
    
    Grid_Debitos_Preenche = SUCESSO
    
    Exit Function
    
Erro_Grid_Debitos_Preenche:

    Grid_Debitos_Preenche = Err
    
    Select Case Err
        
        Case 49735
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143423)

    End Select

    Exit Function

End Function

Private Function Frame_Parcelas_Limpa()
'Limpa os comandos do Frame Parcelas

    'Limpa os Grids
    Call Grid_Limpa(objGridRecebAntecipados)
    objGridRecebAntecipados.iLinhasExistentes = 0
    Call Grid_Limpa(objGridParcelas)
    objGridParcelas.iLinhasExistentes = 0
    Call Grid_Limpa(objGridDebitos)
    objGridDebitos.iLinhasExistentes = 0

    'Limpa os MaskEdBox
    ValorReceber.Caption = ""
    TotalBaixar.Caption = Format(0, "Fixed")
    Historico.Text = ""
    DataBaixa = Format(gdtDataAtual, "dd/mm/yy")

    'Deixa deselecionado
    ContaCorrente.ListIndex = -1

    'Modifica a cor da label pagamento(1) e pagamento(2)
    Recebimento(1).ForeColor = vbButtonText
    Recebimento(2).ForeColor = vbButtonText
    
    Call Ordenacao_Limpa(objGridParcelas)
    
End Function

Private Function Frame_Titulos_Limpa()
'Limpa os comandos do Frame Titulos

    'Limpa os MaskEdBox
    Cliente.Text = ""
    EmissaoInic.PromptInclude = False
    EmissaoInic.Text = ""
    EmissaoInic.PromptInclude = True
    EmissaoFim.PromptInclude = False
    EmissaoFim.Text = ""
    EmissaoFim.PromptInclude = True
    VencInic.PromptInclude = False
    VencInic.Text = ""
    VencInic.PromptInclude = True
    VencFim.PromptInclude = False
    VencFim.Text = ""
    VencFim.PromptInclude = True
    TituloInic.PromptInclude = False
    TituloInic.Text = ""
    TituloInic.PromptInclude = True
    TituloFim.PromptInclude = False
    TituloFim.Text = ""
    TituloFim.PromptInclude = True

    'Deixa a ComboBox sem ítem selecionado
    Filial.ListIndex = -1

End Function

Private Function Saida_Celula_ValorBaixar(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorBaixar As Double
Dim dValorDesconto As Double
Dim dValorSaldo As Double

On Error GoTo Erro_Saida_Celula_ValorBaixar

    Set objGridInt.objControle = ValorBaixar

    'Se ValorBaixar está preenchido
    If Len(Trim(ValorBaixar.Text)) <> 0 Then

        'Verifica se ValorBaixar é válido
        lErro = Valor_Positivo_Critica(ValorBaixar.Text)
        If lErro <> SUCESSO Then Error 29083

        'Lê os dados da tela
        dValorBaixar = CDbl(ValorBaixar.Text)
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col))) <> 0 Then dValorDesconto = CDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col))
        dValorSaldo = CDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Saldo_Col))

        'Verifica se Saldo é maior que o ValorBaixar
        If dValorBaixar > dValorSaldo Then Error 29084

        'Verifica se Desconto é maior ou igual ao ValorBaixar
        If dValorBaixar <= dValorDesconto Then Error 29085

        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, ValorBaixar.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col))
        If lErro <> SUCESSO Then Error 56769
                
        'Formata o Valor na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col) = Format(dValorBaixar, "Standard")

    Else
    
        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, ValorBaixar.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col))
        If lErro <> SUCESSO Then Error 56770
                
    End If
    
    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Baixar_Col)) = GRID_CHECKBOX_ATIVO Then Call Calcula_Total_Receber

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 29086

    Saida_Celula_ValorBaixar = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBaixar:

    Saida_Celula_ValorBaixar = Err

    Select Case Err

        Case 29083, 29086, 56769, 56770
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 29084
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_MAIOR_SALDO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 29085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_MENOR_IGUAL_DESCONTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143424)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorDesconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dValorBaixar As Double
Dim dValorDesconto As Double

On Error GoTo Erro_Saida_Celula_ValorDesconto

    Set objGridParcelas.objControle = ValorDesconto

    'Se Desconto está preenchido
    If Len(Trim(ValorDesconto.Text)) <> 0 Then

        'Verifica se o Desconto é válido
        lErro = Valor_NaoNegativo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then Error 29087

        'Lê os dados da tela
        dValorDesconto = CDbl(ValorDesconto.Text)
        If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col))) <> 0 Then dValorBaixar = CDbl(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col))

        'Verifica se o Desconto é maior ou igual ao ValorBaixado
        If dValorDesconto >= dValorBaixar Then Error 29088

        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), ValorDesconto.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col))
        If lErro <> SUCESSO Then Error 56771
        
        'Formata o Desconto na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col) = Format(dValorDesconto, "Standard")

    Else
    
        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), ValorDesconto.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col))
        If lErro <> SUCESSO Then Error 56772
        
    End If
    
    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Baixar_Col)) = 1 Then Call Calcula_Total_Receber

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 29089

    Saida_Celula_ValorDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorDesconto:

    Saida_Celula_ValorDesconto = Err

    Select Case Err

        Case 29087, 29089, 56771, 56772
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 29088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_MENOR_IGUAL_DESCONTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143425)

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
        If lErro <> SUCESSO Then Error 29090

        If Len(Trim(DataBaixa.Text)) = 0 Then Error 56775
        
        'Lê os dados da tela
        dValorMulta = CDbl(ValorMulta.Text)
        
        dtDataBaixa = CDate(DataBaixa.Text)
        dtDataVencimento = CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col))

        If gobjCRFAT.iPermiteMultaSemAtraso <> MARCADO Then

            'Verifica se a Multa é maior que zero quando a DataBaixa é menor ou igual à Data Vencimento
            If dtDataBaixa <= dtDataVencimento And dValorMulta > 0 Then Error 29091

        End If

        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col), ValorMulta.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col))
        If lErro <> SUCESSO Then Error 56776
        
        'Formata a Multa na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col) = Format(dValorMulta, "Standard")

    Else
    
        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col), ValorMulta.Text, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col))
        If lErro <> SUCESSO Then Error 56777
    
    End If

    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Baixar_Col)) = 1 Then Call Calcula_Total_Receber

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 29092

    Saida_Celula_ValorMulta = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorMulta:

    Saida_Celula_ValorMulta = Err

    Select Case Err

        Case 56775
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_JUROS_DATABAIXA_VAZIA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 29090, 29092, 56776, 56777
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 29091
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_INCOMPATIVEL_DATABAIXA_DATAVENC", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143426)

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
        If lErro <> SUCESSO Then Error 29093

        If Len(Trim(DataBaixa.Text)) = 0 Then Error 56778
        
        'Lê os dados da tela
        dValorJuros = CDbl(ValorJuros.Text)
        
        dtDataBaixa = CDate(DataBaixa.Text)
        dtDataVencimento = CDate(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DataVencimento_Col))

        If gobjCRFAT.iPermiteMultaSemAtraso <> MARCADO Then
        
            'Verifica se a Multa é maior que zero quando a DataBaixa é menor ou igual à Data Vencimento
            If dtDataBaixa <= dtDataVencimento And dValorJuros > 0 Then Error 56779

        End If

        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col), ValorJuros.Text)
        If lErro <> SUCESSO Then Error 56773
                
        'Formata os Juros na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorJuros_Col) = Format(dValorJuros, "Standard")

    Else
    
        lErro = Parcela_RecalculaValorReceber(GridParcelas.Row, GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorDesconto_Col), GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorMulta_Col), ValorJuros.Text)
        If lErro <> SUCESSO Then Error 56774
                
    End If
    
    If StrParaInt(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Baixar_Col)) = 1 Then Call Calcula_Total_Receber

    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 29094

    Saida_Celula_ValorJuros = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorJuros:

    Saida_Celula_ValorJuros = Err

    Select Case Err

        Case 29093, 29094, 56773, 56774
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 56778
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_JUROS_DATABAIXA_VAZIA", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 56779
            lErro = Rotina_Erro(vbOKOnly, "ERRO_JUROS_INCOMPATIVEL_DATABAIXA_DATAVENC", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143427)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dtDataBaixa As Date
Dim iParcelaSelecionada As Integer
Dim objInfoParcRec As ClassInfoParcRec
Dim colBaixaParcReceber As New ColBaixaParcRec
Dim dValorRecebidoParcela As Double, dValorParcelaBaixa As Double
Dim dSomaReceber As Double
Dim dtDataVencimento  As Date
Dim dValorMulta  As Double
Dim objFilialCliente As New ClassFilialCliente
Dim vbResult As VbMsgBoxResult

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Cliente e Filial estão preenchidos e obtem dados
    lErro = FilialCliente_ObtemDaTela(objFilialCliente)
    If lErro <> SUCESSO Then gError 59489
                
    'Verifica se DataBaixa está preenchida
    If Len(DataBaixa.ClipText) = 0 Then gError 29101
    
    If Len(Trim(DataCredito.ClipText)) = 0 Then gError 39015
        
    If StrParaDate(DataBaixa.ClipText) <> StrParaDate(DataCredito.ClipText) Then
        vbResult = Rotina_Aviso(vbYesNo, "AVISO_DATA_CREDITO_DIF_DATA_BAIXA")
        If vbResult = vbNo Then gError ERRO_SEM_MENSAGEM
    End If
        
    'Lê a DataBaixa na tela
    dtDataBaixa = CDate(DataBaixa.Text)

    iParcelaSelecionada = 0

    'Percorre todas as linhas do Grid Parcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcRec = gcolInfoParcRec.Item(iIndice)

        'Se a checkbox Selecionada está marcada
        If GridParcelas.TextMatrix(iIndice, iGrid_Baixar_Col) = GRID_CHECKBOX_ATIVO Then

            'Conta quantas Parcelas estão selecionadas
            iParcelaSelecionada = iParcelaSelecionada + 1

            'Verifica se Valor a Baixar está preenchido
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorBaixar_Col))) = 0 Then gError 29102
            
            'Verifica a integridade da data em relacao a multa
            dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_DataVencimento_Col))
            'Pega o valor da multa
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorMulta_Col))) <> 0 Then
                dValorMulta = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorMulta_Col))
            Else
                dValorMulta = 0
            End If
            
            If gobjCRFAT.iPermiteMultaSemAtraso <> MARCADO Then

                'Verifica se a Multa é maior que zero quando a DataBaixa é menor ou igual à Data Vencimento
                If dtDataBaixa <= dtDataVencimento And dValorMulta > 0 Then Error 49554
            
            End If
            
            dValorParcelaBaixa = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorBaixar_Col))
            dValorRecebidoParcela = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorAReceber_Col))

            'Faz o somatório do Valor total a receber
            dSomaReceber = Round(dSomaReceber + dValorRecebidoParcela, 2)

            'Passa os dados da tela para o Obj
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorDesconto_Col))) <> 0 Then objInfoParcRec.dValorDesconto = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorDesconto_Col))
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorMulta_Col))) <> 0 Then objInfoParcRec.dValorMulta = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorMulta_Col))
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorJuros_Col))) <> 0 Then objInfoParcRec.dValorJuros = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorJuros_Col))
            
'            If giTipoVersao = VERSAO_FULL Then
                objInfoParcRec.iCobrador = StrParaInt(Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobrador_Col)))
'            End If
            
            Call colBaixaParcReceber.Add(0, 0, objInfoParcRec.lNumIntParc, 0, STATUS_LANCADO, objInfoParcRec.dValorMulta, objInfoParcRec.dValorJuros, objInfoParcRec.dValorDesconto, dValorParcelaBaixa, dValorRecebidoParcela, objInfoParcRec.iCobrador)
        
        End If

    Next

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(DataBaixa.Text))
    If lErro <> SUCESSO Then gError 20828

    'Verifica se alguma célula Selecionada do GridParcelas está marcada
    If iParcelaSelecionada < 1 Then gError 29103

    'Se o Tipo de recebimento é Recebimento em Dinheiro
    If Recebimento(0) = True Then

        lErro = Grava_Recebimento_Dinheiro(dSomaReceber, dtDataBaixa, colBaixaParcReceber, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        If lErro <> SUCESSO Then gError 29016

    'Se o Tipo de recebimento é Recebimento Antecipado
    ElseIf Recebimento(1) = True Then

        lErro = Grava_Recebimento_Antecipado(dSomaReceber, dtDataBaixa, colBaixaParcReceber, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        If lErro <> SUCESSO Then gError 29017
        
    'Se o Tipo de recebimento é Débito/Devolução
    ElseIf Recebimento(2) = True Then
    
        lErro = Grava_Recebimento_DebDev(dSomaReceber, dtDataBaixa, colBaixaParcReceber, objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        If lErro <> SUCESSO Then gError 29019
        
    'Se o Tipo de recebimento é Perda
    ElseIf Recebimento(3) = True Then
    
        'Insere BaixaReceber e BaixasParRec correspondentes no BD
        lErro = CF("BaixaRec_Grava_Perda", colBaixaParcReceber, dtDataBaixa, objContabil, HistoricoPerda.Text, StrParaDate(CTBDataContabil.Text), objFilialCliente.lCodCliente, objFilialCliente.iCodFilial)
        If lErro <> SUCESSO Then gError 29176
    
    End If
    
    If ImprimirRecibo.Value = vbChecked Then
    
        lErro = CF("RelRecibo_Prepara", colBaixaParcReceber.Item(1).lNumIntBaixa)
        'If lErro <> SUCESSO Then gError 182868
    
    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 29016, 29017, 29019, 29176, 59489, 20828, ERRO_SEM_MENSAGEM

        Case 29101
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_BAIXA_SEM_PREENCHIMENTO", gErr)

        Case 29102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_PARCELAS_NAO_INFORMADO", gErr)

        Case 29103
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_MARCADO", gErr)
        
        Case 39015
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CREDITO_SEM_PREENCHIMENTO", gErr)
        
        Case 49554
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MULTA_INCOMPATIVEL_DATABAIXA_DATAVENC", gErr)
            DataBaixa.SetFocus
            
        Case 182868
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143428)

    End Select

    Exit Function

End Function

Private Function Limpa_Tela_BaixaRec()
'Limpa a tela

Dim lErro As Long

    'Limpa a coleção de Parcelas
    Set gcolInfoParcRec = Nothing
    
    'Limpa os comandos da tela
    Call Limpa_Tela(Me)

    'Limpa os Labels
    TotalBaixar.Caption = Format(0, "Fixed")
    ValorReceber.Caption = Format(0, "Fixed")
    
    'Limpa os Grid
    Call Grid_Limpa(objGridRecebAntecipados)
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridDebitos)

    'Deixa as ComboBOx sem ítem selecionado
    ContaCorrente.ListIndex = -1
    Filial.ListIndex = -1

    FormaPagto.ListIndex = -1
    
    'Passa a data do sitema para a tela
    DataBaixa = Format(gdtDataAtual, "dd/mm/yy")
    
    DataCredito.PromptInclude = False
    DataCredito.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataCredito.PromptInclude = True

    'Modifica a cor da label pagamento(1) e pagamento(2)
    Recebimento(1).ForeColor = vbButtonText
    Recebimento(2).ForeColor = vbButtonText
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    '########################################
    'Inserido por Wagner
    Call Ordenacao_Limpa(objGridParcelas)
        
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
    ImprimirRecibo.Value = vbUnchecked
    
    TipoDocTodos.Value = True
    Call TipoDocTodos_Click
    
    Call CF2(Me, "BR_Limpa_Tela_BaixaRec")
    '########################################
    
    DescProdInic.Caption = ""
    DescProdFim.Caption = ""
   
    'Zera iAlterado
    iAlterado = 0

End Function

Private Function Carrega_Tab_Parcelas() As Long

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente
Dim iCodFilialCliente As Integer
Dim iFilialCli As Integer
Dim lCodCli As Long
Dim lTituloInic As Long
Dim lTituloFim As Long
Dim dtEmissaoInic As Date
Dim dtEmissaoFim As Date
Dim dtVencInic As Date
Dim dtVencFim As Date, lCodViagem As Long, iFormaPagamento As Integer
Dim objVendedor As New ClassVendedor, iVendedor As Integer

On Error GoTo Erro_Carrega_Tab_Parcelas
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    DoEvents
    
    'Verifica se Cliente e Filial estão preenchidos e obtem dados
    lErro = FilialCliente_ObtemDaTela(objFilialCliente)
    If lErro <> SUCESSO Then Error 59488
                
    lCodCli = objFilialCliente.lCodCliente
    iFilialCli = objFilialCliente.iCodFilial

    dtEmissaoInic = MaskedParaDate(EmissaoInic)
    dtEmissaoFim = MaskedParaDate(EmissaoFim)
    dtVencInic = MaskedParaDate(VencInic)
    dtVencFim = MaskedParaDate(VencFim)
        
    'Se as datas EmissaoFim e EmissaoInic estão preenchidas
    If dtEmissaoInic <> DATA_NULA And dtEmissaoFim <> DATA_NULA Then
        'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
        If dtEmissaoFim < dtEmissaoInic Then Error 29032
    End If

    'Se as datas VencFim e VencInic estão preenchidas
    If dtVencInic <> DATA_NULA And dtVencFim <> DATA_NULA Then
        'Verifica se a EmissaoFim é maior ou igual a EmissaoInic
        If dtVencFim < dtVencInic Then Error 29033
    End If

    'Lê TituloInic e TituloFim que estão na tela
    lTituloInic = StrParaLong(TituloInic.Text)
    lTituloFim = StrParaLong(TituloFim.Text)
        
    'Se TituloFim e TituloInic estão preenchidos
    If Len(Trim(TituloFim.Text)) And Len(Trim(TituloInic.Text)) > 0 Then
        
        'Verifica se TituloFim é maior ou igual que TituloInic
        If lTituloFim < lTituloInic Then Error 29034

    End If

    'Limpa gcolInfoParcRec antes de carregar as novas parcelas
    Set gcolInfoParcRec = Nothing
    Set gcolRecebAntecipado = Nothing
    Set gcolInfoParcRec = New Collection
    
    lCodViagem = StrParaLong(CodigoViagem.Text)
    If FormaPagto.ListIndex <> -1 Then iFormaPagamento = FormaPagto.ItemData(FormaPagto.ListIndex)

    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text

    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
    
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError ERRO_SEM_MENSAGEM

        'Não encontrou o vendedor ==> erro
        If lErro = 25008 Then gError 209208

        iVendedor = objVendedor.iCodigo

    End If
        
    lErro = CF2(Me, "BR_Carrega_Tab_Parcelas", lCodCli, iFilialCli, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, objFilialCliente, gcolInfoParcRec, gcolRecebAntecipado, lCodViagem, iFormaPagamento, iVendedor, Codigo_Extrai(CobradorFiltro.Text))
    If lErro <> SUCESSO Then gError 29035
    
'    'Preenche a Coleção de Parcelas
'    lErro = CF("ParcelasRec_Le_BaixaManualSel", lCodCli, iFilialCli, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, gcolInfoParcRec)
'    If lErro <> SUCESSO Then Error 29035

    'Verifica o número máximo de parcelas
    'If gcolInfoParcRec.Count > NUM_MAX_PARCELAS_BAIXA Then Error 43097

    'Preenche a Coleção de Débitos
    Set gcolDebitos = Nothing
    lErro = CF("DebitosReceber_Le_FilialCliente", objFilialCliente, gcolDebitos)
    If lErro <> SUCESSO And lErro <> 29148 Then Error 29036

    'Verifica o número máximo de débitos
    'If gcolDebitos.Count > NUM_MAX_DEBITOS Then Error 43098

'    'Preenche a Coleção de Recebimentos antecipados
'    Set gcolRecebAntecipado = Nothing
'    lErro = CF("RecebAntecipados_Le_FilialCliente_Baixa", objFilialCliente, gcolRecebAntecipado)
'    If lErro <> SUCESSO And lErro <> 29142 Then Error 29037

    'Verifica o número máximo de recebimentos antecipados
    'If gcolRecebAntecipado.Count > NUM_MAX_RECEB_ANTECIPADOS Then Error 43099
    
    'Limpa o frame Parcelas
    Call Frame_Parcelas_Limpa

    'Passa a data do sistema para a tela
    DataBaixa = Format(gdtDataAtual, "dd/mm/yy")
    DataCredito = Format(gdtDataAtual, "dd/mm/yy")

    'Preenche o GridParcelas
    lErro = Grid_Parcelas_Preenche(gcolInfoParcRec)
    If lErro <> SUCESSO Then Error 49724
    
    'Preenche o GridRecebAntecipados
    lErro = Grid_RecebAntecipados_Preenche(gcolRecebAntecipado)
    If lErro <> SUCESSO Then Error 49737
        
    'Preenche o GridDebitos
    lErro = Grid_Debitos_Preenche(gcolDebitos)
    If lErro <> SUCESSO Then Error 49738
    
    'Formata o Valor a receber na tela
    ValorReceber.Caption = ""
    
    'para forçar o recalculo
    dtDataMultaJuros = DATA_NULA
    
    Call Calcula_Multa_Juros_Desc_Parcelas
    
    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    GL_objMDIForm.MousePointer = vbDefault
    
    Carrega_Tab_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Carrega_Tab_Parcelas:

    GL_objMDIForm.MousePointer = vbDefault

    Carrega_Tab_Parcelas = Err
    
    Select Case Err
    
        Case 29035, 29036, 29037, 49724, 49737, 49738, 59488
            
        Case 29032
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_INICIAL_MAIOR", Err)

        Case 29033
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_INICIAL_MAIOR", Err)
        
        Case 29034
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOINIC_MAIOR_TITULOFIM", Err)
        
        Case 43097
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELAS_SUPERIOR_NUM_MAX_PARCELAS_BAIXA", Err)
            
        Case 43098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DEBITOS_SUPERIOR_NUM_MAX_DEBITOS", Err)
            
        Case 43099
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEBANTECIPADOS_SUPERIOR_NUM_MAX_RECEB_ANTECIPADOS", Err)
            
        Case 209208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", gErr, objVendedor.sNomeReduzido)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143429)

    End Select

    Exit Function

End Function

Private Function Grava_Recebimento_Dinheiro(dSomaReceber As Double, dtDataBaixa As Date, colBaixaParcReceber As ColBaixaParcRec, lCliente As Long, iFilialCli As Integer) As Long
'Valida dados da tela para gravação de uma nova baixa com pagamento em dinheiro

Dim lErro As Long
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objMovCtaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Grava_Recebimento_Dinheiro

    'Verifica se ContaCorrente está preenchida
    If Len(Trim(ContaCorrente.Text)) = 0 Then Error 29104

    'Verifica se ValorReceber está preenchido
    If Len(Trim(ValorReceber.Caption)) = 0 Then Error 29105

    'Verifica se Valor do Recebimento é igual à soma dos valores a receber
    If CDbl(Trim(ValorReceber.Caption)) <> dSomaReceber Then Error 29106

    'Extrai da tela o código da Conta Corrente
    objContaCorrenteInt.iCodigo = Codigo_Extrai(ContaCorrente.Text)

    'Lê os dados da Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objContaCorrenteInt.iCodigo, objContaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 29113
    
    'Se não achou a Conta Corrente --> erro
    If lErro <> SUCESSO Then Error 43109

    'Armazena os dados do recebimento em ObjMovCtaCorrente
    objMovCtaCorrente.iFilialEmpresa = objContaCorrenteInt.iFilialEmpresa
    objMovCtaCorrente.iCodConta = objContaCorrenteInt.iCodigo
    objMovCtaCorrente.iTipo = MOVCCI_RECEBIMENTO_TITULO
    objMovCtaCorrente.iExcluido = NAO_EXCLUIDO
    objMovCtaCorrente.iTipoMeioPagto = RECEBIMENTO_EM_DINHEIRO
    objMovCtaCorrente.dtDataMovimento = MaskedParaDate(DataCredito)
    objMovCtaCorrente.dtDataBaixa = dtDataBaixa
    objMovCtaCorrente.dValor = CDbl(ValorReceber.Caption)
    objMovCtaCorrente.sHistorico = Trim(Historico.Text)
    objMovCtaCorrente.iConciliado = NAO_CONCILIADO
    objMovCtaCorrente.dtDataContabil = IIf(MaskedParaDate(CTBDataContabil) = DATA_NULA, gdtDataAtual, MaskedParaDate(CTBDataContabil))
    
    'Chama BaixaReceber_Grava_Recebimento dentro de transação
    lErro = CF("BaixaRec_Grava_Recebimento", colBaixaParcReceber, objMovCtaCorrente, objContabil, lCliente, iFilialCli)
    If lErro <> SUCESSO Then Error 29159
    
    Grava_Recebimento_Dinheiro = SUCESSO

    Exit Function
    
Erro_Grava_Recebimento_Dinheiro:

    Grava_Recebimento_Dinheiro = Err
            
    Select Case Err
    
        Case 29104
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_RECDINHEIRO_NAO_INFORMADA", Err)

        Case 29105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_RECDINHEIRO_NAO_INFORMADO", Err)

        Case 29106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORREC_DIFERENTE_TOTALPARCELAS", Err, ValorReceber.Caption, dSomaReceber)

        Case 29113, 29159

        Case 43109
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrenteInt.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143430)

    End Select

    Exit Function

End Function

Private Function Grava_Recebimento_Antecipado(dSomaReceber As Double, dtDataBaixa As Date, colBaixaParcReceber As ColBaixaParcRec, lCliente As Long, iFilialCli As Integer) As Long
'Valida dados da tela para gravação de baixa
        
Dim lErro As Long
Dim iIndice As Integer
Dim iRecebAntecipadoSelecionada As Integer
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objTipoMeioPagto As New ClassTipoMeioPagto
Dim objCliente As New ClassCliente
Dim objRecebAntecipado As New ClassAntecipReceb
Dim iCodFilialCli As Integer

On Error GoTo Erro_Grava_Recebimento_Antecipado

    iRecebAntecipadoSelecionada = 0

    'Percorre todas as linhas do Grid RecebAntecipados
    For iIndice = 1 To objGridRecebAntecipados.iLinhasExistentes

        'Se a checkbox Selecionada está marcada
        If GridRecebAntecipados.TextMatrix(iIndice, iGrid_SelecionarRA_Col) = GRID_CHECKBOX_ATIVO Then

            'Soma quantos Recebimentos antecipados estão selecionados
            iRecebAntecipadoSelecionada = iRecebAntecipadoSelecionada + 1
            
            'Verifica se saldo do adiantamento é menor à soma dos valores a receber
            If CDbl(GridRecebAntecipados.TextMatrix(iIndice, iGrid_SaldoRA_Col)) < dSomaReceber Then Error 41625
            
            Exit For

        End If

    Next

    'Verifica se alguma célula Selecionada do Grid RecebAntecipados está marcada
    If iRecebAntecipadoSelecionada < 1 Then Error 29108

    Set objRecebAntecipado = gcolRecebAntecipado.Item(iIndice)
    objRecebAntecipado.dValorBaixado = StrParaDbl(ValorReceber.Caption)
            
    'Insere BaixaReceber e BaixasParcelaRec correspondentes no BD. Apropria RecebAntecipado do valor baixado.
    lErro = CF("BaixaRec_Grava_RecebAntecipado", colBaixaParcReceber, objRecebAntecipado, dtDataBaixa, objContabil, StrParaDate(CTBDataContabil.Text), lCliente, iFilialCli)
    If lErro <> SUCESSO Then Error 29117

    Grava_Recebimento_Antecipado = SUCESSO
    
    Exit Function
    
Erro_Grava_Recebimento_Antecipado:
    
    Grava_Recebimento_Antecipado = Err
    
    Select Case Err
    
        Case 29020, 29115, 29116, 29117
        
        Case 41625
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_RA_MENOR_TOTALRECEBER", Err)
        
        Case 29021
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOMEIOPAGTO_INEXISTENTE", Err, objTipoMeioPagto.iTipo)

        Case 29108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_RA_NAO_MARCADA", Err)

        Case 43106
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrenteInt.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143431)

    End Select
    
    Exit Function
    
End Function

Private Function Grava_Recebimento_DebDev(dSomaReceber As Double, dtDataBaixa As Date, colBaixaParcReceber As ColBaixaParcRec, lCliente As Long, iFilialCli As Integer) As Long
'Valida os dados da tela para gravação de uma nova baixa
        
Dim lErro As Long
Dim iIndice As Integer
Dim iDebitosSelecionada As Integer
Dim iCodFilialCli As Integer
Dim objCliente As New ClassCliente
Dim objDebitoReceber As New ClassDebitoRecCli
        
On Error GoTo Erro_Grava_Recebimento_DebDev

    iDebitosSelecionada = 0

    'Percorre todas as linhas do Grid Débitos
    For iIndice = 1 To objGridDebitos.iLinhasExistentes

        'Se a checkbox Selecionar está marcada
        If GridDebitos.TextMatrix(iIndice, iGrid_SelecionarDB_Col) Then

            'Conta quantos Débitos estão selecionados
            iDebitosSelecionada = iDebitosSelecionada + 1

            'Verifica se saldo do debito é menor à soma dos valores a receber
            If CDbl(GridDebitos.TextMatrix(iIndice, iGrid_SaldoDB_Col)) < dSomaReceber Then Error 29112
        
            Exit For
        
        End If
    
    Next

    'Verifica se alguma célula Selecionada do Grid Débitos está marcada
    If iDebitosSelecionada < 1 Then Error 29111

    Set objDebitoReceber = gcolDebitos.Item(iIndice)
    objDebitoReceber.dValorBaixado = StrParaDbl(ValorReceber.Caption)
    
    'Insere BaixaReceber e BaixasParcelaReceber correspondentes no BD. Apropria DebitoReceber do valor baixado.
    lErro = CF("BaixaRec_Grava_DebitoReceber", colBaixaParcReceber, objDebitoReceber, dtDataBaixa, objContabil, StrParaDate(CTBDataContabil.Text), lCliente, iFilialCli)
    If lErro <> SUCESSO Then Error 29119
    
    Grava_Recebimento_DebDev = SUCESSO
    
    Exit Function
    
Erro_Grava_Recebimento_DebDev:

    Grava_Recebimento_DebDev = Err
    
    Select Case Err

        Case 29110
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBAIXAR_DB_NAO_INFORMADO", Err)

        Case 29111
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_DB_NAO_MARCADA", Err)

        Case 29112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SALDO_DB_MENOR_TOTALRECEBER", Err)

        Case 29118, 29119

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143432)

    End Select
        
    Exit Function

End Function

Public Sub DataCredito_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DataCredito_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataCredito_Validate

    'Se a DataCredito está preenchida
    If Len(DataCredito.ClipText) > 0 Then

        'Verifica se a DataCredito é válida
        lErro = Data_Critica(DataCredito.Text)
        If lErro <> SUCESSO Then Error 46480

    'Se a DataCredito não está preenchida
    Else
    
        Error 46481
        
    End If

    Exit Sub

Erro_DataCredito_Validate:

    Cancel = True


    Select Case Err

        Case 46480

        Case 46481
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143433)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataCredito_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCredito_DownClick

    'Diminui a DataCredito em 1 dia
    lErro = Data_Up_Down_Click(DataCredito, DIMINUI_DATA)
    If lErro <> SUCESSO Then Error 46482

    Exit Sub

Erro_UpDownDataCredito_DownClick:

    Select Case Err

        Case 46482

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143434)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataCredito_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataCredito_UpClick

    'Aumenta a DataCredito em 1 dia
    lErro = Data_Up_Down_Click(DataCredito, AUMENTA_DATA)
    If lErro <> SUCESSO Then Error 46483

    Exit Sub

Erro_UpDownDataCredito_UpClick:

    Select Case Err

        Case 46483

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143435)

    End Select

    Exit Sub

End Sub

Public Sub Cobrador_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Cobrador_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Cobrador_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = Cobrador
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

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
Dim iLinha As Integer, sCampoGlobal As String
Dim iCodigo As Integer, sContaTela As String
Dim objContaCorrenteInt As New ClassContasCorrentesInternas
Dim objInfoParcRec As ClassInfoParcRec
Dim objCarteiraCobrador As New ClassCarteiraCobrador
Dim objCliente As New ClassCliente, objTipoCliente As New ClassTipoCliente
Dim objFilial As New ClassFilialCliente, objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico
    
        Case CLIENTE_COD
            
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56519
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case CLIENTE_NOME
        
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56520
            
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
        
            Else
            
                objMnemonicoValor.colValor.Add ""
                
            End If
        
        Case BAIXAR
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                    
                objMnemonicoValor.colValor.Add CInt(GridParcelas.TextMatrix(iLinha, iGrid_Baixar_Col))
                
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
                If lErro <> SUCESSO Then Error 56521
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 56522
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56523
                
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
                If lErro <> SUCESSO Then Error 56524
                
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
            
        Case NUM_TITULO
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CLng(GridParcelas.TextMatrix(iLinha, iGrid_Numero_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next
            
        Case PARCELA1
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                    
                If Len(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_Parcela_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
                
            Next
            
        Case VALOR_BAIXAR
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If GridParcelas.TextMatrix(iLinha, iGrid_Baixar_Col) = GRID_CHECKBOX_ATIVO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixar_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixar_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case VALOR_DESCONTO
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If GridParcelas.TextMatrix(iLinha, iGrid_Baixar_Col) = GRID_CHECKBOX_ATIVO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case VALOR_MULTA
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If GridParcelas.TextMatrix(iLinha, iGrid_Baixar_Col) = GRID_CHECKBOX_ATIVO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorMulta_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorMulta_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case VALOR_JUROS
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                If GridParcelas.TextMatrix(iLinha, iGrid_Baixar_Col) = GRID_CHECKBOX_ATIVO And Len(Trim(GridParcelas.TextMatrix(iLinha, iGrid_ValorJuros_Col))) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridParcelas.TextMatrix(iLinha, iGrid_ValorJuros_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            
            Next
            
        Case PAGTO_DINHEIRO
            
            objMnemonicoValor.colValor.Add Recebimento(0).Value
            
        Case PAGTO_ANTECIPADO
            
            objMnemonicoValor.colValor.Add Recebimento(1).Value
            
        Case CRED_DEV
            
            objMnemonicoValor.colValor.Add Recebimento(2).Value
                
        Case PERDA
            objMnemonicoValor.colValor.Add Recebimento(3).Value
            
        Case CTA_CONTA_CORRENTE
            
            If Len(ContaCorrente.Text) > 0 Then
                iCodigo = Codigo_Extrai(ContaCorrente.Text)
                
                lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 49725
                
                If objContaCorrenteInt.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objContaCorrenteInt.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56525
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela

            Else
            
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case CTACARTEIRACOBRADOR
        
            'Percorre todas as linhas do Grid Parcelas
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                
                sContaTela = ""
                    
                'Passa a linha do Grid para o Obj
                Set objInfoParcRec = gcolInfoParcRec.Item(iLinha)
            
                objCarteiraCobrador.iCobrador = objInfoParcRec.iCobrador
                objCarteiraCobrador.iCodCarteiraCobranca = objInfoParcRec.iCarteiraCobrador

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
                
                objMnemonicoValor.colValor.Add sContaTela
            
            Next
            
        Case DATA_CREDITO
            If Len(DataCredito.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataCredito.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case CLIENTE_COD_GRID
        
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                objMnemonicoValor.colValor.Add LCodigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col))
            Next
            
        Case CLIENTE_NOME_GRID
        
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                'Preenche NomeReduzido com o fornecedor da tela
                If LCodigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col)) <> 0 Then
                    objCliente.lCodigo = LCodigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col))
                    lErro = CF("Cliente_Le", objCliente)
                    If lErro <> SUCESSO Then Error 56520
                    objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case FILIAL_COD_GRID
            
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                objMnemonicoValor.colValor.Add Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col))
            Next
            
        Case FILIAL_NOME_RED_GRID
            
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col)) <> 0 Then
                    objFilial.iCodFilial = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col))
                    objFilial.lCodCliente = LCodigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col))
                    lErro = CF("FilialCliente_Le", objFilial)
                    If lErro <> SUCESSO Then Error 56521
                    objMnemonicoValor.colValor.Add objFilial.sNome
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
        Case FILIAL_CONTA_GRID
            
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col)) <> 0 Then
                    objFilial.iCodFilial = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col))
                    objFilial.lCodCliente = LCodigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col))
                    lErro = CF("FilialCliente_Le", objFilial)
                    If lErro <> SUCESSO Then Error 56522
                    
                    If objFilial.sContaContabil <> "" Then
                        lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                        If lErro <> SUCESSO Then Error 56523
                    Else
                        sContaTela = ""
                    End If
                    objMnemonicoValor.colValor.Add sContaTela
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
            
            
        Case FILIAL_CGC_CPF_GRID
            
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
                If Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col)) <> 0 Then
                    objFilial.iCodFilial = Codigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Filial_Col))
                    objFilial.lCodCliente = LCodigo_Extrai(GridParcelas.TextMatrix(iLinha, iGrid_Cliente_Col))
                    lErro = CF("FilialCliente_Le", objFilial)
                    If lErro <> SUCESSO Then Error 41980
                    objMnemonicoValor.colValor.Add objFilial.sCgc
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next
                
        Case NOME_CONTA_CORRENTE
            
            If Len(ContaCorrente.Text) > 0 Then
                iCodigo = Codigo_Extrai(ContaCorrente.Text)
                
                lErro = CF("ContaCorrenteInt_Le", iCodigo, objContaCorrenteInt)
                If lErro <> SUCESSO Then Error 49725
                
                objMnemonicoValor.colValor.Add objContaCorrenteInt.sNomeReduzido

            Else
            
                objMnemonicoValor.colValor.Add ""
            
            End If
                
        Case Else
            Error 39695
    
    End Select
    
    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err
    
    Select Case Err
        
        Case 49724, 49725, 49726, 56519 To 56526, 56800
        
        Case 56797, 56799
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRADOR_NAO_CADASTRADO", Err, objCarteiraCobrador.iCobrador)
        
        Case 56801
            Call Rotina_Erro(vbOKOnly, "ERRO_MNEMONICO_INEXISTENTE", Err, objMnemonico.sMnemonico)
        
        Case 39695
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143436)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Calcula_Total_Receber()

Dim iIndice As Integer
Dim dTotalSelecionado As Double

    dTotalSelecionado = 0
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
        
        If StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_Baixar_Col)) <> 0 Then
            
            dTotalSelecionado = dTotalSelecionado + StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorAReceber_Col))
        End If
    Next
    
    If dTotalSelecionado > 0 Then
        ValorReceber.Caption = Format(dTotalSelecionado, "Standard")
        TotalBaixar.Caption = Format(dTotalSelecionado, "Standard")
    Else
        ValorReceber.Caption = ""
        TotalBaixar.Caption = ""
    End If
    
    Exit Sub
    
End Sub

Private Function Parcela_RecalculaValorReceber(iLinha As Integer, sValorBaixar As String, sValorDesconto As String, sValorMulta As String, sValorJuros As String) As Long
'obs: chamada dentro de saida celula

Dim lErro As Long, dValorReceber As Double

On Error GoTo Erro_Parcela_RecalculaValorReceber

    'Lê os dados da tela
    dValorReceber = StrParaDbl(sValorBaixar) - StrParaDbl(sValorDesconto) + StrParaDbl(sValorMulta) + StrParaDbl(sValorJuros)
    
    If dValorReceber < 0 Then Error 56780
    
    'Preenche ValorReceber
    GridParcelas.TextMatrix(iLinha, iGrid_ValorAReceber_Col) = Format(dValorReceber, "Standard")

    Parcela_RecalculaValorReceber = SUCESSO
     
    Exit Function
    
Erro_Parcela_RecalculaValorReceber:

    Parcela_RecalculaValorReceber = Err
     
    Select Case Err
          
        Case 56780
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_RECEBER_NEGATIVO", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143437)
     
    End Select
     
    Exit Function

End Function


'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_BAIXA_PARCELAS_RECEBER_TITULOS
    Set Form_Load_Ocx = Me
    Caption = "Baixa de Parcelas a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BaixaRec"
    
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


Private Function Calcula_Multa_Juros_Desc_Parcela(iLinha As Integer) As Long
'Preenche para a parcela da linha do grid passada como parametro os valores para multa, juros e descontos

Dim lErro As Long
Dim objInfoParcRec As ClassInfoParcRec, objParcelaReceber As New ClassParcelaReceber
Dim dtDataBaixa As Date
Dim dValorMulta As Double
Dim dValorJuros As Double
Dim dValorDesconto As Double
Dim objTituloReceber As New ClassTituloReceber
Dim iDias As Integer, dtDataVenctoReal As Date

On Error GoTo Erro_Calcula_Multa_Juros_Desc_Parcela
    
    If gcolInfoParcRec.Count >= iLinha Then
    
        dtDataBaixa = MaskedParaDate(DataBaixa)
        
        Set objInfoParcRec = gcolInfoParcRec(iLinha)
            
        objParcelaReceber.lNumIntDoc = objInfoParcRec.lNumIntParc
        lErro = CF("ParcelaReceber_Le", objParcelaReceber)
        If lErro <> SUCESSO And lErro <> 46477 Then Error 56783
        If lErro <> SUCESSO Then Error 56784
        
        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", objParcelaReceber.dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then Error 56785

        If dtDataBaixa > dtDataVenctoReal Then
                    
            lErro = CF("Calcula_Multa_Juros_Parcela", objParcelaReceber, dtDataBaixa, dValorMulta, dValorJuros)
            If lErro <> SUCESSO Then Error 56786
            
            GridParcelas.TextMatrix(iLinha, iGrid_ValorMulta_Col) = Format(dValorMulta, "Standard")
            GridParcelas.TextMatrix(iLinha, iGrid_ValorJuros_Col) = Format(dValorJuros, "Standard")
            GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = ""
                   
         Else
                   
            GridParcelas.TextMatrix(iLinha, iGrid_ValorMulta_Col) = ""
            GridParcelas.TextMatrix(iLinha, iGrid_ValorJuros_Col) = ""
            
            lErro = CF("Calcula_Desconto_Parcela", objParcelaReceber, dValorDesconto, dtDataBaixa)
            If lErro <> SUCESSO Then Error 56787
            
            If dValorDesconto > 0 Then
                GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = Format(dValorDesconto, "Standard")
            Else
                GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = ""
            End If
                   
        End If

    End If

    Calcula_Multa_Juros_Desc_Parcela = SUCESSO
    
    Exit Function

Erro_Calcula_Multa_Juros_Desc_Parcela:

    Calcula_Multa_Juros_Desc_Parcela = Err
    
    Select Case Err
    
        Case 56783, 56785, 56786, 56787
        
        Case 56784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143438)
    
    End Select
    
    Exit Function

End Function

Private Function Calcula_Multa_Juros_Desc_Parcelas() As Long
'Preenche para todas as parcelas do grid os valores para multa, juros e descontos

Dim lErro As Long
Dim iLinha As Integer, dtDataBaixa As Date

On Error GoTo Erro_Calcula_Multa_Juros_Desc_Parcelas
    
    If Len(Trim(DataBaixa.ClipText)) <> 0 Then
    
        dtDataBaixa = CDate(DataBaixa.FormattedText)
        
        If dtDataMultaJuros <> dtDataBaixa Then
        
            For iLinha = 1 To objGridParcelas.iLinhasExistentes
            
                lErro = Calcula_Multa_Juros_Desc_Parcela(iLinha)
                If lErro <> SUCESSO Then Error 56788
            
                lErro = Parcela_RecalculaValorReceber(iLinha, GridParcelas.TextMatrix(iLinha, iGrid_ValorBaixar_Col), GridParcelas.TextMatrix(iLinha, iGrid_ValorDesconto_Col), GridParcelas.TextMatrix(iLinha, iGrid_ValorMulta_Col), GridParcelas.TextMatrix(iLinha, iGrid_ValorJuros_Col))
                If lErro <> SUCESSO Then Error 59052
                
            Next
    
            dtDataMultaJuros = dtDataBaixa
            
        End If
    
    End If
    
    Call Calcula_Total_Receber
    
    Calcula_Multa_Juros_Desc_Parcelas = SUCESSO
    
    Exit Function

Erro_Calcula_Multa_Juros_Desc_Parcelas:

    Calcula_Multa_Juros_Desc_Parcelas = Err
    
    Select Case Err
    
        Case 56788, 59052
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143439)
    
    End Select
    
    Exit Function

End Function

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then
            Call LabelCli_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call VendedorLabel_Click
        End If
    
    End If
    
End Sub

Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Function FilialCliente_ObtemDaTela(objFilialCliente As ClassFilialCliente)
'preenche em objFilialCliente o codigo do cliente e sua filial à partir dos campos cliente e filial da tela

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilialCliente As Integer

On Error GoTo Erro_FilialCliente_ObtemDaTela

    'Verifica se Cliente e Filial estão preenchidos
    'Comentado por Wagner
'    If Len(Trim(Cliente.Text)) = 0 Then Error 29029
'    If Len(Trim(Filial.Text)) = 0 Then Error 29031
    
    If Len(Trim(Cliente.Text)) <> 0 Then
    
        'Lê os dados do Cliente
        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilialCliente)
        If lErro <> SUCESSO Then Error 29030
        
        'Passa Cliente que está na tela para o Obj
        objFilialCliente.lCodCliente = objCliente.lCodigo
    
        'Passa a Filial que está na tela para o Obj
        objFilialCliente.iCodFilial = Codigo_Extrai(Filial.Text)
        
    End If

    FilialCliente_ObtemDaTela = SUCESSO
     
    Exit Function
    
Erro_FilialCliente_ObtemDaTela:

    FilialCliente_ObtemDaTela = Err
     
    Select Case Err
          
        Case 29030
        
        Case 29029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 29031
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143440)
     
    End Select
     
    Exit Function
    
End Function

Public Sub mnuGridMarcarTodos_Click()
    Select Case iGridAtual
        'Se for o grid parcelas
        Case GRID_PARCELAS
            Call Grid_Marca_Desmarca(objGridParcelas, iGrid_Baixar_Col, MARCADO)
            Call Calcula_Total_Receber
        'Se for o grid de créditos/ devoluções
         Case GRID_DEVOLUCOES
            Call Grid_Marca_Desmarca(objGridDebitos, iGrid_SelecionarDB_Col, MARCADO)
         'Se for o grid de adiantamentos
         Case GRID_ADIANTAMENTOS
            Call Grid_Marca_Desmarca(objGridRecebAntecipados, iGrid_SelecionarRA_Col, MARCADO)
    End Select
End Sub

Public Sub mnuGridDesmarcarTodos_Click()
    Select Case iGridAtual
        'Se for o grid parcelas
        Case GRID_PARCELAS
            Call Grid_Marca_Desmarca(objGridParcelas, iGrid_Baixar_Col, DESMARCADO)
            ValorReceber.Caption = ""
            TotalBaixar.Caption = ""
        'Se for o grid de créditos/ devoluções
         Case GRID_DEVOLUCOES
            Call Grid_Marca_Desmarca(objGridDebitos, iGrid_SelecionarDB_Col, DESMARCADO)
         'Se for o grid de adiantamentos
         Case GRID_ADIANTAMENTOS
            Call Grid_Marca_Desmarca(objGridRecebAntecipados, iGrid_SelecionarRA_Col, DESMARCADO)
    End Select
End Sub

Public Sub mnuGridConsultaDocOriginal_Click()
'Chama a tela de consulta de Títulos a Receber quando essa opção for selecionada no grid

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim objFilialCliente As New ClassFilialCliente

On Error GoTo Erro_mnuGridConsultaDocOriginal_Click

    'Se o Cliente não foi preenchido => erro
'    If Len(Trim(Cliente.Text)) = 0 Then gError 79908
    
    'Se a filial não foi preenchida => erro
'    If Len(Trim(Filial.Text)) = 0 Then gError 79909
       
    'Verifica em qual grid está o documento que será consultado
    Select Case iGridAtual

     'Se for o grid parcelas
     Case GRID_PARCELAS
         
        'Guarda no obj o parâmetro que será utilizado na obtenção do código do Cliente
        objCliente.lCodigo = LCodigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cliente_Col))
            
        'Obtém o código do Cliente que será passado como parâmetro
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 79910
        
        'Se não encontrou o cliente = > erro
        If lErro = 12293 Then gError 79911
        
        'Guarda no objFilialCliente o código do cliente e da filial
        objFilialCliente.lCodCliente = objCliente.lCodigo
        objFilialCliente.iCodFilial = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Filial_Col))
        
        'Abre a tela de consulta de títulos a pagar
        lErro = Consulta_TituloReceber(objFilialCliente)
        If lErro <> SUCESSO Then gError 79918
    
    'Se for o grid de créditos/ devoluções
     Case GRID_DEVOLUCOES
         
        lErro = FilialCliente_ObtemDaTela(objFilialCliente)
        If lErro <> SUCESSO Then gError 79919
         
         'Abre a tela de cadastro de créditos/devoluções
         lErro = Consulta_DebitosReceber(objFilialCliente)
         If lErro <> SUCESSO Then gError 79919
     
     'Se for o grid de adiantamentos
     Case GRID_ADIANTAMENTOS
     
        lErro = FilialCliente_ObtemDaTela(objFilialCliente)
        If lErro <> SUCESSO Then gError 79920
     
         'Abre a tela para cadastro de adiantamentos
         lErro = Consulta_RecebAntecipados(objFilialCliente)
         If lErro <> SUCESSO Then gError 79920

    End Select
    
    Exit Sub
    
Erro_mnuGridConsultaDocOriginal_Click:

    Select Case gErr
    
        Case 79910
    
        Case 79911
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.lCodigo)

        Case 79918 To 79920
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143441)
            
    End Select
    
    Exit Sub
    
End Sub

Private Function Consulta_TituloReceber(objFilialCliente As ClassFilialCliente) As Long
'Exibe o título referente à parcela selecionada no grid
            
Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
            
On Error GoTo Erro_Consulta_TituloReceber

    'Se não há linha selecionada no grid parcelas => erro
    If GridParcelas.Row <= 0 Then gError 79916
    
    'Se a linha selecionada não contém dados => erro
    If Len(Trim(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col))) = 0 Then gError 79917
    
    'Preenche o obj com os dados que serão utilizados para leitura do título que será consultado
    With objTituloReceber

        .lCliente = objFilialCliente.lCodCliente
        .iFilial = objFilialCliente.iCodFilial
        .iFilialEmpresa = Codigo_Extrai(GridParcelas.TextMatrix(GridParcelas.Row, iGrid_FilialEmpresa_Col))
        .sSiglaDocumento = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Tipo_Col)
        .lNumTitulo = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Numero_Col)

    End With
    
    'Obtém o NumIntDoc que será usado no trata parâmetros da tela
    lErro = CF("TituloReceber_Le_SemNumIntDoc", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 28574 Then gError 79925
    
    'Se não encontrou o Título => erro
    If lErro = 28574 Then gError 79926
    
    'Chama a tela de consulta de títulos a receber
    Call Chama_Tela("TituloReceber_Consulta", objTituloReceber)
    
    Consulta_TituloReceber = SUCESSO
    
    Exit Function

Erro_Consulta_TituloReceber:

    Consulta_TituloReceber = gErr
    
    Select Case gErr
    
        Case 79916
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 79917
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 79925
        
        Case 79926
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", gErr, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143442)
    
    End Select
    
    Exit Function
    
End Function

Private Function Consulta_DebitosReceber(objFilialCliente As ClassFilialCliente) As Long
'Exibe o débito com cliente em sua tela de cadastro

Dim lErro As Long
Dim objDebitoReceber As New ClassDebitoRecCli

On Error GoTo Erro_Consulta_DebitosReceber
    
    'Se não existe linha selecionada no grid de devoluções => erro
    If GridDebitos.Row <= 0 Then gError 79921
    
    'Se a linha selecionada não contém dados => erro
    If Len(Trim(GridDebitos.TextMatrix(GridDebitos.Row, iGrid_NumTituloDB_Col))) = 0 Then gError 79922
    
    'Seleciona o objDebitoReceber referente à linha cujo documento será consultado
    Set objDebitoReceber = gcolDebitos.Item(GridDebitos.Row)
    
    'Chama a tela de cadastro de Débitos a Receber e exibe o débito selecionado
    Call Chama_Tela("DebitosReceb", objDebitoReceber)
    
    Consulta_DebitosReceber = SUCESSO
    
    Exit Function
    
Erro_Consulta_DebitosReceber:

    Consulta_DebitosReceber = gErr
    
    Select Case gErr
    
        Case 79921
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 79922
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 79923
        
        Case 79924
            Call Rotina_Erro(vbOKOnly, "ERRO_DEBITORECCLI_NAO_CADASTRADO", objDebitoReceber.lCliente, objDebitoReceber.iFilial, objDebitoReceber.sSiglaDocumento, objDebitoReceber.lNumTitulo, objDebitoReceber.dtDataEmissao, gErr)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143443)
    
    End Select
    
    Exit Function

End Function

Private Function Consulta_RecebAntecipados(objFilialCliente As ClassFilialCliente) As Long
'Exibe o Adiantamento à Cliente em sua tela de cadastro

Dim lErro As Long
Dim objAntecipReceb As New ClassAntecipReceb
Dim objMovContaCorrente As New ClassMovContaCorrente

On Error GoTo Erro_Consulta_RecebAntecipados
    
    'Se não existe linha selecionada no grid de adiantamentos=> erro
    If GridRecebAntecipados.Row <= 0 Then gError 79929
    
    'Se a linha selecionada não contém dados => erro
    If Len(Trim(GridRecebAntecipados.TextMatrix(GridRecebAntecipados.Row, iGrid_ContaCorrenteRA_Col))) = 0 Then gError 79930
    
    Set objAntecipReceb = gcolRecebAntecipado.Item(GridRecebAntecipados.Row)
    
    'Exibe o Adiantamento em sua tela de cadastro
    Call Chama_Tela("AntecipReceb", objAntecipReceb)
    
    Consulta_RecebAntecipados = SUCESSO
    
    Exit Function
    
Erro_Consulta_RecebAntecipados:

    Consulta_RecebAntecipados = gErr
    
    Select Case gErr
    
        Case 79929
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 79930
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143444)
    
    End Select
    
    Exit Function
    
End Function

'
'Private Function Carrega_Contabilizacao_Subtipo(iCodigoMotivo As Integer) As Long
''Inicializa a contabilidade e carrega os modelos contábeis de acordo com o tipo de baixa selecionado
'
'Dim lErro As Long
'Dim objMotivosBaixa As New ClassMotivosBaixa
'
'On Error GoTo Erro_Carrega_Contabilizacao_Subtipo
'
'    'Define o parâmetro que será passado para MotivosBaixa_Le
'    objMotivosBaixa.iCodigo = iCodigoMotivo
'
'    'Lê a tabela de MotivosBaixa e obtém o subtipo contábil
'    lErro = CF("MotivosBaixa_Le",objMotivosBaixa)
'    If lErro <> SUCESSO And lErro <> 79660 Then gError 79661
'
'    'Se não encontrou o motivo => erro
'    If lErro = 79660 Then gError 79663
'
'    'Inicializa a contabilidade
'    lErro = objContabil.Contabil_Processa_Alteracao_Subtipo(objMotivosBaixa.iSubtipo)
'    If lErro <> SUCESSO Then gError 79662
'
'    Carrega_Contabilizacao_Subtipo = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Contabilizacao_Subtipo:
'
'    Carrega_Contabilizacao_Subtipo = gErr
'
'    Select Case gErr
'
'        Case 79661, 79662
'
'        Case 79663
'            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVOBAIXA_NAO_ENCONTRADO", gErr, objMotivosBaixa.iCodigo)
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143445)
'
'    End Select
'
'    Exit Function
'
'End Function


Public Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134027

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134027

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 143446)

    End Select
    
    Exit Sub

End Sub

Private Sub ClienteItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub ClienteItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub ClienteItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ClienteItem
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub FilialClienteItem_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub FilialClienteItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Private Sub FilialClienteItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Private Sub FilialClienteItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = FilialClienteItem
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Valida_Clientes_Diferentes() As Long

Dim lErro As Long
Dim iInd1 As Integer
Dim iInd2 As Integer
Dim lCli1 As Long
Dim lCli2 As Long

On Error GoTo Erro_Valida_Clientes_Diferentes

    'Para cada linha
    For iInd1 = 1 To objGridParcelas.iLinhasExistentes
        
        'Obtém o cliente
        lCli1 = LCodigo_Extrai(GridParcelas.TextMatrix(iInd1, iGrid_Cliente_Col))
        
        'Se a linha estiver marcada
        If StrParaInt(GridParcelas.TextMatrix(iInd1, iGrid_Baixar_Col)) <> 0 Then
        
            'Para cada outra linha
            For iInd2 = 1 To objGridParcelas.iLinhasExistentes
            
                'Obtém o cliente
                lCli2 = LCodigo_Extrai(GridParcelas.TextMatrix(iInd2, iGrid_Cliente_Col))
            
                'se a linha estiver selecionada
                If StrParaInt(GridParcelas.TextMatrix(iInd2, iGrid_Baixar_Col)) <> 0 Then
            
                    'Se tiver clientes diferentes marcados
                    If lCli1 <> lCli2 Then gError 140295
                    
                End If
                
            Next
            
        End If
        
    Next
    
    Valida_Clientes_Diferentes = SUCESSO

    Exit Function

Erro_Valida_Clientes_Diferentes:

    Valida_Clientes_Diferentes = gErr

    Select Case gErr
    
        Case 140295

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function
'##############################################################################


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

Function BR_Traz_Parcela_Tela(ByVal objCT As Object, ByVal lNumIntParc As Long, ByVal iLinha As Integer) As Long
    BR_Traz_Parcela_Tela = SUCESSO
End Function

Sub GridParcelas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Sub GridRecebAntecipados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Seta o Grid Parcelas como grid atual
    iGridAtual = GRID_ADIANTAMENTOS
        
    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
        Set PopUpMenuGridMD.objTela = Me
        PopUpMenuGridMD.PopupMenu PopUpMenuGridMD.mnuGrid
        Set PopUpMenuGridMD.objTela = Nothing
    End If

End Sub

Sub GridDebitos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Faz com que apareca um PopupMenu o botao direito do mouse acionado sobre o grid

    'Seta o Grid Parcelas como grid atual
    iGridAtual = GRID_DEVOLUCOES
        
    'Verifica se foi o botao direito do mouse que foi pressionado
    If Button = vbRightButton Then
        Set PopUpMenuGridMD.objTela = Me
        PopUpMenuGridMD.PopupMenu PopUpMenuGridMD.mnuGrid
        Set PopUpMenuGridMD.objTela = Nothing
    End If

End Sub

Function BR_Traz_RA_Tela(ByVal objCT As Object, ByVal objRecebAntecipados As ClassAntecipReceb, ByVal iLinha As Integer) As Long
    BR_Traz_RA_Tela = SUCESSO
End Function

Function BR_Carrega_Tab_Parcelas(ByVal objCT As Object, lCodCli As Long, iFilialCli As Integer, dtEmissaoInic As Date, dtEmissaoFim As Date, dtVencInic As Date, dtVencFim As Date, lTituloInic As Long, lTituloFim As Long, ByVal objFilialCliente As ClassFilialCliente, ByVal gcolInfoParcRec As Collection, ByVal gcolRecebAntecipado As Collection, ByVal lCodViagem As Long, ByVal iFormaPagamento As Integer, ByVal iVendedor As Integer, ByVal iCobrador As Integer) As Long

Dim lErro As Long
Dim sTipoDoc As String
Dim sProd_I As String, iProdPreenchido_I As Integer, sProd_F As String, iProdPreenchido_F As Integer

On Error GoTo Erro_BR_Carrega_Tab_Parcelas

    If TipoDocApenas.Value = True Then
        If TipoDocSeleciona.ListIndex = -1 Then gError 182898
        sTipoDoc = SCodigo_Extrai(TipoDocSeleciona.Text)
    Else
        sTipoDoc = ""
    End If
    
    'formata o Produto Inicial
    lErro = CF("Produto_Formata", ProdutoInicial.Text, sProd_I, iProdPreenchido_I)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If iProdPreenchido_I <> PRODUTO_PREENCHIDO Then sProd_I = ""

    'formata o Produto Final
    lErro = CF("Produto_Formata", ProdutoFinal.Text, sProd_F, iProdPreenchido_F)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If iProdPreenchido_F <> PRODUTO_PREENCHIDO Then sProd_F = ""
    
    'Preenche a Coleção de Parcelas
    lErro = CF("ParcelasRec_Le_BaixaManualSel", lCodCli, iFilialCli, dtEmissaoInic, dtEmissaoFim, dtVencInic, dtVencFim, lTituloInic, lTituloFim, gcolInfoParcRec, sTipoDoc, lCodViagem, iFormaPagamento, iVendedor, iCobrador, sProd_I, sProd_F)
    If lErro <> SUCESSO Then gError 182897

    'Preenche a Coleção de Recebimentos antecipados
    lErro = CF("RecebAntecipados_Le_FilialCliente_Baixa", objFilialCliente, gcolRecebAntecipado)
    If lErro <> SUCESSO And lErro <> 29142 Then gError 182898
    
    BR_Carrega_Tab_Parcelas = SUCESSO

    Exit Function

Erro_BR_Carrega_Tab_Parcelas:

    BR_Carrega_Tab_Parcelas = gErr

    Select Case gErr
    
        Case 182897, 182898

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 182899)

    End Select

    Exit Function

End Function

Function BR_Form_Load(ByVal objCT As Object) As Long
    BR_Form_Load = SUCESSO
End Function

Function BR_Limpa_Tela_BaixaRec(ByVal objCT As Object) As Long
    BR_Limpa_Tela_BaixaRec = SUCESSO
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

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    objComboBox.Clear

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 185032
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 185032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185033)

    End Select

    Exit Function

End Function

Public Sub TipoDocApenas_Click()

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = True
    
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDocTodos_Click()

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona.ListIndex = COMBO_INDICE

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDocSeleciona_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigoViagem_evSelecao(obj1 As Object)

Dim lErro As Long

On Error GoTo Erro_objEventoCodigoViagem_evSelecao

    If obj1.lCodigo <> 0 Then
        CodigoViagem.PromptInclude = False
        CodigoViagem.Text = CStr(obj1.lCodigo)
        CodigoViagem.PromptInclude = True
    End If

    Me.Show

    Exit Sub

Erro_objEventoCodigoViagem_evSelecao:

    Select Case gErr

        Case 205311

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205312)

    End Select

    Exit Sub

End Sub

Public Sub LabelViagem_Click()

Dim lErro As Long
'Dim objMapaDeEntrega As New ClassMapaDeEntrega
Dim colSelecao As New Collection

On Error GoTo Erro_LabelViagem_Click

    Call Chama_Tela("MapaDeEntregaLista", colSelecao, Nothing, objEventoCodigoViagem)

    Exit Sub

Erro_LabelViagem_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 205313)

    End Select

    Exit Sub

End Sub

Public Sub FormaPagto_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CodigoViagem_GotFocus()
Dim iFrameAux As Integer
    
    iFrameAux = iFramePrincipalAlterado
    Call MaskEdBox_TrataGotFocus(CodigoViagem, iAlterado)
    iFramePrincipalAlterado = iFrameAux

End Sub

Public Sub CodigoViagem_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub CodigoViagem_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CodigoViagem_Validate

    'Testa para ver se a viagem existe

    Exit Sub

Erro_CodigoViagem_Validate:

    Cancel = True

    Select Case gErr

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 143404)

    End Select

    Exit Sub

End Sub

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = 1

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim dPercComissao As Double

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If Len(Trim(Vendedor.Text)) <> 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Vendedor, objVendedor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209207)
    
    End Select

End Sub

Public Sub VendedorLabel_Click()

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    iAlterado = 0

    Me.Show

    Vendedor.SetFocus 'Inserido por Wagner
    
    Exit Sub

End Sub

Public Sub CobradorFiltro_Change()

    iFramePrincipalAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Function Carrega_Cobrador()

Dim lErro As Long
Dim ColCobrador As New Collection
Dim objCobrador As ClassCobrador

On Error GoTo Erro_Carrega_Cobrador

    CobradorFiltro.AddItem ""
    CobradorFiltro.ItemData(CobradorFiltro.NewIndex) = 0

    lErro = CF("Cobradores_Le_Todos_Filial", ColCobrador)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objCobrador In ColCobrador
    
        If objCobrador.iCodigo <> COBRADOR_PROPRIA_EMPRESA And objCobrador.iInativo <> Inativo Then
        
            CobradorFiltro.AddItem objCobrador.iCodigo & SEPARADOR & objCobrador.sNomeReduzido
            CobradorFiltro.ItemData(CobradorFiltro.NewIndex) = objCobrador.iCodigo
            
        End If

    Next
    
    Carrega_Cobrador = SUCESSO

    Exit Function

Erro_Carrega_Cobrador:

    Carrega_Cobrador = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209207)
    
    End Select
    
    Exit Function
    
End Function

Public Sub NossoNumero_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub NossoNumero_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub NossoNumero_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = NossoNumero
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub objEventoProdutoAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoAte_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138328

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 138329
    
    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO Then gError 138330

    Me.Show

    Exit Sub

Erro_objEventoProdutoAte_evSelecao:

    Select Case gErr

        Case 138328, 138330

        Case 138329
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167047)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProdutoDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProdutoDe_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 138331

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 138332

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO Then gError 138333

    Me.Show

    Exit Sub

Erro_objEventoProdutoDe_evSelecao:

    Select Case gErr

        Case 138331, 138333

        Case 138332
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167048)

    End Select

    Exit Sub

End Sub

Public Sub LabelProdutoAte_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoAte_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoFinal.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoFinal.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 138334

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoAte)

    Exit Sub

Erro_LabelProdutoAte_Click:

    Select Case gErr

        Case 138334

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167049)

    End Select

    Exit Sub

End Sub

Public Sub LabelProdutoDe_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProdutoDe_Click

    'Verifica se o produto foi preenchido
    If Len(ProdutoInicial.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", ProdutoInicial.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 138335

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProdutoDe)

    Exit Sub

Erro_LabelProdutoDe_Click:

    Select Case gErr

        Case 138335

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167050)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoFinal_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoFinal, DescProdFim)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 138377
    
    If lErro <> SUCESSO Then gError 138378

    Exit Sub

Erro_ProdutoFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 138377

        Case 138378
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167068)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProdutoInicial_Validate

    lErro = CF("Produto_Perde_Foco", ProdutoInicial, DescProdInic)
    If lErro <> SUCESSO And lErro <> 27095 Then gError 138379
    
    If lErro <> SUCESSO Then gError 138380

    Exit Sub

Erro_ProdutoInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 138379

        Case 138380
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 167069)

    End Select

    Exit Sub

End Sub

Public Sub ProdutoInicial_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoFinal_Change()

    'Registra que houve alteração
    iAlterado = REGISTRO_ALTERADO
    iFramePrincipalAlterado = REGISTRO_ALTERADO

End Sub
