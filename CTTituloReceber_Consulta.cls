VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTituloRec_Consulta"
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

Public iAlterado As Integer
Dim iFrameRecebimentoAtual As Integer

'Início da contabilidade
Dim objGrid1 As AdmGrid
Public objContabil As New ClassContabil
'Fim da contabilidade

Dim iFrameAtual As Integer

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo
Dim gobjParcRecBaixa As New ClassBaixaParcRec
Dim gobjBaixaReceber As New ClassBaixaReceber

Dim objGridParcelas As AdmGrid
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_VenctoReal_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Status_Col As Integer
Dim iGrid_DescPrev_Col As Integer
    
Dim objGridDesconto As AdmGrid
Dim iGrid_TipoDesconto_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_ValorDesconto_Col As Integer
Dim iGrid_Percentual1_Col As Integer

Dim objGridComissoes As AdmGrid
Dim objGridComissoesEmissao As AdmGrid
Dim iGrid_Vendedor_Col As Integer
Dim iGrid_PercentualComissao_Col As Integer
Dim iGrid_ValorBase_Col As Integer
Dim iGrid_ValorComissao_Col As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoTipoDocumento As AdmEvento
Attribute objEventoTipoDocumento.VB_VarHelpID = -1

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

Public Property Get LabelNatureza() As Object
     Set LabelNatureza = objUserControl.Controls("LabelNatureza")
End Property

Public Property Get LabelNaturezaDesc() As Object
     Set LabelNaturezaDesc = objUserControl.Controls("LabelNaturezaDesc")
End Property

Public Property Get ReajustadoAte() As Object
     Set ReajustadoAte = objUserControl.Controls("ReajustadoAte")
End Property

'Public Property Get LabelDescCpoGenerico1() As Object
'     Set LabelDescCpoGenerico1 = objUserControl.Controls("LabelDescCpoGenerico1")
'End Property

Public Property Get LabelCcl() As Object
     Set LabelCcl = objUserControl.Controls("LabelCcl")
End Property

Public Property Get DescPrev() As Object
     Set DescPrev = objUserControl.Controls("DescPrev")
End Property

Public Property Get ReajustePeriodicidade() As Object
     Set ReajustePeriodicidade = objUserControl.Controls("ReajustePeriodicidade")
End Property

Public Property Get Moeda() As Object
     Set Moeda = objUserControl.Controls("Moeda")
End Property

Public Property Get ReajusteBase() As Object
     Set ReajusteBase = objUserControl.Controls("ReajusteBase")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get BotaoDocOriginal() As Object
     Set BotaoDocOriginal = objUserControl.Controls("BotaoDocOriginal")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get INSSRetido() As Object
     Set INSSRetido = objUserControl.Controls("INSSRetido")
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

Public Property Get StatusParcela() As Object
     Set StatusParcela = objUserControl.Controls("StatusParcela")
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

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get CondPagtoLabel() As Object
     Set CondPagtoLabel = objUserControl.Controls("CondPagtoLabel")
End Property

Public Property Get ValorEmissao() As Object
     Set ValorEmissao = objUserControl.Controls("ValorEmissao")
End Property

Public Property Get ValorBaseEmissao() As Object
     Set ValorBaseEmissao = objUserControl.Controls("ValorBaseEmissao")
End Property

Public Property Get PercentualEmissao() As Object
     Set PercentualEmissao = objUserControl.Controls("PercentualEmissao")
End Property

Public Property Get VendedorEmissao() As Object
     Set VendedorEmissao = objUserControl.Controls("VendedorEmissao")
End Property

Public Property Get GridComissoesEmissao() As Object
     Set GridComissoesEmissao = objUserControl.Controls("GridComissoesEmissao")
End Property

Public Property Get TotalPercentualEmissao() As Object
     Set TotalPercentualEmissao = objUserControl.Controls("TotalPercentualEmissao")
End Property

Public Property Get TotalValorEmissao() As Object
     Set TotalValorEmissao = objUserControl.Controls("TotalValorEmissao")
End Property

Public Property Get LabelTotaisEmissao() As Object
     Set LabelTotaisEmissao = objUserControl.Controls("LabelTotaisEmissao")
End Property

Public Property Get Saldo() As Object
     Set Saldo = objUserControl.Controls("Saldo")
End Property

Public Property Get JurosMensais() As Object
     Set JurosMensais = objUserControl.Controls("JurosMensais")
End Property

Public Property Get PercMulta() As Object
     Set PercMulta = objUserControl.Controls("PercMulta")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get ValorTitulo() As Object
     Set ValorTitulo = objUserControl.Controls("ValorTitulo")
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

Public Property Get FrameRecebimento() As Object
     Set FrameRecebimento = objUserControl.Controls("FrameRecebimento")
End Property

Public Property Get ContaCorrente() As Object
     Set ContaCorrente = objUserControl.Controls("ContaCorrente")
End Property

Public Property Get ValorRecebimento() As Object
     Set ValorRecebimento = objUserControl.Controls("ValorRecebimento")
End Property

Public Property Get Historico() As Object
     Set Historico = objUserControl.Controls("Historico")
End Property

Public Property Get HistoricoPerda() As Object
     Set HistoricoPerda = objUserControl.Controls("HistoricoPerda")
End Property

Public Property Get DataEmissaoCred() As Object
     Set DataEmissaoCred = objUserControl.Controls("DataEmissaoCred")
End Property

Public Property Get SiglaDocumentoCR() As Object
     Set SiglaDocumentoCR = objUserControl.Controls("SiglaDocumentoCR")
End Property

Public Property Get ValorDebito() As Object
     Set ValorDebito = objUserControl.Controls("ValorDebito")
End Property

Public Property Get FilialEmpresaCR() As Object
     Set FilialEmpresaCR = objUserControl.Controls("FilialEmpresaCR")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
End Property

Public Property Get SaldoDebito() As Object
     Set SaldoDebito = objUserControl.Controls("SaldoDebito")
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

Public Property Get CCIntNomeReduzido() As Object
     Set CCIntNomeReduzido = objUserControl.Controls("CCIntNomeReduzido")
End Property

Public Property Get Frame4() As Object
     Set Frame4 = objUserControl.Controls("Frame4")
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

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
End Property

Public Property Get Recebimento() As Object
     Set Recebimento = objUserControl.Controls("Recebimento")
End Property

Public Property Get Sequencial() As Object
     Set Sequencial = objUserControl.Controls("Sequencial")
End Property

Public Property Get UpDownParcela() As Object
     Set UpDownParcela = objUserControl.Controls("UpDownParcela")
End Property

Public Property Get ValorParcBaixa() As Object
     Set ValorParcBaixa = objUserControl.Controls("ValorParcBaixa")
End Property

Public Property Get DataVencParcBaixa() As Object
     Set DataVencParcBaixa = objUserControl.Controls("DataVencParcBaixa")
End Property

Public Property Get Parcela() As Object
     Set Parcela = objUserControl.Controls("Parcela")
End Property

Public Property Get Frame3() As Object
     Set Frame3 = objUserControl.Controls("Frame3")
End Property

Public Property Get Cobrador() As Object
     Set Cobrador = objUserControl.Controls("Cobrador")
End Property

Public Property Get Carteira() As Object
     Set Carteira = objUserControl.Controls("Carteira")
End Property

Public Property Get StatusParc() As Object
     Set StatusParc = objUserControl.Controls("StatusParc")
End Property

Public Property Get SaldoParc() As Object
     Set SaldoParc = objUserControl.Controls("SaldoParc")
End Property

Public Property Get ValorParc() As Object
     Set ValorParc = objUserControl.Controls("ValorParc")
End Property

Public Property Get ValorOriginalParc() As Object
     Set ValorOriginalParc = objUserControl.Controls("ValorOriginalParc")
End Property

Public Property Get DataVenctParcReal() As Object
     Set DataVenctParcReal = objUserControl.Controls("DataVenctParcReal")
End Property

Public Property Get DataVenctoParc() As Object
     Set DataVenctoParc = objUserControl.Controls("DataVenctoParc")
End Property

Public Property Get ValorComissao() As Object
     Set ValorComissao = objUserControl.Controls("ValorComissao")
End Property

Public Property Get ValorBase() As Object
     Set ValorBase = objUserControl.Controls("ValorBase")
End Property

Public Property Get PercentualComissao() As Object
     Set PercentualComissao = objUserControl.Controls("PercentualComissao")
End Property

Public Property Get Vendedor() As Object
     Set Vendedor = objUserControl.Controls("Vendedor")
End Property

Public Property Get GridComissoes() As Object
     Set GridComissoes = objUserControl.Controls("GridComissoes")
End Property

Public Property Get LabelTotaisComissoes() As Object
     Set LabelTotaisComissoes = objUserControl.Controls("LabelTotaisComissoes")
End Property

Public Property Get TotalValorComissao() As Object
     Set TotalValorComissao = objUserControl.Controls("TotalValorComissao")
End Property

Public Property Get TotalPercentualComissao() As Object
     Set TotalPercentualComissao = objUserControl.Controls("TotalPercentualComissao")
End Property

Public Property Get TipoDesconto() As Object
     Set TipoDesconto = objUserControl.Controls("TipoDesconto")
End Property

Public Property Get Percentual1() As Object
     Set Percentual1 = objUserControl.Controls("Percentual1")
End Property

Public Property Get ValorDesconto() As Object
     Set ValorDesconto = objUserControl.Controls("ValorDesconto")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

Public Property Get GridDescontos() As Object
     Set GridDescontos = objUserControl.Controls("GridDescontos")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
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

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

Public Property Get NumeroTitulo() As Object
     Set NumeroTitulo = objUserControl.Controls("NumeroTitulo")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get Status() As Object
     Set Status = objUserControl.Controls("Status")
End Property

Public Property Get Line1() As Object
     Set Line1 = objUserControl.Controls("Line1")
End Property

Public Property Get TipoDocumentoLabel() As Object
     Set TipoDocumentoLabel = objUserControl.Controls("TipoDocumentoLabel")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get ClienteEtiqueta() As Object
     Set ClienteEtiqueta = objUserControl.Controls("ClienteEtiqueta")
End Property

'####################################
'Inserido por Wagner
Public Property Get NossoNumero() As Object
     Set NossoNumero = objUserControl.Controls("NossoNumero")
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
'####################################

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

Public Property Get FilialEmpresa() As Object
     Set FilialEmpresa = objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get LabelFilialEmpresa() As Object
     Set LabelFilialEmpresa = objUserControl.Controls("LabelFilialEmpresa")
End Property

'--- fim dos properties get dos controles da tela

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load

    'Carrega os Tipos de Documento
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then Error 58871

    'Carrega os Tipos de Desconto
    lErro = Carrega_TipoDesconto()
    If lErro <> SUCESSO Then Error 58872

    'Inicializa o frame atual
    iFrameAtual = 1
    iFrameRecebimentoAtual = 0

    'Inicializa os Eventos da Tela
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoTipoDocumento = New AdmEvento

    'Inicializa os Grids da tela
    Set objGridParcelas = New AdmGrid
    Set objGridDesconto = New AdmGrid
    Set objGridComissoes = New AdmGrid
    Set objGridComissoesEmissao = New AdmGrid
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Faz as inicializações particulares ao GridParcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 58873

    'Faz as inicializações particulares ao GridDescontos
    lErro = Inicializa_Grid_Descontos(objGridDesconto)
    If lErro <> SUCESSO Then Error 58874

    'Faz as inicializações particulares ao GridComissoes
    lErro = Inicializa_Grid_Comissoes(objGridComissoes)
    If lErro <> SUCESSO Then Error 58875

    'Faz as inicializações particulares ao GridComissoesEmissao
    lErro = Inicializa_Grid_ComissoesEmissao(objGridComissoesEmissao)
    If lErro <> SUCESSO Then Error 58876

    DataEmissao.Caption = ""

    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade_Consulta(Me, objGrid1, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 58877

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 58871, 58872, 58873, 58874, 58875, 58876, 58877

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158369)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosRecTodos"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objTituloReceber)
    If lErro <> SUCESSO Then Error 58878

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tiRecem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDoc", objTituloReceber.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "Cliente", objTituloReceber.lCliente, 0, "Cliente"
    colCampoValor.Add "Filial", objTituloReceber.iFilial, 0, "Filial"
    colCampoValor.Add "NumTitulo", objTituloReceber.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objTituloReceber.dtDataEmissao, DATA_NULA, "DataEmissao"
    colCampoValor.Add "NumParcelas", objTituloReceber.iNumParcelas, 0, "NumParcelas"
    colCampoValor.Add "Valor", objTituloReceber.dValor, 0, "Valor"
    colCampoValor.Add "ValorIRRF", objTituloReceber.dValorIRRF, 0, "ValorIRRF"
    colCampoValor.Add "ValorINSS", objTituloReceber.dValorINSS, 0, "ValorINSS"
    colCampoValor.Add "INSSRetido", objTituloReceber.iINSSRetido, 0, "INSSRetido"
    colCampoValor.Add "SiglaDocumento", objTituloReceber.sSiglaDocumento, STRING_BUFFER_MAX_TEXTO, "SiglaDocumento"
    colCampoValor.Add "PercJurosDiarios", objTituloReceber.dPercJurosDiario, 0, "PercJurosDiarios"
    colCampoValor.Add "PercMulta", objTituloReceber.dPercMulta, 0, "PercMulta"
    colCampoValor.Add "Status", objTituloReceber.iStatus, 0, "Status"
    colCampoValor.Add "Saldo", objTituloReceber.dSaldo, 0, "Saldo"
    colCampoValor.Add "PISRetido", objTituloReceber.dPISRetido, 0, "PISRetido"
    colCampoValor.Add "ISSRetido", objTituloReceber.dISSRetido, 0, "ISSRetido"
    colCampoValor.Add "COFINSRetido", objTituloReceber.dCOFINSRetido, 0, "COFINSRetido"
    colCampoValor.Add "CSLLRetido", objTituloReceber.dCSLLRetido, 0, "CSLLRetido"
    
    colCampoValor.Add "Moeda", objTituloReceber.iMoeda, 0, "Moeda"
    colCampoValor.Add "Ccl", objTituloReceber.sCcl, STRING_CCL, "Ccl"
''    colCampoValor.Add "CpoGenerico1", objTituloReceber.lCpoGenerico1, 0, "CpoGenerico1"
    colCampoValor.Add "ReajustePeriodicidade", objTituloReceber.iReajustePeriodicidade, 0, "ReajustePeriodicidade"
    colCampoValor.Add "ReajusteBase", objTituloReceber.dtReajusteBase, 0, "ReajusteBase"
    colCampoValor.Add "Natureza", objTituloReceber.sNatureza, STRING_NATMOVCTA_CODIGO, "Natureza"

    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, iFilialEmpresaTela

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 58878

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158370)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber

On Error GoTo Erro_Tela_Preenche

    objTituloReceber.lNumIntDoc = colCampoValor.Item("NumIntDoc").vValor

    If objTituloReceber.lNumIntDoc <> 0 Then

        'Carrega objTituloReceber com os dados passados em colCampoValor
        objTituloReceber.lCliente = colCampoValor.Item("Cliente").vValor
        objTituloReceber.iFilial = colCampoValor.Item("Filial").vValor
        objTituloReceber.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objTituloReceber.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objTituloReceber.dValor = colCampoValor.Item("Valor").vValor
        objTituloReceber.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        objTituloReceber.dValorINSS = colCampoValor.Item("ValorINSS").vValor
        objTituloReceber.iINSSRetido = colCampoValor.Item("INSSRetido").vValor
        objTituloReceber.sSiglaDocumento = colCampoValor.Item("SiglaDocumento").vValor
        objTituloReceber.dPercJurosDiario = colCampoValor.Item("PercJurosDiarios").vValor
        objTituloReceber.dPercMulta = colCampoValor.Item("PercMulta").vValor
        objTituloReceber.iStatus = colCampoValor.Item("Status").vValor
        objTituloReceber.dSaldo = colCampoValor.Item("Saldo").vValor
        objTituloReceber.dPISRetido = colCampoValor.Item("PISRetido").vValor
        objTituloReceber.dISSRetido = colCampoValor.Item("ISSRetido").vValor
        objTituloReceber.dCOFINSRetido = colCampoValor.Item("COFINSRetido").vValor
        objTituloReceber.dCSLLRetido = colCampoValor.Item("CSLLRetido").vValor

        objTituloReceber.iMoeda = colCampoValor.Item("Moeda").vValor
        objTituloReceber.sCcl = colCampoValor.Item("Ccl").vValor
''        objTituloReceber.lCpoGenerico1 = colCampoValor.Item("CpoGenerico1").vValor
        objTituloReceber.iReajustePeriodicidade = colCampoValor.Item("ReajustePeriodicidade").vValor
        objTituloReceber.dtReajusteBase = colCampoValor.Item("ReajusteBase").vValor
        objTituloReceber.sNatureza = colCampoValor.Item("Natureza").vValor

        'Coloca na Tela os Dados de objTituloReceber
        lErro = Traz_TituloReceber_Tela(objTituloReceber)
        If lErro <> SUCESSO Then Error 58879

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 58879 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158371)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objTituloReceber As ClassTituloReceber) As Long
'Recolhe os dados da tela para a memória

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iIndice As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o cliente foi digitado
    If Len(Trim(Cliente.ClipText)) > 0 Then
        objCliente.sNomeReduzido = Cliente.Text

        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then Error 58880

        'Não achou o Cliente --> erro
        If lErro <> SUCESSO Then Error 58881

        'Guarda o código no objTituloReceber
        objTituloReceber.lCliente = objCliente.lCodigo
    End If

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    End If

    'Recolhe os dados obrigatórios da Tela
    If Len(Trim(NumeroTitulo.ClipText)) > 0 Then objTituloReceber.lNumTitulo = CLng(NumeroTitulo.ClipText)

    If Len(Trim(DataEmissao.Caption)) = 0 Then
        objTituloReceber.dtDataEmissao = DATA_NULA
    Else
        objTituloReceber.dtDataEmissao = CDate(DataEmissao.Caption)
    End If

    'Recolhe os dados restantes
    If Len(Trim(ValorTitulo.Caption)) > 0 Then objTituloReceber.dValor = CDbl(ValorTitulo.Caption)
    If Len(Trim(ValorIRRF.Caption)) > 0 Then objTituloReceber.dValorIRRF = CDbl(ValorIRRF.Caption)
    If Len(Trim(ValorINSS.Caption)) > 0 Then objTituloReceber.dValorINSS = CDbl(ValorINSS.Caption)
    objTituloReceber.iINSSRetido = INSSRetido.Value
    objTituloReceber.iFilialEmpresa = iFilialEmpresaTela
    If Len(Trim(Tipo.Text)) > 0 Then objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    If Len(Trim(PISRetido.Caption)) > 0 Then objTituloReceber.dPISRetido = CDbl(PISRetido.Caption)
    If Len(Trim(ISSRetido.Caption)) > 0 Then objTituloReceber.dISSRetido = CDbl(ISSRetido.Caption)
    If Len(Trim(COFINSRetido.Caption)) > 0 Then objTituloReceber.dCOFINSRetido = CDbl(COFINSRetido.Caption)
    If Len(Trim(CSLLRetido.Caption)) > 0 Then objTituloReceber.dCSLLRetido = CDbl(CSLLRetido.Caption)

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 58880 'Tratado na rotina chamada

        Case 58881
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, Cliente.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158372)

    End Select

    Exit Function

End Function


Private Function Carrega_TipoDocumento()
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim iTipo As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then Error 58882

    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
        Tipo.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
        If objTipoDocumento.sSigla = OUTROS_RECEBIMENTOS Then iTipo = iIndice - 1
    
    Next

    Tipo.ListIndex = iTipo

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = Err

    Select Case Err

        Case 58882 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158373)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Status")
    objGridInt.colColuna.Add ("Observação")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    objGridInt.colCampo.Add (StatusParcela.Name)
    objGridInt.colCampo.Add (DescPrev.Name)
    
    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_Col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Status_Col = 4
    iGrid_DescPrev_Col = 5
    
    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 850

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Parcelas = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Descontos(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Descontos

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Percentual")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (TipoDesconto.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (ValorDesconto.Name)
    objGridInt.colCampo.Add (Percentual1.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridDescontos

    'Colunas do Grid
    iGrid_TipoDesconto_Col = 1
    iGrid_Data_Col = 2
    iGrid_ValorDesconto_Col = 3
    iGrid_Percentual1_Col = 4

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_DESCONTOS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridDescontos.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Descontos = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Comissoes(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Comissões

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Vendedor.Name)
    objGridInt.colCampo.Add (PercentualComissao.Name)
    objGridInt.colCampo.Add (ValorBase.Name)
    objGridInt.colCampo.Add (ValorComissao.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridComissoes

    'Colunas do Grid
    iGrid_Vendedor_Col = 1
    iGrid_PercentualComissao_Col = 2
    iGrid_ValorBase_Col = 3
    iGrid_ValorComissao_Col = 4

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_COMISSOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridComissoes.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalPercentualComissao.top = GridComissoes.top + GridComissoes.Height
    TotalPercentualComissao.left = GridComissoes.left
    For iIndice = 0 To 1
        TotalPercentualComissao.left = TotalPercentualComissao.left + GridComissoes.ColWidth(iIndice) + GridComissoes.GridLineWidth + 20
    Next

    TotalPercentualComissao.Width = GridComissoes.ColWidth(2)

    TotalValorComissao.top = TotalPercentualComissao.top
    TotalValorComissao.Width = GridComissoes.ColWidth(4)
    For iIndice = 0 To iGrid_ValorComissao_Col - 1
        TotalValorComissao.left = TotalPercentualComissao.left + TotalPercentualComissao.Width + GridComissoes.ColWidth(iIndice) + GridComissoes.GridLineWidth + 20
    Next

    LabelTotaisComissoes.top = TotalPercentualComissao.top + (TotalPercentualComissao.Height - LabelTotaisComissoes.Height) / 2
    LabelTotaisComissoes.left = TotalPercentualComissao.left - LabelTotaisComissoes.Width

    Inicializa_Grid_Comissoes = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_ComissoesEmissao(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Comissões Emissões

Dim iIndice As Integer

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Vendedor")
    objGridInt.colColuna.Add ("Percentual")
    objGridInt.colColuna.Add ("Valor Base")
    objGridInt.colColuna.Add ("Valor")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (VendedorEmissao.Name)
    objGridInt.colCampo.Add (PercentualEmissao.Name)
    objGridInt.colCampo.Add (ValorBaseEmissao.Name)
    objGridInt.colCampo.Add (ValorEmissao.Name)

    'Grid do GridInterno
    objGridInt.objGrid = GridComissoesEmissao

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_COMISSOES + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridComissoesEmissao.ColWidth(0) = 500

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    'Posiciona os painéis totalizadores
    TotalPercentualEmissao.top = GridComissoesEmissao.top + GridComissoesEmissao.Height
    TotalPercentualEmissao.left = GridComissoesEmissao.left
    For iIndice = 0 To 1
        TotalPercentualEmissao.left = TotalPercentualEmissao.left + GridComissoesEmissao.ColWidth(iIndice) + GridComissoesEmissao.GridLineWidth + 20
    Next

    TotalPercentualEmissao.Width = GridComissoesEmissao.ColWidth(2)

    TotalValorEmissao.top = TotalPercentualEmissao.top
    For iIndice = 0 To iGrid_ValorComissao_Col - 1
        TotalValorEmissao.left = TotalPercentualEmissao.left + TotalPercentualEmissao.Width + GridComissoesEmissao.ColWidth(iIndice) + GridComissoesEmissao.GridLineWidth + 20
    Next

    TotalValorEmissao.Width = GridComissoesEmissao.ColWidth(4)
    LabelTotaisEmissao.top = TotalPercentualEmissao.top + (TotalPercentualEmissao.Height - LabelTotaisEmissao.Height) / 2
    LabelTotaisEmissao.left = TotalPercentualEmissao.left - LabelTotaisEmissao.Width

    Inicializa_Grid_ComissoesEmissao = SUCESSO

    Exit Function

End Function

Public Sub BotaoConsulta_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim objCliente As New ClassCliente

On Error GoTo Erro_BotaoConsulta_Click

    'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then Error 58883

    objCliente.sNomeReduzido = Cliente.Text

    'Lê o codigo através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then Error 58887

    'Não achou o Cliente --> erro
    If lErro <> SUCESSO Then Error 58888

    objTituloReceber.lCliente = objCliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloReceber.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloReceber.iFilialEmpresa = iFilialEmpresaTela

    'Procura o Titulo (Baixados ou não)
    lErro = CF("TitulosRec_Le_Todos", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 58997 Then Error 58889

    If lErro = 58997 Then Error 58890

    'Traz o Titulo para a Tela
    lErro = Traz_TituloReceber_Tela(objTituloReceber)
    If lErro <> SUCESSO Then Error 58891

    Exit Sub

Erro_BotaoConsulta_Click:

    Select Case Err

        Case 58883, 58887, 58889, 58891

        Case 58888
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", Err, objCliente.sNomeReduzido)

        Case 58890
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", Err, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158374)

    End Select

    Exit Sub

End Sub

Public Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim sTela As String
Dim objCliente As New ClassCliente
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoDocOriginal_Click

    'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número)
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 58892

    objCliente.sNomeReduzido = Cliente.Text

    'Lê o codigo do Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 58896

    'Não achou o Cliente --> erro
    If lErro <> SUCESSO Then gError 58897
    
    If iFilialEmpresaTela <> giFilialEmpresa Then gError 185050

    objTituloReceber.lCliente = objCliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloReceber.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloReceber.iFilialEmpresa = iFilialEmpresaTela

    'Procura o Titulo (Baixados ou não)
    lErro = CF("TitulosRec_Le_Todos", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 58997 Then gError 58898

    If lErro = 58997 Then gError 58899

    'Se for do Tipo "OR" dá mensagem pois este Título não tem documento original
    If objTituloReceber.sSiglaDocumento = OUTROS_RECEBIMENTOS Then gError 58900

    'Se for do Tipo Fatura a Receber chama a Tela de Browser Listando todas as Notas da Fatura a Receber que originou este Título
    If objTituloReceber.sSiglaDocumento = TIPODOC_FATURA_A_RECEBER Then

        colSelecao.Add objTituloReceber.lNumIntDoc

        Call Chama_Tela("NF_FaturaReceberLista", colSelecao)

    Else

        'Le o Nome da Tela que originou este Título
        lErro = CF("Titulo_Le_DocumentoOriginal", objTituloReceber.lNumIntDoc, CPR_TITULO_RECEBER, objNFiscal, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError 58901

        If lErro = SUCESSO Then
'            If giTipoVersao = VERSAO_FULL Then
                'Chama a Tela
                Call Chama_Tela(sTela, objNFiscal)
'            ElseIf giTipoVersao = VERSAO_LIGHT Then
'                'Chama a Tela
'                Call Chama_Tela("NFiscalFatura", objNFiscal)
'            End If
        Else
            gError 58943
        End If

    End If

    Exit Sub

Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 58892, 58896, 58898, 58901

        Case 58897
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)

        Case 58899
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", gErr, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)

        Case 58900, 58943
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_SEM_DOCORIGINAL", gErr)

        Case 185050
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_DIFERENTE", gErr, iFilialEmpresaTela, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158375)

    End Select

    Exit Sub

End Sub

Public Sub CTBGridContabil_Click()

    Call objContabil.Contabil_GridContabil_Consulta_Click

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 58957

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then Error 58958

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 58959

        If lErro = 17660 Then Error 58960

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

    End If

    'Não encontrou a STRING
    If lErro = 6731 Then Error 58961

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case Err

        Case 58957, 58959

        Case 58958
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 58960
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 58961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158376)

    End Select

    Exit Sub

End Sub

Public Sub NumeroTitulo_GotFocus()

    Call MaskEdBox_TrataGotFocus(NumeroTitulo)

End Sub

Private Sub objEventoTipoDocumento_evSelecao(obj1 As Object)

Dim objTipoDocumento As ClassTipoDocumento

    Set objTipoDocumento = obj1

    Tipo.Text = objTipoDocumento.sSigla

    Me.Show

    Exit Sub

End Sub

Public Sub ClienteEtiqueta_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objCliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente, Cancel As Boolean

    Set objCliente = obj1

    'Preenche o Cliente com o Cliente selecionado
    Cliente.Text = objCliente.sNomeReduzido

    Call Cliente_Validate(Cancel)

    Me.Show

    Exit Sub

End Sub

Public Sub NumeroLabel_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim colParcelas As New colParcelaPagar
Dim colSelecao As New Collection

On Error GoTo Erro_NumeroLabel_Click

    'Se Cliente não estiver Preenchido ---> Erro
    If Len(Trim(Cliente.Text)) = 0 Then Error 58902

    'Se Filial não estiver Preenchida --> Erro
    If Len(Trim(Filial.Text)) = 0 Then Error 58903

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Error 58904

    'Move os dados da Tela para objtituloreceber e colParcelas
    lErro = Move_Tela_Memoria(objTituloReceber)
    If lErro <> SUCESSO Then Error 58905

    'Adiciona filtros: lCliente, iFilial e Tipo de Documento
    colSelecao.Add objTituloReceber.lCliente
    colSelecao.Add objTituloReceber.iFilial
    colSelecao.Add objTituloReceber.sSiglaDocumento

    'Chama Tela TituloReceberLista
    Call Chama_Tela("TitulosRecTodosLista", colSelecao, objTituloReceber, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 58902
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 58903
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 58904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)

        Case 58905

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158377)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloReceber As ClassTituloReceber

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloReceber = obj1

    'Traz os dados de objTituloReceber para tela
    lErro = Traz_TituloReceber_Tela(objTituloReceber)
    If lErro <> SUCESSO Then Error 58906

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 58906

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158378)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim objTipoDocumento As New ClassTipoDocumento
Dim iIndice As Integer
Dim bAchou As Boolean

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloReceber Is Nothing) Then

        'Lê o Título
        lErro = CF("TituloReceber_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 26061 Then gError 58907

        'Se não achou o Título
        If lErro <> SUCESSO Then

            'Lê o Título a Receber baixado
            lErro = CF("TituloReceberBaixado_Le", objTituloReceber)
            If lErro <> SUCESSO And lErro <> 56570 Then gError 58908

            If lErro = 56570 Then gError 58909

        End If

        'Traz os dados para a Tela
        lErro = Traz_TituloReceber_Tela(objTituloReceber)
        If lErro <> SUCESSO Then gError 58910

        If objTituloReceber.iParcelaTela <> 0 Then
        
            'Torna Frame atual invisível
            Frame1(Opcao.SelectedItem.Index).Visible = False
            If objTituloReceber.iBaixaTela = 0 Then
                iFrameAtual = TAB_Parcelas
            Else
                iFrameAtual = TAB_Baixas
            End If
            'Torna Frame atual visível
            Frame1(iFrameAtual).Visible = True
            Opcao.Tabs.Item(iFrameAtual).Selected = True
            
            Call Opcao_Click
            
            Parcela(1).Caption = objTituloReceber.iParcelaTela - 1
            
            Call UpDownParcela_UpClick(1)
            
            If objTituloReceber.iBaixaTela <> 0 Then
                
                iIndice = 0
                bAchou = False
                For iIndice = 0 To Sequencial.ListCount - 1
                    If StrParaInt(Sequencial.List(iIndice)) = objTituloReceber.iBaixaTela Then
                        Sequencial.ListIndex = iIndice
                        bAchou = True
                        Exit For
                    End If
                Next
                
                If Not bAchou Then gError 189364
            
                Call Sequencial_Click
            End If
        
        
        End If


    End If

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case gErr

        Case 58907, 58908, 58910

        Case 58909
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", gErr, objTituloReceber.lNumIntDoc)

        Case 189364 'ERRO_BAIXA_PARCELA_CANCELADA
            Call Rotina_Erro(vbOKOnly, "ERRO_BAIXA_PARCELA_CANCELADA", gErr, objTituloReceber.iParcelaTela, objTituloReceber.iBaixaTela)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158379)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_Tela_TituloReceber

    Call Limpa_Tab_Baixa

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158380)

    End Select

    Exit Sub

End Sub

Private Function Critica_CamposPreenchidos() As Long

Dim lErro As Long

On Error GoTo Erro_Critica_CamposPreenchidos

    'Se Cliente estiver vazio, erro
    If Len(Trim(Cliente.Text)) = 0 Then Error 58911

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then Error 58912

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Error 58913

    'Verifica se o Titulo está Preenchido
    If Len(Trim(NumeroTitulo.Text)) = 0 Then Error 58914

    Critica_CamposPreenchidos = SUCESSO

    Exit Function

Erro_Critica_CamposPreenchidos:

    Critica_CamposPreenchidos = Err

    Select Case Err

        Case 58911
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 58912
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)

        Case 58913
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)

        Case 58914
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158381)

    End Select

    Exit Function

End Function

Public Sub Sequencial_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim objParcelaReceber As New ClassParcelaReceber
Dim iParcela As Integer

On Error GoTo Erro_Sequencial_Click

    If Sequencial.ListIndex = -1 Then Exit Sub

    'Critica os Campos se estão Preenchidos
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then Error 58940

    'Verifica se já existe alguma Parcela na Tela
    If Len(Trim(Parcela(1).Caption)) > 0 Then
        iParcela = CInt(Parcela(1).Caption)
    Else
        iParcela = 0
    End If

    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloReceber.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloReceber.iFilialEmpresa = iFilialEmpresaTela
    objParcelaReceber.iNumParcela = iParcela

    lErro = CF("ParcelasRec_Le_Todas_NumTitulo", objTituloReceber, objParcelaReceber)
    If lErro <> SUCESSO And lErro <> 60003 Then Error 58915

    If lErro = 60003 Then Error 58916

    'Se for uma Parcela Baixada Traz os dados da Baixa
    If Sequencial.ListCount > 0 Then

        'Carrega os Dados da Baixa (1 como Default)
        Call Carrega_Dados_Baixa(objParcelaReceber.iNumParcela, Sequencial.List(Sequencial.ListIndex), objParcelaReceber.lNumIntTitulo)

    Else
        Call Limpa_Tab_Baixa
    End If

    Exit Sub

Erro_Sequencial_Click:

    Select Case Err

        Case 58915, 58940

        Case 58916
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", Err)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158382)

    End Select

    Exit Sub

End Sub

Public Sub TipoDocumentoLabel_Click()

Dim colSelecao As Collection
Dim objTipoDocumento As New ClassTipoDocumento

    objTipoDocumento.sSigla = SCodigo_Extrai(Tipo.Text)

    Call Chama_Tela("TipoDocTituloRecLista", colSelecao, objTipoDocumento, objEventoTipoDocumento)

End Sub

Private Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 58917

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 58917

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158383)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a Data de Emissão em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 58918

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 58918

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158384)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub NumeroTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumeroTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumeroTitulo.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumeroTitulo.ClipText)
    If lErro <> SUCESSO Then Error 58919

    Exit Sub

Erro_NumeroTitulo_Validate:

    Cancel = True


    Select Case Err

        Case 58919

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158385)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer, bCancel As Boolean
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
        If lErro <> SUCESSO Then Error 58920

        lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
        If lErro <> SUCESSO Then Error 58921

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case Err

        Case 58920, 58921 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158386)

    End Select

    Exit Sub

End Sub

Public Sub Opcao_Click()

Dim lErro As Long

On Error GoTo Erro_Opcao_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

    Frame1(Opcao.SelectedItem.Index).Visible = True
    Frame1(iFrameAtual).Visible = False

    'Armazena novo valor de iFrameAtual
    iFrameAtual = Opcao.SelectedItem.Index
    
    Select Case iFrameAtual
    
        Case TAB_Titulos
            Parent.HelpContextID = IDH_TITULOREC_CONS_TITULO
            
        Case TAB_Parcelas
            Parent.HelpContextID = IDH_TITULOREC_CONS_PARCELAS
        
        Case TAB_Baixas
            Parent.HelpContextID = IDH_TITULOREC_CONS_BAIXA
        
        Case TAB_Contabilizacao
            Parent.HelpContextID = IDH_TITULOREC_CONS_CONTABILIZACAO
                    
    End Select
    
    Exit Sub

Erro_Opcao_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158387)

    End Select

    Exit Sub

End Sub

Private Function Traz_TituloReceber_Tela(objTituloReceber As ClassTituloReceber) As Long
'Traz os dados do Título a Receber para a Tela

Dim lErro As Long
Dim iLinha As Integer
Dim iIndice As Integer
Dim colComEmissao As New colComissao
Dim objComissao As ClassComissao, bCancel As Boolean
Dim lPosicaoSeparador As Long
Dim iIndice1 As Integer, Cancel As Boolean
Dim colParcelaReceber As New colParcelaReceber
Dim objVendedor As New ClassVendedor
Dim sTela As String, objMoedas As New ClassMoedas
Dim sModulo As String, sPeriodicidade As String
Dim objNFiscal As New ClassNFiscal
Dim lNumIntDoc As Long, sCclMascarado As String
Dim objCondicaoPagamento As New ClassCondicaoPagto, objNatMovCta As New ClassNatMovCta
Dim iFilialAux As Integer

On Error GoTo Erro_Traz_TituloReceber_Tela

    'Limpa a tela
    Call Limpa_Tela_TituloReceber
    
    iFilialEmpresaTela = objTituloReceber.iFilialEmpresa
    
    iFilialAux = giFilialEmpresa
    giFilialEmpresa = iFilialEmpresaTela
    
    'Coloca o Cliente na Tela
    Cliente.Text = objTituloReceber.lCliente
    Call Cliente_Validate(Cancel)

    'Coloca a Filial na Tela
    Filial.Text = objTituloReceber.iFilial
    Call Filial_Validate(bCancel)
        
    If objTituloReceber.iStatus = STATUS_BAIXADO Then
        Status.Caption = "Baixado"
    Else
        Status.Caption = "Aberto"
    End If
    
    Saldo.Caption = Format(objTituloReceber.dSaldo, "Standard")

    If objTituloReceber.dtReajustadoAte <> DATA_NULA Then
        ReajustadoAte.Caption = Format(objTituloReceber.dtReajustadoAte, "mm/yyyy")
    Else
        ReajustadoAte.Caption = "  /    "
    End If

    '??? LabelDescCpoGenerico1

    If Len(Trim(objTituloReceber.sCcl)) <> 0 Then
    
        lErro = Mascara_MascararCcl(objTituloReceber.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 130150
        
        LabelCcl.Caption = sCclMascarado
        
    Else
    
        LabelCcl.Caption = ""
        
    End If

    If objTituloReceber.iReajustePeriodicidade <> 0 Then
    
        Select Case objTituloReceber.iReajustePeriodicidade
            
            Case PERIODICIDADE_CPR_MENSAL
                sPeriodicidade = "Mensal"
            
            Case PERIODICIDADE_CPR_BIMESTRAL
                sPeriodicidade = "Bimestral"
            
            Case PERIODICIDADE_CPR_TRIMESTRAL
                sPeriodicidade = "Trimestral"
            
            Case PERIODICIDADE_CPR_SEMESTRAL
                sPeriodicidade = "Semestral"
                
            Case PERIODICIDADE_CPR_ANUAL
                sPeriodicidade = "Anual"
                
            Case Else
                sPeriodicidade = CStr(objTituloReceber.iReajustePeriodicidade)
                
        End Select
        
        ReajustePeriodicidade.Caption = sPeriodicidade
    
    Else
    
        ReajustePeriodicidade.Caption = ""
        
    End If
    
    If objTituloReceber.iMoeda <> MOEDA_REAL Then
        
        objMoedas.iCodigo = objTituloReceber.iMoeda
        lErro = CF("Moedas_Le", objMoedas)
        If lErro <> SUCESSO And lErro <> 108821 Then gError 130151
        If lErro <> SUCESSO Then gError 130152
        Moeda.Caption = objMoedas.iCodigo & SEPARADOR & objMoedas.sNome
        
    Else
        Moeda.Caption = ""
    End If
    
    If objTituloReceber.dtReajusteBase <> DATA_NULA Then
        ReajusteBase.Caption = Format(objTituloReceber.dtReajusteBase, "dd/mm/yyyy")
    Else
        ReajusteBase.Caption = "  /  /    "
    End If
    
    If objTituloReceber.iCondicaoPagto <> 0 Then
    
        'Coloca a condicao de Pagamento na Tela
        objCondicaoPagamento.iCodigo = objTituloReceber.iCondicaoPagto
        
        lErro = CF("CondicaoPagto_Le", objCondicaoPagamento)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 60182
            
        'Se não encontrou --> ERRO
        If lErro = 19205 Then gError 60183
        
        CondicaoPagamento.Caption = objCondicaoPagamento.iCodigo & SEPARADOR & objCondicaoPagamento.sDescReduzida
    
    Else
        
        CondicaoPagamento.Caption = ""
    
    End If
    
    If objTituloReceber.sNatureza <> "" Then
        objNatMovCta.sCodigo = objTituloReceber.sNatureza
        
        lErro = CF("NatMovCta_Le", objNatMovCta)
        If lErro <> SUCESSO And lErro <> 122786 Then Error 32299
        If lErro <> SUCESSO Then Error 32300
        LabelNaturezaDesc.Caption = objNatMovCta.sDescricao
    Else
        LabelNaturezaDesc.Caption = ""
    End If
    
    'Coloca os demais dados do Título a Receber na Tela
    NumeroTitulo.Text = objTituloReceber.lNumTitulo
    If objTituloReceber.dtDataEmissao <> DATA_NULA Then
        DataEmissao.Caption = Format(objTituloReceber.dtDataEmissao, "dd/mm/yyyy")
    Else
        DataEmissao.Caption = "  /  /    "
    End If
    
    Tipo.Text = objTituloReceber.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)
    
    If objTituloReceber.dValor > 0 Then ValorTitulo.Caption = Format(objTituloReceber.dValor, "Standard")
    If objTituloReceber.dValorIRRF > 0 Then ValorIRRF.Caption = Format(objTituloReceber.dValorIRRF, "Standard")
    If objTituloReceber.dValorINSS > 0 Then ValorINSS.Caption = Format(objTituloReceber.dValorINSS, "Standard")
    INSSRetido.Value = Format(objTituloReceber.iINSSRetido, "Standard")
    If objTituloReceber.dPISRetido > 0 Then PISRetido.Caption = Format(objTituloReceber.dPISRetido, "Standard")
    If objTituloReceber.dISSRetido > 0 Then ISSRetido.Caption = Format(objTituloReceber.dISSRetido, "Standard")
    If objTituloReceber.dCOFINSRetido > 0 Then COFINSRetido.Caption = Format(objTituloReceber.dCOFINSRetido, "Standard")
    If objTituloReceber.dCSLLRetido > 0 Then CSLLRetido.Caption = Format(objTituloReceber.dCSLLRetido, "Standard")

    'Frame Para Atrasos de Pagamento
    If objTituloReceber.dPercMulta > 0 Then PercMulta.Caption = Format(objTituloReceber.dPercMulta, "##0.00##%")
    If objTituloReceber.dPercJurosDiario > 0 Then JurosMensais.Caption = Round((objTituloReceber.dPercJurosDiario * 3000), 2) & "%"

    'Lê as Parcelas a Receber vinculadas ao Título
    lErro = CF("ParcelasReceber_Le_Todas", objTituloReceber, colParcelaReceber)
    If lErro <> SUCESSO And lErro <> 58990 Then gError 58922

    If colParcelaReceber.Count > NUM_MAXIMO_PARCELAS Then gError 58923

    'Lê as comissões do Título
    lErro = CF("Comissoes_Le", objTituloReceber.lNumIntDoc, colComEmissao, TIPO_TITULO_RECEBER)
    If lErro <> SUCESSO Then gError 58924

    'Loop de preenchimento do GridComissoesEmissao
    For Each objComissao In colComEmissao

        iIndice = iIndice + 1

        'ler do bd o nome reduzido do vendedor
        objVendedor.iCodigo = objComissao.iCodVendedor

        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 58925

        If lErro <> SUCESSO Then gError 58926

        GridComissoesEmissao.TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido

        GridComissoesEmissao.TextMatrix(iIndice, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
        GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
        GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")

    Next

    'Linhas existentes
    iIndice = 0
    objGridComissoesEmissao.iLinhasExistentes = iIndice

    'Chama Soma_Percentual
    lErro = Soma_Percentual(objGridComissoesEmissao)
    If lErro <> SUCESSO Then gError 58927

    'Chama Soma_Valor
    lErro = Soma_Valor(objGridComissoesEmissao)
    If lErro <> SUCESSO Then gError 58928

    'Carrega o Grid de Parcelas
    lErro = Carrega_GridParcelas(colParcelaReceber)
    If lErro <> SUCESSO Then gError 58929

    Call Traz_Parcela_Tela(colParcelaReceber.Item(1))

    DataVencParcBaixa.Caption = Format(colParcelaReceber.Item(1).dtDataVencimento, "dd/mm/yyyy")
    ValorParcBaixa.Caption = Format(colParcelaReceber.Item(1).dValor, "Standard")

    lErro = Carrega_ComboSequencial(colParcelaReceber.Item(1).lNumIntDoc)
    If lErro <> SUCESSO Then gError 58930

    'Se for uma Parcela Baixada Traz os dados da Baixa
    If Sequencial.ListCount > 0 Then

        'Carrega os Dados da Baixa (1 como Default)
        Call Carrega_Dados_Baixa(1, Sequencial.List(0), colParcelaReceber.Item(1).lNumIntTitulo)

    Else
        Call Limpa_Tab_Baixa
    End If

    lNumIntDoc = 0

    'Le a Tela e o Modulo que originou a contabilidade
    Select Case objTituloReceber.sSiglaDocumento

        Case OUTROS_RECEBIMENTOS

            sTela = "TituloReceber"
            sModulo = MODULO_CONTASARECEBER
            lNumIntDoc = objTituloReceber.lNumIntDoc

        Case Else

            'Verifica se de CPR ou FAT
            lErro = CF("Titulo_Le_DocumentoOriginal", objTituloReceber.lNumIntDoc, CPR_TITULO_RECEBER, objNFiscal, sTela)
            If lErro <> SUCESSO And lErro <> 58942 Then gError 59494

            'Se for de CPR
            If lErro = 58942 Then
                lNumIntDoc = objTituloReceber.lNumIntDoc
                sTela = "TituloReceber"
                sModulo = MODULO_CONTASARECEBER
            End If

    End Select

    'se o titulo foi gerado pelo CR ==> traz a contabilidade associada
    If lNumIntDoc <> 0 Then

        'Traz os dados contábeis para a tela (contabilidade)
        lErro = objContabil.Contabil_Traz_Doc_Tela_Consulta(lNumIntDoc, sModulo, sTela)
        If lErro <> SUCESSO And lErro <> 58885 Then gError 58931

    End If

    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objTituloReceber.objProjetoInfo, objTituloReceber.lNumIntDoc, PRJ_CR_TIPO_TITREC)
    If lErro <> SUCESSO Then gError 181676
    '###################################################

    giFilialEmpresa = iFilialAux
    
    Traz_TituloReceber_Tela = SUCESSO

    Exit Function

Erro_Traz_TituloReceber_Tela:

    giFilialEmpresa = iFilialAux

    Traz_TituloReceber_Tela = gErr

    Select Case gErr

        Case 58922, 58924, 58925, 58927, 58928, 58929, 58930, 58931, 59494, 60182, 130150, 130151, 32299, 32300, 181676 'Tratado na rotina chamada

        Case 130152
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDA_NAO_CADASTRADA", gErr, objMoedas.iCodigo)
        
        Case 58923
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", gErr, colParcelaReceber.Count, NUM_MAXIMO_PARCELAS)

        Case 58926
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)
        
        Case 60183
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagamento.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158388)

    End Select

    Exit Function

End Function

Private Function Soma_Percentual(objGridInt As AdmGrid) As Long
'Faz a soma da coluna de Percentual de ObjGridInt

Dim iIndice As Integer
Dim dSomaPercentual As Double
Dim lComprimento As Long

    dSomaPercentual = 0

    'Loop no Grid
    For iIndice = 1 To objGridInt.iLinhasExistentes

        lComprimento = Len(Trim(objGridInt.objGrid.TextMatrix(iIndice, iGrid_PercentualComissao_Col)))

        'Verifica se Percentual da Comissão está preenchido
        If lComprimento > 0 Then

            'Acumula Percentual em dSomaPercentual
            dSomaPercentual = dSomaPercentual + CDbl(left(objGridInt.objGrid.TextMatrix(iIndice, iGrid_PercentualComissao_Col), lComprimento - 1))

        End If

    Next

    'Mostra na tela o Total Percentual
    If objGridInt.objGrid.Name = GridComissoes.Name Then
        TotalPercentualComissao.Caption = Format(dSomaPercentual / 100, "Percent")
    ElseIf objGridInt.objGrid.Name = GridComissoesEmissao.Name Then
        TotalPercentualEmissao.Caption = Format(dSomaPercentual / 100, "Percent")
    End If

    Soma_Percentual = SUCESSO

    Exit Function

End Function

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Tipo_Validate

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then Exit Sub

    'Verifica se o Tipo foi selecionado
    If Tipo.Text = Tipo.List(Tipo.ListIndex) Then Exit Sub

    'Tenta Localizar o Tipo no Text da Combo
    lErro = CF("SCombo_Seleciona", Tipo)
    If lErro <> SUCESSO And lErro <> 60483 Then Error 26138

    'Se nao encontrar -> Erro
    If lErro = 60483 Then Error 26139

    Exit Sub

Erro_Tipo_Validate:

    Cancel = True


    Select Case Err

        Case 26138

        Case 26139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, Tipo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158389)

    End Select

    Exit Sub

End Sub

Public Sub UpDownParcela_DownClick(Index As Integer)

Dim iIndice As Integer
Dim lErro As Long
Dim iParcela As Integer
Dim objTituloReceber As New ClassTituloReceber
Dim objParcelaReceber As New ClassParcelaReceber

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
        If lErro <> SUCESSO Then Error 58941

        objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
        objTituloReceber.lNumTitulo = CLng(NumeroTitulo.Text)
        objTituloReceber.iFilialEmpresa = iFilialEmpresaTela
        objParcelaReceber.iNumParcela = iParcela - 1

        lErro = CF("ParcelasRec_Le_Todas_NumTitulo", objTituloReceber, objParcelaReceber)
        If lErro <> SUCESSO And lErro <> 60003 Then Error 58944

        If lErro = 60003 Then Error 58945

        'Coloca na Tela os dados da nova parcela a tratar
        lErro = Traz_Parcela_Tela(objParcelaReceber)
        If lErro <> SUCESSO Then Error 58946
        
        DataVencParcBaixa.Caption = Format(objParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
        ValorParcBaixa.Caption = Format(objParcelaReceber.dValor, "Standard")

        lErro = Carrega_ComboSequencial(objParcelaReceber.lNumIntDoc)
        If lErro <> SUCESSO Then Error 58947

        'Se for uma Parcela Baixada Traz os dados da Baixa
        If Sequencial.ListCount > 0 Then

            'Carrega os Dados da Baixa (1 como Default)
            Call Carrega_Dados_Baixa(objParcelaReceber.iNumParcela, Sequencial.List(0), objParcelaReceber.lNumIntTitulo)

        Else
            Call Limpa_Tab_Baixa
        End If

    End If

    Exit Sub

Erro_UpDownParcela_DownClick:

    Select Case Err

        Case 58941, 58944, 58946, 58947 'Tratado na Rotina chamada

        Case 58945
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158390)

    End Select

    Exit Sub

End Sub

Function Carrega_ComboSequencial(lNumIntParcela As Long) As Long

Dim lErro As Long
Dim colSequenciais As New Collection
Dim iIndice  As Integer

On Error GoTo Erro_Carrega_ComboSequencial

    lErro = CF("BaixaParcRec_Le_Sequenciais", lNumIntParcela, colSequenciais)
    If lErro <> SUCESSO Then Error 58948
    
    Sequencial.Clear
    
    If colSequenciais.Count = 0 Then Exit Function

    For iIndice = 1 To colSequenciais.Count

        Sequencial.AddItem (CStr(colSequenciais(iIndice)))

    Next

    Sequencial.ListIndex = 0

    Carrega_ComboSequencial = SUCESSO

    Exit Function

Erro_Carrega_ComboSequencial:

    Carrega_ComboSequencial = Err

    Select Case Err

        Case 58948

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158391)

    End Select

    Exit Function

End Function

Public Sub UpDownParcela_UpClick(Index As Integer)

Dim iIndice As Integer
Dim lErro As Long
Dim iParcela As Integer
Dim objTituloReceber As New ClassTituloReceber
Dim objParcelaReceber As New ClassParcelaReceber

On Error GoTo Erro_UpDownParcela_UpClick

    'Verifica se já existe alguma Parcela na tela
    If Len(Trim(Parcela(1).Caption)) > 0 Then
        iParcela = CInt(Parcela(1).Caption)
    Else
        iParcela = 0
    End If

    'Verifica se existe a Parcela (iParcela+1)
    If iParcela + 1 <= objGridParcelas.iLinhasExistentes Then

        'Critica se os Campos estão Preenchidos(Cliente, Filial, Tipo, Número)
        lErro = Critica_CamposPreenchidos()
        If lErro <> SUCESSO Then Error 58949

        objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
        objTituloReceber.lNumTitulo = CLng(NumeroTitulo.Text)
        objTituloReceber.iFilialEmpresa = iFilialEmpresaTela
        objParcelaReceber.iNumParcela = iParcela + 1

        lErro = CF("ParcelasRec_Le_Todas_NumTitulo", objTituloReceber, objParcelaReceber)
        If lErro <> SUCESSO And lErro <> 60003 Then Error 58953

        If lErro = 60003 Then Error 58954

        'Traz para telas os dados da Parcela seguinte
        lErro = Traz_Parcela_Tela(objParcelaReceber)
        If lErro <> SUCESSO Then Error 58955

        DataVencParcBaixa.Caption = Format(objParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
        ValorParcBaixa.Caption = Format(objParcelaReceber.dValor, "Standard")

        lErro = Carrega_ComboSequencial(objParcelaReceber.lNumIntDoc)
        If lErro <> SUCESSO Then Error 58956

        'Se for uma Parcela Baixada Traz os dados da Baixa
        If Sequencial.ListCount > 0 Then

            'Carrega os Dados da Baixa (1 como Default)
            Call Carrega_Dados_Baixa(objParcelaReceber.iNumParcela, Sequencial.List(0), objParcelaReceber.lNumIntTitulo)

        Else
            Call Limpa_Tab_Baixa
        End If

    End If

    Exit Sub

Erro_UpDownParcela_UpClick:

    Select Case Err

        Case 58949, 58953, 58955, 58956 'Tratado na Rotina chamada

        Case 58954
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PARCELA_REC_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158392)

    End Select

    Exit Sub

End Sub

Private Function Soma_Valor(objGridInt As AdmGrid) As Long

Dim iIndice As Integer
Dim dSomaValor As Double

    dSomaValor = 0

    'Loop no GridComissao
    For iIndice = 1 To objGridInt.iLinhasExistentes

        'Verifica se Valor da Comissão está preenchido
        If Len(Trim(objGridInt.objGrid.TextMatrix(iIndice, iGrid_ValorComissao_Col))) > 0 Then

            'Acumula Valor em dSomaValor
            dSomaValor = dSomaValor + CDbl(objGridInt.objGrid.TextMatrix(iIndice, iGrid_ValorComissao_Col))

        End If

    Next

    'Mostra na tela o Total Valor
    If objGridInt.objGrid.Name = GridComissoes.Name Then
        TotalValorComissao.Caption = Format(dSomaValor, "Standard")
    ElseIf objGridInt.objGrid.Name = GridComissoesEmissao.Name Then
        TotalValorEmissao.Caption = Format(dSomaValor, "Standard")
    End If

    Soma_Valor = SUCESSO

    Exit Function

End Function

Private Sub Limpa_Tela_TituloReceber()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa os Campos da Tela
    Call Limpa_Tela(Me)

    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = objUserControl
    Set gobjTelaProjetoInfo.objTela = Me
    
    Set gobjParcRecBaixa = New ClassBaixaParcRec
    Set gobjBaixaReceber = New ClassBaixaReceber

    'Limpa os campos não limpos em Limpa_Tela
    PercMulta.Caption = ""
    JurosMensais.Caption = ""
    Filial.Clear
    Tipo.Text = ""
    INSSRetido.Value = False
    TotalPercentualComissao.Caption = ""
    TotalPercentualEmissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorEmissao.Caption = ""
    ValorTitulo.Caption = ""
    DataVenctParcReal.Caption = ""
    DataVenctoParc.Caption = ""
    ValorParc.Caption = ""
    ValorOriginalParc.Caption = ""
    ValorINSS.Caption = ""
    ValorIRRF.Caption = ""
    PISRetido.Caption = ""
    ISSRetido.Caption = ""
    COFINSRetido.Caption = ""
    CSLLRetido.Caption = ""
    Parcela(0).Caption = ""
    Parcela(1).Caption = ""
    Saldo.Caption = ""
    SaldoParc.Caption = ""
    Status.Caption = ""

    LabelNaturezaDesc.Caption = ""
    DataEmissao.Caption = ""

    'Limpa os Grids da Tela
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridDesconto)
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridComissoesEmissao)

    ReajustadoAte.Caption = ""
''    LabelDescCpoGenerico1.Caption = ""
    LabelCcl.Caption = ""
    ReajustePeriodicidade.Caption = ""
    Moeda.Caption = ""
    ReajusteBase.Caption = ""

    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade_Consulta

End Sub

Sub Limpa_Tab_Parcelas()
'Limpa o TabParcelas

    'Limpa a parcela
    Parcela(0).Caption = ""
    Parcela(1).Caption = ""
    DataVenctoParc.Caption = ""
    DataVenctParcReal.Caption = ""
    ValorParc.Caption = ""
    ValorOriginalParc.Caption = ""
    SaldoParc.Caption = ""
    StatusParc.Caption = ""
    NossoNumero.Caption = ""
    
    Set gobjParcRecBaixa = New ClassBaixaParcRec
    Set gobjBaixaReceber = New ClassBaixaReceber

    'Limpa os Grids de Desconto e Comissões
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridDesconto)

    'Limpa os Totais
    TotalPercentualComissao.Caption = ""
    TotalValorComissao.Caption = ""

End Sub

Sub Limpa_Tab_Baixa()

    Sequencial.Clear
    DataBaixa.Caption = ""
    ValorPago.Caption = ""
    ValorBaixado = ""
    Desconto.Caption = ""
    Multa.Caption = ""
    Juros.Caption = ""
    ContaCorrente.Caption = ""
    Historico.Caption = ""
    ValorRecebimento.Caption = ""
    DataMovimento.Caption = ""
    MeioPagtoDescricao.Caption = ""
    CCIntNomeReduzido.Caption = ""
    ValorPA.Caption = ""
    DataEmissaoCred.Caption = ""
    NumTitulo.Caption = ""
    SaldoDebito.Caption = ""
    FilialEmpresaCR.Caption = ""
    ValorDebito.Caption = ""
    SiglaDocumentoCR.Caption = ""
    HistoricoPerda.Caption = ""

    Set gobjParcRecBaixa = New ClassBaixaParcRec
    Set gobjBaixaReceber = New ClassBaixaReceber

End Sub

Private Function Traz_Parcela_Tela(objParcelaReceber As ClassParcelaReceber) As Long
'Traz para a tela os dados da Parcela passada

Dim objDesconto As ClassDesconto, colDescontos As New colDesconto
Dim colComissao As New colComissao, objComissao As ClassComissao
Dim objVendedor As New ClassVendedor
Dim iLinha As Integer, iIndice As Integer, lSeparador As Long, lErro As Long
Dim objCobrador As New ClassCobrador
Dim objCarteira As New ClassCarteiraCobranca

On Error GoTo Erro_Traz_Parcela_Tela

    Call Limpa_Tab_Parcelas
    Call Limpa_Tab_Baixa

    'Põe a Parcela na tela
    Parcela(0).Caption = objParcelaReceber.iNumParcela
    Parcela(1).Caption = objParcelaReceber.iNumParcela
    DataVenctoParc.Caption = Format(objParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
    DataVenctParcReal.Caption = Format(objParcelaReceber.dtDataVencimentoReal, "dd/mm/yyyy")
    ValorParc.Caption = Format(objParcelaReceber.dValor, "Standard")
    ValorOriginalParc.Caption = Format(objParcelaReceber.dValorOriginal, "Standard")
    
    '#############################
    'Inserido por Wagner
    NossoNumero.Caption = objParcelaReceber.sNumTitCobrador
    '#############################
    
    'Preenche o Cobrador
    objCobrador.iCodigo = objParcelaReceber.iCobrador
    
    lErro = CF("Cobrador_Le", objCobrador)
    If lErro <> SUCESSO And lErro <> 19294 Then Error 60351
    
    If lErro = 19294 Then Error 60352
    
    Cobrador.Caption = objCobrador.sNomeReduzido
    
    'Preenche a Carteira
    objCarteira.iCodigo = objParcelaReceber.iCarteiraCobranca
    
    lErro = CF("CarteiraDeCobranca_Le", objCarteira)
    If lErro <> SUCESSO And lErro <> 23413 Then Error 60353
    
    If lErro = 23413 Then Error 60354
    
    Carteira.Caption = objCarteira.sDescricao
    
    If objParcelaReceber.iStatus = STATUS_BAIXADO Then
        StatusParc.Caption = "Baixada"
    Else
        StatusParc.Caption = "Aberta"
    End If
    
    If objParcelaReceber.dSaldo > 0 Then
        SaldoParc.Caption = Format(objParcelaReceber.dSaldo, "Standard")
    Else
        SaldoParc.Caption = ""
    End If
        
    'Adiciona em colDesc os descontos da Parcela
    If objParcelaReceber.iDesconto1Codigo <> 0 Then colDescontos.Add objParcelaReceber.iDesconto1Codigo, objParcelaReceber.dtDesconto1Ate, objParcelaReceber.dDesconto1Valor
    If objParcelaReceber.iDesconto2Codigo <> 0 Then colDescontos.Add objParcelaReceber.iDesconto2Codigo, objParcelaReceber.dtDesconto2Ate, objParcelaReceber.dDesconto2Valor
    If objParcelaReceber.iDesconto3Codigo <> 0 Then colDescontos.Add objParcelaReceber.iDesconto3Codigo, objParcelaReceber.dtDesconto3Ate, objParcelaReceber.dDesconto3Valor

    iLinha = 0

    'Preenche o GridDesconto com os Descontos existentes em colDesconto
    For Each objDesconto In colDescontos

            iLinha = iLinha + 1

            If objDesconto.dtData <> DATA_NULA Then GridDescontos.TextMatrix(iLinha, iGrid_Data_Col) = Format(objDesconto.dtData, "dd/mm/yyyy")

            For iIndice = 0 To TipoDesconto.ListCount - 1
                If TipoDesconto.ItemData(iIndice) = objDesconto.iCodigo Then
                    GridDescontos.TextMatrix(iLinha, iGrid_TipoDesconto_Col) = TipoDesconto.List(iIndice)
                    Exit For
                End If
            Next

            If objDesconto.iCodigo = VALOR_FIXO Or objDesconto.iCodigo = VALOR_ANT_DIA Or objDesconto.iCodigo = VALOR_ANT_DIA_UTIL Then
                GridDescontos.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = Format(objDesconto.dValor, "Standard")
            ElseIf objDesconto.iCodigo = Percentual Or objDesconto.iCodigo = PERC_ANT_DIA Or objDesconto.iCodigo = PERC_ANT_DIA_UTIL Then
                GridDescontos.TextMatrix(iLinha, iGrid_Percentual1_Col) = Format(objDesconto.dValor, "Percent")
                GridDescontos.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = Format(objDesconto.dValor * objParcelaReceber.dValor, "Standard")
            End If

    Next

    'Atribui o número de linhas existentes
    objGridDesconto.iLinhasExistentes = iLinha

    'Lê as Comissoes relacionadas a Parcela
    lErro = CF("Comissoes_Le", objParcelaReceber.lNumIntDoc, colComissao, TIPO_PARCELA)
    If lErro <> SUCESSO Then Error 58962

    iLinha = 0

    'Preenche o GridComissoes com as Comissões existentes em colComissao
    For Each objComissao In colComissao

        iLinha = iLinha + 1
        GridComissoes.TextMatrix(iLinha, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
        GridComissoes.TextMatrix(iLinha, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
        GridComissoes.TextMatrix(iLinha, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")
        objVendedor.iCodigo = objComissao.iCodVendedor

        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then Error 58963

        If lErro <> SUCESSO Then Error 58964

        GridComissoes.TextMatrix(iLinha, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido

    Next

    'Atualiza o número de linhas existentes
    objGridComissoes.iLinhasExistentes = iLinha

    'Faz a Soma das colunas Percentual e Valor do Grid de Comissões
    Call Soma_Percentual(objGridComissoes)
    Call Soma_Valor(objGridComissoes)

    Traz_Parcela_Tela = SUCESSO

    Exit Function

Erro_Traz_Parcela_Tela:

    Traz_Parcela_Tela = Err

    Select Case Err

        Case 58962, 58963, 60351, 60353
        
        Case 58964
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)
        
        Case 60352
            Call Rotina_Erro(vbOKOnly, "ERRO_COBRADOR_INEXISTENTE", Err, objVendedor.iCodigo)
        
        Case 60354
            Call Rotina_Erro(vbOKOnly, "ERRO_CARTEIRACOBRANCA_NAO_CADASTRADA", Err)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158393)

    End Select

    Exit Function

End Function

Private Function Carrega_TipoDesconto() As Long
'Carrega os Tipos de Desconto

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    'Lê o código e a descrição de todos os Tipos de Desconto
    lErro = CF("Cod_Nomes_Le", "TiposDeDesconto", "Codigo", "DescReduzida", STRING_TIPOSDEDESCONTO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 58965

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o ítem na List da Combo TipoDesconto
        TipoDesconto.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TipoDesconto.ItemData(TipoDesconto.NewIndex) = objCodDescricao.iCodigo

    Next

    Carrega_TipoDesconto = SUCESSO

    Exit Function

Erro_Carrega_TipoDesconto:

    Carrega_TipoDesconto = Err

    Select Case Err

        Case 58965

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158394)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoTipoDocumento = Nothing

    Set gobjParcRecBaixa = Nothing
    Set gobjBaixaReceber = Nothing

    Set objGridParcelas = Nothing
    Set objGridDesconto = Nothing
    Set objGridComissoes = Nothing
    Set objGridComissoesEmissao = Nothing
    
    Set gobjTelaProjetoInfo = Nothing

    Set objGrid1 = Nothing
    Set objContabil = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub


Private Function Carrega_GridParcelas(colParcelaReceber As colParcelaReceber) As Long
'Carrega os dados das Parcelas no GridParcelas

Dim lErro As Long
Dim iLinha As Integer
Dim objParcelaReceber As ClassParcelaReceber

On Error GoTo Erro_Carrega_GridParcelas

    'Limpa o Grid de Parcelas
    Call Grid_Limpa(objGridParcelas)

    iLinha = 0

    'Preenche as linhas do Grid Parcelas com os dados de cada Parcela
    For Each objParcelaReceber In colParcelaReceber

        'Carrega o GridParcelas
        iLinha = iLinha + 1
        If objParcelaReceber.dtDataVencimento <> DATA_NULA Then GridParcelas.TextMatrix(iLinha, iGrid_Vencimento_Col) = Format(objParcelaReceber.dtDataVencimento, "dd/mm/yyyy")
        If objParcelaReceber.dtDataVencimentoReal <> DATA_NULA Then GridParcelas.TextMatrix(iLinha, iGrid_VenctoReal_Col) = Format(objParcelaReceber.dtDataVencimentoReal, "dd/mm/yyyy")
        GridParcelas.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objParcelaReceber.dValor, "Standard")
        
        If objParcelaReceber.iStatus = STATUS_BAIXADO Then
            GridParcelas.TextMatrix(iLinha, iGrid_Status_Col) = "Baixada"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_Status_Col) = "Aberta"
        End If
        GridParcelas.TextMatrix(iLinha, iGrid_DescPrev_Col) = objParcelaReceber.sObservacao
        
    Next

    objGridParcelas.iLinhasExistentes = iLinha

    Carrega_GridParcelas = SUCESSO

    Exit Function

Erro_Carrega_GridParcelas:

    Carrega_GridParcelas = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158395)

    End Select

    Exit Function

End Function

Private Sub Carrega_Dados_Baixa(iNumParcela As Integer, iSequencialBaixa As Integer, lNumIntTituloReceber As Long)
'Carrega os dados da Baixa

Dim lErro As Long
Dim objParcRecBaixa As New ClassBaixaParcRec
Dim objBaixaReceber As New ClassBaixaReceber

On Error GoTo Erro_Carrega_Dados_Baixa

    'Lê a Parcela no BD
    lErro = CF("BaixaRecCancelar_Le_Parcela", lNumIntTituloReceber, iNumParcela, iSequencialBaixa, objParcRecBaixa)
    If lErro <> SUCESSO And lErro <> 46230 Then Error 58965

    'Se não encontrou a Parcela
    If lErro = SUCESSO Then

        'Se encontrou e Baixa já está cancelada --> Erro
        If objParcRecBaixa.iStatus = STATUS_EXCLUIDO Then Error 58966

        objBaixaReceber.lNumIntBaixa = objParcRecBaixa.lNumIntBaixa

        'Lê  a Baixa
        lErro = CF("BaixaRec_Le", objBaixaReceber)
        If lErro <> SUCESSO And lErro <> 46234 Then Error 58967

        If lErro <> SUCESSO Then Error 58968 'Não encontrou

        'Coloca na tela os dados da Baixa lida
        lErro = Traz_Dados_Baixa(objParcRecBaixa, objBaixaReceber)
        If lErro <> SUCESSO Then Error 58969

    Else
        'Limpa o Tab de Baixa
        Call Limpa_Tab_Baixa

    End If
    
    Set gobjParcRecBaixa = objParcRecBaixa
    Set gobjBaixaReceber = objBaixaReceber

    Exit Sub

Erro_Carrega_Dados_Baixa:

    Select Case Err

        Case 58966
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXA_PARCELA_CANCELADA", Err, iNumParcela, iSequencialBaixa)

        Case 58965, 58967, 58969

        Case 58968
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BAIXAREC_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158396)

    End Select

    Exit Sub

End Sub

Public Sub Recebimento_Click(Index As Integer)

    FrameRecebimento(iFrameRecebimentoAtual).Visible = False
    FrameRecebimento(Index).Visible = True
    iFrameRecebimentoAtual = Index

End Sub

Function Traz_Dados_Baixa(objParcRecBaixa As ClassBaixaParcRec, objBaixaReceber As ClassBaixaReceber) As Long
'Mostra na tela os dados da baixa

Dim lErro As Long

On Error GoTo Erro_Traz_Dados_Baixa

    'Coloca os dados da Baixa na tela
    Desconto.Caption = Format(objParcRecBaixa.dValorDesconto, "Standard")
    ValorPago.Caption = Format(objParcRecBaixa.dValorRecebido, "Standard") 'Alterado por Wagner 18/05/2005
    Multa.Caption = Format(objParcRecBaixa.dValorMulta, "Standard")
    ValorBaixado.Caption = Format(objParcRecBaixa.dValorBaixado, "Standard")
    Juros.Caption = Format(objParcRecBaixa.dValorJuros, "Standard")

    'preenche a data em branco caso seja data nula
    If objBaixaReceber.dtData <> DATA_NULA Then
        DataBaixa.Caption = Format(objBaixaReceber.dtData, "dd/mm/yyyy")
    Else
        DataBaixa.Caption = ""
    End If

    If objBaixaReceber.iMotivo = MOTIVO_RECEBIMENTO Then
        Recebimento(0).Value = True

        'Traz os dados do Recebimento
        lErro = Traz_Dados_Recebimento(objBaixaReceber)
        If lErro <> SUCESSO Then gError 58970

    ElseIf objBaixaReceber.iMotivo = MOTIVO_RECEBTO_ANTECIPADO Then
        Recebimento(1).Value = True
        'Traz os dados do Recebimento antecipado
        lErro = Traz_Dados_Recebimento_Antecipado(objBaixaReceber)
        If lErro <> SUCESSO Then gError 58971

    ElseIf objBaixaReceber.iMotivo = MOTIVO_DEBITO_CLIENTE Then
        Recebimento(2).Value = True

        'Traz os dados do Débito
        lErro = Traz_Dados_Debito_Cliente(objBaixaReceber)
        If lErro <> SUCESSO Then gError 58972

    ElseIf objBaixaReceber.iMotivo = MOTIVO_PERDA Then
        Recebimento(3).Value = True
        HistoricoPerda.Caption = objBaixaReceber.sHistorico
    
    ElseIf objBaixaReceber.iMotivo = MOTIVO_CARTAO_DEBITO_CREDITO Then
        
        Recebimento(4).Value = True
        
        lErro = Traz_Dados_Titulo_Cartao(objBaixaReceber)
        If lErro <> SUCESSO Then gError 183056
    
    End If

    Traz_Dados_Baixa = SUCESSO

    Exit Function

Erro_Traz_Dados_Baixa:

    Traz_Dados_Baixa = gErr

    Select Case gErr

        Case 58970, 58971, 58972, 183056

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158397)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Recebimento(objBaixaReceber As ClassBaixaReceber) As Long
'Mostra os dados do recebimento utilizado para a baixa

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objContaCorrente As New ClassContasCorrentesInternas

On Error GoTo Erro_Traz_Dados_Recebimento

    objMovCCI.lNumMovto = objBaixaReceber.lNumMovCta

    'Lê o Movimento
    lErro = CF("MovContaCorrente_Le", objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then Error 58973

    'Se não encontrou Movimento --> erro
    If lErro = 11893 Then Error 58974

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objContaCorrente)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 58975

    'Se não encontrou a conta Corrente --> erro
    If lErro <> SUCESSO Then Error 58976

    'Coloca os dados na tela
    ContaCorrente.Caption = objContaCorrente.sNomeReduzido
    ValorRecebimento.Caption = Format(objMovCCI.dValor, "Standard")
    Historico.Caption = objMovCCI.sHistorico

    Traz_Dados_Recebimento = SUCESSO

    Exit Function

Erro_Traz_Dados_Recebimento:

    Traz_Dados_Recebimento = Err

    Select Case Err

        Case 58973, 58975

        Case 58974
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovCCI.lNumMovto)

        Case 58976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objContaCorrente.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158398)

    End Select

    Exit Function

End Function

Function Traz_Dados_Recebimento_Antecipado(objBaixaReceber As ClassBaixaReceber) As Long
'exibe dados do adiantamanto utilizado na baixa

Dim lErro As Long
Dim objMovCCI As New ClassMovContaCorrente
Dim objCCI As New ClassContasCorrentesInternas
Dim objAntecipRec As New ClassAntecipReceb

On Error GoTo Erro_Traz_Dados_Recebimento_Antecipado

    objMovCCI.lNumMovto = objBaixaReceber.lNumMovCta

    'Lê o movimento da Baixa
    lErro = CF("MovContaCorrente_Le", objMovCCI)
    If lErro <> SUCESSO And lErro <> 11893 Then Error 58977

    If lErro = 11893 Then Error 58978 'Não encontrou

    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objMovCCI.iCodConta, objCCI)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 58979

    If lErro <> SUCESSO Then Error 58980 'Não encontrou

    objAntecipRec.lNumMovto = objMovCCI.lNumMovto

    lErro = CF("AntecipRec_Le_NumMovto", objAntecipRec)
    If lErro <> AD_SQL_SUCESSO And lErro <> 46252 Then Error 58981

    If lErro = 46252 Then Error 58982

    'Coloca os dados na tela

    'preenche a data em branco se for data_nula
    If objMovCCI.dtDataMovimento <> DATA_NULA Then
        DataMovimento.Caption = Format(objMovCCI.dtDataMovimento, "dd/mm/yyyy")
    Else
        DataMovimento.Caption = ""
    End If

    ValorPA.Caption = Format(objMovCCI.dValor, "Standard")
    CCIntNomeReduzido.Caption = objCCI.sNomeReduzido
    If objMovCCI.iTipoMeioPagto = DINHEIRO Then
        MeioPagtoDescricao.Caption = "Dinheiro"
    ElseIf objMovCCI.iTipoMeioPagto = Cheque Then
        MeioPagtoDescricao.Caption = "Cheque"
    ElseIf objMovCCI.iTipoMeioPagto = BORDERO Then
        MeioPagtoDescricao.Caption = "Borderô"
    End If

    Traz_Dados_Recebimento_Antecipado = SUCESSO

    Exit Function

Erro_Traz_Dados_Recebimento_Antecipado:

    Traz_Dados_Recebimento_Antecipado = Err

    Select Case Err

        Case 58977, 58979, 58981

        Case 58978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTO_NAO_CADASTRADO3", Err, objMovCCI.lNumMovto)

        Case 58980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTACORRENTE_INEXISTENTE", Err, objCCI.iCodigo)

        Case 58982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RECEBIMENTO_ANTECIPADO_INEXISTENTE", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158399)

    End Select

    Exit Function

End Function

Private Function Traz_Dados_Debito_Cliente(objBaixaReceber As ClassBaixaReceber) As Long
'exibe dados do debito utilizado para a baixa

Dim lErro As Long
Dim objDebitoRec As New ClassDebitoRecCli

On Error GoTo Erro_Traz_Dados_Debito_Cliente

    objDebitoRec.lNumIntDoc = objBaixaReceber.lNumIntDoc

    'Lê o debito com o cliente
    lErro = CF("DebitoReceber_Le", objDebitoRec)
    If lErro <> AD_SQL_SUCESSO And 17071 Then Error 58983

    If lErro <> SUCESSO Then Error 58984

    'Coloca os dados na tela
    If objDebitoRec.dtDataEmissao <> DATA_NULA Then
        DataEmissaoCred.Caption = Format(objDebitoRec.dtDataEmissao, "dd/mm/yyyy")
    Else
        DataEmissaoCred.Caption = ""
    End If

    NumTitulo.Caption = objDebitoRec.lNumTitulo
    SaldoDebito.Caption = Format(objDebitoRec.dSaldo, "Standard")
    SiglaDocumentoCR.Caption = objDebitoRec.sSiglaDocumento
    ValorDebito.Caption = Format(objDebitoRec.dValorTotal, "Standard")
    FilialEmpresaCR.Caption = objDebitoRec.iFilialEmpresa

    Traz_Dados_Debito_Cliente = SUCESSO

    Exit Function

Erro_Traz_Dados_Debito_Cliente:

    Traz_Dados_Debito_Cliente = Err

    Select Case Err

        Case 58983

        Case 58984
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DEBITORECCLI_NAO_ENCONTRADO", Err, objDebitoRec.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158400)

    End Select

    Exit Function

End Function

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
    If lErro <> SUCESSO And lErro <> 26061 Then gError 183058

    If lErro <> SUCESSO Then
    
        lErro = CF("TituloReceberBaixado_Le", objTitRec)
        If lErro <> SUCESSO And lErro <> 56570 Then gError 183059
    
    End If

    If lErro = SUCESSO Then
    
        objCliente.lCodigo = objTitRec.lCliente
    
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 183060
        
        If lErro = SUCESSO Then
            
            ClienteTitCartao.Caption = objCliente.sNomeReduzido
            
            sCliente = objCliente.sNomeReduzido
            objFilialCliente.iCodFilial = objTitRec.iFilial
    
            'Pesquisa se existe Filial com o código extraído
            lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
            If lErro <> SUCESSO And lErro <> 17660 Then gError 183061
    
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

        Case 183058 To 183061

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158398)

    End Select

    Exit Function

End Function


'Contabilidade

Public Sub CTBBotaoImprimir_Click()

    Call objContabil.Contabil_BotaoImprimir_Consulta_Click

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TITULOREC_CONS_TITULO
    Set Form_Load_Ocx = Me
    Caption = "Consulta de Títulos a Receber"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TituloReceber_Consulta"

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
            Call ClienteEtiqueta_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call TipoDocumentoLabel_Click
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

Public Sub Cliente_Change()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Change
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134021

    Exit Sub

Erro_Cliente_Change:

    Select Case gErr

        Case 134021

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158401)

    End Select
    
    Exit Sub

End Sub

Public Sub BotaoDif_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim colSelecao As New Collection
Dim iPreenchido As Integer
Dim sSelecao As String
Dim objParcelaRecDif As New ClassParcelasRecDif

On Error GoTo Erro_BotaoDif_Click

    'Move os dados da Tela para objtituloreceber e colParcelas
    lErro = Move_Tela_Memoria(objTituloReceber)
    If lErro <> SUCESSO Then gError 177922
    
    'Verifica se os obj(s) estão preenchidos antes de serem incluídos na coleção
    If objTituloReceber.lCliente <> 0 Then
        sSelecao = "Cliente = ?"
        iPreenchido = 1
        colSelecao.Add (objTituloReceber.lCliente)
    End If
    
    If objTituloReceber.iFilial <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND Filial = ?"
        Else
            iPreenchido = 1
            sSelecao = "Filial = ?"
        End If
        colSelecao.Add (objTituloReceber.iFilial)
    End If
    
    If Len(Trim(objTituloReceber.sSiglaDocumento)) <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND SiglaDocumento = ?"
        Else
            iPreenchido = 1
            sSelecao = "SiglaDocumento = ?"
        End If
        colSelecao.Add (objTituloReceber.sSiglaDocumento)
    End If
    
    If objTituloReceber.lNumTitulo <> 0 Then
        If iPreenchido = 1 Then
            sSelecao = sSelecao & " AND NumTitulo = ?"
        Else
            iPreenchido = 1
            sSelecao = "NumTitulo = ?"
        End If
        colSelecao.Add (objTituloReceber.lNumTitulo)
    End If

    'Chama Tela TituloReceberLista
    Call Chama_Tela("ParcelasRecDifTituloLista", colSelecao, objParcelaRecDif, Nothing, sSelecao)
            
    Exit Sub

Erro_BotaoDif_Click:

    Select Case Err

        Case 177922

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 177923)

    End Select

    Exit Sub
End Sub


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

    objItensPRJCR.sItem = "GR"
    objItensPRJCR.sDescricao = "Geral"

    objItensPRJCR.dValorOriginal = StrParaDbl(ValorTitulo.Caption)

    objItensPRJCR.iTipoOrigem = PRJ_CR_TIPO_TITREC
    
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

Public Sub BotaoImprimirRecibo_Click()

Dim lErro As Long
Dim iParcela As Integer
Dim objParcRecBaixa As New ClassBaixaParcRec
Dim objTituloReceber As New ClassTituloReceber
Dim objParcelaReceber As New ClassParcelaReceber

On Error GoTo Erro_BotaoImprimirRecibo_Click

    If Sequencial.ListIndex = -1 Then gError 202236

    'Critica os Campos se estão Preenchidos
    lErro = Critica_CamposPreenchidos()
    If lErro <> SUCESSO Then gError 182869

    'Verifica se já existe alguma Parcela na Tela
    If Len(Trim(Parcela(1).Caption)) > 0 Then
        iParcela = CInt(Parcela(1).Caption)
    Else
        iParcela = 0
    End If

    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloReceber.lNumTitulo = CLng(NumeroTitulo.Text)
    objTituloReceber.iFilialEmpresa = iFilialEmpresaTela
    objParcelaReceber.iNumParcela = iParcela

    lErro = CF("ParcelasRec_Le_Todas_NumTitulo", objTituloReceber, objParcelaReceber)
    If lErro <> SUCESSO And lErro <> 60003 Then gError 182870

    If lErro = 60003 Then gError 182871

    If Sequencial.ListCount = 0 Then gError 202237

    'Se for uma Parcela Baixada Traz os dados da Baixa
    If Len(Trim(Sequencial.Text)) = 0 Then gError 202236

    'Lê a Parcela no BD
    lErro = CF("BaixaRecCancelar_Le_Parcela", objParcelaReceber.lNumIntTitulo, iParcela, Sequencial.List(Sequencial.ListIndex), objParcRecBaixa)
    If lErro <> SUCESSO And lErro <> 46230 Then gError 182872

    'Se encontrou a Parcela
    If lErro = SUCESSO Then

        'Se encontrou e Baixa já está cancelada --> Erro
        If objParcRecBaixa.iStatus = STATUS_EXCLUIDO Then gError 182873

        lErro = CF("RelRecibo_Prepara", objParcRecBaixa.lNumIntBaixa)
        If lErro <> SUCESSO Then gError 182874

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

Public Sub BotaoCTBBaixa_Click()
    Call CF("Lancamentos_Abre_Tela", TRANSACAOCTBORIGEM_BAIXASREC, gobjBaixaReceber.lNumIntBaixa, TRANSACAOCTBORIGEM_BAIXASPARCREC, gobjParcRecBaixa.lNumIntDoc)
End Sub
