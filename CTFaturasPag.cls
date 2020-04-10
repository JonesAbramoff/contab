VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTFaturasPag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim giTeste As Integer

Dim m_objUserControl As Object

Public gobjInfoUsu As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'inicio contabilidade
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
Private Const VALOR_IR As String = "Valor_IRRF"
Private Const CONTA_DESP_ESTOQUE As String = "Conta_Desp_Estoque"
Private Const CONTA_DESP_EST_FORN As String = "Conta_Desp_Est_Forn"

'fim contabilidade

Public iAlterado As Integer
Private iFrameAtual As Integer
Private iFornecedorAlterado As Integer
Private iValorTotalAlterado As Integer
Private iEmissaoAlterada As Integer
Private iValorIRRFAlterado As Integer
Private sOldFornecedor As String
Private sOldFilial As String
Private sOldNumTitulo As String
Private dtOldDataEmissao As Date

Public objGridNFiscais As AdmGrid
Public iGrid_FilialNF_Col As Integer
Public iGrid_Numero_Col As Integer
Public iGrid_Emissao_Col As Integer
Public iGrid_ValorNF_Col As Integer
Public iGrid_ValorNFBruto_Col As Integer
Public iGrid_Selecionada_Col As Integer
Public iGrid_FilialFornecedor_Col As Integer

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

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1

'obs.: nao remover deste fonte
Private Const NUM_LINHAS_GRID_NF = 9 'numero minimo de linhas visiveis no grid de nfs

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Cobranca = 2
Private Const TAB_Contabilizacao = 3

'--- inicio dos properties get dos controles da tela

Public Property Get ValorNFBruto() As Object
     Set ValorNFBruto = objUserControl.Controls("ValorNFBruto")
End Property

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get SSFrame3() As Object
     Set SSFrame3 = objUserControl.Controls("SSFrame3")
End Property

Public Property Get CobrancaAutomatica() As Object
     Set CobrancaAutomatica = objUserControl.Controls("CobrancaAutomatica")
End Property

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get TipoCobranca() As Object
     Set TipoCobranca = objUserControl.Controls("TipoCobranca")
End Property

Public Property Get Suspenso() As Object
     Set Suspenso = objUserControl.Controls("Suspenso")
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

Public Property Get SSFrame1() As Object
     Set SSFrame1 = objUserControl.Controls("SSFrame1")
End Property

Public Property Get FilialNF() As Object
     Set FilialNF = objUserControl.Controls("FilialNF")
End Property

Public Property Get FilialFornecedor() As Object
     Set FilialFornecedor = objUserControl.Controls("FilialFornecedor")
End Property

Public Property Get BotaoConsultarNFiscal() As Object
     Set BotaoConsultarNFiscal = objUserControl.Controls("BotaoConsultarNFiscal")
End Property

Public Property Get BotaoAtualizarNFs() As Object
     Set BotaoAtualizarNFs = objUserControl.Controls("BotaoAtualizarNFs")
End Property

Public Property Get NFRegistrar() As Object
     Set NFRegistrar = objUserControl.Controls("NFRegistrar")
End Property

Public Property Get Selecionada() As Object
     Set Selecionada = objUserControl.Controls("Selecionada")
End Property

Public Property Get ValorNF() As Object
     Set ValorNF = objUserControl.Controls("ValorNF")
End Property

Public Property Get DataEmissaoNF() As Object
     Set DataEmissaoNF = objUserControl.Controls("DataEmissaoNF")
End Property

Public Property Get NumNotaFiscal() As Object
     Set NumNotaFiscal = objUserControl.Controls("NumNotaFiscal")
End Property

Public Property Get GridNF() As Object
     Set GridNF = objUserControl.Controls("GridNF")
End Property

Public Property Get ValorTotalNFSelecionadas() As Object
     Set ValorTotalNFSelecionadas = objUserControl.Controls("ValorTotalNFSelecionadas")
End Property

Public Property Get NumNFSelecionadas() As Object
     Set NumNFSelecionadas = objUserControl.Controls("NumNFSelecionadas")
End Property

Public Property Get SSFrame2() As Object
     Set SSFrame2 = objUserControl.Controls("SSFrame2")
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

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get ValorTotal() As Object
     Set ValorTotal = objUserControl.Controls("ValorTotal")
End Property

Public Property Get ValorIRRF() As Object
     Set ValorIRRF = objUserControl.Controls("ValorIRRF")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get FornecedorLabel() As Object
     Set FornecedorLabel = objUserControl.Controls("FornecedorLabel")
End Property

Public Property Get CTBGerencial() As Object
     Set CTBGerencial = objUserControl.Controls("CTBGerencial")
End Property

Public Property Get CTBLancAutomatico() As Object
     Set CTBLancAutomatico = objUserControl.Controls("CTBLancAutomatico")
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

Public Property Get CTBListHistoricos() As Object
     Set CTBListHistoricos = objUserControl.Controls("CTBListHistoricos")
End Property

Public Property Get CTBHistorico() As Object
     Set CTBHistorico = objUserControl.Controls("CTBHistorico")
End Property

Public Property Get CTBAglutina() As Object
     Set CTBAglutina = objUserControl.Controls("CTBAglutina")
End Property

Public Property Get CTBBotaoImprimir() As Object
     Set CTBBotaoImprimir = objUserControl.Controls("CTBBotaoImprimir")
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

Public Property Get CTBLabel1() As Object
     Set CTBLabel1 = objUserControl.Controls("CTBLabel1")
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

Public Property Get CTBContaBtn() As Object
     Set CTBContaBtn = objUserControl.Controls("CTBContaBtn")
End Property

Public Property Get CTBCclBtn() As Object
     Set CTBCclBtn = objUserControl.Controls("CTBCclBtn")
End Property

Public Property Get CTBHistBtn() As Object
     Set CTBHistBtn = objUserControl.Controls("CTBHistBtn")
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

'--- fim dos properties get dos controles da tela

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
    iFrameAtual = 1
    iFornecedorAlterado = 0
    iEmissaoAlterada = 0
    iValorIRRFAlterado = 0
    
    Set objEventoFornecedor = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoCondPagto = New AdmEvento

    'Carrega a List da Combo de Tipos de Cobrança
    lErro = Carrega_TipoCobranca()
    If lErro <> SUCESSO Then gError 18597

    'Carrega a List da Combo de Condições de Pagamento
        lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 18600

    'Inicializa o GridParcelas
    Set objGridParcelas = New AdmGrid
    lErro = Inicializa_GridParcelas(objGridParcelas)
    If lErro <> SUCESSO Then gError 18601
    
    'Carrega na combo de Banco de Cobradores
    lErro = Carrega_Bancos()
    If lErro <> SUCESSO Then gError 18601
    
    'Carrega na combo de Portadores
    lErro = Carrega_Portadores()
    If lErro <> SUCESSO Then gError 18601

    'Inicializa o GridNF
    Set objGridNFiscais = New AdmGrid
    lErro = CF2(Me, "FATPAG_Inicializa_GridNF", objGridNFiscais)
    If lErro <> SUCESSO Then gError 18602
    
    'Carrega Motivos de diferença
    lErro = Carrega_MotivoDiferenca()
    If lErro <> SUCESSO Then gError 500060
    
    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASAPAGAR)
    If lErro <> SUCESSO Then gError 184125
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 18597, 18600, 18601, 18602, 184125, 500060

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160143)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colNFPag As New ColNFsPag

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosPag"

    'Lê os dados da Tela Notas Fiscais a Pagar
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas, colNFPag)
    If lErro <> SUCESSO Then gError 18782

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Fornecedor", objTituloPagar.lFornecedor, 0, "Fornecedor"
    colCampoValor.Add "Filial", objTituloPagar.iFilial, 0, "Filial"
    colCampoValor.Add "NumTitulo", objTituloPagar.lNumTitulo, 0, "NumTitulo"
    colCampoValor.Add "DataEmissao", objTituloPagar.dtDataEmissao, 0, "DataEmissao"
    colCampoValor.Add "ValorTotal", objTituloPagar.dValorTotal, 0, "ValorTotal"
    colCampoValor.Add "NumParcelas", objTituloPagar.iNumParcelas, 0, "NumParcelas"
    colCampoValor.Add "NumIntDoc", objTituloPagar.lNumIntDoc, 0, "NumIntDoc"
    colCampoValor.Add "CondicaoPagto", objTituloPagar.iCondicaoPagto, 0, "CondicaoPagto"
    colCampoValor.Add "ValorIRRF", objTituloPagar.dValorIRRF, 0, "ValorIRRF"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "Status", OP_DIFERENTE, STATUS_EXCLUIDO
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    colSelecao.Add "SiglaDocumento", OP_IGUAL, TIPODOC_FATURA_A_PAGAR
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 18782

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160144)

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
        objTituloPagar.dtDataEmissao = colCampoValor.Item("DataEmissao").vValor
        objTituloPagar.iNumParcelas = colCampoValor.Item("NumParcelas").vValor
        objTituloPagar.lNumTitulo = colCampoValor.Item("NumTitulo").vValor
        objTituloPagar.dValorTotal = colCampoValor.Item("ValorTotal").vValor
        objTituloPagar.iCondicaoPagto = colCampoValor.Item("CondicaoPagto").vValor
        objTituloPagar.dValorIRRF = colCampoValor.Item("ValorIRRF").vValor
        
        'Traz os dados da Fatura, passada em objTituloPagar, para a tela
        lErro = Traz_FaturaPagar_Tela(objTituloPagar)
        If lErro <> SUCESSO Then gError 18786

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 18786

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160145)

    End Select

    Exit Sub

End Sub

Public Sub BotaoAtualizarNFs_Click()

    'força a atualizacao das lista de nfs no grid
    Call Atualiza_NotasFiscais(1)
        
End Sub

Public Sub BotaoConsultarNFiscal_Click()

Dim lErro As Long
Dim objNFsPag As New ClassNFsPag
Dim sTela As String
Dim objNFiscal As New ClassNFiscal
Dim colSelecao As New Collection
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_BotaoConsultarNFiscal_Click

    'Verifica se uma linha do Grid foi selecionada
    If GridNF.Row <= 0 Then gError 79387
    
    'Critica se a linha selecionada está preenchida
    If Len(Trim(GridNF.TextMatrix(GridNF.Row, iGrid_Numero_Col))) = 0 Then gError 79388
    
    objFornecedor.sNomeReduzido = Fornecedor.Text

    'Lê o codigo do Fonecedor através do Nome Reduzido
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 79389

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 79390

    objNFsPag.lFornecedor = objFornecedor.lCodigo
    objNFsPag.iFilial = Codigo_Extrai(Filial.Text)
    objNFsPag.lNumNotaFiscal = StrParaLong(GridNF.TextMatrix(GridNF.Row, iGrid_Numero_Col))
    objNFsPag.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF2(Me, "FATPAG_BotaoConsultarNFiscal_Click", objNFsPag)
    If lErro <> SUCESSO Then gError 79389
    
    If Len(Trim(GridNF.TextMatrix(GridNF.Row, iGrid_Emissao_Col))) > 0 Then
        objNFsPag.dtDataEmissao = CDate(GridNF.TextMatrix(GridNF.Row, iGrid_Emissao_Col))
    Else
        objNFsPag.dtDataEmissao = DATA_NULA
    End If
    
    'Procura o Titulo (apenas na tabela de NFiscais em aberto)
    lErro = CF("NFPag_Le_Numero", objNFsPag)
    If lErro <> SUCESSO And lErro <> 18338 Then gError 79391

    'Le o Nome da Tela que originou este Título
    lErro = CF("Titulo_Le_DocumentoOriginal", objNFsPag.lNumIntDoc, CPR_NF_PAGAR, objNFiscal, sTela)
    If lErro <> SUCESSO And lErro <> 58942 Then gError 79392
    
    If lErro = SUCESSO Then
        'Chama a Tela Estoque
        Call Chama_Tela(sTela, objNFiscal)
    Else
        Call Chama_Tela("NFPag_Consulta", objNFsPag)
    End If

    Exit Sub

Erro_BotaoConsultarNFiscal_Click:

    Select Case gErr

        Case 79389, 79391, 79392 'Tratado na Rotina chamada

        Case 79387
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 79388
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_PREENCHIDA", gErr)
        
        Case 79390
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160146)

    End Select

    Exit Sub

End Sub

Public Function FATPAG_BotaoConsultarNFiscal_Click(ByVal objCT As Object, ByVal objNFsPag As ClassNFsPag) As Long
    FATPAG_BotaoConsultarNFiscal_Click = SUCESSO
End Function

Public Sub CondicaoPagamento_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_CondicaoPagamento_Validate

    'Verifica se a Condicao Pagamento foi preenchida
    If Len(Trim(CondicaoPagamento.Text)) = 0 Then Exit Sub

    'Verifica se é uma Condicao Pagamento selecionada
    If CondicaoPagamento.Text = CondicaoPagamento.List(CondicaoPagamento.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(CondicaoPagamento, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18660

    'Se não encontra, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 18661
        If lErro <> SUCESSO Then gError 18663

        'Testa se pode ser usada em Contas a Pagar
        If objCondicaoPagto.iEmPagamento = 0 Then gError 18784

        'Coloca na tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)

        Call Recalcula_Cobranca
        
    End If

    'Não encontrou e é STRING
    If lErro = 6731 Then gError 18664

    Exit Sub

Erro_CondicaoPagamento_Validate:

    Cancel = True
    
    Select Case gErr

       Case 18660, 18661, 18662

       Case 18663
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 18664
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagamento.Text)
        
        Case 18784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_PAGAMENTO", gErr, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160147)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_GotFocus()
Dim iEmissaoAux As Integer
    
    iEmissaoAux = iEmissaoAlterada
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    iEmissaoAlterada = iEmissaoAux

End Sub

Public Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim sFornecedor As String
Dim objFornecedor As New ClassFornecedor
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18624

    'Não encontrou, mas extrai o código
    If lErro = 6730 Then

        'Verifica se Fornecedor foi digitado
        If Len(Trim(Fornecedor.Text)) = 0 Then gError 18626

        sFornecedor = Fornecedor.Text

        objFilialFornecedor.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o codigo extraido
        lErro = CF("FilialFornecedor_Le_NomeRed_CodFilial", sFornecedor, objFilialFornecedor)
        If lErro <> SUCESSO And lErro <> 18272 Then gError 18625
        
        If lErro <> SUCESSO Then gError 18627

        'Coloca na tela
        Filial.Text = iCodigo & SEPARADOR & objFilialFornecedor.sNome

        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        Call Atualiza_NotasFiscais
    
    End If

    'Não encontrou e é STRING
    If lErro = 6731 Then gError 18628

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr
       
       Case 18624, 18625

       Case 18626
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
       
       Case 18627
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALFORNECEDOR", iCodigo, Fornecedor.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisFornecedores", objFilialFornecedor)
            End If

        Case 18628
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALFORNECEDOR_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160148)

    End Select

    Exit Sub

End Sub

Public Sub FilialNF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub FornecedorLabel_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    'Preenche NomeReduzido com o Fornecedor da tela
    If Len(Trim(Fornecedor.Text)) > 0 Then objFornecedor.sNomeReduzido = Fornecedor.Text
    
    'Chama a Tela de com a lista de Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)
    
End Sub

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)

End Sub

Public Sub NumTitulo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor, Cancel As Boolean

    Set objFornecedor = obj1

    'Coloca a descrição do Fornecedor na tela
    Fornecedor.Text = objFornecedor.sNomeReduzido
    
    Call Fornecedor_Validate(Cancel)
    
    Me.Show

    Exit Sub

End Sub

Public Sub NumeroLabel_Click()

Dim objTituloPagar As New ClassTituloPagar
Dim colParcelas As New colParcelaPagar
Dim colNFPagar As New ColNFsPag
Dim colSelecao As New Collection
Dim lErro As Long

On Error GoTo Erro_NumeroLabel_Click
    
    'Verifica se o Fornecedor e a Filial foram informados
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 18603
    If Len(Trim(Filial.Text)) = 0 Then gError 18604

    'Recolhe os dados da tela
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelas, colNFPagar)
    If lErro <> SUCESSO Then gError 18605

    'Adiciona o código do Fornecedor e o código da Filial
    colSelecao.Add objTituloPagar.lFornecedor
    colSelecao.Add objTituloPagar.iFilial

    'Chama a Tela com a Lista de Faturas
    Call Chama_Tela("FaturasPagLista", colSelecao, objTituloPagar, objEventoNumero)

    Exit Sub

Erro_NumeroLabel_Click:

    Select Case gErr

        Case 18603
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)

        Case 18604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 18605

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160149)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTituloPagar As ClassTituloPagar

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTituloPagar = obj1

    'Traz para a Tela os dados da Fatura passada em objTituloPagar
    lErro = Traz_FaturaPagar_Tela(objTituloPagar)
    If lErro <> SUCESSO Then gError 18604

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 18604

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160150)

    End Select

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As Collection

    'Verifica se a condicao está preenchida
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

    'Chama a Tela com a Lista das Condicões de Pagamento disponíveis
    Call Chama_Tela("CondicaoPagtoCPLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

End Sub

Private Sub objEventoCondPagto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCondicaoPagto As ClassCondicaoPagto

On Error GoTo Erro_objEventoCondPagto_evSelecao

    Set objCondicaoPagto = obj1

    'Coloca na Tela a condicao de Pagamento retornada
    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo)
    Call CondicaoPagamento_Validate(bSGECancelDummy)

    Exit Sub

Erro_objEventoCondPagto_evSelecao:

    Select Case gErr

        Case 18605

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160151)

     End Select

     Exit Sub

End Sub

Function Trata_Parametros(Optional objTituloPagar As ClassTituloPagar) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloPagar Is Nothing) Then

        'Lê o Título
        lErro = CF("TituloPagar_Le", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18372 Then gError 18606
        
        If lErro <> SUCESSO Then gError 18384
            
        'Verifica se é Fatura
        If objTituloPagar.sSiglaDocumento <> TIPODOC_FATURA_A_PAGAR Then gError 18607

        'Traz os dados para a tela
        lErro = Traz_FaturaPagar_Tela(objTituloPagar)
        If lErro <> SUCESSO Then gError 18608

    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 18606, 18608

        Case 18607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_FATURAPAGAR", gErr, objTituloPagar.lNumTitulo)

        Case 18384
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURAPAG_NAO_CADASTRADA", gErr, objTituloPagar.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160152)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Traz_FaturaPagar_Tela(objTituloPagar As ClassTituloPagar) As Long
'Coloca na Tela os dados da Fatura passada em objTituloPagar

Dim lErro As Long
Dim colParcelasPag As New colParcelaPagar
Dim objParcelaPagar As ClassParcelaPagar
Dim colNFPagAberta As New ColNFsPag
Dim colNFPagVinculada As New ColNFsPag
Dim objNFPag As ClassNFsPag
Dim iLinha As Integer
Dim iIndice As Integer, bCancel As Boolean
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim dValorTotalParc As Double
Dim dValorTotalOrig As Double

On Error GoTo Erro_Traz_FaturaPagar_Tela

    Call Limpa_Tela_FaturasPag
    
    'Preenche o Cabeçalho
    Fornecedor.Text = objTituloPagar.lFornecedor
    Call Fornecedor_Validate(bCancel)

    Filial.Text = objTituloPagar.iFilial
    Call Filial_Validate(bCancel)

    NumTitulo.Text = objTituloPagar.lNumTitulo
    
    Call DateParaMasked(DataEmissao, objTituloPagar.dtDataEmissao)

    ValorTotal.Text = Format(objTituloPagar.dValorTotal, "Standard")
    ValorIRRF.Text = objTituloPagar.dValorIRRF
          
    If objTituloPagar.iCondicaoPagto <> 0 Then
    
        CondicaoPagamento.Text = CStr(objTituloPagar.iCondicaoPagto)
        Call CondicaoPagamento_Validate(bCancel)
    
    Else
    
        CondicaoPagamento.Text = ""
    
    End If
    
    'Lê as Parcelas a Pagar vinculadas ao Título
    lErro = CF("ParcelasPagar_Le", objTituloPagar, colParcelasPag)
    If lErro <> SUCESSO Then gError 18608

    'Verifica se o número de parcelas no BD é superior ao máximo
    If colParcelasPag.Count > NUM_MAXIMO_PARCELAS Then gError 18787
    
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
        
        For iIndice = 0 To TipoCobranca.ListCount - 1
            If TipoCobranca.ItemData(iIndice) = objParcelaPagar.iTipoCobranca Then
                GridParcelas.TextMatrix(iLinha, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                Exit For
            End If
        Next

        If objParcelaPagar.iStatus = STATUS_SUSPENSO Then
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "1"
        Else
            GridParcelas.TextMatrix(iLinha, iGrid_Suspenso_Col) = "0"
        End If
    
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

        CodigodeBarras.PromptInclude = False
        CodigodeBarras.Text = objParcelaPagar.sCodigoDeBarras
        CodigodeBarras.PromptInclude = True
        GridParcelas.TextMatrix(iLinha, iGrid_CodigoDeBarras_Col) = CodigodeBarras.Text
    
        'Acumula valores e valores originais
        dValorTotalParc = dValorTotalParc + objParcelaPagar.dValor
        dValorTotalOrig = dValorTotalOrig + objParcelaPagar.dValorOriginal
    
    Next

    'Coloca o Numero de Linhas existentes do Grid de Parcelas
    objGridParcelas.iLinhasExistentes = iLinha
    Call Grid_Refresh_Checkbox(objGridParcelas)

    'Atualiza as Notas Fiscais no Grid de Notas Fiscais
    Call Atualiza_NotasFiscais

    'traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objTituloPagar.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 184126

    Diferenca.Text = Format(dValorTotalParc - dValorTotalOrig, "Standard")

    iAlterado = 0
    iFornecedorAlterado = 0
    iValorTotalAlterado = 0
    iEmissaoAlterada = 0
    iValorIRRFAlterado = 0
    
    Traz_FaturaPagar_Tela = SUCESSO

    Exit Function

Erro_Traz_FaturaPagar_Tela:

    Select Case gErr

        Case 18606, 184126

        Case 18787
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", gErr, colParcelasPag.Count, NUM_MAXIMO_PARCELAS)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160153)

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
Dim bCancel As Boolean

On Error GoTo Erro_Fornecedor_Validate

    If iFornecedorAlterado = 1 Then

        If Len(Trim(Fornecedor.Text)) > 0 Then

            'Tenta ler o Fornecedor (NomeReduzido ou Código ou CPF ou CGC)
            lErro = TP_Fornecedor_Le(Fornecedor, objFornecedor, iCodFilial)
            If lErro <> SUCESSO Then gError 26023

            'Lê coleção de códigos, nomes de Filiais do Fornecedor
            lErro = CF("FiliaisFornecedores_Le_Fornecedor", objFornecedor, colCodigoNome)
            If lErro <> SUCESSO Then gError 26024

            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)

            'CODIGO ESPECÍFICO
            'Verifica se chave de TipoFornecedor está preenchida
            If objFornecedor.iTipo > 0 Then

                'le os dados do Tipo de fornecedor
                lErro = TipoFornecedor_Dados(objFornecedor, objTipoFornecedor)
                If lErro <> SUCESSO Then gError 26025

            End If
            
            'Se Cond Pagto de Fornecedor está preenchida
            If objFornecedor.iCondicaoPagto <> 0 Then
              
                objCondicaoPagto.iCodigo = objFornecedor.iCondicaoPagto
                
                'Lê a Cond Pagto
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 26026
                If lErro = 19205 Then gError 26028
                
                If objCondicaoPagto.iEmPagamento = 1 Then
                    'Coloca na Tela
                    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo)
                    Call CondicaoPagamento_Validate(bCancel)
                End If
            'Se Cond Pagto de TipoFornecedor está preenchida
            ElseIf objTipoFornecedor.iCondicaoPagto <> 0 Then
                
                objCondicaoPagto.iCodigo = objTipoFornecedor.iCondicaoPagto
                
                'Lê a Cond Pagto
                lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
                If lErro <> SUCESSO And lErro <> 19205 Then gError 26027
                If lErro = 19205 Then gError 26029
                
                If objCondicaoPagto.iEmPagamento = 1 Then
                    'Coloca na Tela
                    CondicaoPagamento.Text = CStr(objCondicaoPagto.iCodigo)
                    Call CondicaoPagamento_Validate(bCancel)
                End If
                
            End If

            'Atualiza as Notas Fiscais no Grid de Notas Fiscais
            Call Atualiza_NotasFiscais
            
            'FIM CODIGO ESPECÍFICO
            
        ElseIf Len(Trim(Fornecedor.Text)) = 0 Then
            
            'Limpa Combo de Filial
            Filial.Clear

        End If

        iFornecedorAlterado = 0

    End If

    Exit Sub

Erro_Fornecedor_Validate:
    
    Cancel = True

    Select Case gErr

        Case 26023, 26024, 26025, 26026, 26027
        
        Case 26028, 26029
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160154)

    End Select

    Exit Sub

End Sub

Private Function Atualiza_NotasFiscais(Optional iForcada As Integer = 0) As Long
'Atualiza no Grid as Notas Fiscais a Pagar de acordo com os dados passados por parâmetro

Dim lErro As Long
Dim lErro1 As Integer
Dim colNfsPagAberta As New ColNFsPag
Dim colNFsPagVinculada As New ColNFsPag
Dim objFilialFornecedor As New ClassFilialFornecedor
Dim objTituloPagar As New ClassTituloPagar
Dim objFornecedor As New ClassFornecedor
Dim sFornecedor As String
Dim sFilial As String
Dim sNumTitulo As String
Dim dtDataEmissao As Date

On Error GoTo Erro_Atualiza_NotasFiscais

    'Recolhe os dados para fazera atualizacao das Notas Fiscais
    sFornecedor = Fornecedor.Text
    sFilial = Filial.Text
    sNumTitulo = NumTitulo.Text
    dtDataEmissao = MaskedParaDate(DataEmissao)

    'Verifica se os dados da tela continuam os mesmos
    If iForcada = 0 And sFornecedor = sOldFornecedor And sFilial = sOldFilial And sNumTitulo = sOldNumTitulo And dtDataEmissao = dtOldDataEmissao Then Exit Function

    'Verifica se todos estão preenchidos
    If Len(Trim(sFornecedor)) = 0 Or Len(Trim(sFilial)) = 0 Or Len(Trim(sNumTitulo)) = 0 Then Exit Function

    objFornecedor.sNomeReduzido = sFornecedor
    
    'Lê o Fornecedor
    lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
    If lErro <> SUCESSO Then gError 18648

    'Carrega o fornecedor e filial em objFilialFornecedor
    objFilialFornecedor.lCodFornecedor = objFornecedor.lCodigo
    objFilialFornecedor.iCodFilial = Codigo_Extrai(sFilial)

    'Guarda os dados do Título
    objTituloPagar.lFornecedor = objFornecedor.lCodigo
    objTituloPagar.iFilial = objFilialFornecedor.iCodFilial
    objTituloPagar.lNumTitulo = CLng(sNumTitulo)
    objTituloPagar.dtDataEmissao = dtDataEmissao
    objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR
    objTituloPagar.iFilialEmpresa = giFilialEmpresa

    'Lê o Título
    lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18551 Then gError 18640
    
    'se é uma fatura já cadastrada
    If lErro = SUCESSO Then
    
        'Lê as Notas Fiscais Vinculadas a Fatura a Pagar
        lErro = CF("NfsPag_Le_FaturaPagar", objTituloPagar, colNFsPagVinculada)
        If lErro <> SUCESSO And lErro <> 26020 Then gError 18642
    
    Else
    
        lErro1 = CF("NfsPag_Le_FilialFornecedor_Desvinculadas", objFilialFornecedor, colNfsPagAberta)
        If lErro1 <> SUCESSO Then gError 18639

    End If
    
    'Limpa o Grid
    Call Grid_Limpa(objGridNFiscais)
    Call Ordenacao_Limpa(objGridNFiscais)
    ValorTotalNFSelecionadas.Caption = ""
    NumNFSelecionadas.Caption = ""
        
    'joga p/o grid as nfs desvinculadas
    Call Carrega_GridNF(colNfsPagAberta, colNFsPagVinculada)

    'Atualiza os dados
    sOldFornecedor = objFornecedor.sNomeReduzido
    sOldFilial = sFilial
    sOldNumTitulo = sNumTitulo
    dtOldDataEmissao = dtDataEmissao
    
    Call CF2(Me, "FATPAG_Calcula_Dif")

    Atualiza_NotasFiscais = SUCESSO

    Exit Function

Erro_Atualiza_NotasFiscais:

    Atualiza_NotasFiscais = gErr

    Select Case gErr

        Case 18639, 18640, 18642, 18648

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160155)

    End Select

    Exit Function

End Function

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    'Verifica se alguma Filial foi selecionada
    If Filial.ListIndex = -1 Then Exit Sub

    'Atualiza as Notas Fiscais no Grid de Notas Fiscais
    Call Atualiza_NotasFiscais
    
    iAlterado = REGISTRO_ALTERADO
    
    Exit Sub
    
Erro_Filial_Click:

    Select Case gErr
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160156)
        
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
    If lErro <> SUCESSO Then gError 18647

    'Atualiza as Notas Fiscais no Grid de Notas Fiscais
    Call Atualiza_NotasFiscais

    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True


    Select Case gErr

        Case 18647

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160157)

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
                Parent.HelpContextID = IDH_FATURAS_PAGAR_ID
                
            Case TAB_Cobranca
                Parent.HelpContextID = IDH_FATURAS_PAGAR_COBRANCA
                        
            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_NOTA_FISCAL_FATURA_CONTABILIZACAO
        
        End Select
        
    End If

End Sub

Public Sub UpDownEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iEmissaoAlterada = REGISTRO_ALTERADO

End Sub

Public Sub ValorTotal_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorTotalAlterado = 1
    
End Sub

Public Sub ValorTotal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ValorTotal_Validate

    'Verifica se algum valor foi digitado
    If Len(Trim(ValorTotal.ClipText)) = 0 Then Exit Sub

    'Critica o valor
    lErro = Valor_Positivo_Critica(ValorTotal.Text)
    If lErro <> SUCESSO Then gError 18651

    'Põe o valor formatado na tela
    ValorTotal.Text = Format(ValorTotal.Text, "Standard")

    If iValorTotalAlterado = 1 Then
    
        Call Recalcula_Cobranca
        iValorTotalAlterado = 0
        
    End If
    
    Call CF2(Me, "FATPAG_Calcula_Dif")
    
    Exit Sub

Erro_ValorTotal_Validate:

    Cancel = True


    Select Case gErr

        Case 18651

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160158)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate

    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then gError 18652

    If iEmissaoAlterada = 1 Then
        
        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        'Call Atualiza_NotasFiscais
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If

    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case gErr

        Case 18652

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160159)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then gError 18654

    If iEmissaoAlterada = 1 Then
        
        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        Call Atualiza_NotasFiscais
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case gErr

        Case 18654

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160160)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then gError 18655

    If iEmissaoAlterada = 1 Then
        
        'Atualiza as Notas Fiscais no Grid de Notas Fiscais
        Call Atualiza_NotasFiscais
        
        'força o recalculo das parcelas
        Call Recalcula_Cobranca
        
        iEmissaoAlterada = 0
    
    End If

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case gErr

        Case 18655

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160161)

    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagamento_Click()

Dim lErro As Long

On Error GoTo Erro_CondicaoPagamento_Click

    iAlterado = REGISTRO_ALTERADO

    Call Recalcula_Cobranca
    
    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160162)

      End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long
Dim dValor As Double
Dim dValorFinal As Double
Dim dtDataEmissao As Date
Dim dtDataVenctoReal As Date
Dim iIndice As Integer, dValorIRRF As Double

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o Grid Parcelas
    Call Grid_Limpa(objGridParcelas)

    dValorIRRF = StrParaDbl(ValorIRRF)
    
    'Valor a Pagar
    If Len(Trim(ValorTotal)) > 0 Then dValor = CDbl(ValorTotal) - dValorIRRF

    'Se Valor a Pagar for positivo
    If dValor > 0 Then

        objCondicaoPagto.dValorTotal = dValor
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then gError 18417

        'Número de Parcelas
        objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        
            'Coloca valor original igual a valor como default
            GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If
    
    'Valor a Pagar
    If Len(Trim(ValorTotal)) > 0 Then dValorFinal = CDbl(ValorTotal) - dValorIRRF - StrParaDbl(Diferenca.Text)

    'Se Valor a Pagar for positivo
    If dValorFinal > 0 And Abs(dValorFinal - dValor) > DELTA_VALORMONETARIO Then

        objCondicaoPagto.dValorTotal = dValorFinal
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then gError 18417

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        
            If MotivoDiferenca.ListCount > 0 Then
                MotivoDiferenca.ListIndex = 0
                GridParcelas.TextMatrix(iIndice, iGrid_MotivoDiferenca_Col) = MotivoDiferenca.Text
            End If
        Next

    End If

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataEmissao.ClipText)) > 0 Then

        dtDataEmissao = CDate(DataEmissao.Text)
        objCondicaoPagto.dtDataEmissao = dtDataEmissao
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then gError 18441

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
            
            lErro = CF("Fatura_Preenche_Cobranca_Cust", objGridParcelas, TipoCobranca, iIndice, iGrid_Cobranca_Col, Fornecedor.Text)
            If lErro <> SUCESSO Then gError 18443

        Next

    End If

    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = gErr

    Select Case gErr

        Case 18417, 18441, 18443

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160163)

    End Select

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 18613

        If objGridInt.objGrid Is GridParcelas Then

            'Verifica qual e coluna atual
            Select Case objGridInt.objGrid.Col

                'faz a Critica da Data de Vencimento e Gera a Data de Vencimento Real
                Case iGrid_Vencimento_Col
                    lErro = Saida_Celula_Vencimento(objGridInt)
                    If lErro <> SUCESSO Then gError 18613

                'faz o tratamento relacionado ao valor da Parcela
                Case iGrid_ValorParcela_Col
                    lErro = Saida_Celula_Valor(objGridInt)
                    If lErro <> SUCESSO Then gError 18614

                'faz o tratamento relacionado ao Tipo de Cobrança
                Case iGrid_Cobranca_Col
                    lErro = Saida_Celula_Cobranca(objGridInt)
                    If lErro <> SUCESSO Then gError 18615
                    
                Case iGrid_Banco_Col
                    lErro = Saida_Celula_Banco(objGridInt)
                    If lErro <> SUCESSO Then gError 18615
            
                Case iGrid_Portador_Col
                    lErro = Saida_Celula_Portador(objGridInt)
                    If lErro <> SUCESSO Then gError 18615
            
                Case iGrid_CodigoDeBarras_Col
                    lErro = Saida_Celula_CodigodeBarras(objGridInt)
                    If lErro <> SUCESSO Then gError 18615

                Case iGrid_ValorOriginal_Col
                    lErro = Saida_Celula_ValorOriginal(objGridInt)
                    If lErro <> SUCESSO Then gError 500050
                
                'Motivo Diferença
                Case iGrid_MotivoDiferenca_Col
                    lErro = Saida_Celula_MotivoDiferenca(objGridInt)
                    If lErro <> SUCESSO Then gError 500051
                    
            End Select
       
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 18617

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 18613, 18614, 18615, 18616, 500050, 500051

        Case 18617
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a critica da celula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = ValorParcela

    'Verifica se o valor parcela foi digitda
    If Len(ValorParcela.ClipText) > 0 Then
        
        'Verifica se o Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then gError 18666
        
        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then CobrancaAutomatica.Value = vbUnchecked
        
        If GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) = GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorParcela_Col) Then GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) = ValorParcela.Text
        
        'No caso de ser uma linha nova incrementa-se o número de linhas existente
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            For iIndice = 0 To TipoCobranca.ListCount - 1
                If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                    Exit For
                End If
            Next
        
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 18667

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 18666, 18667
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160164)

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
        If lErro <> SUCESSO Then gError 18613

        dtDataVencimento = CDate(DataVencimento.Text)

        'Se Data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            
            dtDataEmissao = CDate(DataEmissao.Text)
            If dtDataVencimento < DataEmissao Then gError 18614
        
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
        
        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then gError 18615

        'Coloca a  Data de Vencimento Real na tela
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        'Incrementa o número de linhas existentes caso esteja sendo digitado em uma lina nao existente
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        
            For iIndice = 0 To TipoCobranca.ListCount - 1
                If TipoCobranca.ItemData(iIndice) = TIPO_COBRANCA_CARTEIRA Then
                    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Cobranca_Col) = TipoCobranca.List(iIndice)
                    Exit For
                End If
            Next
        
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 18616

    Saida_Celula_Vencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_Vencimento:

    Saida_Celula_Vencimento = gErr

    Select Case gErr

        Case 18613, 18615, 18616
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18614
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", gErr, DataVencimento.Text, GridParcelas.Row, DataEmissao.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160165)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cobranca(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_Saida_Celula_Cobranca

    Set objGridInt.objControle = TipoCobranca

    'Verifica se o Tipo de Cobrança foi preenchido
    If Len(Trim(TipoCobranca.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoCobranca.Text <> TipoCobranca.List(TipoCobranca.ListIndex) Then

            'Tenta Selecionar na combo o Tipo de Cobrança
            lErro = Combo_Seleciona(TipoCobranca, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 18665

            'Não foi encontrado, mas retornou código
            If lErro = 6730 Then gError 18788
            
            'Não foi encontrado a STRING
            If lErro = 6731 Then gError 18789

        End If
    
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 18668

    Saida_Celula_Cobranca = SUCESSO

    Exit Function

Erro_Saida_Celula_Cobranca:

    Saida_Celula_Cobranca = gErr

    Select Case gErr

        Case 18665
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", gErr, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18668
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 18788
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_CADASTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 18789
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOCOBRANCA_NAO_ENCONTRADO", gErr, TipoCobranca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160166)

    End Select

    Exit Function

End Function

Public Sub Selecionada_Click()

Dim dValorTotalNFSel As Double
Dim dValorNF As Double
Dim iNumSelecionadas As Integer
Dim iSelecionada As Integer

    'Verifica se o click foi dado em uma linha não preenchida do Grid
    If GridNF.Row > objGridNFiscais.iLinhasExistentes Then Exit Sub
    
    'Recolhe os valores da tela
    iSelecionada = CInt(GridNF.TextMatrix(GridNF.Row, iGrid_Selecionada_Col))
    dValorNF = CDbl(GridNF.TextMatrix(GridNF.Row, iGrid_ValorNF_Col))
    
    'Recolhe os valores da tela
    iNumSelecionadas = CInt(NumNFSelecionadas.Caption)
    dValorTotalNFSel = CDbl(ValorTotalNFSelecionadas.Caption)
    
    'atualiza os valores dos acumuladores
    dValorTotalNFSel = dValorTotalNFSel + IIf(iSelecionada = 1, dValorNF, -dValorNF)
    iNumSelecionadas = iNumSelecionadas + IIf(iSelecionada = 1, 1, -1)

    'Atualiza os totais na Tela
    ValorTotalNFSelecionadas.Caption = Format(dValorTotalNFSel, "Standard")
    NumNFSelecionadas.Caption = CStr(iNumSelecionadas)
    
    Call CF2(Me, "FATPAG_Calcula_Dif")

    Exit Sub

End Sub

Public Function FATPAG_Calcula_Dif(ByVal objCT) As Long
    FATPAG_Calcula_Dif = SUCESSO
End Function

Public Sub NFRegistrar_Click()
    
    'Chama a tela de Notas Fiscais
    Call Chama_Tela("NFPag")
    
End Sub

Public Sub BotaoGravar_Click()
'Dispara a gravação de uma Fatura

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'chama a rotina de gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 18669

    'Limpa os campos da tela
    Call Limpa_Tela_FaturasPag

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 18669

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160167)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim dValorParcelas As Double
Dim iIndice As Integer
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelaPagar As New colParcelaPagar
Dim colNFPag As New ColNFsPag
Dim dtDataVencimento As Date, dtDataEmissao As Date
Dim dValorTotal As Double
Dim dValorIRRF As Double
Dim dSomaParcOriginal As Double

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.ClipText)) = 0 Then gError 18670

    'Verifica se a Filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 18671

    'Verifica se o Número do Título foi preenchido
    If Len(Trim(NumTitulo.ClipText)) = 0 Then gError 18673

    'Verifica se o Valor Total foi preenchido
    If Len(Trim(ValorTotal.ClipText)) = 0 Then gError 18684
    
    'Verifica se exitem parcelas no Grid para gravar
    If objGridParcelas.iLinhasExistentes = 0 Then gError 18686
    
    dtDataEmissao = MaskedParaDate(DataEmissao)
    dValorParcelas = 0
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        'Verifica se a Data de Vencimento foi informada
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 18687
    
        dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        
        'Verifica se a Data de Vencimento é maior ou igual à Data de Emissao
        If dtDataEmissao <> DATA_NULA And dtDataVencimento < dtDataEmissao Then gError 18688
        
        'Verifica se as parcelas estão ordenadas pela Data de Vencimento
        If iIndice > 1 Then
            If dtDataVencimento < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then gError 18689
        End If
        
        'Verifica se o Valor das Parcelas foi informado
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) = 0 Then gError 18690
             
        'Verifica se Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        If lErro <> SUCESSO Then gError 18691
        
        'Verifica se Valor Original da Parcela foi preenchido
        If StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col)) = 0 Then 'gError 500057
            GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col) = GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col)
        End If
        
        'Verifica se o Motivo da Diferença foi preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_MotivoDiferenca_Col))) = 0 And (StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col)) <> StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col))) Then gError 500058
        
        'Se Valor e Valor Original são iguais e Motivo Diferença foi preenchido, erro
        If StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col)) = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col)) And Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_MotivoDiferenca_Col))) > 0 Then gError 500064
        
        'Acumula o Valor Original das Parcelas
        dSomaParcOriginal = dSomaParcOriginal + StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col))

        'Acumula Valor Parcela em dSomaParcelas
        dValorParcelas = dValorParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        
    Next
    
    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
    
    dValorTotal = CDbl(ValorTotal.Text)
    
    'Compara Soma das Parcelas com Valor a Pagar e o Valor de diferença
    If Format(dValorParcelas, "0.00") <> Format(dValorTotal - dValorIRRF, "0.00") Then gError 18692
        
    'Compara a Soma dos Valores Originais com o Valor à Pagar
    If Format(dSomaParcOriginal, "0.00") <> Format(dValorTotal - StrParaDbl(Diferenca.Text) - dValorIRRF, "0.00") Then gError 500059
    
    'Verifica se Valor Total = soma Notas Fiscais selecionadas
    If Abs((StrParaDbl(ValorTotal.Text) - StrParaDbl(Diferenca.Text)) - StrParaDbl(ValorTotalNFSelecionadas.Caption)) > DELTA_VALORMONETARIO Then gError 18693
    
    'Move dados da Tela para objTituloPagar, colParcelaPagar e colNFPag
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelaPagar, colNFPag)
    If lErro <> SUCESSO Then gError 18694
    
    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    If StrParaDate(DataEmissao.Text) <> DATA_NULA Then
        lErro = objContabil.Contabil_Testa_Data(StrParaDate(DataEmissao.Text))
        If lErro <> SUCESSO Then gError 182905
    End If
    
    'Grava no BD
    lErro = CF("FaturaPagar_Grava", objTituloPagar, colParcelaPagar, colNFPag, objContabil)
    If lErro <> SUCESSO Then gError 18695
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 18670
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
        
        Case 18671
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
                
        Case 18673
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
        
        Case 18684
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORTOTAL_NAO_INFORMADO", gErr)
        
        Case 18686
            Call Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", gErr)

        Case 18687
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", gErr, iIndice)
            
        Case 18688
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", gErr, dtDataVencimento, DataEmissao.Text, iIndice)
        
        Case 18689
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_NAO_ORDENADA", gErr)
        
        Case 18690
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_NAO_INFORMADO", gErr, iIndice)
            
        Case 18691, 18694, 18695, 182905
        
        Case 18692
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_INVALIDA", gErr, dValorParcelas, dValorTotal)
        
        Case 18693
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_NFS_SELECIONADAS_INVALIDA", gErr, ValorTotalNFSelecionadas.Caption, ValorTotal.Text)
    
        Case 19392
            Call Rotina_Erro(vbOKOnly, "ERRO_NUM_MAX_NFS_SELEC_EXCEDIDO", gErr, NUM_MAXIMO_NF_VINCULADA_FATURA)
        
        Case 500057
            Call Rotina_Erro(vbOKOnly, "ERRO_VALORORIGINALPARC_NAO_INFORMADO", gErr, iIndice)
        
        Case 500058
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCAPARC_NAO_INFORMADO", gErr, iIndice)
        
        Case 500059
            Call Rotina_Erro(vbOKOnly, "ERRO_SOMA_VALORORIGINAL_DIFERENTE_VALORTOTAL", gErr, dSomaParcOriginal, dValorTotal - StrParaDbl(Diferenca.Text))
        
        Case 500064
            Call Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCA_INFORMADO_ERRADO", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160168)
            
    End Select
        
    Exit Function
        
End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTituloPagar As New ClassTituloPagar
Dim colParcelaPagar As New colParcelaPagar
Dim colNFPagar As New ColNFsPag
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) = 0 Then gError 18752
    
    'Verifica se Filial foi preenchida
    If Len(Trim(Filial.Text)) = 0 Then gError 18753
    
    'Verifica se o número do Título foi preenchido
    If Len(Trim(NumTitulo.Text)) = 0 Then gError 18754

    'Guarda os dados da tela em objTituloPagar, colParcelaPagar, colNFPagar
    lErro = Move_Tela_Memoria(objTituloPagar, colParcelaPagar, colNFPagar)
    If lErro <> SUCESSO Then gError 18755
    
    'Tenta ler Fatura como Baixada
    lErro = CF("TituloPagarBaixado_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18556 Then gError 18758
    
    'Se encontrou --> Erro
    If lErro = SUCESSO Then gError 18759
    
    'Verifica se a Fatura está cadastrada
    lErro = CF("TituloPagar_Le_Numero", objTituloPagar)
    If lErro <> SUCESSO And lErro <> 18551 Then gError 18760
        
    'Se não encontrou --> Erro
    If lErro <> SUCESSO Then gError 18761

    'Pede a confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_FATURAPAGAR", objTituloPagar.lNumTitulo)

    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui a Fatura do BD (incluindo dados contabeis (contabilidade))
    lErro = CF("FaturaPagar_Exclui", objTituloPagar, objContabil)
    If lErro <> SUCESSO Then gError 18762
    
    'Limpa a tela
    Call Limpa_Tela_FaturasPag
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 18752
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            
        Case 18753
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
        
        Case 18754
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
    
        Case 18755, 18758, 18760, 18762
        
        Case 18759
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURAPAG_BAIXADA_EXCLUSAO", gErr, objTituloPagar.lNumTitulo)
            
        Case 18761
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FATURAPAG_NAO_CADASTRADA2", gErr, objTituloPagar.lNumTitulo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160169)
            
    End Select

    Exit Sub

End Sub

Public Sub BotaoLimpar_Click()
    
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Confirma o pedido de limpeza da tela
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 18751
    
    'Lima a Tela
    Call Limpa_Tela_FaturasPag
       
    Exit Sub

Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 18751
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160170)
    
    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iEmissaoAlterada = 1
    
End Sub

Private Function Carrega_TipoCobranca() As Long
'Carrega na combobox os Tipos de Cobrança

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoCobranca

    'Lê o nome e o codigo de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 18598

    'Carrega na combo de Tipos de Cobrança os Tipos que estão em colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao

        TipoCobranca.AddItem CInt(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        TipoCobranca.ItemData(TipoCobranca.NewIndex) = objCodigoDescricao.iCodigo

    Next

    Carrega_TipoCobranca = SUCESSO

    Exit Function

Erro_Carrega_TipoCobranca:

    Carrega_TipoCobranca = gErr

    Select Case gErr

        Case 18598

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160171)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_CondicaoPagamento() As Long
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCodDescricao As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê as Condicões de Pagamento utilizadas em Contas a Pagar
'    lErro = CF("CondicoesPagto_Le_Pagamento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then gError 18599
'
'    'Carrega na Combo as Condições de Pagamento retornadas em colCod_DescReduzida
'    For Each objCodDescricao In colCod_DescReduzida
'
'        CondicaoPagamento.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
'        CondicaoPagamento.ItemData(CondicaoPagamento.NewIndex) = objCodDescricao.iCodigo
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
'        Case 18599
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160172)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Inicializa_GridParcelas(objGridInt As AdmGrid) As Long
'Faz as Inicializações no Grid de Parcelas a Pagar

    'Indica o Form do Grid
    Set objGridInt.objForm = Me

    'Indica os nomes das colunas
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
    
    'Indica os campos relacionados a cada coluna
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
    
    'Inicializa os valores das colunas
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
    
    'Indica o Grid à que se referem os dados
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do Grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do Grid
    objGridInt.iLinhasVisiveis = 8

    'Determina a largura da coluna 0
    GridParcelas.ColWidth(0) = 900

    'Indica a largura do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama a rotina que faz as demais inicializações
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridParcelas = SUCESSO

    Exit Function

End Function

Public Function FATPAG_Inicializa_GridNF(ByVal objCT As Object, objGridInt As AdmGrid) As Long
    FATPAG_Inicializa_GridNF = Inicializa_GridNF(objGridInt)
End Function

Private Function Inicializa_GridNF(objGridInt As AdmGrid) As Long
'Faz as Inicializações no Grid de Notas Fiscais Fatura

    'Indica o Form do Grid
    Set objGridInt.objForm = Me

    'Indica os nomes das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Filial ")
    objGridInt.colColuna.Add ("Filial Forn")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Valor NF")
    objGridInt.colColuna.Add ("Valor FAT")
    objGridInt.colColuna.Add ("Selecionada")
    
    'Indica os campos relacionados a cada coluna
    objGridInt.colCampo.Add (FilialNF.Name)
    objGridInt.colCampo.Add (FilialFornecedor.Name)
    objGridInt.colCampo.Add (NumNotaFiscal.Name)
    objGridInt.colCampo.Add (DataEmissaoNF.Name)
    objGridInt.colCampo.Add (ValorNFBruto.Name)
    objGridInt.colCampo.Add (ValorNF.Name)
    objGridInt.colCampo.Add (Selecionada.Name)
    
    
    'Inicializa os valores das colunas
    iGrid_FilialNF_Col = 1
    iGrid_FilialFornecedor_Col = 2
    iGrid_Numero_Col = 3
    iGrid_Emissao_Col = 4
    iGrid_ValorNFBruto_Col = 5
    iGrid_ValorNF_Col = 6
    iGrid_Selecionada_Col = 7

    'Indica o Grid ao qual faz referencia
    objGridInt.objGrid = GridNF

    'Linhas visíveis do Grid
    objGridInt.iLinhasVisiveis = NUM_LINHAS_GRID_NF

    'Todas as linhas do Grid
    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1

    'Indica a largura da coluna 0
    GridNF.ColWidth(0) = 300

    'Indica a largura automática do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR

    'Chama a rotina que faz as demais inicializações
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridNF = SUCESSO

    Exit Function

End Function

Public Sub BotaoFechar_Click()
    
    Unload Me

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
   
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoNumero = Nothing
    Set objEventoFornecedor = Nothing
    Set objEventoCondPagto = Nothing
    
    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
    Set objGridParcelas = Nothing
    Set objGridNFiscais = Nothing
    Set objContabil = Nothing
    Set objGrid1 = Nothing

End Sub

Public Sub GridNF_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridNFiscais, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscais, iAlterado)
    End If
    
    Call Ordenacao_ClickGrid(objGridNFiscais)

End Sub

Public Sub GridNF_EnterCell()
    Call Grid_Entrada_Celula(objGridNFiscais, iAlterado)
End Sub

Public Sub GridNF_GotFocus()
    Call Grid_Recebe_Foco(objGridNFiscais)
End Sub

Public Sub GridNF_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridNFiscais, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridNFiscais, iAlterado)
    End If

End Sub

Public Sub GridNF_LeaveCell()
        Call Saida_Celula(objGridNFiscais)
End Sub

Public Sub GridNF_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridNFiscais)
End Sub

Public Sub GridNF_RowColChange()
    Call Grid_RowColChange(objGridNFiscais)
End Sub

Public Sub GridNF_Scroll()
    Call Grid_Scroll(objGridNFiscais)
End Sub

Public Sub GridParcelas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_EnterCell()
    
    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    
End Sub

Public Sub GridParcelas_GotFocus()
    
    Call Grid_Recebe_Foco(objGridParcelas)
    
End Sub

Public Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Public Sub GridParcelas_LeaveCell()
    Call Saida_Celula(objGridParcelas)
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

Private Sub Limpa_Tela_FaturasPag()

Dim lErro As Long

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa os campos comuns nas telas
    Call Limpa_Tela(Me)

    'Limpa os Grids
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridNFiscais)
    Call Ordenacao_Limpa(objGridNFiscais)

    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
    
    'Limpa os campos não limpos em Limpa_Tela
    Filial.Clear
    CondicaoPagamento.Text = ""
    ValorTotalNFSelecionadas.Caption = ""
    NumNFSelecionadas.Caption = ""
        
    sOldFornecedor = ""
    
    Call CF2(Me, "FATPAG_Limpa_Tela_FaturasPag")
    
    iAlterado = 0
    
End Sub

Public Sub DataVencimento_Change()

    iAlterado = REGISTRO_ALTERADO
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
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataVencimentoReal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub DataVencimentoReal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub DataVencimentoReal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimentoReal
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub TipoCobranca_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoCobranca_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub TipoCobranca_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub TipoCobranca_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub TipoCobranca_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = TipoCobranca
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorParcela_Change()
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

Private Sub Carrega_GridNF(colNFPag As ColNFsPag, colNFsPagVinculada As ColNFsPag)
'Coloca no Grid de Notas Fiscais as Notas passadas em colNFPag

Dim iLinha As Integer
Dim iIndice As Integer
Dim dValor As Double
Dim iNumNfsSel As Integer
Dim dValorNFsSel As Double
Dim lErro As Long, objNFsPag As ClassNFsPag

On Error GoTo Erro_Carrega_GridNF
  
    'Recoloca o Numero de Linhas caso seja maior que o Numero de Linhas Visiveis na Tela que é 6
    If GridNF.Rows < colNFPag.Count + colNFsPagVinculada.Count + 1 Then
        GridNF.Rows = colNFPag.Count + colNFsPagVinculada.Count + 1
'    Else
'        GridNF.Rows = NUM_LINHAS_GRID_NF + 1
    End If
       
    'Renicializa
    Call Grid_Inicializa(objGridNFiscais)
    
    iIndice = 0

    'Prenche os GriNF com as Nfs de colNFsPagVinculada
    For iLinha = 1 To colNFsPagVinculada.Count

        iIndice = iIndice + 1

        Set objNFsPag = colNFsPagVinculada(iIndice)

        GridNF.TextMatrix(iLinha, iGrid_FilialNF_Col) = objNFsPag.iFilialEmpresa
        GridNF.TextMatrix(iLinha, iGrid_Numero_Col) = objNFsPag.lNumNotaFiscal
        GridNF.TextMatrix(iLinha, iGrid_Emissao_Col) = IIf(objNFsPag.dtDataEmissao <> DATA_NULA, Format(objNFsPag.dtDataEmissao, "dd/mm/yyyy"), "")
        GridNF.TextMatrix(iLinha, iGrid_FilialFornecedor_Col) = objNFsPag.iFilial
        
        lErro = CF2(Me, "FATPAG_Carrega_GridNF", objNFsPag, iLinha)
        If lErro <> SUCESSO Then gError 202147
        
        'Calcula o Valor da Nota Fiscal
        dValor = objNFsPag.dValorTotal - (objNFsPag.dValorIRRF + objNFsPag.dCOFINSRetido + objNFsPag.dCSLLRetido + objNFsPag.dISSRetido + objNFsPag.dPISRetido + StrParaDbl(IIf(objNFsPag.iINSSRetido = MARCADO, objNFsPag.dValorINSS, 0)))
        GridNF.TextMatrix(iLinha, iGrid_ValorNF_Col) = Format(dValor, "Standard")
        GridNF.TextMatrix(iLinha, iGrid_ValorNFBruto_Col) = Format(objNFsPag.dValorTotal, "Standard")
        If colNFsPagVinculada(iIndice).lNumIntTitPag <> 0 Then
            GridNF.TextMatrix(iLinha, iGrid_Selecionada_Col) = "1"
            iNumNfsSel = iNumNfsSel + 1
            dValorNFsSel = dValorNFsSel + dValor
        End If

    Next

    iIndice = 0

    'Prenche os GriNF com as Nfs de colNFPag
    For iLinha = (colNFsPagVinculada.Count + 1) To colNFPag.Count + colNFsPagVinculada.Count

        iIndice = iIndice + 1

        GridNF.TextMatrix(iLinha, iGrid_FilialNF_Col) = colNFPag(iIndice).iFilialEmpresa
        GridNF.TextMatrix(iLinha, iGrid_Numero_Col) = colNFPag(iIndice).lNumNotaFiscal
        GridNF.TextMatrix(iLinha, iGrid_Emissao_Col) = IIf(colNFPag(iIndice).dtDataEmissao <> DATA_NULA, Format(colNFPag(iIndice).dtDataEmissao, "dd/mm/yyyy"), "")
        GridNF.TextMatrix(iLinha, iGrid_FilialFornecedor_Col) = colNFPag(iIndice).iFilial

        lErro = CF2(Me, "FATPAG_Carrega_GridNF", colNFPag(iIndice), iLinha)
        If lErro <> SUCESSO Then gError 202147

        'Calcula o Valor da Nota Fiscal
'        dValor = colNFPag(iIndice).dValorTotal - colNFPag(iIndice).dValorIRRF
'        GridNF.TextMatrix(iLinha, iGrid_ValorNF_Col) = Format(dValor, "Standard")

        'Calcula o Valor da Nota Fiscal
        Set objNFsPag = colNFPag(iIndice)
        dValor = objNFsPag.dValorTotal - (objNFsPag.dValorIRRF + objNFsPag.dCOFINSRetido + objNFsPag.dCSLLRetido + objNFsPag.dISSRetido + objNFsPag.dPISRetido + StrParaDbl(IIf(objNFsPag.iINSSRetido = MARCADO, objNFsPag.dValorINSS, 0)))
        GridNF.TextMatrix(iLinha, iGrid_ValorNF_Col) = Format(dValor, "Standard")
        GridNF.TextMatrix(iLinha, iGrid_ValorNFBruto_Col) = Format(objNFsPag.dValorTotal, "Standard")

        If colNFPag(iIndice).lNumIntTitPag <> 0 Then
            GridNF.TextMatrix(iLinha, iGrid_Selecionada_Col) = "1"
            iNumNfsSel = iNumNfsSel + 1
            dValorNFsSel = dValorNFsSel + dValor
        End If

    Next
   
    objGridNFiscais.iLinhasExistentes = colNFPag.Count + colNFsPagVinculada.Count

    'Inicializa os contadores e acumuladores dos valores
    'das Notas Fiscais
    NumNFSelecionadas.Caption = CStr(iNumNfsSel)
    ValorTotalNFSelecionadas.Caption = Format(dValorNFsSel, "Standard")
    
    Call CF2(Me, "FATPAG_Calcula_Dif")
    
    'Faz um Refresh nas CheckBox's do Grid de Notas Fiscais
    Call Grid_Refresh_Checkbox(objGridNFiscais)
    
    Exit Sub

Erro_Carrega_GridNF:

    Select Case gErr
    
        Case 202147
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 202148)

    End Select

    Exit Sub
    
 End Sub
 
 Public Function FATPAG_Carrega_GridNF(ByVal objCT As Object, ByVal objNFsPag As ClassNFsPag, ByVal iLinha As Integer) As Long
    FATPAG_Carrega_GridNF = SUCESSO
End Function

Public Sub NumNotaFiscal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumNotaFiscal_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Public Sub NumNotaFiscal_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Public Sub NumNotaFiscal_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub DataEmissaoNF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataEmissaoNF_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Public Sub DataEmissaoNF_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Public Sub DataEmissaoNF_Validate(Cancel As Boolean)

Dim lErro As Long
    
    Set objGridParcelas.objControle = DataEmissaoNF
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Selecionada_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Public Sub Selecionada_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Public Sub Selecionada_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridNFiscais.objControle = Selecionada
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorNF_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ValorNF_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridNFiscais)
End Sub

Public Sub ValorNF_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridNFiscais)
End Sub

Public Sub ValorNF_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridParcelas.objControle = ValorNF
    lErro = Grid_Campo_Libera_Foco(objGridNFiscais)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Function Move_Tela_Memoria(objTituloPagar As ClassTituloPagar, colParcelaPagar As colParcelaPagar, colNFPag As ColNFsPag) As Long
'Recolhe os dados da Tela e passa para objTituloPagar, colParcelaPagar, colNFPag

Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer, sData As String
Dim lErro As Long
Dim lNumNotaFiscal As Long
Dim dtDataEmissao As Date
Dim objNFsPag As ClassNFsPag

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o Fornecedor foi preenchido
    If Len(Trim(Fornecedor.Text)) > 0 Then
        objFornecedor.sNomeReduzido = Fornecedor.Text
        
        lErro = CF("Fornecedor_Le_NomeReduzido", objFornecedor)
        If lErro <> SUCESSO And lErro <> 6681 Then gError 18747
        
        If lErro <> SUCESSO Then gError 26019

        objTituloPagar.lFornecedor = objFornecedor.lCodigo
    Else
        objTituloPagar.lFornecedor = 0
    End If
    
    'Verifica se a Filial foi preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloPagar.iFilial = Codigo_Extrai(Filial.Text)
    Else
        objTituloPagar.iFilial = 0
    End If
    
    'Se estiver preenchido preenche com o NumTitulo da tela
    If Len(Trim(NumTitulo.ClipText)) > 0 Then objTituloPagar.lNumTitulo = CLng(NumTitulo.ClipText)
    
    'Preenche objTituloPagar com Data Emissão. Se não foi digitada preenche com DATA_NULA.
    If Len(Trim(DataEmissao.ClipText)) > 0 Then
        objTituloPagar.dtDataEmissao = CDate(DataEmissao.Text)
    Else
        objTituloPagar.dtDataEmissao = DATA_NULA
    End If
    
    If Len(Trim(ValorTotal.ClipText)) > 0 Then objTituloPagar.dValorTotal = CDbl(ValorTotal.Text)
    If Len(Trim(ValorIRRF.ClipText)) > 0 Then objTituloPagar.dValorIRRF = CDbl(ValorIRRF.Text)
    
    'Carrega os dados restantes da tela
    objTituloPagar.iNumParcelas = objGridParcelas.iLinhasExistentes
    objTituloPagar.sSiglaDocumento = TIPODOC_FATURA_A_PAGAR
    objTituloPagar.iFilialEmpresa = giFilialEmpresa
    objTituloPagar.iStatus = STATUS_LANCADO
    objTituloPagar.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)
    
    'Move os dados do Grid para colParcelaPagar
    lErro = Move_GridParcelas_Memoria(colParcelaPagar)
    If lErro <> SUCESSO Then gError 18748
       
    'Loop de adição das linhas selecionadas de GridNF à colNFPag
    For iIndice = 1 To objGridNFiscais.iLinhasExistentes
    
        'Se Selecionada, adiciona Nota Fiscal a colNFPag
        If GridNF.TextMatrix(iIndice, iGrid_Selecionada_Col) = "1" Then
            lNumNotaFiscal = CLng(GridNF.TextMatrix(iIndice, iGrid_Numero_Col))
            sData = GridNF.TextMatrix(iIndice, iGrid_Emissao_Col)
            If Len(Trim(sData)) > 0 Then
                dtDataEmissao = CDate(sData)
            Else
                dtDataEmissao = DATA_NULA
            End If
            
            Set objNFsPag = colNFPag.Add(0, StrParaInt(GridNF.TextMatrix(iIndice, iGrid_FilialNF_Col)), objTituloPagar.lFornecedor, StrParaInt(GridNF.TextMatrix(iIndice, iGrid_FilialFornecedor_Col)), lNumNotaFiscal, dtDataEmissao, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        
            lErro = CF2(Me, "FATPAG_Move_Tela_Memoria", objNFsPag, iIndice)
            If lErro <> SUCESSO Then gError 18748
        
        End If
    Next

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 18747, 18748
        
        Case 26019
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objFornecedor.sNomeReduzido)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 160173)
        
    End Select
        
    Exit Function

End Function

Public Function FATPAG_Move_Tela_Memoria(ByVal objCT As Object, ByVal objNFsPag As ClassNFsPag, ByVal iLinha As Integer) As Long
    FATPAG_Move_Tela_Memoria = SUCESSO
End Function

Private Function Move_GridParcelas_Memoria(colParcelas As colParcelaPagar)
'Move para colParcelas os dados existentes no GridParcelas

Dim iIndice As Integer
Dim objParcelaPag As ClassParcelaPagar
 
    'Loop de preenchimento de colParcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        Set objParcelaPag = New ClassParcelaPagar
        
        'Preenchimento de objParcelaPag com linha do GridParcelas
        objParcelaPag.iNumParcela = iIndice
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) > 0 Then objParcelaPag.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))) > 0 Then objParcelaPag.dtDataVencimentoReal = CDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))) > 0 Then objParcelaPag.dValor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorParcela_Col))
        
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))) = 0 Then
            objParcelaPag.iTipoCobranca = TIPO_COBRANCA_CARTEIRA
        Else
            objParcelaPag.iTipoCobranca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_Cobranca_Col))
        End If
        
        objParcelaPag.dValorOriginal = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_ValorOriginal_Col))
        objParcelaPag.iMotivoDiferenca = Codigo_Extrai(GridParcelas.TextMatrix(iIndice, iGrid_MotivoDiferenca_Col))
        
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
        
        'Adição de objParcelaPag a colParcelas
        With objParcelaPag
            Call colParcelas.Add(0, 0, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, 0, .dValor, .iPortador, 1, .iTipoCobranca, .iBancoCobrador, .sNossoNumero, .sCodigoDeBarras, "", .dValorOriginal, .iMotivoDiferenca)
        End With
    Next
        
    Move_GridParcelas_Memoria = SUCESSO
    
    Exit Function
                
End Function

Private Function TipoFornecedor_Dados(objFornecedor As ClassFornecedor, objTipoFornecedor As ClassTipoFornecedor) As Long
'Lê os dados de TipoFornecedor ligado a objFornecedor.
    
Dim lErro As Long

On Error GoTo Erro_TipoFornecedor_Dados

    objTipoFornecedor.iCodigo = objFornecedor.iTipo
    
    'Lê o TipoFornecedor a partir do código
    lErro = CF("TipoFornecedor_Le", objTipoFornecedor)
    If lErro <> SUCESSO And lErro <> 12765 Then gError 26031
    If lErro <> SUCESSO Then gError 26030
   
    TipoFornecedor_Dados = SUCESSO

    Exit Function

Erro_TipoFornecedor_Dados:

    TipoFornecedor_Dados = gErr

    Select Case gErr

        Case 26030
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_FORNECEDOR_NAO_CADASTRADO", gErr, objTipoFornecedor.iCodigo)
        
        Case 26031

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160174)

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
        If lErro <> SUCESSO And lErro <> 19205 Then gError 18657
        If lErro = 19205 Then gError 18659
    
        'Preenche o GridParcelas
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then gError 18658

    End If
    
    Exit Sub
     
Erro_Recalcula_Cobranca:

    Select Case gErr
          
        Case 18657, 18658

        Case 18659
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", gErr, objCondicaoPagto.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160175)
     
    End Select
     
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_FATURAS_PAGAR_ID
    Set Form_Load_Ocx = Me
    Caption = "Faturas a Pagar"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "FaturasPag"
    
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
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is CondicaoPagamento Then
            Call CondPagtoLabel_Click
        End If
    
    End If
    
End Sub

Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Public Sub Fornecedor_Preenche()
'por Jorge Specian - Para localizar pela parte digitada do Nome
'Reduzido do Fornecedor através da CF Fornecedor_Pesquisa_NomeReduzido em RotinasCPR.ClassCPRSelect'

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objFornecedor As Object
    
On Error GoTo Erro_Fornecedor_Preenche
    
    Set objFornecedor = Fornecedor
    
    lErro = CF("Fornecedor_Pesquisa_NomeReduzido", objFornecedor, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134046

    Exit Sub

Erro_Fornecedor_Preenche:

    Select Case gErr

        Case 134046

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 160176)

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
    
    'If iValorIRRFAlterado <> 0 Then
    
        Call Recalcula_Cobranca
        iValorIRRFAlterado = 0
        
    'End If
    
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
Dim objTituloPagar As New ClassTituloPagar, colParcelaPagar As New colParcelaPagar, colNFPag  As New ColNFsPag

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

        Case VALOR_IR
            If Len(Trim(ValorIRRF.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(ValorIRRF.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
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
            
            lErro = Move_Tela_Memoria(objTituloPagar, colParcelaPagar, colNFPag)
            If lErro <> SUCESSO Then gError 39567
        
            lErro = CF("FP_Calcula_Mnemonico_Cust", objMnemonicoValor, objTituloPagar, colParcelaPagar, colNFPag)
            If lErro <> SUCESSO Then gError 36229
            
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

Public Function FP_Calcula_Mnemonico_Cust(objMnemonicoValor As ClassMnemonicoValor, objTitPag As ClassTituloPagar, colParcelaPagar As colParcelaPagar, colNFPag As ColNFsPag) As Long
    FP_Calcula_Mnemonico_Cust = SUCESSO
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

Public Function FATPAG_Limpa_Tela_FaturasPag(ByVal objCT) As Long
    FATPAG_Limpa_Tela_FaturasPag = SUCESSO
End Function

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

Public Sub comboCobrador_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub comboPortador_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub comboPortador_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub comboCobrador_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub comboCobrador_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub comboPortador_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridParcelas)
End Sub

Public Sub comboCobrador_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
End Sub

Public Sub comboPortador_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)
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

Private Function Saida_Celula_ValorOriginal(objGridInt As AdmGrid) As Long
'Faz a crítica da celula ValorOriginal do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim dColunaSoma As Double

On Error GoTo Erro_Saida_Celula_ValorOriginal

    Set objGridInt.objControle = ValorOriginal

    'Verifica se valor original está preenchido
    If Len(Trim(ValorOriginal.Text)) > 0 Then
    
        'Critica se valor original é positivo
        lErro = Valor_Positivo_Critica(ValorOriginal.Text)
        If lErro <> SUCESSO Then gError 500052

        ValorOriginal.Text = Format(ValorOriginal.Text, "Standard")
        If ValorOriginal.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_ValorOriginal_Col) Then CobrancaAutomatica.Value = vbUnchecked
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 500053

    Saida_Celula_ValorOriginal = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorOriginal:

    Saida_Celula_ValorOriginal = gErr

    Select Case gErr

        Case 500052, 500053
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
            If lErro <> SUCESSO And lErro <> 12250 Then gError 500054
    
            'Se não encontrou o Motivo de Diferença nem com o Código nem Descrição ==> erro
            If lErro <> SUCESSO Then gError 500056
            
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 500055

    Saida_Celula_MotivoDiferenca = SUCESSO

    Exit Function

Erro_Saida_Celula_MotivoDiferenca:

    Saida_Celula_MotivoDiferenca = gErr

    Select Case gErr

        Case 500054, 500055
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 500056
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MOTIVODIFERENCA_NAO_ENCONTRADO", gErr, MotivoDiferenca.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)

    End Select

    Exit Function

End Function

Public Sub Diferenca_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Diferenca_Validate
    
    'Se diferença foi preenchida
    If Len(Trim(Diferenca.ClipText)) > 0 Then
    
        'Critica seu valor
        lErro = Valor_Critica(Diferenca.Text)
        If lErro <> SUCESSO Then gError 500062
    
    End If
    
    Call Recalcula_Cobranca
    
    Exit Sub
    
Erro_Diferenca_Validate:

    Cancel = True
    
    Select Case gErr
        
        Case 500062
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$)
        
    End Select
    
    Exit Sub

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
    If lErro <> SUCESSO Then gError 500061

   'Preenche ComboBox com código e descrição dos Bancos
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

        Case 500061 'Tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function

End Function

