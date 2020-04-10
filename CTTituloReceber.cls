VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTTituloReceber"
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

Dim gobjTelaProjetoInfo As ClassTelaPRJInfo

'Início da contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1
Private WithEvents objEventoNatureza As AdmEvento
Attribute objEventoNatureza.VB_VarHelpID = -1

Private Const CLIENTE_COD As String = "Cliente_Codigo"
Private Const CLIENTE_NOME As String = "Cliente_Nome"
Private Const FILIAL_COD As String = "FilialCli_Codigo"
Private Const FILIAL_NOME_RED As String = "FilialCli_Nome"
Private Const FILIAL_CONTA As String = "FilialCli_Conta_Ctb"
Private Const FILIAL_CGC_CPF As String = "FilialCli_CGC_CPF"
Private Const EMISSAO1 As String = "Data_Emissao"
Private Const VALOR1 As String = "Valor_Total"
Private Const TIPO1 As String = "Tipo_Documento"
Private Const VALOR_IR As String = "Valor_IRRF"
Private Const VALOR_INSS As String = "Valor_INSS"
Private Const INSS_RETIDO As String = "INSS_Retido"
Private Const NUMERO1 As String = "Numero_Titulo"
Private Const TIPOCLI_CONTA As String = "TipoCli_Conta_Ctb"
Private Const PIS_RETIDO As String = "PIS_Retido"
Private Const ISS_RETIDO As String = "ISS_Retido"
Private Const COFINS_RETIDO As String = "COFINS_Retido"
Private Const CSLL_RETIDO As String = "CSLL_Retido"

'Vinculados a Parcelas
Private Const CHEQUEPRE_BANCO As String = "ChequePre_Banco"
Private Const CHEQUEPRE_AGENCIA As String = "ChequePre_Agencia"
Private Const CHEQUEPRE_CONTA As String = "ChequePre_Conta"
Private Const CHEQUEPRE_NUMERO As String = "ChequePre_Numero"
Private Const CHEQUEPRE_VALOR As String = "ChequePre_Valor"

'Fim da contabilidade

Public iAlterado As Integer
Dim iClienteAlterado As Integer
Dim iFrameAtual As Integer
Dim giTrazendoTitulo As Integer

Dim objGridParcelas As AdmGrid
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Vencimento_Col As Integer
Dim iGrid_VenctoReal_Col As Integer
Dim iGrid_Valor_Col As Integer

'Inserido por Wagner
'#################
Dim iGrid_Previsao_Col As Integer
Dim iGrid_DescPrev_Col As Integer
'#################

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

Private colParcelaReceber As colParcelaReceber
Private colcolComissao As colcolComissao
Private colcolDesconto As colcolDesconto
Private colChequePre As colChequePre


'Constantes que verificam se os Valores dos campos foram alterados
Private iValorAlterado As Integer
Private iDataEmissaoAlterada As Integer
Private iValorINSSAlterado As Integer
Private iValorIRRFAlterado As Integer
Private iPISRetidoAlterado As Integer
Private iISSRetidoAlterado As Integer
Private iCOFINSRetidoAlterado As Integer
Private iCSLLRetidoAlterado As Integer

Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCondPagto As AdmEvento
Attribute objEventoCondPagto.VB_VarHelpID = -1
Private WithEvents objEventoVendedores As AdmEvento
Attribute objEventoVendedores.VB_VarHelpID = -1
Private WithEvents objEventoVendedoresParc As AdmEvento
Attribute objEventoVendedoresParc.VB_VarHelpID = -1
Private WithEvents objEventoTipoDocumento As AdmEvento
Attribute objEventoTipoDocumento.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Parcelas = 2
Private Const TAB_Contabilizacao = 3

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Frame6() As Object
     Set Frame6 = objUserControl.Controls("Frame6")
End Property

Public Property Get ReajustePeriodicidade() As Object
     Set ReajustePeriodicidade = objUserControl.Controls("ReajustePeriodicidade")
End Property

Public Property Get Moeda() As Object
     Set Moeda = objUserControl.Controls("Moeda")
End Property

Public Property Get UpDownReajusteBase() As Object
     Set UpDownReajusteBase = objUserControl.Controls("UpDownReajusteBase")
End Property

Public Property Get ReajusteBase() As Object
     Set ReajusteBase = objUserControl.Controls("ReajusteBase")
End Property

Public Property Get LabelMoeda() As Object
     Set LabelMoeda = objUserControl.Controls("LabelMoeda")
End Property

Public Property Get Frame5() As Object
     Set Frame5 = objUserControl.Controls("Frame5")
End Property

Public Property Get Frame2() As Object
     Set Frame2 = objUserControl.Controls("Frame2")
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

Public Property Get SSFrame6() As Object
     Set SSFrame6 = objUserControl.Controls("SSFrame6")
End Property

Public Property Get INSSRetido() As Object
     Set INSSRetido = objUserControl.Controls("INSSRetido")
End Property

Public Property Get ValorINSS() As Object
     Set ValorINSS = objUserControl.Controls("ValorINSS")
End Property

Public Property Get SSFrame2() As Object
     Set SSFrame2 = objUserControl.Controls("SSFrame2")
End Property

Public Property Get BotaoProjetos() As Object
     Set BotaoProjetos = objUserControl.Controls("BotaoProjetos")
End Property

Public Property Get Etapa() As Object
     Set Etapa = objUserControl.Controls("Etapa")
End Property

Public Property Get BotaoLimparFAT() As Object
     Set BotaoLimparFAT = objUserControl.Controls("BotaoLimparFAT")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get Filial() As Object
     Set Filial = objUserControl.Controls("Filial")
End Property

Public Property Get NumTitulo() As Object
     Set NumTitulo = objUserControl.Controls("NumTitulo")
End Property

Public Property Get Cliente() As Object
     Set Cliente = objUserControl.Controls("Cliente")
End Property

Public Property Get UpDownEmissao() As Object
     Set UpDownEmissao = objUserControl.Controls("UpDownEmissao")
End Property

Public Property Get DataEmissao() As Object
     Set DataEmissao = objUserControl.Controls("DataEmissao")
End Property

Public Property Get Valor() As Object
     Set Valor = objUserControl.Controls("Valor")
End Property

Public Property Get Ccl() As Object
     Set Ccl = objUserControl.Controls("Ccl")
End Property

Public Property Get Natureza() As Object
     Set Natureza = objUserControl.Controls("Natureza")
End Property

Public Property Get Projeto() As Object
     Set Projeto = objUserControl.Controls("Projeto")
End Property

Public Property Get LabelProjeto() As Object
     Set LabelProjeto = objUserControl.Controls("LabelProjeto")
End Property

Public Property Get LabelNaturezaDesc() As Object
     Set LabelNaturezaDesc = objUserControl.Controls("LabelNaturezaDesc")
End Property

Public Property Get LabelNatureza() As Object
     Set LabelNatureza = objUserControl.Controls("LabelNatureza")
End Property

Public Property Get CclLabel() As Object
     Set CclLabel = objUserControl.Controls("CclLabel")
End Property

Public Property Get NumeroFAT() As Object
     Set NumeroFAT = objUserControl.Controls("NumeroFAT")
End Property

Public Property Get TipoDocumentoLabel() As Object
     Set TipoDocumentoLabel = objUserControl.Controls("TipoDocumentoLabel")
End Property

Public Property Get LabelFilial() As Object
     Set LabelFilial = objUserControl.Controls("LabelFilial")
End Property

Public Property Get NumeroLabel() As Object
     Set NumeroLabel = objUserControl.Controls("NumeroLabel")
End Property

Public Property Get ClienteEtiqueta() As Object
     Set ClienteEtiqueta = objUserControl.Controls("ClienteEtiqueta")
End Property

Public Property Get SSFrame7() As Object
     Set SSFrame7 = objUserControl.Controls("SSFrame7")
End Property

Public Property Get BotaoVendedores() As Object
     Set BotaoVendedores = objUserControl.Controls("BotaoVendedores")
End Property

Public Property Get VendedorEmissao() As Object
     Set VendedorEmissao = objUserControl.Controls("VendedorEmissao")
End Property

Public Property Get PercentualEmissao() As Object
     Set PercentualEmissao = objUserControl.Controls("PercentualEmissao")
End Property

Public Property Get ValorBaseEmissao() As Object
     Set ValorBaseEmissao = objUserControl.Controls("ValorBaseEmissao")
End Property

Public Property Get ValorEmissao() As Object
     Set ValorEmissao = objUserControl.Controls("ValorEmissao")
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

Public Property Get BotaoVendedoresParc() As Object
     Set BotaoVendedoresParc = objUserControl.Controls("BotaoVendedoresParc")
End Property

Public Property Get SSFrame3() As Object
     Set SSFrame3 = objUserControl.Controls("SSFrame3")
End Property

Public Property Get CondicaoPagamento() As Object
     Set CondicaoPagamento = objUserControl.Controls("CondicaoPagamento")
End Property

Public Property Get CobrancaAutomatica() As Object
     Set CobrancaAutomatica = objUserControl.Controls("CobrancaAutomatica")
End Property

Public Property Get Previsao() As Object
     Set Previsao = objUserControl.Controls("Previsao")
End Property

Public Property Get DescPrev() As Object
     Set DescPrev = objUserControl.Controls("DescPrev")
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

Public Property Get BotaoEditarDif() As Object
     Set BotaoEditarDif = objUserControl.Controls("BotaoEditarDif")
End Property

Public Property Get BotaoDif() As Object
     Set BotaoDif = objUserControl.Controls("BotaoDif")
End Property

Public Property Get SSFrame4() As Object
     Set SSFrame4 = objUserControl.Controls("SSFrame4")
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

Public Property Get TotalPercentualComissao() As Object
     Set TotalPercentualComissao = objUserControl.Controls("TotalPercentualComissao")
End Property

Public Property Get TotalValorComissao() As Object
     Set TotalValorComissao = objUserControl.Controls("TotalValorComissao")
End Property

Public Property Get LabelTotaisComissoes() As Object
     Set LabelTotaisComissoes = objUserControl.Controls("LabelTotaisComissoes")
End Property

Public Property Get SSFrame1() As Object
     Set SSFrame1 = objUserControl.Controls("SSFrame1")
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

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get BotaoLimpar() As Object
     Set BotaoLimpar = objUserControl.Controls("BotaoLimpar")
End Property

Public Property Get BotaoExcluir() As Object
     Set BotaoExcluir = objUserControl.Controls("BotaoExcluir")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get Opcao() As Object
     Set Opcao = objUserControl.Controls("Opcao")
End Property

'--- fim dos properties get dos controles da tela

Private Function Atualiza_DatasParcelas(objCondicaoPagto As ClassCondicaoPagto) As Long
'Atualiza as datas de vencto das Parcelas

Dim lErro As Long, iIndice As Integer
Dim dtDataEmissao As Date
Dim dtDataVenctoReal As Date

On Error GoTo Erro_Atualiza_DatasParcelas

    'Se Data Emissão estiver preenchida
    If Len(Trim(DataEmissao.ClipText)) > 0 And IsDate(DataEmissao.Text) Then

        dtDataEmissao = CDate(DataEmissao.Text)
        
        objCondicaoPagto.dtDataEmissao = dtDataEmissao
        
        'Calcula Datas de Vencimento das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, False, True)
        If lErro <> SUCESSO Then Error 26155

        'Loop de preenchimento do Grid Parcelas com Datas de Vencimento
        For iIndice = 1 To objCondicaoPagto.iNumeroParcelas

            'Coloca Data de Vencimento no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dtVencimento, "dd/mm/yyyy")

            'Calcula Data Vencimento Real
            lErro = CF("DataVencto_Real", objCondicaoPagto.colParcelas(iIndice).dtVencimento, dtDataVenctoReal)
            If lErro <> SUCESSO Then Error 26156

            'Coloca Data de Vencimento Real no Grid Parcelas
            GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        Next

    End If

    Atualiza_DatasParcelas = SUCESSO
    
    Exit Function
    
Erro_Atualiza_DatasParcelas:

    Atualiza_DatasParcelas = Err
    
    Select Case Err
    
        Case 26155, 26156
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175168)
    
    End Select
    
    Exit Function
    
End Function

Private Function Atualiza_ValorParcelas(objCondicaoPagto As ClassCondicaoPagto) As Long
'Atualiza os valores das Parcelas em funcao dos valores da tela

Dim lErro As Long
Dim dValorINSSRetido As Double
Dim dValorIRRF As Double, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double, dISSRetido As Double
Dim dValorReceber As Double
Dim iIndice As Integer
Dim iNumParcelas As Integer

On Error GoTo Erro_Atualiza_ValorParcelas

    iNumParcelas = objGridParcelas.iLinhasExistentes
    
    If IsNumeric(Valor) Then
    
        If IsNumeric(ValorINSS.Text) Then dValorINSSRetido = StrParaDbl(IIf(INSSRetido.Value, ValorINSS.Text, 0))
        If Len(Trim(ValorIRRF.Text)) > 0 And IsNumeric(ValorIRRF.Text) Then dValorIRRF = CDbl(ValorIRRF)
        If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
        If Len(Trim(ISSRetido.Text)) <> 0 And IsNumeric(ISSRetido.Text) Then dISSRetido = CDbl(ISSRetido.Text)
        If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
        If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
        
        dValorReceber = CDbl(Valor) - (dValorIRRF + dPISRetido + dISSRetido + dCOFINSRetido + dCSLLRetido) - dValorINSSRetido

    End If
    
    'Se Valor a Receber for positivo
    If dValorReceber > 0 Then

        objCondicaoPagto.dValorTotal = dValorReceber
        
        'Calcula os valores das Parcelas
        lErro = CF("CondicaoPagto_CalculaParcelas", objCondicaoPagto, True, False)
        If lErro <> SUCESSO Then Error 26154

        'Coloca os valores das Parcelas no Grid Parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col) = Format(objCondicaoPagto.colParcelas(iIndice).dValor, "Standard")
        Next

    End If

    Atualiza_ValorParcelas = SUCESSO

    Exit Function
    
Erro_Atualiza_ValorParcelas:

    Atualiza_ValorParcelas = Err
    
    Select Case Err
    
        Case 26154
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175169)
    
    End Select
    
    Exit Function
    
End Function

Private Function Inicializa_ComissaoParcela(iParcela As Integer, colComissao As colComissao, objVendedor As ClassVendedor, objComissaoNF As ClassComissaoNF) As Long
'Inicializa as Comissões das Parcelas

Dim lErro As Long
Dim objComissao As New ClassComissao

On Error GoTo Erro_Inicializa_ComissaoParcela

    lErro = CF("ComissaoAutomatica_Obter_Parcela", objVendedor, StrParaDbl(GridParcelas.TextMatrix(iParcela, iGrid_Valor_Col)), objComissaoNF, objComissao)
    If lErro <> SUCESSO Then Error 59015
    
    'Se vendedor e percentual não forem nulos adiciona a Comissão na coleção
    If objComissaoNF.iCodVendedor <> 0 And objComissao.dPercentual <> 0 Then
    
        colComissao.Add 0, giFilialEmpresa, 0, 0, 0, objComissao.iCodVendedor, DATA_NULA, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor, DATA_NULA
        
    End If
            
    Inicializa_ComissaoParcela = SUCESSO
    
    Exit Function
    
Erro_Inicializa_ComissaoParcela:

    Inicializa_ComissaoParcela = Err
            
    Select Case Err
    
        Case 59015
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175170)
            
    End Select
    
    Exit Function
                    
End Function

Private Sub Agencia_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Banco_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Banco_GotFocus()
    
    'Call MaskEdBox_TrataGotFocus(Banco, iAlterado)

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
    If lErro <> SUCESSO Then gError 79306

    objCliente.sNomeReduzido = Cliente.Text

    'Lê o codigo do Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objCliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 79307

    'Não achou o Cliente --> erro
    If lErro <> SUCESSO Then gError 79308

    objTituloReceber.lCliente = objCliente.lCodigo
    objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    objTituloReceber.lNumTitulo = CLng(NumTitulo.Text)
    objTituloReceber.iFilialEmpresa = giFilialEmpresa

    'Procura o Titulo (Baixados ou não)
    lErro = CF("TitulosRec_Le_Todos", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 58997 Then gError 79309

    If lErro = 58997 Then gError 79310

    'Se for do Tipo "OR" dá mensagem pois este Título não tem documento original
    If objTituloReceber.sSiglaDocumento = OUTROS_RECEBIMENTOS Then gError 79311

    'Se for do Tipo Fatura a Receber chama a Tela de Browser Listando todas as Notas da Fatura a Receber que originou este Título
    If objTituloReceber.sSiglaDocumento = TIPODOC_FATURA_A_RECEBER Then

        colSelecao.Add objTituloReceber.lNumIntDoc

        Call Chama_Tela("NF_FaturaReceberLista", colSelecao)

    Else

        'Le o Nome da Tela que originou este Título
        lErro = CF("Titulo_Le_DocumentoOriginal", objTituloReceber.lNumIntDoc, CPR_TITULO_RECEBER, objNFiscal, sTela)
        If lErro <> SUCESSO And lErro <> 58942 Then gError 79312

        If lErro = SUCESSO Then
'            If giTipoVersao = VERSAO_FULL Then
                'Chama a Tela
                Call Chama_Tela(sTela, objNFiscal)
'            ElseIf giTipoVersao = VERSAO_LIGHT Then
'                'Chama a Tela
'                Call Chama_Tela("NFiscalFatura", objNFiscal)
'            End If
        Else
            gError 79313
        End If

    End If

    Exit Sub

Erro_BotaoDocOriginal_Click:

    Select Case gErr

        Case 79306, 79307, 79309, 79312

        Case 79308
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, objCliente.sNomeReduzido)

        Case 79310
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", gErr, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)

        Case 79311, 79313
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_REC_SEM_DOCORIGINAL", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175171)

    End Select

    Exit Sub

End Sub
        
Private Function Critica_CamposPreenchidos() As Long

Dim lErro As Long

On Error GoTo Erro_Critica_CamposPreenchidos

    'Se Cliente estiver vazio, erro
    If Len(Trim(Cliente.Text)) = 0 Then gError 79302

    'Se Filial estiver vazia, erro
    If Len(Trim(Filial.Text)) = 0 Then gError 79303

    'Verifica se o Tipo foi preenchido
    If Len(Trim(Tipo.Text)) = 0 Then gError 79304

    'Verifica se o Titulo está Preenchido
    If Len(Trim(NumTitulo.Text)) = 0 Then gError 79305

    Critica_CamposPreenchidos = SUCESSO

    Exit Function

Erro_Critica_CamposPreenchidos:

    Critica_CamposPreenchidos = gErr

    Select Case gErr

        Case 79302
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 79303
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)

        Case 79304
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)

        Case 79305
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175172)

    End Select

    Exit Function

End Function

Public Sub BotaoEditarDif_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber, objParcelaReceber As New ClassParcelaReceber
Dim iParcela As Integer
Dim objParcelaRecDif As New ClassParcelasRecDif
Dim colComissaoEmissao As New colComissao

On Error GoTo Erro_BotaoEditarDif_Click

    iParcela = GridParcelas.Row
    
    If iParcela <> 0 Then
    
        'Move os dados da Tela para objtituloreceber e colParcelas
        lErro = Move_Tela_Memoria(objTituloReceber, colComissaoEmissao)
        If lErro <> SUCESSO Then gError 177922
            
        'Verifica se os obj(s) estão preenchidos
        If objTituloReceber.iFilialEmpresa <> 0 And objTituloReceber.dtDataEmissao <> DATA_NULA And Len(Trim(objTituloReceber.sSiglaDocumento)) <> 0 And objTituloReceber.lNumTitulo <> 0 Then
        
            objParcelaReceber.iNumParcela = iParcela
            
            lErro = CF("ParcelaReceber_Le_NumTitulo", objTituloReceber, objParcelaReceber)
            If lErro <> SUCESSO And lErro <> 46477 Then gError 130516
            
        End If

    End If
    
    If objParcelaReceber.lNumIntDoc <> 0 Then
        objParcelaRecDif.lNumIntParc = objParcelaReceber.lNumIntDoc
        Call Chama_Tela("ParcelasRecDif", objParcelaRecDif)
    Else
        Call Chama_Tela("ParcelasRecDif")
    End If
    
    Exit Sub

Erro_BotaoEditarDif_Click:

    Select Case Err

        Case 130516, 177922

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 177923)

    End Select

    Exit Sub

End Sub

Public Sub BotaoLimparFAT_Click()

    NumeroFAT.Caption = ""
    
End Sub

Public Sub CobrancaAutomatica_Click()

    iAlterado = REGISTRO_ALTERADO
    
    Call Recalcula_Parcelas_Comissoes

End Sub

Private Sub ContaCorrente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataDeposito_Change()
    
    iAlterado = REGISTRO_ALTERADO
    
End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)
    
End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1
    
End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim sMascaraCcl As String

On Error GoTo Erro_Form_Load
    
    If giTipoVersao = VERSAO_LIGHT Then
        
        Opcao.Tabs.Remove (TAB_Contabilizacao)
    
    End If
    
'    If giTipoVersao = VERSAO_LIGHT Then
'
'        LabelFilial.Visible = False
'        Filial.Visible = False
'
'    End If
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Carrega os Tipos de Documento
    lErro = Carrega_TipoDocumento()
    If lErro <> SUCESSO Then Error 26032

    'Carrega as Condições de Pagamento
    lErro = CF("Carrega_CondicaoPagamento", CondicaoPagamento, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 26038

    'Carrega os Tipos de Desconto
    lErro = Carrega_TipoDesconto()
    If lErro <> SUCESSO Then Error 26040

    'Carrega a Lista de Vendedores
'    lErro = Carrega_Vendedores()
'    If lErro <> SUCESSO Then Error 26042

    'carrega a combo de Moedas
    lErro = Carrega_Moeda()
    If lErro <> SUCESSO Then Error 26042
    
    'Inicializa Máscara de Ccl
    sMascaraCcl = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCcl)
    If lErro <> SUCESSO Then Error 26042

    Ccl.Mask = sMascaraCcl
    
    ReajustePeriodicidade.ListIndex = 0
        
    'Inicializa o frame atual
    iFrameAtual = 1

    'Inicializa os Eventos da Tela
    Set objEventoCliente = New AdmEvento
    Set objEventoNumero = New AdmEvento
    Set objEventoCondPagto = New AdmEvento
    Set objEventoVendedores = New AdmEvento
    Set objEventoVendedoresParc = New AdmEvento
    Set objEventoTipoDocumento = New AdmEvento
    Set objEventoCcl = New AdmEvento
    Set objEventoNatureza = New AdmEvento
    
    'Inicializa os Grids da tela
    Set objGridParcelas = New AdmGrid
    Set objGridDesconto = New AdmGrid
    Set objGridComissoes = New AdmGrid
    Set objGridComissoesEmissao = New AdmGrid
        
    'Inicializa as coleções globais da tela
    Set colParcelaReceber = New colParcelaReceber
    Set colcolComissao = New colcolComissao

    Set colChequePre = New colChequePre

    Set colcolDesconto = New colcolDesconto

    'Faz as inicializações particulares ao GridParcelas
    lErro = Inicializa_Grid_Parcelas(objGridParcelas)
    If lErro <> SUCESSO Then Error 26048

    'Faz as inicializações particulares ao GridDescontos
    lErro = Inicializa_Grid_Descontos(objGridDesconto)
    If lErro <> SUCESSO Then Error 26049

    'Faz as inicializações particulares ao GridComissoes
    lErro = Inicializa_Grid_Comissoes(objGridComissoes)
    If lErro <> SUCESSO Then Error 26050

    'Faz as inicializações particulares ao GridComissoesEmissao
    lErro = Inicializa_Grid_ComissoesEmissao(objGridComissoesEmissao)
    If lErro <> SUCESSO Then Error 26051
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
    
    'Inicialização da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_CONTASARECEBER)
    If lErro <> SUCESSO Then Error 39580
    
    'Inicializa a mascara de Natureza
    lErro = Inicializa_Mascara_Natureza()
    If lErro <> SUCESSO Then Error 32296
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 26032, 26038, 26040, 26042, 26048, 26049, 26050, 26051, 39580, 32296

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175173)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim colComissaoEmissao As New colComissao
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TitulosRec"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objTituloReceber, colComissaoEmissao)
    If lErro <> SUCESSO Then Error 26375

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
    colCampoValor.Add "SiglaDocumento", objTituloReceber.sSiglaDocumento, STRING_SIGLA_DOCUMENTO, "SiglaDocumento"
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
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
        
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 26375

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175174)

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
        If lErro <> SUCESSO Then Error 26376

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 26376

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175175)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoDocumento()
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim iTipo As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento
Dim iFatATivo As Integer

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then Error 26033
    
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

        Case 26033

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175176)

    End Select

    Exit Function

End Function
'
'Private Function Carrega_CondicaoPagamento() As Long
''Carrega as Condições de Pagamento
'
'Dim lErro As Long
'Dim colCod_DescReduzida As New AdmColCodigoNome
'Dim objCod_DescReduzida As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_CondicaoPagamento
'
'    'Lê o código e a descrição reduzida de todas as Condições de Pagamento
'    lErro = CF("CondicoesPagto_Le_Recebimento", colCod_DescReduzida)
'    If lErro <> SUCESSO Then Error 26039
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
'    Carrega_CondicaoPagamento = Err
'
'    Select Case Err
'
'        Case 26039
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175177)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Carrega_TipoDesconto() As Long
'Carrega os Tipos de Desconto

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome

On Error GoTo Erro_Carrega_TipoDesconto

    'Lê o código e a descrição de todos os Tipos de Desconto
    lErro = CF("Cod_Nomes_Le", "TiposDeDesconto", "Codigo", "DescReduzida", STRING_TIPOSDEDESCONTO_DESCRICAO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 26041

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

        Case 26041

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175178)

    End Select

    Exit Function

End Function

'Private Function Carrega_Vendedores() As Long
''Carrega os Vendedores
'
'Dim lErro As Long
'Dim colCodigoNome As New AdmColCodigoNome
'Dim objCodNome As AdmCodigoNome
'
'On Error GoTo Erro_Carrega_Vendedores
'
'    'Lê o Código e o Nome Reduzido dos Vendedores
'    lErro = CF("Cod_Nomes_Le", "Vendedores", "Codigo", "NomeReduzido", STRING_VENDEDOR_NOME_REDUZIDO, colCodigoNome)
'    If lErro <> SUCESSO Then Error 26043
'
'    For Each objCodNome In colCodigoNome
'
'        'Adiciona o ítem na List de Vendedores
'        Vendedores.AddItem objCodNome.iCodigo & SEPARADOR & objCodNome.sNome
'        Vendedores.ItemData(Vendedores.NewIndex) = objCodNome.iCodigo
'
'    Next
'
'    Carrega_Vendedores = SUCESSO
'
'    Exit Function
'
'Erro_Carrega_Vendedores:
'
'    Carrega_Vendedores = Err
'
'    Select Case Err
'
'        Case 26043
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175179)
'
'    End Select
'
'    Exit Function
'
'End Function

Private Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Parcelas

    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Vencto Real")
    objGridInt.colColuna.Add ("Valor")
    
    'Inserido por Wagner
    '####################
    objGridInt.colColuna.Add ("Previsão")
    objGridInt.colColuna.Add ("Observação")
    '##########################

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (DataVencimentoReal.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
    
    'Inserido por Wagner
    '####################
    objGridInt.colCampo.Add (Previsao.Name)
    objGridInt.colCampo.Add (DescPrev.Name)
    '##########################

    'Colunas do Grid
    iGrid_Parcela_Col = 0
    iGrid_Vencimento_Col = 1
    iGrid_VenctoReal_Col = 2
    iGrid_Valor_Col = 3

    'Inserido por Wagner
    '####################
    iGrid_Previsao_Col = 4
    iGrid_DescPrev_Col = 5
    '##########################

    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARCELAS + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 900

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
    GridDescontos.ColWidth(0) = 200

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
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
    GridComissoes.ColWidth(0) = 200

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

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
    GridComissoesEmissao.ColWidth(0) = 750

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

Public Sub BotaoVendedores_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection, lErro As Long
    
On Error GoTo Erro_BotaoVendedores_Click

    If (GridComissoesEmissao.Row = 0) Then Error 56594
    
    'Chama a tela de browse de Vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedores)

    Exit Sub

Erro_BotaoVendedores_Click:

    Select Case Err

        Case 56594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175180)

    End Select

    Exit Sub

End Sub

'Private Sub DataDeposito_GotFocus()
'
'    'Call MaskEdBox_TrataGotFocus(DataDeposito, iAlterado)
'
'End Sub

'''Private Sub DataDeposito_Validate(Cancel As Boolean)
'''
'''Dim lErro As Long
'''
'''On Error GoTo Erro_DataDeposito_Validate
'''
'''    'Critica a data digitada
'''    lErro = Data_Critica(DataDeposito.Text)
'''    If lErro <> SUCESSO Then Error 59726
'''
'''    Exit Sub
'''
'''Erro_DataDeposito_Validate:
'''
'''    Cancel = True
'''
'''    Select Case Err
'''
'''        Case 59726
'''
'''        Case Else
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175181)
'''
'''    End Select
'''
'''    Exit Sub
'''
'''End Sub

Public Sub DataEmissao_GotFocus()
Dim iDataAux As Integer
    
    iDataAux = iDataEmissaoAlterada
    Call MaskEdBox_TrataGotFocus(DataEmissao, iAlterado)
    iDataEmissaoAlterada = iDataAux

End Sub


Public Sub INSSRetido_Click()

On Error GoTo Erro_INSSRetido_Click

    If giTrazendoTitulo = 1 Then Exit Sub
    
    iAlterado = REGISTRO_ALTERADO
    
    If Len(Trim(ValorINSS)) <> 0 Then Call Recalcula_Parcelas_Comissoes
    
    Exit Sub
    
Erro_INSSRetido_Click:

    Select Case Err
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175182)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO

End Sub

'Private Sub Numero_GotFocus()
'
'   ' Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
'
'End Sub

Public Sub Moeda_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NumTitulo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NumTitulo, iAlterado)

End Sub

Private Sub objEventoTipoDocumento_evSelecao(obj1 As Object)

Dim objTipoDocumento As ClassTipoDocumento

    Set objTipoDocumento = obj1
    
    Tipo.Text = objTipoDocumento.sSigla
    
    Call Tipo_Validate(bSGECancelDummy)
    
    Me.Show
    
    Exit Sub

End Sub

Private Sub objEventoVendedores_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_objEventoVendedores_evSelecao
    
    If GridComissoesEmissao.Row >= GridComissoesEmissao.FixedRows Then

        Set objVendedor = obj1
        
        lErro = VendedorEmissao_Linha_Preenche(objVendedor)
        If lErro <> SUCESSO Then Error 59020
        
    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoVendedores_evSelecao:

    Select Case Err
    
    Case 59020
    
    Case Else
        lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175183)
            
    End Select
        
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
Dim colSelecao As New Collection
Dim colComissaoEmissao As New colComissao
Dim iPreenchido As Integer
Dim sSelecao As String

On Error GoTo Erro_NumeroLabel_Click

    'Move os dados da Tela para objtituloreceber e colParcelas
    lErro = Move_Tela_Memoria(objTituloReceber, colComissaoEmissao)
    If lErro <> SUCESSO Then Error 26055
    
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

    'Chama Tela TituloReceberLista
    Call Chama_Tela("TituloReceberLista", colSelecao, objTituloReceber, objEventoNumero, sSelecao)
            
    Exit Sub

Erro_NumeroLabel_Click:

    Select Case Err

        Case 26055

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175184)

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
    If lErro <> SUCESSO Then Error 26056

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show
    
    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case Err

        Case 26056

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175185)

    End Select

    Exit Sub

End Sub

Public Sub CondPagtoLabel_Click()

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection

    'Se Condição de Pagto estiver preenchida, extrai o código
    objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)

    'Chama a Tela CondicoesPagamentoCRLista
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondPagto)

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

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175186)

     End Select

     Exit Sub

End Sub

Function Trata_Parametros(Optional objTituloReceber As ClassTituloReceber) As Long

Dim lErro As Long
Dim objTipoDocumento As New ClassTipoDocumento

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Título
    If Not (objTituloReceber Is Nothing) Then

        'Lê o Título
        lErro = CF("TituloReceber_Le", objTituloReceber)
        If lErro <> SUCESSO And lErro <> 26061 Then Error 26062
        
        'Se não achou o Título
        If lErro <> SUCESSO Then Error 26063

        'Traz os dados para a Tela
        lErro = Traz_TituloReceber_Tela(objTituloReceber)
        If lErro <> SUCESSO Then Error 26064

    End If
    
    iAlterado = 0
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 26062, 26064

        Case 26063
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", Err, objTituloReceber.lNumIntDoc)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175187)

    End Select
    
    iAlterado = 0
    
    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se há alterações e quer salvá-las
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 42787
    
    'Limpa a Tela
    Call Limpa_Tela_TituloReceber
    Call Limpa_Colecoes
    
    iAlterado = 0
    
    Exit Sub
    
Erro_BotaoLimpar_Click:

    Select Case Err
    
        Case 42787
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175188)
    
    End Select
    
    Exit Sub

End Sub

Public Sub CondicaoPagamento_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CondicaoPagamento_Click()

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_CondicaoPagamento_Click

    If giTrazendoTitulo = 1 Then Exit Sub
    
    'Verifica se alguma Condição foi selecionada
    If CondicaoPagamento.ListIndex = -1 Then Exit Sub

    'Passa o código da Condição para objCondicaoPagto
    objCondicaoPagto.iCodigo = CondicaoPagamento.ItemData(CondicaoPagamento.ListIndex)

    'Lê Condição a partir do código
    lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
    If lErro <> SUCESSO And lErro <> 19205 Then Error 26145
    
    'Se não achou a Condição de Pagamento --> erro
    If lErro = 19205 Then Error 26146

    'Preenche o GridParcelas
    Call Recalcula_Parcelas_Comissoes

    iAlterado = REGISTRO_ALTERADO

    Exit Sub

Erro_CondicaoPagamento_Click:

    Select Case Err

        Case 26145, 26147

        Case 26146
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175189)

      End Select

    Exit Sub

End Sub

Public Sub CondicaoPagamento_Validate(bCancel As Boolean)

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 26148

    'Se não encontra valor que contém CÓDIGO, mas extrai o código
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Lê Condicao Pagamento no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 26149
        
        'Se não achou a Condição de Pagamento --> erro
        If lErro = 19205 Then Error 26150

        'Testa se pode ser usada em Contas à Receber
        If objCondicaoPagto.iEmRecebimento = 0 Then Error 26151

        'Coloca na Tela
        CondicaoPagamento.Text = CondPagto_Traz(objCondicaoPagto)

        'nao pode recalcular se estiver trazendo do bd
        If giTrazendoTitulo = 0 Then
        
            'Preenche o GridParcelas
            Call Recalcula_Parcelas_Comissoes

        End If
        
    End If

    'Não encontrou o valor que era STRING
    If lErro = 6731 Then Error 26153

    Exit Sub

Erro_CondicaoPagamento_Validate:

    bCancel = True
    
    Select Case Err

       Case 26148, 26149, 26152

       Case 26150
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAOPAGTO", iCodigo)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 26151
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", Err, objCondicaoPagto.iCodigo)

        Case 26153
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", Err, CondicaoPagamento.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175190)

    End Select

    Exit Sub

End Sub

Public Sub DataEmissao_Change()

    iAlterado = REGISTRO_ALTERADO
    iDataEmissaoAlterada = REGISTRO_ALTERADO

End Sub

Public Sub DataEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataEmissao_Validate
    
    'se a data não foi alterada ------> Fim
    If iDataEmissaoAlterada <> REGISTRO_ALTERADO Then Exit Sub
    
    'Critica a data digitada
    lErro = Data_Critica(DataEmissao.Text)
    If lErro <> SUCESSO Then Error 26140
        
    'força o recalculo das datas das parcelas
    Call Recalcula_Datas_Parcelas
    
    iDataEmissaoAlterada = 0
    
    Exit Sub

Erro_DataEmissao_Validate:

    Cancel = True
    
    Select Case Err

        Case 26140
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175191)

    End Select

    Exit Sub

End Sub

Public Sub Opcao_GotFocus()

Dim lErro As Long

On Error GoTo Erro_Opcao_GotFocus
  
'    lErro = Trata_Mudanca_Frame()
'    If lErro <> SUCESSO Then Error 26478
  
    Exit Sub
  
Erro_Opcao_GotFocus:

    Select Case Err

        Case 26478
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175192)
            
    End Select
    
    Exit Sub

End Sub

Public Sub Percentual1_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PercentualComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ReajustePeriodicidade_Click()

    iAlterado = REGISTRO_ALTERADO

    If ReajustePeriodicidade.ListIndex = 0 Then
        Moeda.ListIndex = -1
        Call DateParaMasked(ReajusteBase, DATA_NULA)
    End If

End Sub

Public Sub TipoDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub TipoDocumentoLabel_Click()

Dim colSelecao As Collection
Dim objTipoDocumento As New ClassTipoDocumento

    objTipoDocumento.sSigla = SCodigo_Extrai(Tipo.Text)

    Call Chama_Tela("TipoDocTituloRecLista", colSelecao, objTipoDocumento, objEventoTipoDocumento)

End Sub

'''Private Sub UpDataDeposito_DownClick()
'''
'''Dim lErro As Long
'''Dim sData As String
'''
'''On Error GoTo Erro_UpDataDeposito_DownClick
'''
'''    'Diminui a data
'''    lErro = Data_Up_Down_Click(DataDeposito, DIMINUI_DATA)
'''    If lErro Then Error 26374
'''
'''    Exit Sub
'''
'''Erro_UpDataDeposito_DownClick:
'''
'''    Select Case Err
'''
'''        Case 26374
'''
'''        Case Else
'''             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175193)
'''
'''    End Select
'''
'''    Exit Sub
'''
'''End Sub
'''
'''Private Sub UpDataDeposito_UpClick()
'''
'''Dim lErro As Long
'''Dim sData As String
'''
'''On Error GoTo Erro_UpDataDeposito_UpClick
'''
'''    'Aumenta a Data de deposito em um dia
'''    lErro = Data_Up_Down_Click(DataDeposito, AUMENTA_DATA)
'''    If lErro Then Error 26142
'''
'''    Exit Sub
'''
'''Erro_UpDataDeposito_UpClick:
'''
'''    Select Case Err
'''
'''        Case 26142
'''
'''        Case Else
'''             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175194)
'''
'''    End Select
'''
'''    Exit Sub
'''
'''End Sub

Public Sub UpDownEmissao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(DataEmissao, DIMINUI_DATA)
    If lErro Then Error 26141

    Exit Sub

Erro_UpDownEmissao_DownClick:

    Select Case Err

        Case 26141

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175195)

    End Select

    Exit Sub

End Sub

Public Sub UpDownEmissao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownEmissao_UpClick

    'Aumenta a Data de Emissão em um dia
    lErro = Data_Up_Down_Click(DataEmissao, AUMENTA_DATA)
    If lErro Then Error 26142

    Exit Sub

Erro_UpDownEmissao_UpClick:

    Select Case Err

        Case 26142

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175196)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub NumTitulo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub NumTitulo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumTitulo_Validate

    'Verifica se o Numero foi preenchido
    If Len(Trim(NumTitulo.ClipText)) = 0 Then Exit Sub

    'Critica se é Long positivo
    lErro = Long_Critica(NumTitulo.ClipText)
    If lErro <> SUCESSO Then Error 26137
    
    Exit Sub

Erro_NumTitulo_Validate:

    Cancel = True


    Select Case Err

        Case 26137

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175197)

    End Select

    Exit Sub

End Sub

Public Sub Cliente_Change()

   iAlterado = REGISTRO_ALTERADO
   iClienteAlterado = 1

    Call Cliente_Preenche

End Sub

Public Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer, bCancel As Boolean
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    If iClienteAlterado = 1 Then

        'Verifica se o Cliente está preenchido
        If Len(Trim(Cliente.Text)) > 0 Then

            lErro = TP_Cliente_Le(Cliente, objCliente, iCodFilial)
            If lErro <> SUCESSO Then Error 26135

            lErro = CF("FiliaisClientes_Le_Cliente", objCliente, colCodigoNome)
            If lErro <> SUCESSO Then Error 26136
            
            'Preenche ComboBox de Filiais
            Call CF("Filial_Preenche", Filial, colCodigoNome)

            'Seleciona filial na Combo Filial
            Call CF("Filial_Seleciona", Filial, iCodFilial)
            
            'Seleciona a Condição de Pagamento default
            If objCliente.iCondicaoPagto <> 0 Then
                CondicaoPagamento.Text = objCliente.iCondicaoPagto
                Call CondicaoPagamento_Validate(bCancel)
            End If
            
        'Se não estiver preenchido
        ElseIf Len(Trim(Cliente.Text)) = 0 Then

            'Limpa a Combo de Filiais
            Filial.Clear
            
        End If

        'Coloca clientes em atraso em vermelho
        If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
        iClienteAlterado = 0

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case Err

        Case 26135
            
        Case 26136

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175198)

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

    lErro = Trata_Mudanca_Frame()
    If lErro <> SUCESSO Then Error 26478
    
    'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
    If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao
        
    Select Case iFrameAtual
    
        Case TAB_Identificacao
            Parent.HelpContextID = IDH_TITULOS_RECEBER_ID
            
        Case TAB_Parcelas
            Parent.HelpContextID = IDH_TITULOS_RECEBER_PARCELAS
            
        Case TAB_Contabilizacao
            Parent.HelpContextID = IDH_TITULOS_RECEBER_CONTABILIZACAO
                    
    End Select
        
    Exit Sub
    
Erro_Opcao_Click:

    Select Case Err

        Case 26478
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175199)
            
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
Dim lPosicaoSeparador As Long, objVendedor As New ClassVendedor
Dim iIndice1 As Integer, Cancel As Boolean, sCclMascarado As String
Dim sNaturezaEnxuta As String

On Error GoTo Erro_Traz_TituloReceber_Tela

    giTrazendoTitulo = 1
    
    Call Limpa_Colecoes
    
    'Limpa a tela
    Call Limpa_Tela_TituloReceber
    
    'Lê o Título à Receber
    lErro = CF("TituloReceber_Le", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26061 Then gError 43069
    
    'Não encontrou o Título à Receber --> erro
    If lErro = 26061 Then gError 43070
    
    'Coloca o Cliente na Tela
    Cliente.Text = objTituloReceber.lCliente
    Call Cliente_Validate(Cancel)

    'Coloca a Filial na Tela
    Filial.Text = objTituloReceber.iFilial
    Call Filial_Validate(bSGECancelDummy)
        
    'Coloca os demais dados do Título a Receber na Tela
    NumTitulo.Text = objTituloReceber.lNumTitulo
    NumeroFAT.Caption = objTituloReceber.lNumTitulo
    Call DateParaMasked(DataEmissao, objTituloReceber.dtDataEmissao)
    
    Tipo.Text = objTituloReceber.sSiglaDocumento
    Call Tipo_Validate(bSGECancelDummy)
    
    Valor.Text = Format(objTituloReceber.dValor, "Standard")
    ValorIRRF.Text = Format(objTituloReceber.dValorIRRF, "Standard")
    ValorINSS.Text = Format(objTituloReceber.dValorINSS, "Standard")
    INSSRetido.Value = IIf(objTituloReceber.iINSSRetido <> 0, vbChecked, vbUnchecked)
    PISRetido.Text = Format(objTituloReceber.dPISRetido, "Standard")
    ISSRetido.Text = Format(objTituloReceber.dISSRetido, "Standard")
    COFINSRetido.Text = Format(objTituloReceber.dCOFINSRetido, "Standard")
    CSLLRetido.Text = Format(objTituloReceber.dCSLLRetido, "Standard")
    
    If Len(Trim(objTituloReceber.sCcl)) <> 0 Then
    
        sCclMascarado = String(STRING_CCL, 0)
        
        lErro = Mascara_RetornaCclEnxuta(objTituloReceber.sCcl, sCclMascarado)
        If lErro <> SUCESSO Then gError 22930
    
        Ccl.PromptInclude = False
        Ccl.Text = sCclMascarado
        Ccl.PromptInclude = True
        
    Else
        
        Ccl.PromptInclude = False
        Ccl.Text = ""
        Ccl.PromptInclude = True
    
    End If
    
    If objTituloReceber.iCondicaoPagto <> 0 Then
    
        CondicaoPagamento.Text = objTituloReceber.iCondicaoPagto
        Call CondicaoPagamento_Validate(bCancel)
        
    Else
        CondicaoPagamento.Text = ""
    End If
    
    'Lê as Parcelas a Receber vinculadas ao Título
    lErro = CF("ParcelasReceber_Le", objTituloReceber, colParcelaReceber)
    If lErro <> SUCESSO Then gError 26380

    If colParcelaReceber.Count > NUM_MAXIMO_PARCELAS Then gError 26381

    'Lê as comissões do Título
    lErro = CF("Comissoes_Le", objTituloReceber.lNumIntDoc, colComEmissao, TIPO_TITULO_RECEBER)
    If lErro <> SUCESSO Then gError 26164
       
    'Carrega o Grid de Parcelas
    lErro = Carrega_GridParcelas()
    If lErro <> SUCESSO Then gError 26379
    
    iIndice = 0
    
    'Loop de preenchimento do GridComissoesEmissao
    For Each objComissao In colComEmissao
    
        iIndice = iIndice + 1
        
        'ler do bd o nome reduzido do vendedor
        objVendedor.iCodigo = objComissao.iCodVendedor
        lErro = CF("Vendedor_Le", objVendedor)
        If lErro <> SUCESSO And lErro <> 12582 Then gError 59145
        If lErro <> SUCESSO Then gError 59146
        
        GridComissoesEmissao.TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        
        GridComissoesEmissao.TextMatrix(iIndice, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
        GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
        GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")
    
    Next

    objGridComissoesEmissao.iLinhasExistentes = iIndice
    
    'Chama Soma_Percentual
    lErro = Soma_Percentual(objGridComissoesEmissao)
    If lErro <> SUCESSO Then gError 26128

    'Chama Soma_Valor
    lErro = Soma_Valor(objGridComissoesEmissao)
    If lErro <> SUCESSO Then gError 26129
    
    If Len(Trim(objTituloReceber.sNatureza)) <> 0 Then
    
        sNaturezaEnxuta = String(STRING_NATMOVCTA_CODIGO, 0)
    
        lErro = Mascara_RetornaItemEnxuto(SEGMENTO_NATMOVCTA, objTituloReceber.sNatureza, sNaturezaEnxuta)
        If lErro <> SUCESSO Then gError 32297
    
        Natureza.PromptInclude = False
        Natureza.Text = sNaturezaEnxuta
        Natureza.PromptInclude = True
        
    Else
    
        Natureza.PromptInclude = False
        Natureza.Text = ""
        Natureza.PromptInclude = True
        
    End If
    
    Call Natureza_Validate(bSGECancelDummy)
    
    'trazer dados de reajuste
    If objTituloReceber.iReajustePeriodicidade = 0 Then
        ReajustePeriodicidade.ListIndex = 0
    Else
        Call Combo_Seleciona_ItemData(ReajustePeriodicidade, objTituloReceber.iReajustePeriodicidade)
        
        For iIndice = 0 To Moeda.ListCount - 1
            
            If objTituloReceber.iMoeda = Codigo_Extrai(Moeda.List(iIndice)) Then
                
                Moeda.ListIndex = iIndice
                
                Exit For
                         
            End If
                         
        Next
        
        Call DateParaMasked(ReajusteBase, objTituloReceber.dtReajusteBase)
    End If
    
    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objTituloReceber.lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39581

    '###################################################
    'Inserido por Wagner 04/08/2006
    lErro = gobjTelaProjetoInfo.Traz_Dados_Tela(objTituloReceber.objProjetoInfo, objTituloReceber.lNumIntDoc, PRJ_CR_TIPO_TITREC)
    If lErro <> SUCESSO Then gError 181676
    '###################################################

    iAlterado = 0
    giTrazendoTitulo = 0
    iValorAlterado = 0
    iDataEmissaoAlterada = 0
    iValorINSSAlterado = 0
    iValorIRRFAlterado = 0
    iPISRetidoAlterado = 0
    iISSRetidoAlterado = 0
    iCOFINSRetidoAlterado = 0
    iCSLLRetidoAlterado = 0

    Traz_TituloReceber_Tela = SUCESSO

    Exit Function

Erro_Traz_TituloReceber_Tela:

    Traz_TituloReceber_Tela = gErr

    giTrazendoTitulo = 0
    
    Select Case gErr

        Case 26128, 26129, 26379, 26380, 59145, 22930, 32297, 181676

        Case 26381
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUM_MAXIMO_PARCELAS_ULTRAPASSADO", gErr, colParcelaReceber.Count, NUM_MAXIMO_PARCELAS)
        
        Case 39581, 43069
        
        Case 43070
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO", gErr, objTituloReceber.lNumIntDoc)

        Case 59146
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", gErr, objVendedor.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175200)

    End Select

    Exit Function

End Function

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

Public Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentes As Integer
Dim iLinhaAtual As Integer

    iLinhasExistentes = objGridParcelas.iLinhasExistentes
    iLinhaAtual = GridParcelas.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
    
   If (objGridParcelas.iLinhasExistentes < iLinhasExistentes) Then

        If colParcelaReceber.Count >= iLinhaAtual Then colParcelaReceber.Remove (iLinhaAtual)
        If colcolDesconto.Count >= iLinhaAtual Then colcolDesconto.Remove (iLinhaAtual)
        If colcolComissao.Count >= iLinhaAtual Then colcolComissao.Remove (iLinhaAtual)

    End If

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
    
    Call Exibe_Dados_Parcela
    
    objGridParcelas.iLinhaAntiga = GridParcelas.Row

End Sub

Public Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Public Sub GridComissoes_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_EnterCell()

    Call Grid_Entrada_Celula(objGridComissoes, iAlterado)

End Sub

Public Sub GridComissoes_GotFocus()

    Call Grid_Recebe_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentes As Integer
Dim iParcela As Integer
Dim iComissao As Integer

    iLinhasExistentes = objGridComissoes.iLinhasExistentes
    iComissao = GridComissoes.Row

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoes)
    
    If iLinhasExistentes > objGridComissoes.iLinhasExistentes Then
            
        'Chama Soma_Percentual
        lErro = Soma_Percentual(objGridComissoes)
        If lErro <> SUCESSO Then Error 43067

        'Chama Soma_Valor
        lErro = Soma_Valor(objGridComissoes)
        If lErro <> SUCESSO Then Error 43068
    
        iParcela = GridParcelas.Row
        
        If iParcela = 0 Then Exit Sub
        
        colcolComissao.Item(iParcela).colComissao.Remove iComissao
        
    End If
    
    Exit Sub
    
Erro_GridComissoes_KeyDown:

    Select Case Err
    
        Case 43067, 43068

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175201)

    End Select

    Exit Sub

End Sub

Public Sub GridComissoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoes, iAlterado)
    End If

End Sub

Public Sub GridComissoes_LeaveCell()

    Call Saida_Celula(objGridComissoes)

End Sub

Public Sub GridComissoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoes)

End Sub

Public Sub GridComissoes_RowColChange()

    Call Grid_RowColChange(objGridComissoes)

End Sub

Public Sub GridComissoes_Scroll()

    Call Grid_Scroll(objGridComissoes)

End Sub

Public Sub GridDescontos_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridDesconto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDesconto, iAlterado)
    End If

End Sub

Public Sub GridDescontos_EnterCell()

    Call Grid_Entrada_Celula(objGridDesconto, iAlterado)

End Sub

Public Sub GridDescontos_GotFocus()

    Call Grid_Recebe_Foco(objGridDesconto)

End Sub

Public Sub GridDescontos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridDesconto)

End Sub

Public Sub GridDescontos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridDesconto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridDesconto, iAlterado)
    End If

End Sub
Public Sub GridDescontos_LeaveCell()

    Call Saida_Celula(objGridDesconto)

End Sub

Public Sub GridDescontos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridDesconto)

End Sub

Public Sub GridDescontos_RowColChange()

    Call Grid_RowColChange(objGridDesconto)

End Sub

Public Sub GridDescontos_Scroll()

    Call Grid_Scroll(objGridDesconto)

End Sub

Public Sub GridComissoesEmissao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissoesEmissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoesEmissao, iAlterado)
    End If

End Sub

Public Sub GridComissoesEmissao_EnterCell()

    Call Grid_Entrada_Celula(objGridComissoesEmissao, iAlterado)

End Sub

Public Sub GridComissoesEmissao_GotFocus()

    Call Grid_Recebe_Foco(objGridComissoesEmissao)

End Sub

Public Sub GridComissoesEmissao_KeyDown(KeyCode As Integer, Shift As Integer)

Dim lErro As Long
Dim iLinhasExistentes As Integer

On Error GoTo Erro_GridComissoesEmissao_KeyDown

    iLinhasExistentes = objGridComissoesEmissao.iLinhasExistentes

    Call Grid_Trata_Tecla1(KeyCode, objGridComissoesEmissao)
    
    If iLinhasExistentes > objGridComissoesEmissao.iLinhasExistentes Then
    
        'Chama Soma_Percentual
        lErro = Soma_Percentual(objGridComissoesEmissao)
        If lErro <> SUCESSO Then Error 43065

        'Chama Soma_Valor
        lErro = Soma_Valor(objGridComissoesEmissao)
        If lErro <> SUCESSO Then Error 43066

    End If
    
    Exit Sub

Erro_GridComissoesEmissao_KeyDown:

    Select Case Err
    
        Case 43065, 43066

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175202)

    End Select

    Exit Sub
    
End Sub

Public Sub GridComissoesEmissao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissoesEmissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissoesEmissao, iAlterado)
    End If

End Sub

Public Sub GridComissoesEmissao_LeaveCell()

    Call Saida_Celula(objGridComissoesEmissao)

End Sub

Public Sub GridComissoesEmissao_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridComissoesEmissao)

End Sub

Public Sub GridComissoesEmissao_RowColChange()

    Call Grid_RowColChange(objGridComissoesEmissao)

End Sub

Public Sub GridComissoesEmissao_Scroll()

    Call Grid_Scroll(objGridComissoesEmissao)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Tratamento de saída de célula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 39582

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name

            'Se for o GridParcelas
            Case GridParcelas.Name

                lErro = Saida_Celula_GridParcelas(objGridInt)
                If lErro <> SUCESSO Then Error 26064

            'Se for o GridDescontos
            Case GridDescontos.Name

                lErro = Saida_Celula_GridDescontos(objGridInt)
                If lErro <> SUCESSO Then Error 26065

            'Se for o GridComissoes
            Case GridComissoes.Name

                lErro = Saida_Celula_GridComissoes(objGridInt)
                If lErro <> SUCESSO Then Error 26066

            'Se for o GridComissoesEmissao
            Case GridComissoesEmissao.Name

                lErro = Saida_Celula_GridComissoesEmissao(objGridInt)
                If lErro <> SUCESSO Then Error 26067

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 26068

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 26064, 26065, 26066, 26067, 26068, 39582

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175203)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridParcelas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridParcelas

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Critica a Data de Vencimento e gera a Data de Vencto Reral
        Case iGrid_Vencimento_Col
            lErro = Saida_Celula_Vencimento(objGridInt)
            If lErro <> SUCESSO Then gError 26069

        'Faz a critica do valor da Parcela
        Case iGrid_Valor_Col
            lErro = Saida_Celula_Valor(objGridInt)
            If lErro <> SUCESSO Then gError 26070
        
        '##################################################
        'INSERIDO POR WAGNER
        Case iGrid_DescPrev_Col
            lErro = Saida_Celula_DescPrev(objGridInt)
            If lErro <> SUCESSO Then gError 129967
        
        Case iGrid_Previsao_Col
            lErro = Saida_Celula_Previsao(objGridInt)
            If lErro <> SUCESSO Then gError 129968
        'FIM
        '#################################################

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 26071

    Saida_Celula_GridParcelas = SUCESSO

    Exit Function

Erro_Saida_Celula_GridParcelas:

    Saida_Celula_GridParcelas = gErr

    Select Case gErr

        '##################################################
        'ALTERADO POR WAGNER
        Case 26069, 26070, 129967, 129968
        '##################################################

        Case 26071
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175204)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridDescontos(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridDescontos

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_TipoDesconto_Col
            'Faz a crítica do tipo de desconto
            lErro = Saida_Celula_TipoDesconto(objGridInt)
            If lErro <> SUCESSO Then Error 26072

        Case iGrid_Data_Col
            'Faz a crítica da Data
            lErro = Saida_Celula_Data(objGridInt)
            If lErro <> SUCESSO Then Error 26073

        Case iGrid_ValorDesconto_Col
            'Faz a crítica do Valor do desconto
            lErro = Saida_Celula_ValorDesconto(objGridInt)
            If lErro <> SUCESSO Then Error 26074

        Case iGrid_Percentual1_Col
            'Faz a crítica do Percentual do desconto
            lErro = Saida_Celula_Percentual(objGridInt)
            If lErro <> SUCESSO Then Error 26075

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26076

    Saida_Celula_GridDescontos = SUCESSO

    Exit Function

Erro_Saida_Celula_GridDescontos:

    Saida_Celula_GridDescontos = Err

    Select Case Err

        Case 26072, 26073, 26074, 26075

        Case 26076
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175205)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComissoes(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComissoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vendedor_Col
            'Faz a crítica do vendedor
            lErro = Saida_Celula_Vendedor(objGridInt)
            If lErro <> SUCESSO Then Error 26077

        Case iGrid_PercentualComissao_Col
            'Faz a crítica do Percentual da comissao
            lErro = Saida_Celula_PercentualComissao(objGridInt)
            If lErro <> SUCESSO Then Error 26078

        Case iGrid_ValorBase_Col
            'Faz a crítica do Valor Base
            lErro = Saida_Celula_ValorBase(objGridInt)
            If lErro <> SUCESSO Then Error 26079

        Case iGrid_ValorComissao_Col
            'Faz a crítica do Valor da Comissao
            lErro = Saida_Celula_ValorComissao(objGridInt)
            If lErro <> SUCESSO Then Error 26080

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26081

    Saida_Celula_GridComissoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComissoes:

    Saida_Celula_GridComissoes = Err

    Select Case Err

        Case 26077, 26078, 26079, 26080

        Case 26081
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175206)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComissoesEmissao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComissoesEmissao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        Case iGrid_Vendedor_Col
            'Faz a crítica do vendedor
            lErro = Saida_Celula_VendedorEmissao(objGridInt)
            If lErro <> SUCESSO Then Error 26082

        Case iGrid_PercentualComissao_Col
            'Faz a crítica do Percentual da comissao
            lErro = Saida_Celula_PercentualEmissao(objGridInt)
            If lErro <> SUCESSO Then Error 26083

        Case iGrid_ValorBase_Col
            'Faz a crítica do Valor Base
            lErro = Saida_Celula_ValorBaseEmissao(objGridInt)
            If lErro <> SUCESSO Then Error 26084

        Case iGrid_ValorComissao_Col
            'Faz a crítica do Valor da Comissao
            lErro = Saida_Celula_ValorEmissao(objGridInt)
            If lErro <> SUCESSO Then Error 26085

    End Select

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26086

    Saida_Celula_GridComissoesEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComissoesEmissao:

    Saida_Celula_GridComissoesEmissao = Err

    Select Case Err

        Case 26082, 26083, 26084, 26085

        Case 26086
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175207)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Vencimento(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Vencimento que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataEmissao As Date
Dim dtDataVencimento As Date
Dim dtDataVenctoReal As Date
Dim sDataVencimento As String
Dim colDesconto As New colDesconto
Dim colComissao As New colComissao

On Error GoTo Erro_Saida_Celula_Vencimento

    Set objGridInt.objControle = DataVencimento

    'Verifica se Data de Vencimento esta preenchida
    If Len(Trim(DataVencimento.ClipText)) > 0 Then

        'Critica a data
        lErro = Data_Critica(DataVencimento.Text)
        If lErro <> SUCESSO Then Error 26087

         dtDataVencimento = CDate(DataVencimento.Text)

        'Se data de Emissao estiver preenchida verificar se a Data de Vencimento é maior que a Data de Emissão
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            dtDataEmissao = CDate(DataEmissao.Text)
            If dtDataVencimento < DataEmissao Then Error 26088
        End If

        sDataVencimento = Format(dtDataVencimento, "dd/mm/yyyy")
        
        'Calcula a Data de Vencimento Real
        lErro = CF("DataVencto_Real", dtDataVencimento, dtDataVenctoReal)
        If lErro <> SUCESSO Then Error 26089

        'Coloca data de Vencimento Real no Grid
        GridParcelas.TextMatrix(GridParcelas.Row, iGrid_VenctoReal_Col) = Format(dtDataVenctoReal, "dd/mm/yyyy")

        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            colParcelaReceber.Add 0, 0, GridParcelas.Row, STATUS_ABERTO, DATA_NULA, DATA_NULA, 0, 0, 1, 1, 1, " ", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, "", 0
            colcolDesconto.Add colDesconto
            colcolComissao.Add colComissao
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    If sDataVencimento <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Vencimento_Col) Then CobrancaAutomatica.Value = vbUnchecked
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26090

    Saida_Celula_Vencimento = SUCESSO

    Exit Function

Erro_Saida_Celula_Vencimento:

    Saida_Celula_Vencimento = Err

    Select Case Err

        Case 26087, 26089, 26090
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26088
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", Err, DataVencimento.Text, GridParcelas.Row, DataEmissao.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175208)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Valor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim colDesconto As New colDesconto
Dim colComissao As New colComissao

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = ValorParcela

    'Verifica se valor está preenchido
    If Len(ValorParcela.ClipText) > 0 Then
    
        'Critica se valor é positivo
        lErro = Valor_Positivo_Critica(ValorParcela.Text)
        If lErro <> SUCESSO Then Error 26091

        ValorParcela.Text = Format(ValorParcela.Text, "Standard")
        If ValorParcela.Text <> GridParcelas.TextMatrix(GridParcelas.Row, iGrid_Valor_Col) Then CobrancaAutomatica.Value = vbUnchecked
        
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            colParcelaReceber.Add 0, 0, GridParcelas.Row, STATUS_ABERTO, DATA_NULA, DATA_NULA, 0, 0, 1, 1, 1, " ", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, "", 0
            colcolDesconto.Add colDesconto
            colcolComissao.Add colComissao
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
        colParcelaReceber.Item(GridParcelas.Row).dSaldo = StrParaDbl(ValorParcela.Text)
        colParcelaReceber.Item(GridParcelas.Row).dValor = StrParaDbl(ValorParcela.Text)
        colParcelaReceber.Item(GridParcelas.Row).dValorOriginal = StrParaDbl(ValorParcela.Text)
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26092

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = Err

    Select Case Err

        Case 26091, 26092
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175209)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Desconto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim dValor As Double
Dim dPercentual As Double
Dim iParcela As Integer

On Error GoTo Erro_Saida_Celula_TipoDesconto

    Set objGridInt.objControle = TipoDesconto

    iParcela = GridParcelas.Row
    
    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoDesconto.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoDesconto.ListIndex = -1 Then

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(TipoDesconto, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then Error 26093
        
            'Não foi encontrado
            If lErro = 25085 Then Error 26094
            
            If lErro = 25086 Then Error 26095

        End If
        
        iCodigo = Codigo_Extrai(TipoDesconto.Text)
        
        If iCodigo = VALOR_FIXO Or iCodigo = VALOR_ANT_DIA Or iCodigo = VALOR_ANT_DIA_UTIL Then
            GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Percentual1_Col) = ""
        Else
            dValor = StrParaDbl(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_ValorDesconto_Col))
            If dValor > 0 And iParcela > 0 Then
                dPercentual = dValor / colParcelaReceber(iParcela).dValor
                If dPercentual > 0 Then
                    GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Percentual1_Col) = Format(dPercentual, "Percent")
                Else
                    GridDescontos.TextMatrix(GridDescontos.Row, iGrid_ValorDesconto_Col) = ""
                End If
            Else
                GridDescontos.TextMatrix(GridDescontos.Row, iGrid_ValorDesconto_Col) = ""
            End If
            
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridDescontos.Row - GridDescontos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    Else
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Percentual1_Col) = ""
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_Data_Col) = ""
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_ValorDesconto_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26174

    Saida_Celula_TipoDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoDesconto:

    Saida_Celula_TipoDesconto = Err

    Select Case Err

        Case 26093
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26094
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO", Err, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26095
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPODESCONTO_NAO_ENCONTRADO1", Err, TipoDesconto.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26174
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175210)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Data(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = Data

    'Verifica se a Data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then

        'Faz a Crítica da Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then Error 26096

        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26413

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = Err

    Select Case Err

        Case 26096, 26413
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175211)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ValorDesconto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Desconto do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iParcela As Integer


On Error GoTo Erro_Saida_Celula_ValorDesconto

    Set objGridInt.objControle = ValorDesconto

    iParcela = GridParcelas.Row
    
    'Verifica se o Valor do Desconto foi preenchido
    If Len(Trim(ValorDesconto.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(ValorDesconto.Text)
        If lErro <> SUCESSO Then Error 26097

        If iParcela > 0 Then
            If colParcelaReceber(iParcela).dValor < StrParaDbl(ValorDesconto.Text) Then Error 51631
        End If
        
        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26175


    Saida_Celula_ValorDesconto = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorDesconto:

    Saida_Celula_ValorDesconto = Err

    Select Case Err

        Case 26175
        
        Case 51631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_MENOR_DESCONTO", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26097
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175212)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Percentual(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim dPercentual As Double
Dim lErro As Long
Dim dValor As Double
Dim iParcela As Integer

On Error GoTo Erro_Saida_Celula_Percentual

    Set objGridInt.objControle = Percentual1

    If Len(Trim(Percentual1.ClipText)) > 0 Then

        'Verifica se o Percentual foi preenchido
        lErro = Porcentagem_Critica(Percentual1.Text)
        If lErro <> SUCESSO Then Error 26098

        'Formata o Percentual
        Percentual1.Text = Format(Percentual1.Text, "Standard")
        
        iParcela = GridParcelas.Row
        
        If iParcela > 0 Then
            dPercentual = StrParaDbl(Percentual1.Text) / 100
            dValor = colParcelaReceber(iParcela).dValor * dPercentual
            GridDescontos.TextMatrix(GridDescontos.Row, iGrid_ValorDesconto_Col) = Format(dValor, "Standard")
        End If

        'Acrescenta uma linha no Grid se for o caso
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    Else
    
        GridDescontos.TextMatrix(GridDescontos.Row, iGrid_ValorDesconto_Col) = ""
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26176

    Saida_Celula_Percentual = SUCESSO

    Exit Function

Erro_Saida_Celula_Percentual:

    Saida_Celula_Percentual = Err

    Select Case Err

        Case 26098, 26176
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175213)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Vendedor(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim iParcela As Integer
Dim objVendedor As New ClassVendedor
Dim vbMsgRes As VbMsgBoxResult
Dim dPercentual As Double
Dim dValorParcela As Double

On Error GoTo Erro_Saida_Celula_Vendedor

    Set objGridInt.objControle = Vendedor

    'Verifica se vendedor está preenchido
    If Len(Trim(Vendedor.Text)) > 0 Then

        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Grid(Vendedor, objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then Error 26099

        If lErro = 25018 Then Error 26100

        If lErro = 25020 Then Error 26101

        'Loop no GridComissoes
        For iIndice = 1 To objGridComissoes.iLinhasExistentes

            'Verifica se Vendedor comparece em outra linha
            If iIndice <> GridComissoes.Row Then If GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido Then Error 26102

        Next
        
        iParcela = GridParcelas.Row
        
        If iParcela = 0 Then Error 56594
        
        If objVendedor.sNomeReduzido <> GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col) Then
        
            'Acrescenta uma linha no Grid se for o caso
            If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
                objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
                If iParcela <> 0 Then
                    colcolComissao.Item(iParcela).colComissao.Add 0, 0, 0, 0, 0, objVendedor.iCodigo, DATA_NULA, 0, 0, 0, DATA_NULA
                End If
            End If

            Call Vendedor_Linha_Preenche(objVendedor, iParcela)

        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26114

    Saida_Celula_Vendedor = SUCESSO

    Exit Function

Erro_Saida_Celula_Vendedor:

    Saida_Celula_Vendedor = Err

    Select Case Err

        Case 26099

        Case 26100 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = Vendedor.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 26101 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = CDbl(Vendedor.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 26102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err, objVendedor.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 26114
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
               
        Case 56594
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175214)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula PercentualComissoes do grid que está deixando de ser o corrente

Dim lErro As Long
Dim iParcela As Integer
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_PercentualComissao

    Set objGridInt.objControle = PercentualComissao

    'Verifica se o percentual está preenchido
    If Len(PercentualComissao.ClipText) > 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(PercentualComissao.Text)
        If lErro <> SUCESSO Then Error 26103

        dPercentual = CDbl(PercentualComissao.Text)

        'Mostra na tela o percentual formatado
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col) = Format(dPercentual, "Standard")
        
        dPercentual = dPercentual / 100
        
        'Verifica se valor base correspondente esta preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))

           'Calcula o valor da comissao
           dValorComissao = dPercentual * dValorBase

           'Coloca o valor da comissao na tela
           GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        End If

        iParcela = GridParcelas.Row

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            If iParcela <> 0 Then
                colcolComissao.Item(iParcela).colComissao.Add 0, 0, 0, 0, 0, 0, DATA_NULA, dPercentual, 0, 0, DATA_NULA
            End If
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26104

    'Chama Soma_Percentual
    lErro = Soma_Percentual(objGridInt)
    If lErro <> SUCESSO Then Error 26105

    'Chama Soma_Valor
    lErro = Soma_Valor(objGridInt)
    If lErro <> SUCESSO Then Error 26106

    Saida_Celula_PercentualComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualComissao:

    Saida_Celula_PercentualComissao = Err

    Select Case Err

        Case 26103, 26104
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26105, 26106

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175215)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorBase(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Base do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lTamanho As Long

On Error GoTo Erro_Saida_Celula_ValorBase

    Set objGridInt.objControle = ValorBase

    'Verifica se valor base está preenchido
    If Len(ValorBase.ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(ValorBase.Text)
        If lErro <> SUCESSO Then Error 26107

        dValorBase = CDbl(ValorBase.Text)

        'Mostra na tela o ValorBase formatado
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col) = Format(dValorBase, "Standard")

        lTamanho = Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col)))

        'Verifica se percentual comissao está preenchido
        If lTamanho > 0 Then

            dPercentual = PercentParaDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col))

            'Calcula o valor da comissao
            dValorComissao = dPercentual * dValorBase

            'Mostra na tela o valor da comissao
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26108

    'Chama SomaValor
    lErro = Soma_Valor(objGridInt)
    If lErro <> SUCESSO Then Error 26109

    Saida_Celula_ValorBase = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBase:

    Saida_Celula_ValorBase = Err

    Select Case Err

        Case 26107, 26108
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26109

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175216)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorComissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lComprimento As Long

On Error GoTo Erro_Saida_Celula_ValorComissao

    Set objGridInt.objControle = ValorComissao

    'Verifica se valor está preenchido
    If Len(ValorComissao.ClipText) > 0 Then

        'Critica se valor da comissao é positivo
        lErro = Valor_Positivo_Critica(ValorComissao.Text)
        If lErro <> SUCESSO Then Error 26110

        dValorComissao = CDbl(ValorComissao.Text)

        'Mostra na tela o Valor
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        'Verifica se valor base correspondente está preenchido
        If Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col))

            'Verifica se o valor base é menor que o valor corrente
            If dValorComissao > dValorBase Then Error 26355
                        
            'Verifica se percentual comissao correspondente está preenchido
            lComprimento = Len(Trim(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col)))
            
            If lComprimento > 0 Then dPercentual = PercentParaDbl(GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col))

            If Format(dPercentual * dValorBase, "Standard") <> dValorComissao Then

                dPercentual = dValorComissao / dValorBase

                'Mostra o percentual da comissao na tela
                GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col) = Format(dPercentual, "Percent")

            End If

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoes.Row - GridComissoes.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26111

    'Chama Soma_Percentual
    lErro = Soma_Percentual(objGridInt)
    If lErro <> SUCESSO Then Error 26112

    'Chama Soma_Valor
    lErro = Soma_Valor(objGridInt)
    If lErro <> SUCESSO Then Error 26113

    Saida_Celula_ValorComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorComissao:

    Saida_Celula_ValorComissao = Err

    Select Case Err

        Case 26110, 26111
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26112, 26113

        Case 26355
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175217)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_VendedorEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Vendedor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim iIndice As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim dValor As Double
Dim dPercentual As Double

On Error GoTo Erro_Saida_Celula_VendedorEmissao

    Set objGridInt.objControle = VendedorEmissao

    'Verifica se vendedor está preenchido
    If Len(Trim(VendedorEmissao.Text)) > 0 Then

        'Verifica se Vendedor existe
        lErro = TP_Vendedor_Grid(VendedorEmissao, objVendedor)
        If lErro <> SUCESSO And lErro <> 25018 And lErro <> 25020 Then Error 26115

        If lErro = 25018 Then Error 26116

        If lErro = 25020 Then Error 26117

        lErro = VendedorEmissao_Linha_Preenche(objVendedor)
        If lErro <> SUCESSO Then Error 59021
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26118

    Saida_Celula_VendedorEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_VendedorEmissao:

    Saida_Celula_VendedorEmissao = Err

    Select Case Err

        Case 26115

        Case 26116 'Não encontrou nome reduzido de vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Preenche objVendedor com nome reduzido
                objVendedor.sNomeReduzido = VendedorEmissao.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 26117 'Não encontrou codigo do vendedor no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_VENDEDOR")

            If vbMsgRes = vbYes Then

                'Prenche objVendedor com codigo
                objVendedor.iCodigo = CDbl(VendedorEmissao.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)

                'Chama a tela de Vendedores
                Call Chama_Tela("Vendedores", objVendedor)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 26102
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err, objVendedor.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 26118, 59021
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175218)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercentualEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Comissões do grid que está deixando de ser o corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double

On Error GoTo Erro_Saida_Celula_PercentualEmissao

    Set objGridInt.objControle = PercentualEmissao

    'Verifica se o percentual está preenchido
    If Len(PercentualEmissao.ClipText) > 0 Then

        'Critica se é porcentagem
        lErro = Porcentagem_Critica(PercentualEmissao.Text)
        If lErro <> SUCESSO Then Error 26119

        dPercentual = CDbl(PercentualEmissao.Text)

        'Mostra na tela o percentual formatado
        GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_PercentualComissao_Col) = Format(dPercentual, "Standard")

        dPercentual = dPercentual / 100

        'Verifica se valorbase correspondente esta preenchido
        If Len(Trim(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorBase_Col))

           'Calcula o valorcomissao
           dValorComissao = dPercentual * dValorBase

           'Coloca o valorcomissoes na tela
           GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoesEmissao.Row - GridComissoesEmissao.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26120

    'Chama Soma_Percentual
    lErro = Soma_Percentual(objGridInt)
    If lErro <> SUCESSO Then Error 26121

    'Chama Soma_Valor
    lErro = Soma_Valor(objGridInt)
    If lErro <> SUCESSO Then Error 26122

    Saida_Celula_PercentualEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_PercentualEmissao:

    Saida_Celula_PercentualEmissao = Err

    Select Case Err

        Case 26119, 26120
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26121, 26122

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175219)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorBaseEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor Base do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBaseEmissao As Double
Dim dValorComissao As Double
Dim lComprimento As Long

On Error GoTo Erro_Saida_Celula_ValorBaseEmissao

    Set objGridInt.objControle = ValorBaseEmissao

    'Verifica se valor base está preenchido
    If Len(ValorBaseEmissao.ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(ValorBaseEmissao.Text)
        If lErro <> SUCESSO Then Error 26123

        dValorBaseEmissao = CDbl(ValorBaseEmissao.Text)

        'Mostra na tela o ValorBase formatado
        GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorBase_Col) = Format(dValorBaseEmissao, "Standard")

        lComprimento = Len(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_PercentualComissao_Col))
        
        'Verifica se percentual comissao está preenchido
        If lComprimento > 0 Then

            dPercentual = PercentParaDbl(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_PercentualComissao_Col))

            'Calcula o valor da comissao
            dValorComissao = dPercentual * dValorBaseEmissao

            'Mostra na tela o valor da comissao
            GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoesEmissao.Row - GridComissoesEmissao.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26124

    'Chama SomaValor
    lErro = Soma_Valor(objGridInt)
    If lErro <> SUCESSO Then Error 26125

    Saida_Celula_ValorBaseEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorBaseEmissao:

    Saida_Celula_ValorBaseEmissao = Err

    Select Case Err

        Case 26123, 26124
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26125

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175220)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ValorEmissao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Valor do grid que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentual As Double
Dim dValorBase As Double
Dim dValorComissao As Double
Dim lComprimento As Long

On Error GoTo Erro_Saida_Celula_ValorEmissao

    Set objGridInt.objControle = ValorEmissao

    'Verifica se valor está preenchido
    If Len(ValorEmissao.ClipText) > 0 Then

        'Critica se valor base é positivo
        lErro = Valor_Positivo_Critica(ValorEmissao.Text)
        If lErro <> SUCESSO Then Error 26126

        dValorComissao = CDbl(ValorEmissao.Text)

        'Mostra na tela o Valor
        GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorComissao_Col) = Format(dValorComissao, "Standard")

        'Verifica se valor base correspondente está preenchido
        If Len(Trim(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorBase_Col))) > 0 Then

            dValorBase = CDbl(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorBase_Col))

            lComprimento = Len(Trim(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_PercentualComissao_Col)))

            'Verifica se percentual comissao correspondente está preenchido
            If lComprimento > 0 Then dPercentual = PercentParaDbl(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_PercentualComissao_Col))

            If Format(dPercentual * dValorBase, "Standard") <> dValorComissao Then

                dPercentual = dValorComissao / dValorBase

                'Mostra o percentual da comissao na tela
                GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_PercentualComissao_Col) = Format(dPercentual, "Percent")

            End If
          
        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridComissoesEmissao.Row - GridComissoesEmissao.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 26127

    'Chama Soma_Percentual
    lErro = Soma_Percentual(objGridInt)
    If lErro <> SUCESSO Then Error 26128

    'Chama Soma_Valor
    lErro = Soma_Valor(objGridInt)
    If lErro <> SUCESSO Then Error 26129

    Saida_Celula_ValorEmissao = SUCESSO

    Exit Function

Erro_Saida_Celula_ValorEmissao:

    Saida_Celula_ValorEmissao = Err

    Select Case Err

        Case 26126, 26127
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 26128, 26129

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175221)

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

Public Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tipo_Click()
    
    iAlterado = REGISTRO_ALTERADO

    If SCodigo_Extrai(Tipo.Text) = TIPODOC_FATURA_A_RECEBER Then
        NumeroFAT.Visible = True
        NumTitulo.Visible = False
        BotaoLimparFAT.Visible = True
    Else
        NumTitulo.Visible = True
        NumeroFAT.Visible = False
        BotaoLimparFAT.Visible = False
    End If

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
    If lErro <> SUCESSO And lErro <> 60483 Then Error 26138

    'Se nao encontrar -> Erro
    If lErro = 60483 Then Error 26139
    
    If SCodigo_Extrai(Tipo.Text) = TIPODOC_FATURA_A_RECEBER Then
        NumeroFAT.Visible = True
        NumTitulo.Visible = False
        BotaoLimparFAT.Visible = True
    Else
        NumTitulo.Visible = True
        NumeroFAT.Visible = False
        BotaoLimparFAT.Visible = False
    End If
    
    Exit Sub

Erro_Tipo_Validate:

    Cancel = True


    Select Case Err

        Case 26138

        Case 26139
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_CADASTRADO", Err, Tipo.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175222)

    End Select

    Exit Sub

End Sub

'Private Sub UpDownParcela_DownClick()
'
'Dim iIndice As Integer
'Dim lErro As Long
'Dim iParcela As Integer
'
'On Error GoTo Erro_UpDownParcela_DownClick
'
'    'Banco.SetFocus
'
'    DoEvents
'
'    'Verifica se existem Parcelas no Grid
'    If objGridParcelas.iLinhasExistentes = 0 Then
'        Parcela.Caption = ""
'        Exit Sub
'    End If
'
'    'Verifica se já existe alguma Parcela na Tela
'    If Len(Trim(Parcela.Caption)) > 0 Then
'        iParcela = CInt(Parcela.Caption)
'    Else
'        iParcela = 0
'    End If
'
'    'Verifica se existe uma Parcela inferior a Parcela da Tela
'    If iParcela - 1 > 0 Then
'
'        'Recolhe os dados da Parcela da Tela
'        lErro = Recolhe_Parcela_Tela(iParcela)
'        If lErro <> SUCESSO Then Error 26178
'
'        'Coloca na Tela os dados da nova parcela a tratar
'        lErro = Traz_Parcela_Tela(iParcela - 1)
'        If lErro <> SUCESSO Then Error 26179
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownParcela_DownClick:
'
'    Select Case Err
'
'        Case 26178, 26179
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175223)
'
'    End Select
'
'    Exit Sub
'
'End Sub

'Private Sub UpDownParcela_UpClick()
'
'Dim iIndice As Integer
'Dim lErro As Long
'Dim iParcela As Integer
'
'On Error GoTo Erro_UpDownParcela_UpClick
'
'   ' Banco.SetFocus
'
'    DoEvents
'
'    'Verifica se existe alguma parcela no GridParcelas
'    If objGridParcelas.iLinhasExistentes = 0 Then
'        Parcela.Caption = ""
'        Exit Sub
'    End If
'
'    'Verifica se já existe alguma Parcela na tela
'    If Len(Trim(Parcela.Caption)) > 0 Then
'        iParcela = CInt(Parcela.Caption)
'    Else
'        iParcela = 0
'    End If
'
'    'Verifica se existe a Parcela (iParcela+1)
'    If iParcela + 1 <= objGridParcelas.iLinhasExistentes Then
'
'        If iParcela <> 0 Then
'            'Recolhe os dados da parcela que está na tela
'            lErro = Recolhe_Parcela_Tela(iParcela)
'            If lErro <> SUCESSO Then Error 26382
'        End If
'
'
'        'Traz para telas os dados da Parcela seguinte
'        lErro = Traz_Parcela_Tela(iParcela + 1)
'        If lErro <> SUCESSO Then Error 26383
'
'    End If
'
'    Exit Sub
'
'Erro_UpDownParcela_UpClick:
'
'    Select Case Err
'
'        Case 26382, 26383
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175224)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Public Sub Valor_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim objComissao As ClassComissao
Dim iIndice As Integer

On Error GoTo Erro_Valor_Validate
    
    'Verifica se o Valor foi Alterado
    If iValorAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(Valor.ClipText)) <> 0 Then

        'Critica se é valor positivo
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then Error 26414
    
        'Põe o valor formatado na tela
        Valor.Text = Format(Valor.Text, "Standard")
    
    End If
    
    Call Recalcula_Parcelas_Comissoes
    
    iValorAlterado = 0
    
    Exit Sub

Erro_Valor_Validate:

    Cancel = True


    Select Case Err

        Case 41601
        
        Case 26414
            
        Case 49751
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175225)

    End Select

    Exit Sub

End Sub

Public Sub ValorBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ValorCheque_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'''Private Sub ValorCheque_Validate(Cancel As Boolean)
'''
'''Dim lErro As Long
'''
'''On Error GoTo Erro_ValorCheque_Validate
'''
'''    'Verifica se algum valor foi digitado
'''    If Len(Trim(ValorCheque.ClipText)) = 0 Then Exit Sub
'''
'''    'Critica se é valor positivo
'''    lErro = Valor_Positivo_Critica(ValorCheque.Text)
'''    If lErro <> SUCESSO Then Error 43064
'''
'''    'Põe o valor formatado na tela
'''    ValorCheque.Text = Format(ValorCheque.Text, "Standard")
'''
'''    Exit Sub
'''
'''Erro_ValorCheque_Validate:
'''
'''    Cancel = True
'''
'''
'''    Select Case Err
'''
'''        Case 43064
'''
'''        Case Else
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175226)
'''
'''    End Select
'''
'''    Exit Sub
'''
'''End Sub

Public Sub ValorComissao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorDesconto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorIRRF_Change()

    iValorIRRFAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorIRRF_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_ValorIRRF_Validate

    If iValorIRRFAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(ValorIRRF.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorIRRF.Text)
        If lErro <> SUCESSO Then Error 26144

        'Põe o valor formatado na tela
        ValorIRRF.Text = Format(ValorIRRF.Text, "Standard")

    End If
    
    Call Recalcula_Parcelas_Comissoes
    
    iValorIRRFAlterado = 0
     
    Exit Sub

Erro_ValorIRRF_Validate:

    Cancel = True


    Select Case Err

        Case 26144, 26422, 29021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175227)

    End Select

    Exit Sub

End Sub

Public Sub ValorINSS_Change()

    iAlterado = REGISTRO_ALTERADO
    iValorINSSAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ValorINSS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_ValorINSS_Validate

    If iValorINSSAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(ValorINSS.ClipText)) <> 0 Then
    
        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ValorINSS.Text)
        If lErro <> SUCESSO Then Error 26144
    
        'Põe o valor formatado na tela
        ValorINSS.Text = Format(ValorINSS.Text, "Standard")
        
    End If
    
    If INSSRetido.Value = vbChecked Then Call Recalcula_Parcelas_Comissoes
    
    iValorINSSAlterado = 0
    
    Exit Sub

Erro_ValorINSS_Validate:

    Cancel = True


    Select Case Err

        Case 26144, 49752

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175228)

    End Select

    Exit Sub

End Sub

Public Sub ValorParcela_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub ValorParcela_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub ValorParcela_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorParcela
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Vendedor_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Vendedor_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub Vendedor_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Vendedor
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub VendedorEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoesEmissao)

End Sub

Public Sub VendedorEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoesEmissao)

End Sub

Public Sub VendedorEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = VendedorEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoesEmissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub PercentualComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub PercentualComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub PercentualComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub PercentualEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoesEmissao)

End Sub

Public Sub PercentualEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoesEmissao)

End Sub

Public Sub PercentualEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = PercentualEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoesEmissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorBase_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorBase_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorBase_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBase
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorBaseEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoesEmissao)

End Sub

Public Sub ValorBaseEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoesEmissao)

End Sub

Public Sub ValorBaseEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorBaseEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoesEmissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorComissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoes)

End Sub

Public Sub ValorComissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoes)

End Sub

Public Sub ValorComissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorComissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoes)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ValorEmissao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridComissoesEmissao)

End Sub

Public Sub ValorEmissao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissoesEmissao)

End Sub

Public Sub ValorEmissao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorEmissao
    lErro = Grid_Campo_Libera_Foco(objGridComissoesEmissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub TipoDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Public Sub TipoDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Public Sub TipoDesconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = TipoDesconto
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
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

    Set objGridComissoes.objControle = DataVencimento
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Data_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Public Sub Data_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Data
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub


Public Sub ValorDesconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Public Sub ValorDesconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Public Sub ValorDesconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = ValorDesconto
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Percentual1_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridDesconto)

End Sub

Public Sub Percentual1_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridDesconto)

End Sub

Public Sub Percentual1_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Percentual1
    lErro = Grid_Campo_Libera_Foco(objGridDesconto)
    If lErro <> SUCESSO Then Cancel = True
    
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

Public Sub Filial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Filial_Click()

Dim lErro As Long

On Error GoTo Erro_Filial_Click

    iAlterado = REGISTRO_ALTERADO
       
    If Filial.ListIndex = -1 Then Exit Sub
    
    'Verifica se Cliente, Filial e Valor estão preenchidos e se o Grid de Comissões Emissão está vazio
    If Len(Trim(Cliente.ClipText)) > 0 And Len(Trim(Filial.Text)) > 0 And Len(Trim(Valor.ClipText)) > 0 And objGridComissoesEmissao.iLinhasExistentes = 0 Then
    
        'Chama Inicializa_ComissaoEmissao
        lErro = Inicializa_ComissaoEmissao()
        If lErro <> SUCESSO Then Error 26473
        
    End If
    
    Exit Sub
    
Erro_Filial_Click:

    Select Case Err
    
        Case 26473
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175229)
            
    End Select

    Exit Sub

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
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then Error 26130

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then Error 26131

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then Error 26132

        If lErro = 17660 Then Error 26133

        'Coloca na tela a Filial lida
        Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome

        'ESPECÍFICO DESSA TELA
        'Verifica se Cliente, Filial e Valor estão preenchidos e se o Grid de Comissões Emissão está vazio
        If Len(Trim(Cliente.ClipText)) > 0 And Len(Trim(Valor.ClipText)) > 0 And objGridComissoesEmissao.iLinhasExistentes = 0 Then
        
            'Chama Inicializa_ComissaoEmissao
            lErro = Inicializa_ComissaoEmissao()
            If lErro <> SUCESSO Then Error 26474
            
        End If
        'FIM ESPECÍFICO
    
    End If

    'Não encontrou a STRING
    If lErro = 6731 Then Error 26134

    Exit Sub

Erro_Filial_Validate:

    Cancel = True


    Select Case Err

        Case 26474

        Case 26130, 26132

        Case 26131
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)

        Case 26133
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

                If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            Else
            End If

        Case 26134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", Err, Filial.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175230)

    End Select

    Exit Sub

End Sub

Private Function GridParcelas_Preenche(objCondicaoPagto As ClassCondicaoPagto) As Long
'Calcula valores e datas de vencimento de Parcelas a partir da Condição de Pagamento e preenche GridParcelas

Dim lErro As Long

On Error GoTo Erro_GridParcelas_Preenche

    'Limpa o Grid de Parcelas
    Call Grid_Limpa(objGridParcelas)
    
    Call Limpa_Colecoes
    
    objGridParcelas.iLinhasExistentes = objCondicaoPagto.iNumeroParcelas

    lErro = Atualiza_ValorParcelas(objCondicaoPagto)
    If lErro <> SUCESSO Then Error 26421

    lErro = Atualiza_DatasParcelas(objCondicaoPagto)
    If lErro <> SUCESSO Then Error 26421
        
    lErro = Trata_TabParcela()
    If lErro <> SUCESSO Then Error 19140
    
    GridParcelas_Preenche = SUCESSO

    Exit Function

Erro_GridParcelas_Preenche:

    GridParcelas_Preenche = Err

    Select Case Err

        Case 19140, 26421

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175231)

    End Select

End Function

'Private Sub Vendedores_DblClick()
'
'Dim lPosicaoSeparador As Long
'Dim objVendedor As New ClassVendedor
'Dim lErro As Long
'Dim dPercentual As Double
'Dim dValor As Double
'Dim iParcela As Integer
'Dim dValorParcela As Double
'
'On Error GoTo Erro_Vendedores_DblClick
'
'    'Se a célula do Grid for a de vendedores
'    If GridComissoes.Col = iGrid_Vendedor_Col Then
'
'        'Coloca no Grid o Vendedor selecionado
'        lPosicaoSeparador = InStr(Vendedores.Text, SEPARADOR)
'        GridComissoes.TextMatrix(GridComissoes.Row, GridComissoes.Col) = Mid(Vendedores.Text, lPosicaoSeparador + 1)
'        Vendedor.Text = Mid(Vendedores.Text, lPosicaoSeparador + 1)
'
'        objVendedor.iCodigo = Codigo_Extrai(Vendedores.Text)
'
'        lErro = CF("Vendedor_Le", objVendedor)
'        If lErro <> SUCESSO And lErro <> 12582 Then Error 51340
'        If lErro <> SUCESSO Then Error 51341
'
'        iParcela = Gridparcelas.row
'
'        If iParcela = 0 Then Exit Sub
'
'        If GridComissoes.Row > objGridComissoes.iLinhasExistentes Then
'
'            objGridComissoes.iLinhasExistentes = objGridComissoes.iLinhasExistentes + 1
'            colcolComissao.Item(iParcela).colComissao.Add 0, 0, 0, 0, 0, objVendedor.iCodigo, DATA_NULA, 0, 0, 0, DATA_NULA
'
'        End If
'
'        Call Vendedor_Linha_Preenche(objVendedor, iParcela)
'
'    End If
'
'    Exit Sub
'
'Erro_Vendedores_DblClick:
'
'    Select Case Err
'
'        Case 51340
'
'        Case 51341
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175232)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub Limpa_Tela_TituloReceber()
    
Dim lErro As Long
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa os Campos da Tela
    Call Limpa_Tela(Me)
    
    Set gobjTelaProjetoInfo = New ClassTelaPRJInfo
    Set gobjTelaProjetoInfo.objUserControl = Me
    Set gobjTelaProjetoInfo.objTela = Me
    
    'Limpa os campos não limpos em Limpa_Tela
    NumeroFAT.Caption = ""
    Filial.Clear
    Tipo.Text = ""
    INSSRetido.Value = False
    CondicaoPagamento.Text = ""
    TotalPercentualComissao.Caption = ""
    TotalPercentualEmissao.Caption = ""
    TotalValorComissao.Caption = ""
    TotalValorEmissao.Caption = ""
    'Parcela.Caption = ""
        
    Natureza.PromptInclude = False
    Natureza.Text = ""
    Natureza.PromptInclude = True
    LabelNaturezaDesc.Caption = ""
    
    DataEmissao.PromptInclude = False
    DataEmissao.Text = Format(gdtDataAtual, "dd/mm/yy")
    DataEmissao.PromptInclude = True
        
    'Limpa os Grids da Tela
    Call Grid_Limpa(objGridParcelas)
    Call Grid_Limpa(objGridDesconto)
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridComissoesEmissao)
    
    ReajustePeriodicidade.ListIndex = 0
    
    'Limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade
       
    'Coloca clientes em atraso em vermelho
    If giCliAtrasoDestacar <> 0 Then Call CF("Cliente_EmAtraso_Destacar", Cliente)
        
End Sub

Private Function Recolhe_Parcela_Tela(iParcela As Integer) As Long
'Recolhe da Tela os dados da Parcela

Dim iIndice As Integer
Dim lErro As Long
Dim objDesconto As New ClassDesconto
Dim colDesconto As New colDesconto
Dim objComissao As New ClassComissao
Dim colComissao As New colComissao
Dim lTamanho As Long
Dim objVendedor As New ClassVendedor
Dim objChequepre As New ClassChequePre

On Error GoTo Erro_Recolhe_Parcela_Tela

    If iParcela <= objGridParcelas.iLinhasExistentes Then

        'Loop de armazenamento dos dados em colDesconto
        For iIndice = 1 To objGridDesconto.iLinhasExistentes
            
            Set objDesconto = New ClassDesconto
            
            'Recolhe os Dados do Grid
            If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))) > 0 Then objDesconto.iCodigo = Codigo_Extrai(GridDescontos.TextMatrix(iIndice, iGrid_TipoDesconto_Col))
            
            If objDesconto.iCodigo = VALOR_FIXO Or objDesconto.iCodigo = VALOR_ANT_DIA Or objDesconto.iCodigo = VALOR_ANT_DIA_UTIL Then
                If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_ValorDesconto_Col))) > 0 Then objDesconto.dValor = CDbl(GridDescontos.TextMatrix(iIndice, iGrid_ValorDesconto_Col))
            ElseIf objDesconto.iCodigo = Percentual Or objDesconto.iCodigo = PERC_ANT_DIA Or objDesconto.iCodigo = PERC_ANT_DIA_UTIL Then
                lTamanho = Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Percentual1_Col)))
                If lTamanho > 0 Then objDesconto.dValor = PercentParaDbl(GridDescontos.TextMatrix(iIndice, iGrid_Percentual1_Col))
            End If
            
            If Len(Trim(GridDescontos.TextMatrix(iIndice, iGrid_Data_Col))) > 0 Then
                objDesconto.dtData = CDate(GridDescontos.TextMatrix(iIndice, iGrid_Data_Col))
            Else
                objDesconto.dtData = DATA_NULA
            End If
            
            ' Adiciona em colDesconto
            colDesconto.Add objDesconto.iCodigo, objDesconto.dtData, objDesconto.dValor
        Next
        
        'Guarda em colcolDesconto
        Set colcolDesconto(iParcela).colDesconto = colDesconto
        
        'Verifica se existem comissões no GridComissoes
        If objGridComissoes.iLinhasExistentes > 0 Then
        
            'Loop de armazenamento do dados em colComissao
            For iIndice = 1 To objGridComissoes.iLinhasExistentes
                
                Set objComissao = New ClassComissao
            
                'Recolhe os Dados do GridComissao
                lTamanho = Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_PercentualComissao_Col)))
            
                If lTamanho > 0 Then objComissao.dPercentual = PercentParaDbl(GridComissoes.TextMatrix(iIndice, iGrid_PercentualComissao_Col))
                If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_ValorBase_Col))) > 0 Then objComissao.dValorBase = CDbl(GridComissoes.TextMatrix(iIndice, iGrid_ValorBase_Col))
                If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_ValorComissao_Col))) > 0 Then objComissao.dValor = CDbl(GridComissoes.TextMatrix(iIndice, iGrid_ValorComissao_Col))
                If Len(Trim(GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col))) > 0 Then
                    objVendedor.sNomeReduzido = GridComissoes.TextMatrix(iIndice, iGrid_Vendedor_Col)
                    lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
                    If lErro <> SUCESSO And lErro <> 25008 Then Error 56590
                    If lErro <> SUCESSO Then Error 56591
                    objComissao.iCodVendedor = objVendedor.iCodigo
                End If
                
                'Adiciona em colComissao
                colComissao.Add 0, giFilialEmpresa, 0, 0, 0, objComissao.iCodVendedor, DATA_NULA, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor, DATA_NULA
                
            Next
            
            'Guarda em colcolDescolto
            Set colcolComissao(iParcela).colComissao = colComissao
    
        End If
    
        'Janaina
        'Recolhe os Dados do ChequePre
'''        If Len(Trim(Banco.ClipText)) > 0 Then objChequePre.iBanco = CInt(Banco.Text)
'''        If Len(Trim(Agencia.ClipText)) > 0 Then objChequePre.sAgencia = Agencia.Text
'''        If Len(Trim(ContaCorrente.ClipText)) > 0 Then objChequePre.sContaCorrente = ContaCorrente.Text
'''        If Len(Trim(Numero.ClipText)) > 0 Then objChequePre.lNumero = CLng(Numero.Text)
'''        If Len(Trim(ValorCheque.ClipText)) > 0 Then objChequePre.dValor = CDbl(ValorCheque.Text)
'''        If Len(Trim(DataDeposito.ClipText)) > 0 Then
'''            objChequePre.dtDataDeposito = CDate(DataDeposito.Text)
'''        Else
'''            objChequePre.dtDataDeposito = DATA_NULA
'''        End If
        
'''        'Guarda em colChequePre
'''        colchequepre(iParcela).dtDataDeposito = objChequePre.dtDataDeposito
'''        colchequepre(iParcela).dValor = objChequePre.dValor
'''        colchequepre(iParcela).iBanco = objChequePre.iBanco
'''        colchequepre(iParcela).lNumero = objChequePre.lNumero
'''        colchequepre(iParcela).sAgencia = objChequePre.sAgencia
'''        colchequepre(iParcela).sContaCorrente = objChequePre.sContaCorrente
'''        colchequepre(iParcela).iFilialEmpresa = giFilialEmpresa
        'Janaina
    End If
    
    Recolhe_Parcela_Tela = SUCESSO
    
    Exit Function
    
Erro_Recolhe_Parcela_Tela:

    Recolhe_Parcela_Tela = Err
     
    Select Case Err
          
        Case 56590
        
        Case 56591
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175233)
     
    End Select
     
    Exit Function

End Function

Sub Limpa_Tab_Parcelas()
'Limpa o TabParcelas

    'Limpa a parcela
    'Parcela.Caption = " "

    'Limpa os Grids de Desconto e Comissões
    Call Grid_Limpa(objGridComissoes)
    Call Grid_Limpa(objGridDesconto)

'Janaina
    'Limpa os Campos do ChequePre
'''    Banco.Text = ""
'''    ContaCorrente.Text = ""
'''    Numero.Text = ""
'''    Agencia.Text = ""
'''    ValorCheque.Text = ""
'''    DataDeposito.PromptInclude = False
'''    DataDeposito.Text = ""
'''    DataDeposito.PromptInclude = True
'Janaina
    Exit Sub

End Sub

Private Function Traz_Parcela_Tela(iParcela As Integer) As Long
'Traz para a tela os dados da Parcela passada

Dim objDesconto As ClassDesconto, colDescontos As colDesconto
Dim colComissao As colComissao, objComissao As ClassComissao
Dim objChequepre As ClassChequePre, objVendedor As New ClassVendedor
Dim iLinha As Integer, iIndice As Integer, lSeparador As Long, lErro As Long

On Error GoTo Erro_Traz_Parcela_Tela

    If iParcela <> 0 Then

                
        Call Limpa_Tab_Parcelas
    
        'Põe a Parcela na tela
        'Parcela.Caption = iParcela
        
        If colcolDesconto.Count >= iParcela Then
        
            'Pega a coleção de descontos correspondente a Parcela
            Set colDescontos = colcolDesconto(iParcela).colDesconto
            
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
                    GridDescontos.TextMatrix(iLinha, iGrid_ValorDesconto_Col) = Format(objDesconto.dValor * colParcelaReceber(iParcela).dValor, "Standard")
                End If
                
            Next
            
            'Atribui o número de linhas existentes
            objGridDesconto.iLinhasExistentes = iLinha
        
        End If
        
        If colcolComissao.Count >= iParcela Then
        
            'Pega a coleção de Comissões correspondente a Parcela
            Set colComissao = colcolComissao(iParcela).colComissao
            
            iLinha = 0
            
            'Preenche o GridComissoes com as Comissões existentes em colComissao
            For Each objComissao In colComissao
                
                iLinha = iLinha + 1
                GridComissoes.TextMatrix(iLinha, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
                GridComissoes.TextMatrix(iLinha, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
                GridComissoes.TextMatrix(iLinha, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")
                objVendedor.iCodigo = objComissao.iCodVendedor
                lErro = CF("Vendedor_Le", objVendedor)
                If lErro <> SUCESSO And lErro <> 12582 Then Error 56592
                If lErro <> SUCESSO Then Error 56593
                
                GridComissoes.TextMatrix(iLinha, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
            
            Next
            
            'Atualiza o número de linhas existentes
            objGridComissoes.iLinhasExistentes = iLinha
        
        End If
        
        'Faz a Soma das colunas Percentual e Valor do Grid de Comissões
        Call Soma_Percentual(objGridComissoes)
        Call Soma_Valor(objGridComissoes)
    
        'Janaina
'''        'Pega os dados do ChequePre relacionado a Parcela
'''        Set objChequePre = colchequepre(iParcela)
'''
'''        'Coloca na Tela os dados do ChequePre
'''        If objChequePre.iBanco > 0 Then Banco.Text = objChequePre.iBanco
'''        If Len(Trim(objChequePre.sAgencia)) > 0 Then Agencia.Text = objChequePre.sAgencia
'''        If objChequePre.lNumero > 0 Then Numero.Text = objChequePre.lNumero
'''        If Len(Trim(objChequePre.sContaCorrente)) > 0 Then ContaCorrente.Text = objChequePre.sContaCorrente
'''        If objChequePre.dValor > 0 Then ValorCheque.Text = objChequePre.dValor
'''        DataDeposito.PromptInclude = False
'''        If objChequePre.dtDataDeposito <> DATA_NULA Then DataDeposito.Text = Format(objChequePre.dtDataDeposito, "dd/mm/yy")
'''        DataDeposito.PromptInclude = True
        'Janaina
        
    End If
    
    Traz_Parcela_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Parcela_Tela:

    Traz_Parcela_Tela = Err
     
    Select Case Err
          
        Case 56592
        
        Case 56593
            Call Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO", Err, objVendedor.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175234)
     
    End Select
     
    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 42786
    
    'Limpa a Tela
    Call Limpa_Tela_TituloReceber
    Call Limpa_Colecoes
    
    iAlterado = 0

    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 42786
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 175235)
            
    End Select
    
    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long
'Grava o Título à Receber

Dim lErro As Long, iIndice As Integer
Dim dSomaParcelas As Double, dtDataVencimento As Date
Dim dValorIRRF As Double, dValorINSSRetido As Double, dPISRetido As Double, dCOFINSRetido As Double, dCSLLRetido As Double, dISSRetido As Double
Dim dValorTitulo As Double, iParcela As Integer
Dim colComissaoEmissao As New colComissao
Dim lTamanho As Long, vbMsgRes As VbMsgBoxResult
Dim objTituloReceber As New ClassTituloReceber

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos obrigatórios estão preenchidos
    If Len(Trim(Cliente.ClipText)) = 0 Then gError 26181
    If Len(Trim(Filial.Text)) = 0 Then gError 26182
    
    If SCodigo_Extrai(Tipo.Text) <> TIPODOC_FATURA_A_RECEBER Then
        If Len(Trim(NumTitulo.ClipText)) = 0 Then gError 26183
    End If
    
    If Len(Trim(Valor.ClipText)) = 0 Then gError 26184
    If Len(Trim(Tipo.Text)) = 0 Then gError 26185
    
    'Verifica se GridParcelas foi preenchido
    If objGridParcelas.iLinhasExistentes = 0 Then gError 26186

    dSomaParcelas = 0

    'Loop no GridParcelas
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        'Verifica se DataVencimento foi preenchida
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) = 0 Then gError 26187
        dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        
        'Verifica se DataVencimento é maior ou igual a DataEmissao
        If Len(Trim(DataEmissao.ClipText)) > 0 Then
            If dtDataVencimento < CDate(DataEmissao) Then gError 26188
        End If
        
        'Verifica a ordenação das Datas de Vencimento das Parcelas
        If iIndice > 1 Then
        
            If dtDataVencimento < CDate(GridParcelas.TextMatrix(iIndice - 1, iGrid_Vencimento_Col)) Then
            
                vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_DATAVENCIMENTO_NAO_ORDENADA")
                
                'Se não confirmar, sai
                If vbMsgRes = vbNo Then gError 26189
            
            End If
            
        End If

        'Verifica se Valor da Parcela foi preenchido
        If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))) = 0 Then gError 26190
             
        'Verifica se Valor da Parcela é positivo
        lErro = Valor_Positivo_Critica(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        If lErro <> SUCESSO Then gError 26191

        'Acumula Valor Parcela em dSomaParcelas
        dSomaParcelas = dSomaParcelas + CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
    Next

    'Valor a Receber
    If Len(Trim(ValorIRRF.Text)) > 0 Then dValorIRRF = CDbl(ValorIRRF)
    If Len(Trim(ValorINSS.Text)) > 0 Then dValorINSSRetido = CDbl(IIf(INSSRetido.Value, ValorINSS.Text, 0))
    If INSSRetido.Value = 0 Then
        If Len(Trim(ValorINSS.Text)) > 0 Then If CDbl(ValorINSS) > CDbl(Valor) Then gError 26192
    End If
    If Len(Trim(PISRetido.Text)) <> 0 And IsNumeric(PISRetido.Text) Then dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(ISSRetido.Text)) <> 0 And IsNumeric(ISSRetido.Text) Then dISSRetido = CDbl(ISSRetido.Text)
    If Len(Trim(COFINSRetido.Text)) <> 0 And IsNumeric(COFINSRetido.Text) Then dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.Text)) <> 0 And IsNumeric(CSLLRetido.Text) Then dCSLLRetido = CDbl(CSLLRetido.Text)
    
    dValorTitulo = CDbl(Valor.Text) - (dValorIRRF + dPISRetido + dISSRetido + dCOFINSRetido + dCSLLRetido) - dValorINSSRetido
    If dValorTitulo <= 0 Then gError 26193
    If Format(dValorTitulo, "0.00") <> Format(dSomaParcelas, "0.00") Then gError 26194

    'Verifica se existe alguma Parcela no TabParcela
    If GridParcelas.Row > 0 Then
        'Recolhe os dados da Parcela que está na Tela
        iParcela = GridParcelas.Row
        lErro = Recolhe_Parcela_Tela(iParcela)
        If lErro <> SUCESSO Then gError 26195
    End If
    
    'Loop de Validação dos dados do GridComissoesEmissao
    For iIndice = 1 To objGridComissoesEmissao.iLinhasExistentes
    
        'Verifica se o Vendedor foi informado
        If Len(Trim(GridComissoesEmissao.TextMatrix(iIndice, iGrid_Vendedor_Col))) = 0 Then gError 26217
        
        'Verifica se o Percentual foi informado
        lTamanho = Len(Trim(GridComissoesEmissao.TextMatrix(iIndice, iGrid_PercentualComissao_Col)))
        If lTamanho = 0 Then gError 26218
        'Verifica se o valor do Percentual é positivo ou zero
        lErro = Valor_NaoNegativo_Critica(left(GridComissoesEmissao.TextMatrix(iIndice, iGrid_PercentualComissao_Col), lTamanho - 1))
        If lErro <> SUCESSO Then gError 26219
        
        'Verifica se  Valor Base e Valor Comissao foram digitados e se sÃo positivos
        If Len(Trim(GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorBase_Col))) = 0 Then gError 26220
        lErro = Valor_NaoNegativo_Critica(GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorBase_Col))
        If lErro <> SUCESSO Then gError 26221
        
        If Len(Trim(GridComissoesEmissao.TextMatrix(iIndice, iGrid_Valor_Col))) = 0 Then gError 26222
        lErro = Valor_NaoNegativo_Critica(GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorComissao_Col))
        If lErro <> SUCESSO Then gError 26223
    Next
       
    'Move os Dados da Tela Para a Memoria
    lErro = Move_Tela_Memoria(objTituloReceber, colComissaoEmissao)
    If lErro <> SUCESSO Then gError 26216
    
    objTituloReceber.dSaldo = dValorTitulo
    objTituloReceber.dtDataRegistro = gdtDataHoje
    objTituloReceber.dPercJurosDiario = gobjCRFAT.dPercJurosDiario
    objTituloReceber.dPercMulta = gobjCRFAT.dPercMulta
    
    'Valida os dados particulares de cada Parcela
    lErro = Valida_Dados_Parcelas()
    If lErro <> SUCESSO Then gError 26385
    
    'Valida os dados das comissões
    lErro = Valida_Comissoes(objTituloReceber, colComissaoEmissao)
    If lErro <> SUCESSO Then gError 26465
   
    If objTituloReceber.iReajustePeriodicidade <> 0 Then
        If objTituloReceber.dtReajusteBase = DATA_NULA Or objTituloReceber.iMoeda = 0 Then gError 32300
    End If
    
    '###############################################
    'Inserido por Wagner 10/08/2006
    lErro = gobjTelaProjetoInfo.Valida_Dados
    If lErro <> SUCESSO Then gError 181786
    '###############################################
    
    'Passa os dados para a Rotina de Gravação mais interna(grava contabilidade também)
    'Janaina
    lErro = CF("TituloReceber_Grava", objTituloReceber, colComissaoEmissao, colParcelaReceber, colcolComissao, colcolDesconto, objContabil)
    If lErro <> SUCESSO Then gError 26224
    'Janaina
    
    GL_objMDIForm.MousePointer = vbDefault
    
    If SCodigo_Extrai(Tipo.Text) = TIPODOC_FATURA_A_RECEBER Then
            
        vbMsgRes = Rotina_Aviso(vbOKOnly, "AVISO_INFORMA_NUMERO_FATURA", objTituloReceber.lNumTitulo)
    
    End If
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 32300
            lErro = Rotina_Erro(vbOKOnly, "ERRO_REAJUSTE_DADOS_INCOMPLETOS", gErr)
        
        Case 26181
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 26182
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
    
        Case 26183
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", gErr)
            
        Case 26184
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO1", gErr)
            
        Case 26185
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            
        Case 26186
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AUSENCIA_PARCELAS_GRAVAR", gErr)
            
        Case 26187
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_NAO_INFORMADA", gErr, iIndice)
        
        Case 26188
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAVENCIMENTO_PARCELA_MENOR", gErr, GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col), DataEmissao.Text, iIndice)
            
        Case 26189, 181786
            
        Case 26190
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_PARCELA_NAO_INFORMADA", gErr, iIndice)
                 
        Case 26191, 26195, 26219, 26221, 26223, 26216, 26224, 26385, 26465
        
        Case 26192
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORINSS_MAIOR", gErr, ValorINSS.Text, Valor.Text)
        
        Case 26193
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORTITULO_MENOS_IMPOSTOS", gErr)
            
        Case 26194
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_PARCELAS_DIFERENTE", gErr, dSomaParcelas, dValorTitulo)
            
        Case 26217
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_COMISSAO_NAO_INFORMADO", gErr, iIndice)
        
        Case 26218
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_COMISSAO_NAO_INFORMADO", gErr, iIndice)
        
        Case 26220
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_COMISSAO_NAO_INFORMADO", gErr, iIndice)
        
        Case 26222
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_NAO_INFORMADO", gErr, iIndice)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175236)
            
    End Select
    
    Exit Function

End Function
    
Function Move_Tela_Memoria(objTituloReceber As ClassTituloReceber, colComissaoEmissao As colComissao) As Long
'Recolhe os dados da tela para a memória

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim colDesconto As colDesconto
Dim iIndice As Integer
Dim sCclFormatada As String, iCclPreenchida As Integer
Dim sNaturezaFormatada As String
Dim iNaturezaPreenchida As Integer
Dim objProjetoInfo As Object

On Error GoTo Erro_Move_Tela_Memoria

    'Verifica se o cliente foi digitado
    If Len(Trim(Cliente.ClipText)) > 0 Then
        objCliente.sNomeReduzido = Cliente.Text
    
        'Lê o codigo através do Nome Reduzido
        lErro = CF("Cliente_Le_NomeReduzido", objCliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 26369
        
        'Não achou o Cliente --> erro
        If lErro <> SUCESSO Then gError 26370
        
        'Guarda o código no objTituloReceber
        objTituloReceber.lCliente = objCliente.lCodigo
    End If

    'Verifica se a Filial está preenchida
    If Len(Trim(Filial.Text)) > 0 Then
        objTituloReceber.iFilial = Codigo_Extrai(Filial.Text)
    End If
    
    If SCodigo_Extrai(Tipo.Text) = TIPODOC_FATURA_A_RECEBER Then
    
        If Len(Trim(NumeroFAT.Caption)) > 0 Then
            objTituloReceber.lNumTitulo = CLng(NumeroFAT.Caption)
        Else
            objTituloReceber.lNumTitulo = 0
        End If
    Else
        If Len(Trim(NumTitulo.ClipText)) > 0 Then objTituloReceber.lNumTitulo = CLng(NumTitulo.ClipText)
    End If
        
    If Len(Trim(DataEmissao.ClipText)) = 0 Then
        objTituloReceber.dtDataEmissao = DATA_NULA
    Else
        objTituloReceber.dtDataEmissao = CDate(DataEmissao.Text)
    End If
    
    'Verifica preenchimento da Condição de Pagamento
    objTituloReceber.iCondicaoPagto = CondPagto_Extrai(CondicaoPagamento)

    'Recolhe os dados restantes
    If Len(Trim(Valor.ClipText)) > 0 Then objTituloReceber.dValor = CDbl(Valor.Text)
    objTituloReceber.iNumParcelas = objGridParcelas.iLinhasExistentes
    If Len(Trim(ValorIRRF.Text)) > 0 Then objTituloReceber.dValorIRRF = CDbl(ValorIRRF.Text)
    If Len(Trim(PISRetido.ClipText)) > 0 Then objTituloReceber.dPISRetido = CDbl(PISRetido.Text)
    If Len(Trim(ISSRetido.ClipText)) > 0 Then objTituloReceber.dISSRetido = CDbl(ISSRetido.Text)
    If Len(Trim(COFINSRetido.ClipText)) > 0 Then objTituloReceber.dCOFINSRetido = CDbl(COFINSRetido.Text)
    If Len(Trim(CSLLRetido.ClipText)) > 0 Then objTituloReceber.dCSLLRetido = CDbl(CSLLRetido.Text)
    If Len(Trim(ValorINSS.Text)) > 0 Then objTituloReceber.dValorINSS = CDbl(ValorINSS.Text)
    objTituloReceber.iINSSRetido = INSSRetido.Value
    objTituloReceber.iFilialEmpresa = giFilialEmpresa
    If Len(Trim(Tipo.Text)) > 0 Then objTituloReceber.sSiglaDocumento = SCodigo_Extrai(Tipo.Text)
    
    lErro = CF("Ccl_Formata", Ccl.Text, sCclFormatada, iCclPreenchida)
    If lErro <> SUCESSO Then gError 22940

    If iCclPreenchida = CCL_PREENCHIDA Then
        objTituloReceber.sCcl = sCclFormatada
    Else
        objTituloReceber.sCcl = ""
    End If
    
    sNaturezaFormatada = String(STRING_NATMOVCTA_CODIGO, 0)
    
    'Coloca no formato do BD
    lErro = CF("Item_Formata", SEGMENTO_NATMOVCTA, Natureza.Text, sNaturezaFormatada, iNaturezaPreenchida)
    If lErro <> SUCESSO Then gError 32298
    
    objTituloReceber.sNatureza = sNaturezaFormatada

    'Recolhe as Parcelas do GridParcelas
    lErro = Move_GridParcelas_Memoria()
    If lErro <> SUCESSO Then gError 26371

    Call Atualiza_Colecoes
    
    'Verifica se existe alguma parcela na Tela e recolhe
    If GridParcelas.Row > 0 Then Call Recolhe_Parcela_Tela(GridParcelas.Row)
    
    'Passa os descontos de colColDesconto para cada uma das parcelas correspondentes
    For iIndice = 1 To colParcelaReceber.Count
    
        Set colDesconto = colcolDesconto(iIndice).colDesconto
            
        'Verifica se existe algum desconto cadastrado
        If colDesconto.Count > 0 Then
            colParcelaReceber(iIndice).iDesconto1Codigo = colDesconto(1).iCodigo
            colParcelaReceber(iIndice).dtDesconto1Ate = colDesconto(1).dtData
            colParcelaReceber(iIndice).dDesconto1Valor = colDesconto(1).dValor
            colParcelaReceber(iIndice).dtDesconto2Ate = DATA_NULA
            colParcelaReceber(iIndice).dtDesconto3Ate = DATA_NULA
            If colDesconto.Count > 1 Then
                colParcelaReceber(iIndice).iDesconto2Codigo = colDesconto(2).iCodigo
                colParcelaReceber(iIndice).dtDesconto2Ate = colDesconto(2).dtData
                colParcelaReceber(iIndice).dDesconto2Valor = colDesconto(2).dValor
                If colDesconto.Count > 2 Then
                    colParcelaReceber(iIndice).iDesconto3Codigo = colDesconto(3).iCodigo
                    colParcelaReceber(iIndice).dtDesconto3Ate = colDesconto(3).dtData
                    colParcelaReceber(iIndice).dDesconto3Valor = colDesconto(3).dValor
                End If
            End If
        End If
    Next
    
    'Recolhe os dados do GridComissoesEmissao
    lErro = Move_GridComissoesEmissao_Memoria(colComissaoEmissao)
    If lErro <> SUCESSO Then gError 26372
    
    'obtem dados de reajuste
    If ReajustePeriodicidade.ListIndex = -1 Or ReajustePeriodicidade.ListIndex = 0 Then
        objTituloReceber.iReajustePeriodicidade = 0
        objTituloReceber.iMoeda = 0
        objTituloReceber.dtReajusteBase = DATA_NULA
    Else
        objTituloReceber.iReajustePeriodicidade = ReajustePeriodicidade.ItemData(ReajustePeriodicidade.ListIndex)
        If Moeda.ListIndex = 0 Or Moeda.ListIndex = -1 Then
            objTituloReceber.iMoeda = MOEDA_REAL
        Else
            objTituloReceber.iMoeda = Codigo_Extrai(Moeda.List(Moeda.ListIndex))
        End If
        objTituloReceber.dtReajusteBase = MaskedParaDate(ReajusteBase)
    End If
    
    '##################################################
    'Inserido por Wagner 09/08/2006
    lErro = gobjTelaProjetoInfo.Move_Tela_Memoria(objProjetoInfo, PRJ_CR_TIPO_TITREC)
    If lErro <> SUCESSO Then gError 181761
    
    Set objTituloReceber.objProjetoInfo = objProjetoInfo
    '##################################################
    
    Move_Tela_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr
    
    Select Case gErr
    
        Case 26369, 22940, 32298, 181761
        
        Case 26370
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_CADASTRADO1", gErr, Cliente.Text)
            
        Case 26371, 26372
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175237)
            
    End Select
    
    Exit Function

End Function

Private Function Move_GridParcelas_Memoria() As Long
'Move para a memória os dados existentes no Grid

Dim iIndice As Integer
Dim objParcelaReceber As ClassParcelaReceber
Dim lErro As Long
     
    Set colParcelaReceber = New colParcelaReceber
    
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
    
        Set objParcelaReceber = New ClassParcelaReceber
        
        'Preenche objParcelaReceber com a linha do GridParcelas
        objParcelaReceber.iNumParcela = iIndice
        If Len(Trim((GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col)))) <> 0 Then objParcelaReceber.dtDataVencimento = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
        If Len(Trim((GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col)))) <> 0 Then objParcelaReceber.dtDataVencimentoReal = CDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))

        objParcelaReceber.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        objParcelaReceber.dValorOriginal = objParcelaReceber.dValor
        
        '#####################################################
        'INSERIDO POR WAGNER
        objParcelaReceber.iPrevisao = StrParaInt(GridParcelas.TextMatrix(iIndice, iGrid_Previsao_Col))
        objParcelaReceber.sObservacao = GridParcelas.TextMatrix(iIndice, iGrid_DescPrev_Col)
        '#####################################################
               
        'Adiciona objParcelaReceber à coleção colParcelas
        With objParcelaReceber
            colParcelaReceber.Add .lNumIntDoc, .lNumIntTitulo, .iNumParcela, .iStatus, .dtDataVencimento, .dtDataVencimentoReal, .dSaldo, .dValor, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, .iPrevisao, .sObservacao, .dValorOriginal
        End With
    Next
        
    Move_GridParcelas_Memoria = SUCESSO
    
    Exit Function
                
End Function

Function Move_GridComissoesEmissao_Memoria(colComissaoEmissao As colComissao) As Long
'Guarda os dados do GridComissoesEmissao em colComissaoEmissao

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedor As New ClassVendedor
Dim objComissao As ClassComissao
Dim lTamanho As Long

On Error GoTo Erro_Move_GridComissoesEmissao_Memoria

    'Loop de recolhimento dos dados da tela
    For iIndice = 1 To objGridComissoesEmissao.iLinhasExistentes
    
        Set objComissao = New ClassComissao
        
        'Vendedor
        objVendedor.sNomeReduzido = GridComissoesEmissao.TextMatrix(iIndice, iGrid_Vencimento_Col)
        
        If objVendedor.sNomeReduzido <> "" Then
        
            'Lê o Nome Reduzido de Vendedor
            lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
            If lErro <> SUCESSO And lErro <> 25008 Then Error 59075
            If lErro <> SUCESSO Then Error 59076
            
            objComissao.iCodVendedor = objVendedor.iCodigo
            
        End If
        
        'Pencentual
        lTamanho = Len(Trim(GridComissoesEmissao.TextMatrix(iIndice, iGrid_PercentualComissao_Col)))
        If lTamanho > 0 Then objComissao.dPercentual = PercentParaDbl(GridComissoesEmissao.TextMatrix(iIndice, iGrid_PercentualComissao_Col))
        
        'Valor Base
        objComissao.dValorBase = StrParaDbl(GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorBase_Col))
        
        'Valor Comissão
        objComissao.dValor = StrParaDbl(GridComissoesEmissao.TextMatrix(iIndice, iGrid_ValorComissao_Col))
        
        'Guarda em colComissaoEmissao
        colComissaoEmissao.Add 0, giFilialEmpresa, STATUS_LIBERADO, TIPO_TITULO_RECEBER, 0, objComissao.iCodVendedor, DATA_NULA, objComissao.dPercentual, objComissao.dValorBase, objComissao.dValor, DATA_NULA
    
    Next
    
    Move_GridComissoesEmissao_Memoria = SUCESSO
 
    Exit Function

Erro_Move_GridComissoesEmissao_Memoria:

    Move_GridComissoesEmissao_Memoria = Err
     
    Select Case Err
          
        Case 59075
        
        Case 59076
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_NAO_CADASTRADO1", Err, objVendedor.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175238)
     
    End Select
     
    Exit Function

End Function

Public Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTituloReceber As New ClassTituloReceber
Dim vbMsgRes As VbMsgBoxResult
Dim colComissaoEmissao As New colComissao

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se campos identificadores estão preenchidos
    If Len(Trim(Cliente.ClipText)) = 0 Then Error 43060
    If Len(Trim(Filial.Text)) = 0 Then Error 43061
    If Len(Trim(NumTitulo.Text)) = 0 Then Error 42776
    If Len(Trim(Tipo.Text)) = 0 Then Error 42777
    
    'Move dados da Tela para objTituloReceber e colParcelas
    lErro = Move_Tela_Memoria(objTituloReceber, colComissaoEmissao)
    If lErro <> SUCESSO Then Error 42778
    
    'Verifica se campos identificadores correspondem a Título Baixado
    lErro = CF("TituloReceberBaixado_Le_Numero", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26438 Then Error 42781
    
    'Se corresponder, erro
    If lErro = SUCESSO Then Error 42782
    
    'Lê Título Receber a partir dos campos identificadores
    lErro = CF("TituloReceber_Le_Numero", objTituloReceber)
    If lErro <> SUCESSO And lErro <> 26434 Then Error 42783
    
    'Se não encontrou  o Título à Receber --> erro
    If lErro <> SUCESSO Then Error 42784
    
    'Pede confirmação da exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_TITULORECEBER", objTituloReceber.lNumTitulo)
    
    'Se não confirmar, sai
    If vbMsgRes = vbNo Then
        GL_objMDIForm.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Exclui o Titulo (inclusive a sua parte contábil)
    lErro = CF("TituloReceber_Exclui", objTituloReceber, objContabil)
    If lErro <> SUCESSO Then Error 42785
    
    'Limpa a Tela
    Call Limpa_Tela_TituloReceber
    
    iAlterado = 0
        
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub
    
Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 42776
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMTITULO_NAO_PREENCHIDO", Err)
           
        Case 42777
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", Err)
           
        Case 42783, 42781, 42785, 42778
        
        Case 42782
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_BAIXADO_EXCLUSAO", Err, objTituloReceber.lNumTitulo, objTituloReceber.sSiglaDocumento)
            
        Case 42784
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULORECEBER_NAO_CADASTRADO1", Err, objTituloReceber.sSiglaDocumento, objTituloReceber.lNumTitulo)
        
        Case 43060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", Err)
        
        Case 43061
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", Err)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175239)
            
    End Select
    
    Exit Sub
            
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoNatureza = Nothing
    Set objEventoCcl = Nothing
    
    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoCondPagto = Nothing
    Set objEventoVendedores = Nothing
    Set objEventoVendedoresParc = Nothing
    Set objEventoTipoDocumento = Nothing
    
    'Eventos associados à contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set colParcelaReceber = Nothing
    Set colChequePre = Nothing
    Set colcolDesconto = Nothing
    Set colcolComissao = Nothing
    
    Set objGridParcelas = Nothing
    Set objGridDesconto = Nothing
    Set objGridComissoes = Nothing
    Set objGridComissoesEmissao = Nothing
    
    Set objGrid1 = Nothing
    Set objContabil = Nothing
    
    Set gobjTelaProjetoInfo = Nothing
    
   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
End Sub

Private Function Trata_TabParcela() As Long
'Faz o tratamento necessário no caso de o Tab de Parcelas ser selecionado

Dim lErro As Long
Dim iIndice As Integer
Dim colDesconto As colDesconto
'Dim objComissao As New ClassComissao
Dim colComissao As colComissao
Dim objVendedor As New ClassVendedor
Dim objCliente As New ClassCliente
Dim objTipoCliente As New ClassTipoCliente
Dim objFilialCliente As New ClassFilialCliente
Dim iParcela As Integer
Dim dValor As Double
Dim dtData As Date
Dim dtDataReal As Date, objComissaoNF As New ClassComissaoNF

On Error GoTo Erro_Trata_TabParcela
    
    If objGridParcelas.iLinhasExistentes > colParcelaReceber.Count Then
             
        lErro = CF("ComissaoAutomatica_Obter_Info", Cliente.Text, Codigo_Extrai(Filial.Text), objVendedor, objComissaoNF)
        If lErro <> SUCESSO Then Error 59016
        
        'Adiciona em colParcelaReceber as Parcelas que ainda não possuia
        For iIndice = colParcelaReceber.Count + 1 To objGridParcelas.iLinhasExistentes
    
            Set colComissao = New colComissao
            Set colDesconto = New colDesconto
            
            dtData = DATA_NULA
            dtDataReal = DATA_NULA
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))) > 0 Then
                dtData = CDate(GridParcelas.TextMatrix(iIndice, iGrid_Vencimento_Col))
                dtDataReal = CDate(GridParcelas.TextMatrix(iIndice, iGrid_VenctoReal_Col))
            End If
            
            lErro = CF("Parcela_GeraDescontoPadrao", colDesconto, dtData)
            If lErro <> SUCESSO Then Error 26171
    
            'Chama Inicializa_ComissaoParcela
            lErro = Inicializa_ComissaoParcela(iIndice, colComissao, objVendedor, objComissaoNF)
            If lErro <> SUCESSO Then Error 26180
            
            If Len(Trim(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))) > 0 Then dValor = CDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
            
            '##################################
            'ALTERADO POR WAGNER
            colParcelaReceber.Add 0, 0, iIndice, STATUS_ABERTO, dtData, dtDataReal, 0, dValor, 1, 1, 1, " ", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, "", dValor
            '##################################
            
            colcolDesconto.Add colDesconto
            colcolComissao.Add colComissao
            'Janaina
            'colchequepre.Add 0, 0, 0, 0, "", "", 0, DATA_NULA, 0, 0
            'Janaina
        Next
    Else
        Call Atualiza_Colecoes
    End If
    
    'Se não existir Parcelas então limpa-se a tab de parcelas
    If colParcelaReceber.Count = 0 Then Limpa_Tab_Parcelas
        
    'Veririca se existe alguma parcela no Tab
    If GridParcelas.Row > 0 Then
        iParcela = GridParcelas.Row
        'Verifica se a parcela da tela for maior que o número de parcelas existentes
        If iParcela > objGridParcelas.iLinhasExistentes Then iParcela = 1
        Call Traz_Parcela_Tela(iParcela)
    Else
        If colParcelaReceber.Count > 0 Then iParcela = 1
        Call Traz_Parcela_Tela(iParcela)
    End If
        
    'Chama Soma_Percentual
    lErro = Soma_Percentual(objGridComissoes)
    If lErro <> SUCESSO Then Error 26419

    'Chama Soma_Valor
    lErro = Soma_Valor(objGridComissoes)
    If lErro <> SUCESSO Then Error 26420
    
    Trata_TabParcela = SUCESSO
    
    Exit Function
    
Erro_Trata_TabParcela:

    Trata_TabParcela = Err
    
    Select Case Err
    
        Case 26171, 26180, 26419, 26420, 59016
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175240)
            
    End Select
    
    Exit Function
        
End Function

Private Function Carrega_GridParcelas() As Long
'Carrega os dados das Parcelas no GridParcelas

Dim lErro As Long
Dim iLinha As Integer
Dim objParcelaReceber As ClassParcelaReceber
'Dim objChequePre As ClassChequePre
Dim colDesc As colDesconto
Dim colCom As colComissao

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

        '############################################
        'INSERIDO POR WAGNER
        GridParcelas.TextMatrix(iLinha, iGrid_Previsao_Col) = objParcelaReceber.iPrevisao
        GridParcelas.TextMatrix(iLinha, iGrid_DescPrev_Col) = objParcelaReceber.sObservacao
        '############################################

        Set colDesc = New colDesconto
        Set colCom = New colComissao
        'Set objChequePre = New ClassChequePre
        
        'Adiciona em colDesc os descontos da Parcela
        If objParcelaReceber.iDesconto1Codigo <> 0 Then colDesc.Add objParcelaReceber.iDesconto1Codigo, objParcelaReceber.dtDesconto1Ate, objParcelaReceber.dDesconto1Valor
        If objParcelaReceber.iDesconto2Codigo <> 0 Then colDesc.Add objParcelaReceber.iDesconto2Codigo, objParcelaReceber.dtDesconto2Ate, objParcelaReceber.dDesconto2Valor
        If objParcelaReceber.iDesconto3Codigo <> 0 Then colDesc.Add objParcelaReceber.iDesconto3Codigo, objParcelaReceber.dtDesconto3Ate, objParcelaReceber.dDesconto3Valor

        'Lê as Comissoes relacionadas a Parcela
        lErro = CF("Comissoes_Le", objParcelaReceber.lNumIntDoc, colCom, TIPO_PARCELA)
        If lErro <> SUCESSO Then Error 26161

        'Janaina
        'objChequePre.dtDataDeposito = DATA_NULA
        
'''        'Verifica se existe um ChequePre Relacionado à Parcela
'''        If objParcelaReceber.lNumIntCheque <> 0 Then
'''            objChequePre.lNumIntCheque = objParcelaReceber.lNumIntCheque
'''
'''            'Lê o ChequePre
'''            lErro = CF("ChequePre_Le",objChequePre)
'''            If lErro <> SUCESSO And lErro <> 17642 Then Error 26162
'''            If lErro = 17642 Then Error 26163
'''
'''        End If
        'Janaina
        
        'Adiciona nas coleções globais os demais dados das Parcelas
        colcolDesconto.Add colDesc
        colcolComissao.Add colCom
        'colchequepre.Add objChequePre.lNumIntCheque, objChequePre.lCliente, objChequePre.iFilial, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero, objChequePre.dtDataDeposito, objChequePre.dValor, objChequePre.lNumBordero

    Next
       
    objGridParcelas.iLinhasExistentes = iLinha
    
    '#################################################
    'INSERIDO POR WAGNER
    Call Grid_Refresh_Checkbox(objGridParcelas)
    '#################################################
    
    Carrega_GridParcelas = SUCESSO

    Exit Function
    
Erro_Carrega_GridParcelas:

    Carrega_GridParcelas = Err
    
    Select Case Err
    
        Case 26161, 26162
        
        Case 26163
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_NAO_CADASTRADO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175241)
            
    End Select
    
    Exit Function
            
End Function

Private Function Valida_Dados_Parcelas() As Long
'Valida os Dados da Parcela para a Gravação

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objComissao As ClassComissao
Dim objDesconto As ClassDesconto
Dim dDesconto As Double
Dim colDesconto As colDesconto
Dim colComissao As colComissao
'Dim objChequePre As ClassChequePre

On Error GoTo Erro_Valida_Dados_Parcelas

    'Para cada Parcela existente
    For iIndice = 1 To objGridParcelas.iLinhasExistentes

        Set colComissao = colcolComissao(iIndice).colComissao
        
        iIndice2 = 0
        
        'Loop de validação os dados das comissões
        For Each objComissao In colComissao

            iIndice2 = iIndice2 + 1
            If objComissao.iCodVendedor = 0 Then Error 26196
            If objComissao.dPercentual < 0 Then Error 26197
            If objComissao.dValorBase < 0 Then Error 26198
            If objComissao.dValor < 0 Then Error 26199

        Next
            
        Set colDesconto = colcolDesconto(iIndice).colDesconto
        
        iIndice2 = 0
        
        'Loop de validação dos descontos
        For Each objDesconto In colDesconto
    
            iIndice2 = iIndice2 + 1
            If objDesconto.dtData = DATA_NULA Then Error 26200
            If objDesconto.dValor <= 0 Then Error 26201
            If objDesconto.iCodigo <= 0 Then Error 26202
            If objDesconto.dtData > colParcelaReceber(iIndice).dtDataVencimentoReal Then Error 59147
            
            If iIndice2 > 1 Then If objDesconto.dtData <= colDesconto(iIndice2 - 1).dtData Then Error 18982
            
        Next
        'Janaina
'''        Set objChequePre = colchequepre(iIndice)

        'Valida os dados do ChequePre
'''        If objChequePre.dtDataDeposito <> DATA_NULA Or objChequePre.dValor > 0 Or objChequePre.iBanco > 0 Or objChequePre.lNumero > 0 Or Len(Trim(objChequePre.sAgencia)) > 0 Or Len(Trim(objChequePre.sContaCorrente)) > 0 Then
'''            If objChequePre.dtDataDeposito = DATA_NULA Then Error 26203
'''            If objChequePre.dValor <= 0 Then Error 26204
'''            If objChequePre.iBanco <= 0 Then Error 26205
'''            If objChequePre.lNumero <= 0 Then Error 26206
'''            If Len(Trim(objChequePre.sAgencia)) = 0 Then Error 26207
'''            If Len(Trim(objChequePre.sContaCorrente)) = 0 Then Error 26208
'''
'''            'Calcula o valor da Parcela como desconto aplicado
'''            lErro = CF("Parcela_Com_Desconto",colDesconto, colParcelaReceber(iIndice).dValor, objChequePre.dtDataDeposito, colParcelaReceber(iIndice).dtDataVencimento, dDesconto)
'''            If lErro <> SUCESSO Then Error 26214

'''            'Verifica se o valor do ChequePre é igual ao valor da Parcela com o desconto aplicado
'''            If Format(objChequePre.dValor, "0.00") <> Format(dDesconto, "0.00") Then Error 26215
'''            'Janaina
'''
'''            If objChequePre.dtDataDeposito <> colParcelaReceber(iIndice).dtDataVencimento And objChequePre.dtDataDeposito <> colParcelaReceber(iIndice).dtDataVencimentoReal And Format(colParcelaReceber(iIndice).dValor, "0.00") = Format(objChequePre.dValor, "0.00") Then Error 56589
'''
'''        End If
'''
    Next

    Valida_Dados_Parcelas = SUCESSO
    
    Exit Function
    
Erro_Valida_Dados_Parcelas:

    Valida_Dados_Parcelas = Err
    
    Select Case Err
    
        Case 18982
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATAS_DESCONTOS_DESORDENADAS", Err, iIndice)
    
        Case 26196
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)
    
        Case 26197
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PERCENTUAL_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)
            
        Case 26198
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORBASE_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)
            
        Case 26199
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_COMISSAO_PARCELA_NAO_INFORMADO", Err, iIndice2, iIndice)
        
        Case 26200
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESCONTO_PARCELA_NAO_PREENCHIDA", Err, iIndice2, iIndice)
                  
        Case 26201
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_DESCONTO_PARCELA_NAO_PREENCHIDO", Err, iIndice2, iIndice)
            
        Case 26202
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_DESCONTO_PARCELA_NAO_PREENCHIDO", Err, iIndice2, iIndice)
            
        Case 26203
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATADEPOSITO_CHQPRE_PARCELA_NAO_PREENCHIDA", Err, iIndice)
        
        Case 26204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_CHQPRE_PARCELA_NAO_PREENCHIDO", Err, iIndice)
        
        Case 26205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BANCO_CHQPRE_PARCELA_NAO_PREENCHIDO", Err, iIndice)
        
        Case 26206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMERO_CHQPRE_PARCELA_NAO_PREENCHIDO", Err, iIndice)
            
        Case 26207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_AGENCIA_CHQPRE_PARCELA_NAO_PREENCHIDA", Err, iIndice)
            
        Case 26208
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CONTA_CHQPRE_PARCELA_NAO_PREENCHIDA", Err, iIndice)
            
        Case 26214
        
        Case 26215
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALOR_CHQPRE_DIFERENTE_DESCONTO", Err, iIndice, Format(dDesconto, "0.00"))
        
        Case 56589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_CHQPRE_DIFERENTE", Err, iIndice)
        
        Case 59147
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_DESC_APOS_VCTO", Err, iIndice)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175242)
            
    End Select
    
    Exit Function
    
End Function

Private Function Valida_Comissoes(objTituloReceber As ClassTituloReceber, colComissaoEmissao As colComissao) As Long
'Verifica se a soma dos valores de todas as comissoes (tanto da emissao como das parcelas) é menor que o Valor do Título.
'e verifica também, para cada parcela, se a soma das comissões é menor do que o valor da parcela

Dim lErro As Long
Dim iIndice As Integer
Dim dComissoesEmissao As Double
Dim dComissoesParcela As Double
Dim dComissoesTotal As Double
Dim objComissao As ClassComissao
Dim colComissao As colComissao
Dim dValorTotal As Double

On Error GoTo Erro_Valida_Comissoes

    dComissoesEmissao = 0
    dComissoesParcela = 0

    'Acumula os valores das comissões da Emissão
    For Each objComissao In colComissaoEmissao
        dComissoesEmissao = dComissoesEmissao + objComissao.dValor
    Next
    
    dComissoesTotal = dComissoesEmissao
    
    'Loop que soma as comissoes das Parcelas e Acumula na Comissão total
    For iIndice = 1 To colParcelaReceber.Count
    
        Set colComissao = colcolComissao(iIndice).colComissao
        
        dComissoesParcela = 0
           
        'Loop que acumula os valores das comissões para cada parcela
        For Each objComissao In colComissao
            dComissoesParcela = dComissoesParcela + objComissao.dValor
        Next
        
        'Verifica se o valor das comissões é menor que o valor da Parcela
        If dComissoesParcela >= colParcelaReceber(iIndice).dValor Then Error 26466
        
        'Acumula o valor de todas as Comissões
        dComissoesTotal = dComissoesTotal + dComissoesParcela
        
    Next
     
    dValorTotal = objTituloReceber.dValor + IIf(objTituloReceber.iINSSRetido, objTituloReceber.dValorINSS, 0)

    'Verifica se o valor de todas as Comissões é menor que o valor do Título
    If dComissoesTotal >= dValorTotal Then Error 26467
    
    Valida_Comissoes = SUCESSO
    
    Exit Function
    
Erro_Valida_Comissoes:

    Valida_Comissoes = Err
    
    Select Case Err
    
        Case 26466
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_COMISSOES_PARCELA", Err, iIndice)
    
        Case 26467
            lErro = Rotina_Erro(vbOKOnly, "ERRO_SOMA_COMISSOES_EMISSAO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175243)
            
    End Select

    Exit Function

End Function

Private Function Inicializa_ComissaoEmissao() As Long
'Inicializa as Comissões de Emissão

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim objComissao As New ClassComissao, iCodFilial As Integer

On Error GoTo Erro_Inicializa_ComissaoEmissao

    If Len(Trim(Valor.Text)) <> 0 Then
    
        iCodFilial = Codigo_Extrai(Filial.Text)
        lErro = CF("ComissaoAutomatica_Obter_Emissao", Cliente.Text, iCodFilial, objVendedor, objComissao, StrParaDbl(Valor.Text), 0, 0, 0, 0)
        If lErro <> SUCESSO Then Error 59014
        
        'Limpa o Grid
        Call Grid_Limpa(objGridComissoesEmissao)
        
        'Mostra na tela o Vendedor,Percentual Comissão, Valor Base e Valor Comissão
        If objVendedor.sNomeReduzido <> "" Then
        
            GridComissoesEmissao.TextMatrix(1, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
            GridComissoesEmissao.TextMatrix(1, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
            GridComissoesEmissao.TextMatrix(1, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
            GridComissoesEmissao.TextMatrix(1, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")
            
            objGridComissoesEmissao.iLinhasExistentes = 1
            
        End If
        
        'Chama Soma_Percentual
        lErro = Soma_Percentual(objGridComissoesEmissao)
        If lErro <> SUCESSO Then Error 26476
        
        'Chama Soma_Valor
        lErro = Soma_Valor(objGridComissoesEmissao)
        If lErro <> SUCESSO Then Error 26477
    
    End If
    
    Inicializa_ComissaoEmissao = SUCESSO
    
    Exit Function
    
Erro_Inicializa_ComissaoEmissao:

    Inicializa_ComissaoEmissao = Err
    
    Select Case Err
    
        Case 26476, 26477, 59014
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175244)
            
    End Select
        
    Exit Function
    
End Function

Private Sub Atualiza_Colecoes()
'Ajusta as colecoes globais para terem os elementos de acordo com o numero de parcelas do grid do 1o frame, acrescentando ou removendo dados

Dim iIndice As Integer
Dim colComissao As colComissao
Dim colDesconto As colDesconto

'Janaina
'''    'Iguala o número de ChequesPre ao Número de Parcelas existentes
'''    If colParcelaReceber.Count > colchequepre.Count Then
'''        For iIndice = colchequepre.Count + 1 To colParcelaReceber.Count
'''            colchequepre.Add 0, 0, 0, 0, "", "", 0, DATA_NULA, 0, 0
'''        Next
'''    ElseIf colParcelaReceber.Count < colchequepre.Count Then
'''        Do While colParcelaReceber.Count < colchequepre.Count
'''            colchequepre.Remove (colchequepre.Count)
'''        Loop
'''    End If
'Janaina
                
    'Iguala o número de coleções de Comissões ao número de Parcelas
    If colcolComissao.Count < colParcelaReceber.Count Then
        For iIndice = colcolComissao.Count + 1 To colParcelaReceber.Count
            Set colComissao = New colComissao
            colcolComissao.Add colComissao
        Next
    ElseIf colcolComissao.Count > colParcelaReceber.Count Then
        Do While colcolComissao.Count > colParcelaReceber.Count
            colcolComissao.Remove (colcolComissao.Count)
        Loop
    End If
    
    'Iguala o número de coleções de Descontos ao número de Parcelas
    If colcolDesconto.Count < colParcelaReceber.Count Then
        For iIndice = colcolDesconto.Count + 1 To colParcelaReceber.Count
            Set colDesconto = New colDesconto
            colcolDesconto.Add colDesconto
        Next
    ElseIf colcolDesconto.Count > colParcelaReceber.Count Then
        Do While colcolDesconto.Count > colParcelaReceber.Count
            colcolDesconto.Remove (colcolDesconto.Count)
        Loop
    End If
    
    Exit Sub
    
End Sub

Private Sub Limpa_Colecoes()

    'Reinicializa as coleções
    Set colParcelaReceber = New colParcelaReceber
    Set colcolComissao = New colcolComissao
    Set colcolDesconto = New colcolDesconto
    'Janaina
    'Set colchequepre = New colchequepre
    'Janaina
    
    Exit Sub

End Sub

Private Function Trata_Mudanca_Frame() As Long

Dim lErro As Long
Dim lPosicaoSeparador As Long

On Error GoTo Erro_Trata_Mudanca_Frame
 
    'Se o Tab selecionado foi de Parcelas
    If Opcao.SelectedItem.Index = TAB_Parcelas Then

        'Cham Trata_TabParcela
        lErro = Trata_TabParcela()
        If lErro <> SUCESSO Then Error 26378
    Else
        
        If GridParcelas.Row <> 0 Then
            lErro = Recolhe_Parcela_Tela(GridParcelas.Row)
            If lErro <> SUCESSO Then Error 15991
        End If
        
    End If
    
    Trata_Mudanca_Frame = SUCESSO
    
    Exit Function
    
Erro_Trata_Mudanca_Frame:

    Trata_Mudanca_Frame = Err

    Select Case Err
    
        Case 15991, 26378
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175245)
            
    End Select
    
    Exit Function
    
End Function

'Início Contabilidade

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

'****
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
Dim objCliente As New ClassCliente, objTipoCliente As New ClassTipoCliente
Dim objFilial As New ClassFilialCliente
Dim iNumParcelas As Integer, sContaTela As String
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
        
        Case CLIENTE_COD
            
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 39583
                
                objMnemonicoValor.colValor.Add objCliente.lCodigo
                
            Else
                
                objMnemonicoValor.colValor.Add 0
                
            End If
            
        Case CLIENTE_NOME
        
            'Preenche NomeReduzido com o Cliente da tela
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 39584
            
                objMnemonicoValor.colValor.Add objCliente.sRazaoSocial
        
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
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 39585
                
                objMnemonicoValor.colValor.Add objFilial.sNome
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case FILIAL_CONTA
            
            If Len(Filial.Text) > 0 Then
                
                objFilial.iCodFilial = Codigo_Extrai(Filial.Text)
                lErro = CF("FilialCliente_Le_NomeRed_CodFilial", Cliente.Text, objFilial)
                If lErro <> SUCESSO Then Error 39586
                
                If objFilial.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objFilial.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 56500
                
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
                If lErro <> SUCESSO Then Error 39587
                
                objMnemonicoValor.colValor.Add objFilial.sCgc
            
            Else
                
                objMnemonicoValor.colValor.Add ""
            
            End If
            
        Case TIPOCLI_CONTA
            If Len(Trim(Cliente.Text)) > 0 Then
                
                objCliente.sNomeReduzido = Cliente.Text
                lErro = CF("Cliente_Le_NomeReduzido", objCliente)
                If lErro <> SUCESSO Then Error 56501
                
                objTipoCliente.iCodigo = objCliente.iTipo
                lErro = CF("TipoCliente_Le", objTipoCliente)
                If lErro <> SUCESSO Then Error 56502
                
                If objTipoCliente.sContaContabil <> "" Then
                
                    lErro = Mascara_RetornaContaTela(objTipoCliente.sContaContabil, sContaTela)
                    If lErro <> SUCESSO Then Error 41971
                
                Else
                
                    sContaTela = ""
                    
                End If
                
                objMnemonicoValor.colValor.Add sContaTela
                
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
            If Len(DataEmissao.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(DataEmissao.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If
            
        Case VALOR1
            If Len(Trim(Valor.Text)) > 0 Then
                objMnemonicoValor.colValor.Add CDbl(Valor.Text)
            Else
                objMnemonicoValor.colValor.Add 0
            End If
            
        Case TIPO1
            If Len(Tipo.Text) > 0 Then
                objMnemonicoValor.colValor.Add SCodigo_Extrai(Tipo.Text)
            Else
                objMnemonicoValor.colValor.Add ""
            End If
                    
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
                
        Case INSS_RETIDO
            objMnemonicoValor.colValor.Add INSSRetido.Value
        
        Case Else
            Error 39588
            
    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err
            
        Case 39583, 39584, 39585, 39586, 39587, 56500, 56501, 56502
        
        Case 39588
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175246)
            
    End Select
        
    Exit Function
    
End Function

Sub Rotina_Grid_Enable(iLinha As Integer, objControle As Object, iCaminho As Integer)
   
Dim iTipo As Integer

    'Pesquisa a controle da coluna em questão
    Select Case objControle.Name

        Case Data.Name
        
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then
                Data.Enabled = True
            Else
                Data.Enabled = False
            End If
        
        Case ValorDesconto.Name
        
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then
                    
                iTipo = Codigo_Extrai(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))
                
                If iTipo = VALOR_ANT_DIA Or iTipo = VALOR_ANT_DIA_UTIL Or iTipo = VALOR_FIXO Then
                    objControle.Enabled = True
                Else
                    objControle.Enabled = False
                End If
            Else
                objControle.Enabled = False
            End If
                
        
        Case Percentual1.Name
                
            If Len(Trim(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))) > 0 Then
                
                iTipo = Codigo_Extrai(GridDescontos.TextMatrix(GridDescontos.Row, iGrid_TipoDesconto_Col))
                
                If iTipo = PERC_ANT_DIA Or iTipo = PERC_ANT_DIA_UTIL Or iTipo = Percentual Then
                    objControle.Enabled = True
                Else
                    objControle.Enabled = False
                End If
            Else
                objControle.Enabled = False
            End If
    
    End Select
    
    Exit Sub

End Sub

Function Carrega_Colecoes_Parcela() As Long

Dim lErro As Long
Dim objParcelaRec As ClassParcelaReceber
Dim colComissao As colComissao
Dim objChequepre As New ClassChequePre

On Error GoTo Erro_Carrega_Colecoes_Parcela

    For Each objParcelaRec In colParcelaReceber
        
        Set colComissao = New colComissao
            
        lErro = CF("Comissoes_Le", objParcelaRec.lNumIntDoc, colComissao, TIPO_PARCELA)
        If lErro <> SUCESSO Then Error 46548
        
        colcolComissao.Add colComissao
        
            'Janaina
''''        If objParcelaRec.lNumIntCheque <> 0 Then
''''
''''            objChequePre.lNumIntCheque = objParcelaRec.lNumIntCheque
''''
''''            lErro = CF("ChequePre_Le",objChequePre)
''''            If lErro <> SUCESSO And lErro <> 17642 Then Error 46549
''''            If lErro <> SUCESSO Then Error 46550
''''
''''            colchequepre.Add objChequePre.lNumIntCheque, objChequePre.lCliente, objChequePre.iFilial, objChequePre.iBanco, objChequePre.sAgencia, objChequePre.sContaCorrente, objChequePre.lNumero, objChequePre.dtDataDeposito, objChequePre.dValor, objChequePre.lNumBordero
''''        Else
''''            colchequepre.Add 0, 0, 0, 0, "", "", 0, DATA_NULA, 0, 0
''''        End If
            'Janaina
    Next
    
    Carrega_Colecoes_Parcela = SUCESSO
    
    Exit Function
    
Erro_Carrega_Colecoes_Parcela:

    Carrega_Colecoes_Parcela = Err
    
    Select Case Err
    
        Case 46548, 46549
        
        Case 46550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CHEQUEPRE_NAO_CADASTRADO", Err, objChequepre.lNumIntCheque)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175247)
            
    End Select
    
    Exit Function
    
End Function

Private Sub Recalcula_Parcelas_Comissoes()
'gera novas parcelas e comissoes, inclusive comissoes na emissao

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Recalcula_Parcelas_Comissoes

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 And Len(Trim(Valor.Text)) <> 0 Then
    
        'Passa o código da Condição para objCondicaoPagto
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
    
        'Lê Condição a partir do código
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 56740
        If lErro = 19205 Then Error 56741
    
        lErro = GridParcelas_Preenche(objCondicaoPagto)
        If lErro <> SUCESSO Then Error 56738
        
        'Verifica se Cliente, Filial e Valor estão preenchidos
        If Len(Trim(Cliente.ClipText)) > 0 And Len(Trim(Filial.Text)) > 0 Then
        
            'Chama Inicializa_ComissaoEmissao
            lErro = Inicializa_ComissaoEmissao()
            If lErro <> SUCESSO Then Error 26475
            
        End If
        
    End If
    
    Exit Sub
     
Erro_Recalcula_Parcelas_Comissoes:

    Select Case Err
          
        Case 26475, 56738, 56739, 56740
        
        Case 56741
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175248)
     
    End Select
     
    Exit Sub

End Sub

Private Sub Recalcula_Datas_Parcelas()
'recalcula novas datas p/parcelas

Dim lErro As Long
Dim objCondicaoPagto As New ClassCondicaoPagto

On Error GoTo Erro_Recalcula_Datas_Parcelas

    If CobrancaAutomatica.Value = vbChecked And Len(Trim(CondicaoPagamento.Text)) <> 0 And Len(Trim(Valor.Text)) <> 0 Then
    
        'Passa o código da Condição para objCondicaoPagto
        objCondicaoPagto.iCodigo = CondPagto_Extrai(CondicaoPagamento)
    
        'Lê Condição a partir do código
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then Error 56742
        If lErro = 19205 Then Error 56743
    
        lErro = Atualiza_DatasParcelas(objCondicaoPagto)
        If lErro <> SUCESSO Then Error 56734
        
    End If
    
    Exit Sub
     
Erro_Recalcula_Datas_Parcelas:

    Select Case Err
          
        Case 56742, 56744
        
        Case 56743
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_CADASTRADA", Err, objCondicaoPagto.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175249)
     
    End Select
     
    Exit Sub

End Sub

Private Function VendedorEmissao_Linha_Preenche(objVendedor As ClassVendedor) As Long
'gera defaults p/linha de grid de comissoes na emissao apos preenchimento do vendedor

Dim lErro As Long, iCodFilial As Integer, objComissao As New ClassComissao
Dim iIndice As Integer

On Error GoTo Erro_VendedorEmissao_Linha_Preenche

    If UCase(GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_Vendedor_Col)) <> UCase(objVendedor.sNomeReduzido) Then
    
        'Loop no GridComissoes
        For iIndice = 1 To objGridComissoesEmissao.iLinhasExistentes
    
            'Verifica se Vendedor comparece em outra linha
            If iIndice <> GridComissoesEmissao.Row Then If UCase(GridComissoesEmissao.TextMatrix(iIndice, iGrid_Vendedor_Col)) = UCase(objVendedor.sNomeReduzido) Then Error 49748
    
        Next
    
        iCodFilial = Codigo_Extrai(Filial.Text)
        lErro = CF("ComissaoAutomatica_Obter_Emissao", Cliente.Text, iCodFilial, objVendedor, objComissao, StrParaDbl(Valor.Text), 0, 0, 0, 0)
        If lErro <> SUCESSO Then Error 59017
        
        If objComissao.iCodVendedor <> 0 Then
                        
            GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
            VendedorEmissao.Text = objVendedor.sNomeReduzido
            
            GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
            GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
            GridComissoesEmissao.TextMatrix(GridComissoesEmissao.Row, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")
            
            If GridComissoesEmissao.Row > objGridComissoesEmissao.iLinhasExistentes Then objGridComissoesEmissao.iLinhasExistentes = objGridComissoesEmissao.iLinhasExistentes + 1
        
            Call Soma_Percentual(objGridComissoesEmissao)
            Call Soma_Valor(objGridComissoesEmissao)
    
        End If
        
    End If
    
    VendedorEmissao_Linha_Preenche = SUCESSO
    
    Exit Function
     
Erro_VendedorEmissao_Linha_Preenche:

    VendedorEmissao_Linha_Preenche = Err
    
    Select Case Err
          
        Case 49748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VENDEDOR_JA_EXISTENTE", Err, objVendedor.sNomeReduzido)
    
        Case 59017
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175250)
     
    End Select
     
    Exit Function

End Function

Sub Vendedor_Linha_Preenche(objVendedor As ClassVendedor, iParcela As Integer)
'gera defaults p/linha de grid de comissoes na baixa de parcela apos preenchimento do vendedor

Dim lErro As Long
Dim dValorParcela As Double, iCodFilial As Integer
Dim dPercentual As Double, objComissao As New ClassComissao
Dim objComissaoNF As New ClassComissaoNF

On Error GoTo Erro_Vendedor_Linha_Preenche

    iCodFilial = Codigo_Extrai(Filial.Text)
    lErro = CF("ComissaoAutomatica_Obter_Info", Cliente.Text, iCodFilial, objVendedor, objComissaoNF)
    If lErro <> SUCESSO Then Error 59018
    
    If objComissaoNF.iCodVendedor <> 0 Then
        
        lErro = CF("ComissaoAutomatica_Obter_Parcela", objVendedor, StrParaDbl(GridParcelas.TextMatrix(iParcela, iGrid_Valor_Col)), objComissaoNF, objComissao)
        If lErro <> SUCESSO Then Error 59019
    
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col) = Format(objComissao.dPercentual, "Percent")
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col) = Format(objComissao.dValorBase, "Standard")
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = Format(objComissao.dValor, "Standard")
    
    Else
    
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_PercentualComissao_Col) = ""
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorBase_Col) = ""
        GridComissoes.TextMatrix(GridComissoes.Row, iGrid_ValorComissao_Col) = ""
    
    End If

    Call Soma_Percentual(objGridComissoes)
    Call Soma_Valor(objGridComissoes)
    
    Exit Sub
     
Erro_Vendedor_Linha_Preenche:

    Select Case Err
              
        Case 59018, 59019
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175251)
     
    End Select
     
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TITULOS_RECEBER_ID
    Set Form_Load_Ocx = Me
    Caption = "Títulos a Receber"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TituloReceber"
    
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
        ElseIf Me.ActiveControl Is Ccl Then
            Call CclLabel_Click
        ElseIf Me.ActiveControl Is NumTitulo Then
            Call NumeroLabel_Click
        ElseIf Me.ActiveControl Is Tipo Then
            Call TipoDocumentoLabel_Click
        ElseIf Me.ActiveControl Is VendedorEmissao Then
            Call BotaoVendedores_Click
        ElseIf Me.ActiveControl Is CondicaoPagamento Then
            Call CondPagtoLabel_Click
        '#########################################
        'Inserido por Wagner 10/08/2006
        ElseIf Me.ActiveControl Is Projeto Then
            Call LabelProjeto_Click
        ElseIf Me.ActiveControl Is Vendedor Then
            Call BotaoVendedoresParc_Click
        '#########################################
        End If
    
    End If
    
End Sub





































































'Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
'   'Call Controle_DragDrop(Label3, Source, X, Y)
'End Sub
'
'Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   'Call Controle_MouseDown(Label3, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
'   'Call Controle_DragDrop(Label4, Source, X, Y)
'End Sub
'
'Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   'Call Controle_MouseDown(Label4, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
'   'Call Controle_DragDrop(Label6, Source, X, Y)
'End Sub
'
'Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   'Call Controle_MouseDown(Label6, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
'   'Call Controle_DragDrop(Label7, Source, X, Y)
'End Sub
'
'Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  ' Call Controle_MouseDown(Label7, Button, Shift, X, Y)
'End Sub

'Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
'  'Call Controle_DragDrop(Label10, Source, X, Y)
'End Sub
'
'Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   'Call Controle_MouseDown(Label10, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
'   'Call Controle_DragDrop(Label11, Source, X, Y)
'End Sub
'
'Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   'Call Controle_MouseDown(Label11, Button, Shift, X, Y)
'End Sub







'Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label25, Source, X, Y)
'End Sub
'
'Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label25, Button, Shift, X, Y)
'End Sub
'
'Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Label9, Source, X, Y)
'End Sub
'
'Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
'End Sub
'
'Private Sub Parcela_DragDrop(Source As Control, X As Single, Y As Single)
'   Call Controle_DragDrop(Parcela, Source, X, Y)
'End Sub
'
'Private Sub Parcela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'   Call Controle_MouseDown(Parcela, Button, Shift, X, Y)
'End Sub


Public Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Public Sub PISRetido_Change()

    iPISRetidoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub PISRetido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_PISRetido_Validate

    If iPISRetidoAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(PISRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(PISRetido.Text)
        If lErro <> SUCESSO Then Error 26144

        'Põe o valor formatado na tela
        PISRetido.Text = Format(PISRetido.Text, "Standard")

    End If
    
    Call Recalcula_Parcelas_Comissoes
    
    iPISRetidoAlterado = 0
     
    Exit Sub

Erro_PISRetido_Validate:

    Cancel = True

    Select Case Err

        Case 26144, 26422, 29021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175252)

    End Select

    Exit Sub

End Sub

Public Sub COFINSRetido_Change()

    iCOFINSRetidoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub COFINSRetido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_COFINSRetido_Validate

    If iCOFINSRetidoAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(COFINSRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(COFINSRetido.Text)
        If lErro <> SUCESSO Then Error 26144

        'Põe o valor formatado na tela
        COFINSRetido.Text = Format(COFINSRetido.Text, "Standard")

    End If
    
    Call Recalcula_Parcelas_Comissoes
    
    iCOFINSRetidoAlterado = 0
     
    Exit Sub

Erro_COFINSRetido_Validate:

    Cancel = True

    Select Case Err

        Case 26144, 26422, 29021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175253)

    End Select

    Exit Sub

End Sub

Public Sub CSLLRetido_Change()

    iCSLLRetidoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CSLLRetido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_CSLLRetido_Validate

    If iCSLLRetidoAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(CSLLRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(CSLLRetido.Text)
        If lErro <> SUCESSO Then Error 26144

        'Põe o valor formatado na tela
        CSLLRetido.Text = Format(CSLLRetido.Text, "Standard")

    End If
    
    Call Recalcula_Parcelas_Comissoes
    
    iCSLLRetidoAlterado = 0
     
    Exit Sub

Erro_CSLLRetido_Validate:

    Cancel = True

    Select Case Err

        Case 26144, 26422, 29021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175254)

    End Select

    Exit Sub

End Sub

'############################################################
'INSERIDO POR WAGNER
Public Sub DescPrev_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub DescPrev_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub DescPrev_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub DescPrev_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = DescPrev
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Previsao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Previsao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridParcelas)

End Sub

Public Sub Previsao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridParcelas)

End Sub

Public Sub Previsao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissoes.objControle = Previsao
    lErro = Grid_Campo_Libera_Foco(objGridParcelas)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Function Saida_Celula_DescPrev(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DescPrev

    Set objGridInt.objControle = DescPrev

    GridParcelas.TextMatrix(GridParcelas.Row, iGrid_DescPrev_Col) = DescPrev.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129448

    Saida_Celula_DescPrev = SUCESSO

    Exit Function

Erro_Saida_Celula_DescPrev:

    Saida_Celula_DescPrev = gErr

    Select Case gErr

        Case 129448
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 175255)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Previsao(objGridInt As AdmGrid) As Long
'Faz a crítica da celula Previsao do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer
Dim colDesconto As New colDesconto
Dim colComissao As New colComissao

On Error GoTo Erro_Saida_Celula_Previsao

    Set objGridInt.objControle = Previsao

    'Verifica se valor está preenchido
    If Previsao.Value <> 0 Then
           
        'Acrescenta uma linha no Grid se for o caso
        If GridParcelas.Row - GridParcelas.FixedRows = objGridInt.iLinhasExistentes Then
            colParcelaReceber.Add 0, 0, GridParcelas.Row, STATUS_ABERTO, DATA_NULA, DATA_NULA, 0, 0, 1, 1, 1, " ", 0, 0, 0, 0, 0, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, DATA_NULA, 0, 0, 0, 0, 0, 0, "", 0
            colcolDesconto.Add colDesconto
            colcolComissao.Add colComissao
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 129456

    Saida_Celula_Previsao = SUCESSO

    Exit Function

Erro_Saida_Celula_Previsao:

    Saida_Celula_Previsao = gErr

    Select Case gErr

        Case 129456
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175256)

    End Select

    Exit Function

End Function

Public Sub Cliente_Preenche()

Static sNomeReduzidoParte As String '*** rotina para trazer cliente
Dim lErro As Long
Dim objCliente As Object
    
On Error GoTo Erro_Cliente_Preenche
    
    Set objCliente = Cliente
    
    lErro = CF("Cliente_Pesquisa_NomeReduzido", objCliente, sNomeReduzidoParte)
    If lErro <> SUCESSO Then gError 134022

    Exit Sub

Erro_Cliente_Preenche:

    Select Case gErr

        Case 134022

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175257)

    End Select
    
    Exit Sub

End Sub

Function Carrega_Moeda() As Long

Dim lErro As Long
Dim objMoeda As ClassMoedas
Dim colMoedas As New Collection

On Error GoTo Erro_Carrega_Moeda
    
    lErro = CF("Moedas_Le_Todas", colMoedas) 'leo colocar CF
    If lErro <> SUCESSO Then gError 103371
    
    'se não existem moedas cadastradas
    If colMoedas.Count = 0 Then gError 103372
    
    For Each objMoeda In colMoedas
    
        If objMoeda.iCodigo <> MOEDA_REAL Then Moeda.AddItem objMoeda.iCodigo & SEPARADOR & objMoeda.sNome
    
    Next

    Carrega_Moeda = SUCESSO
    
    Exit Function
    
Erro_Carrega_Moeda:

    Carrega_Moeda = gErr
    
    Select Case gErr
    
        Case 103371
        
        Case 103372
            Call Rotina_Erro(vbOKOnly, "ERRO_MOEDAS_NAO_CADASTRADAS", gErr, Error)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175258)
    
    End Select

End Function

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

Public Sub ReajusteBase_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ReajusteBase_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ReajusteBase_Validate
    
    'Critica a data digitada
    lErro = Data_Critica(ReajusteBase.Text)
    If lErro <> SUCESSO Then Error 26140
        
    Exit Sub

Erro_ReajusteBase_Validate:

    Cancel = True
    
    Select Case Err

        Case 26140
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175261)

    End Select

    Exit Sub

End Sub

Public Sub UpDownReajusteBase_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownReajusteBase_DownClick

    'Diminui a data
    lErro = Data_Up_Down_Click(ReajusteBase, DIMINUI_DATA)
    If lErro Then Error 26141

    Exit Sub

Erro_UpDownReajusteBase_DownClick:

    Select Case Err

        Case 26141

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175262)

    End Select

    Exit Sub

End Sub

Public Sub UpDownReajusteBase_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownReajusteBase_UpClick

    'Aumenta a Data de Emissão em um dia
    lErro = Data_Up_Down_Click(ReajusteBase, AUMENTA_DATA)
    If lErro Then Error 26142

    Exit Sub

Erro_UpDownReajusteBase_UpClick:

    Select Case Err

        Case 26142

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175263)

    End Select

    Exit Sub

End Sub

Public Sub Natureza_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LabelNatureza_Click()

    Dim objNatMovCta As New ClassNatMovCta
    Dim colSelecao As New Collection

    objNatMovCta.sCodigo = Natureza.ClipText
    
    colSelecao.Add NATUREZA_TIPO_RECEBIMENTO
    
    Call Chama_Tela("NatMovCtaLista", colSelecao, objNatMovCta, objEventoNatureza, "Tipo = ?")

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
        lErro = CF("Natureza_Critica", objNatMovCta, NATUREZA_TIPO_RECEBIMENTO)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175264)
        
    End Select

    Exit Sub
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175265)
        
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
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175266)

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
Dim colComissaoEmissao As New colComissao

On Error GoTo Erro_BotaoDif_Click

    'Move os dados da Tela para objtituloreceber e colParcelas
    lErro = Move_Tela_Memoria(objTituloReceber, colComissaoEmissao)
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

'###########################################################################
'Inserido por Wagner 11/08/2006
Public Sub BotaoVendedoresParc_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection, lErro As Long
    
On Error GoTo Erro_BotaoVendedoresParc_Click

    If (GridComissoes.Row = 0) Or (GridParcelas.Row = 0) Then Error 56594
    
    'Chama a tela de browse de Vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedoresParc)

    Exit Sub

Erro_BotaoVendedoresParc_Click:

    Select Case gErr

        Case 56594
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175180)

    End Select

    Exit Sub

End Sub

Private Sub objEventoVendedoresParc_evSelecao(obj1 As Object)

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedor As ClassVendedor

On Error GoTo Erro_objEventoVendedoresParc_evSelecao
    
    If GridComissoes.Row >= GridComissoes.FixedRows Then

        Set objVendedor = obj1
        
        If GridComissoes.Row > objGridComissoes.iLinhasExistentes Then

            objGridComissoes.iLinhasExistentes = objGridComissoes.iLinhasExistentes + 1
            colcolComissao.Item(GridParcelas.Row).colComissao.Add 0, 0, 0, 0, 0, objVendedor.iCodigo, DATA_NULA, 0, 0, 0, DATA_NULA

        End If

        If Me.ActiveControl Is Vendedor Then
            Vendedor.Text = objVendedor.sNomeReduzido
        Else
            GridComissoes.TextMatrix(GridComissoes.Row, iGrid_Vendedor_Col) = objVendedor.sNomeReduzido
        End If

        Call Vendedor_Linha_Preenche(objVendedor, GridParcelas.Row)

    End If
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoVendedoresParc_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175183)
            
    End Select
        
    Exit Sub
    
End Sub

Private Sub Exibe_Dados_Parcela()

Dim iIndice As Integer
Dim lErro As Long
Dim iParcela As Integer

On Error GoTo Erro_Exibe_Dados_Parcela

    If GridParcelas.Row <> 0 And GridParcelas.Row <= objGridParcelas.iLinhasExistentes Then

        SSFrame1.Enabled = True
        SSFrame1.Caption = "Descontos (Parcela " & CStr(GridParcelas.Row) & ")"
    
        SSFrame4.Enabled = True
        SSFrame4.Caption = "Comissões (Parcela " & CStr(GridParcelas.Row) & ")"

        DoEvents
    
        If objGridParcelas.iLinhaAntiga <> 0 Then
            'Recolhe os dados da parcela que está na tela
            lErro = Recolhe_Parcela_Tela(objGridParcelas.iLinhaAntiga)
            If lErro <> SUCESSO Then gError 26382
        End If
     
        'Traz para telas os dados da Parcela seguinte
        lErro = Traz_Parcela_Tela(GridParcelas.Row)
        If lErro <> SUCESSO Then gError 26383
        
    Else
    
        SSFrame1.Enabled = False
        SSFrame1.Caption = "Descontos (Selecione uma parcela)"
    
        SSFrame4.Enabled = False
        SSFrame4.Caption = "Comissões (Selecione uma parcela)"
    
    End If
    
    Exit Sub
    
Erro_Exibe_Dados_Parcela:

    Select Case gErr
    
        Case 26382, 26383
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 175224)
            
    End Select
    
    Exit Sub
    
End Sub

'#####################################
'Inserido por Wagner 03/08/2006
Public Sub BotaoProjetos_Click()
    Call gobjTelaProjetoInfo.BotaoProjetos_Click
End Sub

Public Sub LabelProjeto_Click()
    Call gobjTelaProjetoInfo.LabelProjeto_Click
End Sub

Public Sub Projeto_GotFocus()
    Call MaskEdBox_TrataGotFocus(Projeto, iAlterado)
End Sub

Public Sub Projeto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Projeto_Validate(Cancel As Boolean)
    Call gobjTelaProjetoInfo.Projeto_Validate(Cancel)
End Sub

Function Obter_ItensPRJCR(ByVal colItensPRJ As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objItensPRJCR As ClassItensPRJCR

On Error GoTo Erro_Obter_ItensPRJCR

    Set objItensPRJCR = New ClassItensPRJCR

    objItensPRJCR.sItem = "GR"
    objItensPRJCR.sDescricao = "Geral"

    objItensPRJCR.dValorOriginal = StrParaDbl(Valor.Text)

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

'###########################################################################
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

Public Sub ISSRetido_Change()

    iISSRetidoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ISSRetido_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iIndice As Integer
Dim objComissao As ClassComissao

On Error GoTo Erro_ISSRetido_Validate

    If iISSRetidoAlterado = 0 Then Exit Sub
    
    'Verifica se algum valor foi digitado
    If Len(Trim(ISSRetido.ClipText)) <> 0 Then

        'Critica se é valor não negativo
        lErro = Valor_NaoNegativo_Critica(ISSRetido.Text)
        If lErro <> SUCESSO Then Error 26144

        'Põe o valor formatado na tela
        ISSRetido.Text = Format(ISSRetido.Text, "Standard")

    End If
    
    Call Recalcula_Parcelas_Comissoes
    
    iISSRetidoAlterado = 0
     
    Exit Sub

Erro_ISSRetido_Validate:

    Cancel = True

    Select Case Err

        Case 26144, 26422, 29021

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175252)

    End Select

    Exit Sub

End Sub


