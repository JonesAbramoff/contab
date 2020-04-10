VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTInventarioLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

Public gobjAnotacao As ClassAnotacoes

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

Public iAlterado As Integer
Dim iLinhaAntiga As Integer
Dim iFrameAtual As Integer
Dim sCodigoAntigo As String
Dim colItensNumIntDoc As Collection


Public objGrid As AdmGrid

Public iGrid_Sequencial_Col As Integer
Public iGrid_Produto_Col As Integer
Public iGrid_DescricaoItem_Col As Integer
Public iGrid_Almoxarifado_Col As Integer
Public iGrid_Tipo_Col As Integer
Public iGrid_UnidadeMed_Col As Integer
Public iGrid_Atualiza_Col As Integer
Public iGrid_Quantidade_Col As Integer
Public iGrid_QuantEstoque_Col As Integer
Public iGrid_QuantDiferenca_Col As Integer
Public iGrid_CustoUnitario_Col As Integer
Public iGrid_ValorDiferenca_col As Integer
Public iGrid_ContaContabil_Col As Integer
Public iGrid_Etiqueta_Col As Integer
Public iGrid_ContaAjuste_Col As Integer
Public iGrid_LoteProduto_Col As Integer
Public iGrid_FilialOP_Col As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoAlmoxPadrao As AdmEvento
Attribute objEventoAlmoxPadrao.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoEstoque As AdmEvento
Attribute objEventoEstoque.VB_VarHelpID = -1
Private WithEvents objEventoContaContabil As AdmEvento
Attribute objEventoContaContabil.VB_VarHelpID = -1
Private WithEvents objEventoLoteInv As AdmEvento
Attribute objEventoLoteInv.VB_VarHelpID = -1
Private WithEvents objEventoRastroLote As AdmEvento
Attribute objEventoRastroLote.VB_VarHelpID = -1

'Constantes públicas dos tabs
Private Const TAB_Lancamentos = 1
Private Const TAB_Contabilizacao = 2

'Contante que informa qual é o Mnemonico global
Const CTAAJUSTEINV As String = "CtaAjusteInv"

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property


Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
End Property

Public Property Get Lote() As Object
     Set Lote = objUserControl.Controls("Lote")
End Property

Public Property Get BotaoEstoque() As Object
     Set BotaoEstoque = objUserControl.Controls("BotaoEstoque")
End Property

Public Property Get BotaoProdutos() As Object
     Set BotaoProdutos = objUserControl.Controls("BotaoProdutos")
End Property

Public Property Get DescricaoItem() As Object
     Set DescricaoItem = objUserControl.Controls("DescricaoItem")
End Property

Public Property Get UnidadeMed() As Object
     Set UnidadeMed = objUserControl.Controls("UnidadeMed")
End Property

Public Property Get Etiqueta() As Object
     Set Etiqueta = objUserControl.Controls("Etiqueta")
End Property

Public Property Get Tipo() As Object
     Set Tipo = objUserControl.Controls("Tipo")
End Property

Public Property Get BotaoPlanoConta() As Object
     Set BotaoPlanoConta = objUserControl.Controls("BotaoPlanoConta")
End Property

Public Property Get ContaAjuste() As Object
     Set ContaAjuste = objUserControl.Controls("ContaAjuste")
End Property

Public Property Get ContaContabil() As Object
     Set ContaContabil = objUserControl.Controls("ContaContabil")
End Property

Public Property Get ValorDiferenca() As Object
     Set ValorDiferenca = objUserControl.Controls("ValorDiferenca")
End Property

Public Property Get QuantDiferenca() As Object
     Set QuantDiferenca = objUserControl.Controls("QuantDiferenca")
End Property

Public Property Get QuantEstoque() As Object
     Set QuantEstoque = objUserControl.Controls("QuantEstoque")
End Property

Public Property Get CustoUnitario() As Object
     Set CustoUnitario = objUserControl.Controls("CustoUnitario")
End Property

Public Property Get Quantidade() As Object
     Set Quantidade = objUserControl.Controls("Quantidade")
End Property

Public Property Get Produto() As Object
     Set Produto = objUserControl.Controls("Produto")
End Property

Public Property Get Almoxarifado() As Object
     Set Almoxarifado = objUserControl.Controls("Almoxarifado")
End Property

Public Property Get UpDownData() As Object
     Set UpDownData = objUserControl.Controls("UpDownData")
End Property

Public Property Get Data() As Object
     Set Data = objUserControl.Controls("Data")
End Property

'hora
Public Property Get HORA() As Object
     Set HORA = objUserControl.Controls("Hora")
End Property


Public Property Get AlmoxPadrao() As Object
     Set AlmoxPadrao = objUserControl.Controls("AlmoxPadrao")
End Property

Public Property Get GridItens() As Object
     Set GridItens = objUserControl.Controls("GridItens")
End Property

Public Property Get CodigoLabel() As Object
     Set CodigoLabel = objUserControl.Controls("CodigoLabel")
End Property

Public Property Get AlmoxPadraoLabel() As Object
     Set AlmoxPadraoLabel = objUserControl.Controls("AlmoxPadraoLabel")
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

Public Property Get FilialOP() As Object
     Set FilialOP = objUserControl.Controls("FilialOP")
End Property

Public Property Get LoteProduto() As Object
     Set LoteProduto = objUserControl.Controls("LoteProduto")
End Property

Public Property Get Atualiza() As Object
     Set Atualiza = objUserControl.Controls("Atualiza")
End Property


Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridItens.Row = 0 Then gError 43747

    If GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = "" Then gError 43748

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 43750

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 43747
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43748
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 43750

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161893)

    End Select

    Exit Sub

End Sub

Public Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaAjuste_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ContaAjuste_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub ContaAjuste_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub ContaAjuste_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaAjuste
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula_ContaAjuste(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaAjuste

    Set objGrid.objControle = ContaAjuste

    If Len(Trim(ContaAjuste.ClipText)) > 0 Then

        'Verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaAjuste.Text, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 52070

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 52071

            ContaAjuste.PromptInclude = False
            ContaAjuste.Text = sContaMascarada
            ContaAjuste.PromptInclude = True

        'Se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaAjuste.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 52072

            'Conta não cadastrada
            If lErro = 5700 Then gError 52073

        End If

    Else

        ContaAjuste.PromptInclude = False
        ContaAjuste.Text = ""
        ContaAjuste.PromptInclude = True

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52074

    Saida_Celula_ContaAjuste = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaAjuste:

    Saida_Celula_ContaAjuste = gErr

    Select Case gErr

        Case 52070, 52072, 52074
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52071
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52073
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaAjuste.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161894)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iFrameAtual = 1

    Set objEventoCodigo = New AdmEvento
    Set objEventoLoteInv = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoContaContabil = New AdmEvento
    Set objEventoRastroLote = New AdmEvento

    'Preenche list de ComboBox Tipo com o ItemData Correspondente
    Call CargaCombo_Tipo(Tipo)

    'Preenche a Data
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Inicializa a Máscara do Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 41063

    'le a mascara das conta
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 49572

    'Inicializa Mascara Conta Ajuste
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaAjuste)
    If lErro <> SUCESSO Then gError 52069

    'Formato para a Quantidade
    Quantidade.Format = FORMATO_ESTOQUE
    QuantEstoque.Format = FORMATO_ESTOQUE
    QuantDiferenca.Format = FORMATO_ESTOQUE

    'Formato para o Custo
    CustoUnitario.Format = FORMATO_CUSTO
    ValorDiferenca.Format = FORMATO_CUSTO

    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 78574
    
    'Inicializa o GridItens
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridItens(objGrid)
    If lErro <> SUCESSO Then gError 41064

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 41063, 41064, 49572, 52069, 78574

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161895)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 78573

    'Carrega a combo de Filial Empresa com código e nome
    For Each objCodigoNome In colCodigoNome
        FilialOP.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
        FilialOP.ItemData(FilialOP.NewIndex) = objCodigoNome.iCodigo
    Next

    Carrega_FilialOP = SUCESSO

    Exit Function

Erro_Carrega_FilialOP:

    Carrega_FilialOP = gErr

    Select Case gErr

        Case 78573

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161896)

    End Select

    Exit Function

End Function

Private Sub CargaCombo_Tipo(objComboTipo As ComboBox)
'inicializa a combo com os tipos de estoque possiveis

Dim lErro As Long

On Error GoTo Erro_CargaCombo_Tipo

    With objComboTipo

        .AddItem STRING_QUANT_DISPONIVEL_NOSSA
        .ItemData(.NewIndex) = TIPO_QUANT_DISPONIVEL_NOSSA
        .AddItem STRING_QUANT_RECEB_INDISP
        .ItemData(.NewIndex) = TIPO_QUANT_RECEB_INDISP
        .AddItem STRING_QUANT_OUTRAS_INDISP
        .ItemData(.NewIndex) = TIPO_QUANT_OUTRAS_INDISP
        .AddItem STRING_QUANT_DEFEIT
        .ItemData(.NewIndex) = TIPO_QUANT_DEFEIT
        .AddItem STRING_QUANT_3_CONSIG
        .ItemData(.NewIndex) = TIPO_QUANT_3_CONSIG
        .AddItem STRING_QUANT_3_DEMO
        .ItemData(.NewIndex) = TIPO_QUANT_3_DEMO
        .AddItem STRING_QUANT_3_CONSERTO
        .ItemData(.NewIndex) = TIPO_QUANT_3_CONSERTO
        .AddItem STRING_QUANT_3_OUTRAS
        .ItemData(.NewIndex) = TIPO_QUANT_3_OUTRAS
        .AddItem STRING_QUANT_3_BENEF
        .ItemData(.NewIndex) = TIPO_QUANT_3_BENEF
        .AddItem STRING_QUANT_DISPONIVEL_NOSSA_CI
        .ItemData(.NewIndex) = TIPO_QUANT_DISPONIVEL_NOSSA_CI
        .AddItem STRING_QUANT_RECEB_INDISP_CI
        .ItemData(.NewIndex) = TIPO_QUANT_RECEB_INDISP_CI
        .AddItem STRING_QUANT_OUTRAS_INDISP_CI
        .ItemData(.NewIndex) = TIPO_QUANT_OUTRAS_INDISP_CI
        .AddItem STRING_QUANT_DEFEIT_CI
        .ItemData(.NewIndex) = TIPO_QUANT_DEFEIT_CI
        .AddItem STRING_QUANT_3_CONSIG_CI
        .ItemData(.NewIndex) = TIPO_QUANT_3_CONSIG_CI

    End With

    Exit Sub

Erro_CargaCombo_Tipo:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161897)

    End Select

    Exit Sub

End Sub

Public Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Tipo Estoque")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Lote \ O.P.")
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Atualiza")
    objGridInt.colColuna.Add ("Quant. Encontrada")
    objGridInt.colColuna.Add ("Quant. Estoque")
    objGridInt.colColuna.Add ("Diferença-Quant.")
    objGridInt.colColuna.Add ("Custo Unitário")
    objGridInt.colColuna.Add ("Diferença-Valor ")
    objGridInt.colColuna.Add ("Conta de Estoque")
    objGridInt.colColuna.Add ("Conta de Ajuste")
    objGridInt.colColuna.Add ("Etiqueta")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
    objGridInt.colCampo.Add (LoteProduto.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (Atualiza.Name)
    objGridInt.colCampo.Add (Quantidade.Name)
    objGridInt.colCampo.Add (QuantEstoque.Name)
    objGridInt.colCampo.Add (QuantDiferenca.Name)
    objGridInt.colCampo.Add (CustoUnitario.Name)
    objGridInt.colCampo.Add (ValorDiferenca.Name)
    objGridInt.colCampo.Add (ContaContabil.Name)
    objGridInt.colCampo.Add (ContaAjuste.Name)
    objGridInt.colCampo.Add (Etiqueta.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_DescricaoItem_Col = 2
    iGrid_Almoxarifado_Col = 3
    iGrid_Tipo_Col = 4
    iGrid_UnidadeMed_Col = 5
    iGrid_LoteProduto_Col = 6
    iGrid_FilialOP_Col = 7
    iGrid_Atualiza_Col = 8
    iGrid_Quantidade_Col = 9
    iGrid_QuantEstoque_Col = 10
    iGrid_QuantDiferenca_Col = 11
    iGrid_CustoUnitario_Col = 12
    iGrid_ValorDiferenca_col = 13
    iGrid_ContaContabil_Col = 14
    iGrid_ContaAjuste_Col = 15
    iGrid_Etiqueta_Col = 16

    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA  "
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim objInventario As New ClassInventario

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "InventPendente"

    If Len(Trim(Codigo.Text)) > 0 Then objInventario.sCodigo = Codigo.Text

    If Len(Trim(Data.ClipText)) > 0 Then
        objInventario.dtData = CDate(Data.Text)
    Else
        objInventario.dtData = DATA_NULA
    End If

    If Len(Trim(HORA.ClipText)) > 0 Then
        objInventario.dtHora = CDate(HORA.Text)
    Else
        objInventario.dtHora = 0
    End If

    If Len(Trim(Lote.ClipText)) > 0 Then objInventario.iLote = Lote.Text

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objInventario.sCodigo, STRING_INVENTARIO_CODIGO, "Codigo"
    colCampoValor.Add "Data", objInventario.dtData, 0, "Data"
    colCampoValor.Add "Lote", objInventario.iLote, 0, "Lote"
'hora
    colCampoValor.Add "Hora", CDbl(objInventario.dtHora), 0, "Hora"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161898)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do Banco de Dados

Dim lErro As Long
Dim objInventario As New ClassInventario

On Error GoTo Erro_Tela_Preenche

    'Passa os dados da coleção para objReserva
    objInventario.sCodigo = colCampoValor.Item("Codigo").vValor
    objInventario.dtData = colCampoValor.Item("Data").vValor
    objInventario.dtHora = colCampoValor.Item("Hora").vValor
    objInventario.iLote = colCampoValor.Item("Lote").vValor
    objInventario.iFilialEmpresa = giFilialEmpresa

    lErro = Preenche_Tela(objInventario)
    If lErro <> SUCESSO Then gError 41065

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 41065

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161899)

    End Select

    Exit Sub

End Sub

Public Function Preenche_Tela(objInventario As ClassInventario) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Preenche_Tela

    'Limpa a tela sem Fechar o Comando de setas
    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'Lê os ítens do Inventario
    lErro = CF("InventarioPendenteItens_Le", objInventario)
    If lErro <> SUCESSO And lErro <> 41152 Then gError 41066

    'Coloca os Dados na Tela
    Codigo.Text = objInventario.sCodigo

    If objInventario.dtData <> DATA_NULA Then
        Data.PromptInclude = False
        Data.Text = Format(objInventario.dtData, "dd/mm/yy")
        Data.PromptInclude = True

    Else
        Data.PromptInclude = False
        Data.Text = ""
        Data.PromptInclude = True

    End If

'hora
    HORA.PromptInclude = False
    'este teste está correto
    If objInventario.dtData <> DATA_NULA Then HORA.Text = Format(objInventario.dtHora, "hh:mm:ss")
    HORA.PromptInclude = True

    Lote.PromptInclude = False
    Lote.Text = CStr(objInventario.iLote)
    Lote.PromptInclude = True

    lErro = Preenche_GridItens(objInventario.colItens)
    If lErro <> SUCESSO Then gError 41067

    iAlterado = 0

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 41066, 41067

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161900)

    End Select

    Exit Function

End Function

Public Function Preenche_GridItens(colItens As ColItemInventario) As Long

Dim iIndice As Integer, iPosicao As Integer
Dim lErro As Long
Dim sProdutoMascarado As String, sCclMascarado As String
Dim sContaEnxuta As String
Dim objItemInventario As ClassItemInventario
Dim sContaAjusteEnxuta As String
Dim objFilialOP As New AdmFiliais

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Preenche GridItens
    For Each objItemInventario In colItens

        iIndice = iIndice + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemInventario.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 41068

'        lErro = Mascara_MascararProduto(objItemInventario.sProduto, sProdutoMascarado)
'        If lErro <> SUCESSO Then gError 41068

        If objItemInventario.sContaContabilEst <> "" Then

            lErro = Mascara_RetornaContaEnxuta(objItemInventario.sContaContabilEst, sContaEnxuta)
            If lErro <> SUCESSO Then gError 49578

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True

            GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col) = ContaContabil.Text

        End If

        sContaAjusteEnxuta = String(STRING_CONTA, 0)

        If objItemInventario.sContaContabilInv <> "" Then

            lErro = Mascara_RetornaContaEnxuta(objItemInventario.sContaContabilInv, sContaAjusteEnxuta)
            If lErro <> SUCESSO Then gError 52075

            ContaAjuste.PromptInclude = False
            ContaAjuste.Text = sContaAjusteEnxuta
            ContaAjuste.PromptInclude = True

            GridItens.TextMatrix(iIndice, iGrid_ContaAjuste_Col) = ContaAjuste.Text

        End If

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
        GridItens.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objItemInventario.sProdutoDesc
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemInventario.sSiglaUM
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemInventario.dQuantidade)

        For iPosicao = 0 To Tipo.ListCount - 1
            If Tipo.ItemData(iPosicao) = objItemInventario.iTipo Then
                GridItens.TextMatrix(iIndice, iGrid_Tipo_Col) = Tipo.List(iPosicao)
                Exit For
            End If
        Next

        GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemInventario.sAlmoxarifadoNomeRed
        GridItens.TextMatrix(iIndice, iGrid_Etiqueta_Col) = objItemInventario.sEtiqueta

        GridItens.TextMatrix(iIndice, iGrid_Etiqueta_Col) = objItemInventario.sEtiqueta
    
        If Len(Trim(objItemInventario.sLote)) > 0 Then GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col) = objItemInventario.sLote
        
        If objItemInventario.iFilialOP > 0 Then
            
            objFilialOP.iCodFilial = objItemInventario.iFilialOP
            
            'Le a Filial Empresa da OP para pegar a descrição
            lErro = CF("FilialEmpresa_Le", objFilialOP)
            If lErro <> SUCESSO Then gError 78470
            
            GridItens.TextMatrix(iIndice, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
                
        End If
        
        If Len(Trim(objItemInventario.sLote)) > 0 Then
        
            For iPosicao = 0 To Atualiza.ListCount - 1
                If Atualiza.ItemData(iPosicao) = objItemInventario.iAtualizaSoLote Then
                    GridItens.TextMatrix(iIndice, iGrid_Atualiza_Col) = Atualiza.List(iPosicao)
                    Exit For
                End If
            Next
        
        End If
        
        GridItens.Row = iIndice

        'Preenche o custo de acordo com tipo e o Produto se é compravel ou produzivel
        lErro = Atualiza_Custo_Inventario(iIndice, objItemInventario)
        If lErro <> SUCESSO Then gError 49648
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
            'Calcula a diferenca da quantidade e a diferenca do valor
            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), iIndice)
            If lErro <> SUCESSO Then gError 49649
        Else
            'Calcula a Quantidade Disponível
            lErro = QuantLote_Calcula(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), iIndice, GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78906
        End If

    Next

    objGrid.iLinhasExistentes = colItens.Count

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 41069

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 41069, 49648, 49649, 78470

        Case 41068
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemInventario.sProduto)

        Case 49578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objItemInventario.sContaContabilEst)

        Case 52075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objItemInventario.sContaContabilInv)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161901)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long
'Faz a saida de celula da Filial da Ordem de Produção

Dim lErro As Long
Dim objFilialOP As New AdmFiliais
Dim iCodigo As Integer
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objRastroLote As New ClassRastreamentoLote
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_FilialOP

    Set objGridInt.objControle = FilialOP

    If Len(Trim(FilialOP.Text)) <> 0 Then
            
        'Verifica se é uma FilialOP selecionada
        If FilialOP.Text <> FilialOP.List(FilialOP.ListIndex) Then
        
            'Tenta selecionar na combo
            lErro = Combo_Seleciona(FilialOP, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 78481
    
            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then
    
                objFilialOP.iCodFilial = iCodigo
    
                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 78482
        
                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 78483
        
                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome
            
            End If
    
            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 78484
                    
        End If
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) > 0 Then
                    
            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 78583
                                
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                objRastroLote.sCodigo = GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col)
                objRastroLote.sProduto = sProdutoFormatado
                objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)
            
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78584
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78585
                            
                'Calcula a Quantidade Disponível
                lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(FilialOP.Text))
                If lErro <> SUCESSO Then gError 78914
                            
            End If
            
        End If
            
    Else
        
        'se o lote também não está preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
        
            'Calcula a Quantidade Disponível nesse Almoxarifado
            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 78915
    
        Else
        
            GridItens.TextMatrix(GridItens.Row, iGrid_QuantEstoque_Col) = ""
    
        End If
    
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78485

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 78481, 78482, 78485, 78580, 78583, 78584, 78914, 78915
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78483
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78484
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78581
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78585
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161902)

    End Select

    Exit Function

End Function
'Maristela(Inicio)
Public Function Saida_Celula_Atualiza(objGridInt As AdmGrid) As Long
'Faz a saida de celula de Atualiza

Dim lErro As Long
Dim iAtualiza As Integer

On Error GoTo Erro_Saida_Celula_Atualiza

    Set objGridInt.objControle = Atualiza
    
    GridItens.TextMatrix(GridItens.Row, iGrid_Atualiza_Col) = Atualiza.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 90833

    Saida_Celula_Atualiza = SUCESSO

    Exit Function

Erro_Saida_Celula_Atualiza:

    Saida_Celula_Atualiza = gErr

    Select Case gErr

        Case 90833
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161903)

    End Select

    Exit Function

End Function
'Maristela(Fim)

Public Function Saida_Celula_LoteProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iPosicao As Integer

On Error GoTo Erro_Saida_Celula_LoteProduto

    Set objGridInt.objControle = LoteProduto
    
    If Len(Trim(LoteProduto.Text)) > 0 Then
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78548
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78549
            
        If lErro = 28030 Then gError 78550
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then 'Alterado por Wagner
                
                objRastroLote.sCodigo = LoteProduto.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78551
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78552
                
                'Calcula a Quantidade Disponível
                lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, LoteProduto.Text, Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78910
                
            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                
                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))) > 0 Then
                    
                    objRastroLote.sCodigo = LoteProduto.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))
                    
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 78556
                    
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 78557
                
                    'Calcula a Quantidade Disponível
                    lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, LoteProduto.Text, Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78911
                    
                    For iPosicao = 0 To Atualiza.ListCount - 1
                        If Atualiza.ItemData(iPosicao) = INVENTARIO_ATUALIZA_LOTE_ESTOQUE Then
                            GridItens.TextMatrix(GridItens.Row, iGrid_Atualiza_Col) = Atualiza.List(iPosicao)
                            Exit For
                        End If
                    Next
        
                End If
                
            End If
        
        End If
        
    Else
        'Calcula a Quantidade Disponível nesse Almoxarifado
        lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
        If lErro <> SUCESSO Then gError 78912
        
        Atualiza.ListIndex = -1
        GridItens.TextMatrix(GridItens.Row, iGrid_Atualiza_Col) = Atualiza.Text

    End If
            
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78558

    Saida_Celula_LoteProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_LoteProduto:

    Saida_Celula_LoteProduto = gErr

    Select Case gErr

        Case 78548, 78549, 78551, 78553, 78556, 78558, 78910, 78911, 78912
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78550
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 78552, 78557
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 78554
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78555
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161904)

    End Select

    Exit Function

End Function

Public Sub LoteProduto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LoteProduto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub LoteProduto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub LoteProduto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = LoteProduto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'Maristela(Inicio)
Public Sub Atualiza_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Atualiza_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Atualiza_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Atualiza_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Atualiza
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'Maristela(Fim)

Public Sub FilialOP_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOP_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub FilialOP_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub FilialOP_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = FilialOP
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub CodigoLabel_Click()

Dim lErro As Long
Dim objInventario As New ClassInventario
Dim colSelecao As New Collection

    objInventario.sCodigo = Trim(Codigo.Text)

    If Len(Trim(Lote.ClipText)) > 0 Then objInventario.iLote = CInt(Lote.Text)

    Call Chama_Tela("InventarioLoteLista", colSelecao, objInventario, objEventoCodigo)

End Sub


Public Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid)

End Sub

Public Sub Lote_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Lote, iAlterado)

End Sub

Public Sub Lote_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objInvLote As New ClassInvLote
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Lote_Validate

    If Len(Trim(Lote.ClipText)) > 0 Then

        objInvLote.iLote = CInt(Lote.Text)
        objInvLote.iFilialEmpresa = giFilialEmpresa

        lErro = CF("InvLote_Le", objInvLote)
        If lErro <> SUCESSO And lErro <> 41177 Then gError 41211

        'se o lote já foi atualizado ==> erro
        If lErro = SUCESSO Then gError 41212

        lErro = CF("InvLotePendente_Le", objInvLote)
        If lErro <> SUCESSO And lErro <> 41181 Then gError 41213

        'se o lote não existe, pergunta se quer cadastrar
        If lErro = 41181 Then gError 41214

    End If

    Exit Sub

Erro_Lote_Validate:

    Cancel = True


    Select Case gErr

        Case 41211, 41213

        Case 41212
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INV_LOTE_PEND_CADASTRADO", gErr, objInvLote.iLote)

        Case 41214
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_LOTEINV", objInvLote.iLote)

            If vbMsg = vbYes Then
                Call Chama_Tela("LoteEst", objInvLote)
            Else
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161905)

    End Select

    Exit Sub

End Sub

Public Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objInventario As ClassInventario

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objInventario = obj1

    lErro = Preenche_Tela(objInventario)
    If lErro <> SUCESSO Then gError 41070

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 41070

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161906)

    End Select

    Exit Sub

End Sub

Public Sub LoteLabel_Click()

Dim objInvLote As New ClassInvLote
Dim colSelecao As New Collection

    If Len(Trim(Lote.ClipText)) > 0 Then objInvLote.iLote = CInt(Lote.Text)

    Call Chama_Tela("InvLotePendenteLista", colSelecao, objInvLote, objEventoLoteInv)

End Sub

Public Sub objEventoContaContabil_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPlanoConta As ClassPlanoConta
Dim sContaEnxuta As String

On Error GoTo Erro_objEventoContaContabil_evSelecao

    Set objPlanoConta = obj1

    If objPlanoConta.sConta <> "" Then

        If GridItens.Col <> iGrid_ContaContabil_Col And GridItens.Col <> iGrid_ContaAjuste_Col Then
            Me.Show
            Exit Sub
        End If
        
        lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaEnxuta)
        If lErro <> SUCESSO Then gError 55435
        
        If GridItens.Col = iGrid_ContaContabil_Col Then
            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True
        Else
            ContaAjuste.PromptInclude = False
            ContaAjuste.Text = sContaEnxuta
            ContaAjuste.PromptInclude = True
        End If

        GridItens.TextMatrix(GridItens.Row, GridItens.Col) = objGrid.objControle.Text
    
    End If

    Me.Show

    Exit Sub

Erro_objEventoContaContabil_evSelecao:

    Select Case gErr

        Case 55435
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161907)

    End Select

    Exit Sub

End Sub

Public Sub objEventoLoteInv_evSelecao(obj1 As Object)

Dim objInvLote As ClassInvLote

    Set objInvLote = obj1

    Lote.PromptInclude = False
    Lote.Text = CStr(objInvLote.iLote)
    Lote.PromptInclude = True

    Me.Show

End Sub

Public Sub AlmoxPadraoLabel_Click()

Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colSelecao As New Collection

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxPadrao)

End Sub

Public Sub objEventoAlmoxPadrao_evSelecao(obj1 As Object)

Dim objAlmoxarifado As ClassAlmoxarifado

    Set objAlmoxarifado = obj1

    'Preenche AlmoxPadrao com o nome reduzido de objAlmoxarifado
    AlmoxPadrao.Text = objAlmoxarifado.sNomeReduzido

    Me.Show

End Sub

Public Sub BotaoProdutos_Click()

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim sSelecao As String

On Error GoTo Erro_BotaoProdutos_Click

    If GridItens.Row = 0 Then gError 43713

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 41071

        If iPreenchido = PRODUTO_PREENCHIDO Then objProduto.sCodigo = sProduto

    End If
    sSelecao = "ControleEstoque<>?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    
    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 41071

        Case 43713
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161908)

    End Select
    
    Exit Sub

End Sub

Public Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim dCusto As Double
Dim sTipo As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 41072

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 41073


            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 41074

            If lErro = 28030 Then gError 41075

            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True

            If Not (Me.ActiveControl Is Produto) Then
            
                'Preenche a Linha do Grid
                GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = sProdutoMascarado
    
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 41076
    
                sTipo = GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col)
    
                'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
                lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
                If lErro <> SUCESSO Then gError 49606
    
                GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)

                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
                    lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
                    If lErro <> SUCESSO Then gError 41077
                Else
                    'Calcula a Quantidade Disponível
                    lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78907
                End If
                
                lErro = Preenche_ContaContabil()
                If lErro <> SUCESSO Then gError 52244
    
            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 41072, 41074, 41076, 41077, 49606, 52244, 78907

        Case 41073
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case 41075
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161909)

    End Select

    Exit Sub

End Sub

Public Function Preenche_ContaAjuste() As Long
'preenche a conta de ajuste com o que vem da tabela de MnemonicoCTBValor

Dim lErro As Long
Dim objMnemonicoCTBValor As New ClassMnemonicoCTBValor
Dim sContaEnxuta As String
Dim sContaFormatada As String
Dim iContaPreenchida As Integer

On Error GoTo Erro_Preenche_ContaAjuste

    objMnemonicoCTBValor.sMnemonico = CTAAJUSTEINV

    lErro = CF("MnemonicoCTBValor_Le", objMnemonicoCTBValor)
    If lErro <> SUCESSO And lErro <> 39690 Then gError 52078

    If Trim(objMnemonicoCTBValor.sValor) <> "" Then

        ContaAjuste.PromptInclude = False
        ContaAjuste.Text = objMnemonicoCTBValor.sValor
        ContaAjuste.PromptInclude = True

        GridItens.TextMatrix(GridItens.Row, iGrid_ContaAjuste_Col) = ContaAjuste.Text

    End If

    Preenche_ContaAjuste = SUCESSO

    Exit Function

Erro_Preenche_ContaAjuste:

    Preenche_ContaAjuste = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161910)

    End Select

    Exit Function

End Function

Public Function Preenche_ContaContabil() As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre ----> nao tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabil

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col))) > 0 And Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then
    
        'preenche o objEstoqueProduto
        objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col)
    
        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO Then gError 49589
    
        'Formata o Produto para BD
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 49590
    
        objEstoqueProduto.sProduto = sProdutoFormatado
        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
    
        lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 49991 Then gError 49591
    
        If lErro = SUCESSO Then
    
            lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
            If lErro <> SUCESSO Then gError 49600
    
            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True
    
            GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text
    
        End If

    End If
    
    Preenche_ContaContabil = SUCESSO

    Exit Function

Erro_Preenche_ContaContabil:

    Preenche_ContaContabil = gErr

    Select Case gErr

        Case 49589, 49590, 49591

        Case 49600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161911)

    End Select

    Exit Function

End Function

Public Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long
 
Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sAlmoxarifadoPadrao As String

On Error GoTo Erro_ProdutoLinha_Preenche

    'preenche a conta de ajuste com o que vem da tabela de MnemonicoCTBValor
    lErro = Preenche_ContaAjuste()
    If lErro <> SUCESSO Then gError 52080

    'Unidade de Medida
    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = objProduto.sSiglaUMEstoque

    'Descricao
    GridItens.TextMatrix(GridItens.Row, iGrid_DescricaoItem_Col) = objProduto.sDescricao

    'Almoxarifado
    '(Utiliza Almoxarifado Padrão caso esteja preenchido)
    If Len(Trim(AlmoxPadrao.ClipText)) > 0 Then
    
        lErro = CF("EstoqueProduto_TestaAssociacao", Produto.Text, AlmoxPadrao)
        If lErro = SUCESSO Then
            GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = AlmoxPadrao.Text
        Else
            GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = ""
        End If
        
    Else
        
        'le o Nome reduzido do almoxarifado Padrão do Produto em Questão
        lErro = CF("AlmoxarifadoPadrao_Le_NomeReduzido", objProduto.sCodigo, sAlmoxarifadoPadrao)
        If lErro <> SUCESSO Then gError 52242

        GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao
    
    End If

    'disp nossa vai ser sempre a 1a entao o indice vai ser sempre zero
     GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = Tipo.List(0)

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGrid.iLinhasExistentes Then
    
        If objGrid.iLinhasExistentes + 1 = GridItens.Rows Then
            GridItens.Rows = GridItens.Rows + NUM_MAX_ITENS_INVENTARIO
        End If
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 52080, 52242

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161912)

    End Select

    Exit Function

End Function

Public Sub BotaoEstoque_Click()

Dim lErro As Long
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoEstoque_Click

    If GridItens.Row = 0 Then gError 43714

    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41078

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado

        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 43735
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 41078

        Case 43714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43735
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161913)

    End Select

    Exit Sub

End Sub

Public Sub objEventoEstoque_evselecao(obj1 As Object)

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEstoqueProduto As ClassEstoqueProduto
Dim sCodProduto As String

On Error GoTo Erro_objEventoEstoque_evselecao

    Set objEstoqueProduto = obj1

    If (GridItens.Row > 0) Then

        sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

        lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 41079

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
                'Calcula a Quantidade Disponível nesse Almoxarifado
                lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
                If lErro <> SUCESSO Then gError 41080
            Else
                'Calcula a Quantidade Disponível
                lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78908
            End If

            'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
            Call Preenche_ContaContabil
            
        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 41079, 41080, 78908

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161914)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objInventario As ClassInventario) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objInventario Is Nothing Then

        'Lê o Inventario no Banco de Dados
        lErro = CF("InventarioPendente_Le", objInventario)
        If lErro <> SUCESSO And lErro <> 41157 Then gError 41081

        If lErro = SUCESSO Then 'Se ele existe

            lErro = Preenche_Tela(objInventario)
            If lErro <> SUCESSO Then gError 41082

        ElseIf lErro = 41157 Then

            'Se ele não existe exibe apenas o código
            Codigo.Text = objInventario.sCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 41081, 41082

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161915)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub AlmoxPadrao_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado

On Error GoTo Erro_AlmoxPadrao_Validate

    'Verifica se Almoxarifado Padrão foi informado
    If Len(Trim(AlmoxPadrao.ClipText)) > 0 Then

        lErro = TP_Almoxarifado_Filial_Le(AlmoxPadrao, objAlmoxarifado, 0)
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 41083

        If lErro = 25136 Then gError 41084

        If lErro = 25143 Then gError 41085

    End If

    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True
    
    Select Case gErr

        Case 41083

        Case 41084, 41085
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161916)

    End Select

    Exit Sub

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 41086

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True


    Select Case gErr

        Case 41086

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161917)

    End Select

    Exit Sub

End Sub

'hora
Public Sub Hora_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(HORA, iAlterado)

End Sub

'hora
Public Sub Hora_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

'hora
Public Sub Hora_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Hora_Validate

    'Verifica se a hora foi digitada
    If Len(Trim(HORA.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Hora_Critica(HORA.Text)
    If lErro <> SUCESSO Then gError 89804

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 89804

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161918)

    End Select

    Exit Sub

End Sub

Public Sub Tipo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 41087

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 41087

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161919)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 41088

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 41088

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161920)

    End Select

    Exit Sub

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)


Dim lErro As Long
Dim iIndice As Integer
Dim sUnidadeMed As String
Dim sCodProduto As String
Dim objProduto As New ClassProduto
Dim objClasseUM As New ClassClasseUM
Dim objUnidadeDeMedida As ClassUnidadeDeMedida
Dim colSiglas As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridItens.TextMatrix(iLinha, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41089

    If objControl.Name = "Produto" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If

    ElseIf objControl.Name = "UnidadeMed" Then

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 41090

            If lErro = 28030 Then gError 41091

            objClasseUM.iClasse = objProduto.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 41092

            'Guardo o valor da Unidade de Medida da Linha
            sUnidadeMed = GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col)

            'Limpar as Unidades utilizadas anteriormente
            UnidadeMed.Clear

            For Each objUnidadeDeMedida In colSiglas
                UnidadeMed.AddItem objUnidadeDeMedida.sSigla

            Next

            'Tento selecionar na Combo a Unidade anterior
            If UnidadeMed.ListCount <> 0 Then
                For iIndice = 0 To UnidadeMed.ListCount - 1

                    If UnidadeMed.List(iIndice) = sUnidadeMed Then
                        UnidadeMed.ListIndex = iIndice
                        Exit For
                    End If
                Next
            End If

        End If

    ElseIf objControl.Name = "CustoUnitario" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO And (GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_DEMO Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_CONSERTO Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_OUTRAS Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_BENEF Or _
        GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_DISPONIVEL_NOSSA_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_RECEB_INDISP_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_OUTRAS_INDISP_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_DEFEIT_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_CONSIG_CI) Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = "Quantidade" Or objControl.Name = "Almoxarifado" Or objControl.Name = "Tipo" Or objControl.Name = "Etiqueta" Or objControl.Name = "ContaContabil" Or objControl.Name = "ContaAjuste" Then

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If

    ElseIf objControl.Name = "LoteProduto" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78490
    
            If lErro = 28030 Then gError 78491
        
            If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        Else
            objControl.Enabled = False
        End If
    
    ElseIf objControl.Name = "FilialOP" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78492
    
            If lErro = 28030 Then gError 78493
        
            If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        Else
            objControl.Enabled = False
        End If
    
    ElseIf objControl.Name = "Atualiza" Then
        
       If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
            objControl.Enabled = False
        Else
            objControl.Enabled = True
        End If
    
    End If

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case 41089, 41090, 41092, 78490, 78492, 92476

        Case 41091, 78491, 78493, 92477
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161921)

    End Select

    Exit Sub

End Sub

Public Function QuantEstoque_Calcula(sProduto As String, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objProduto As New ClassProduto
Dim dCustoMedioAtual As Double
Dim sTipo As String

On Error GoTo Erro_QuantEstoque_Calcula

    objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)

    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41094

    objProduto.sCodigo = sProdutoFormatado

    sTipo = GridItens.TextMatrix(iLinha, iGrid_Tipo_Col)

    If iLinha >= GridItens.FixedRows And Len(Trim(objAlmoxarifado.sNomeReduzido)) > 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then

        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 41095

        If lErro = 28030 Then gError 41096

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 41097

        If lErro = 25056 Then gError 41098

        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        objEstoqueProduto.sProduto = sProdutoFormatado

        'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 41099

        If lErro = 21306 Then gError 55469

        sUnidadeMed = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)

        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 41100

        Select Case sTipo

            Case STRING_QUANT_DISPONIVEL_NOSSA, STRING_QUANT_DISPONIVEL_NOSSA_CI
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque((objEstoqueProduto.dQuantDispNossa + objEstoqueProduto.dQuantReservada) / dFator)
            Case STRING_QUANT_RECEB_INDISP, STRING_QUANT_RECEB_INDISP_CI
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantRecIndl / dFator)
            Case STRING_QUANT_OUTRAS_INDISP, STRING_QUANT_OUTRAS_INDISP_CI
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantInd / dFator)
            Case STRING_QUANT_DEFEIT, STRING_QUANT_DEFEIT_CI
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantDefeituosa / dFator)
            Case STRING_QUANT_3_CONSIG, STRING_QUANT_3_CONSIG_CI
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque((objEstoqueProduto.dQuantConsig3 + objEstoqueProduto.dQuantReservadaConsig) / dFator)
            Case STRING_QUANT_3_DEMO
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantDemo3 / dFator)
            Case STRING_QUANT_3_CONSERTO
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantConserto3 / dFator)
            Case STRING_QUANT_3_OUTRAS
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantOutras3 / dFator)
            Case STRING_QUANT_3_BENEF
                GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantBenef3 / dFator)
            Case Else
                Error 55432
                
        End Select

        'Calcula a Quantidade diferença
        If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))) > 0 Then
            GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col) = Formata_Estoque(CDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) - CDbl(GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col)))

            'Se o custo unitário estiver preenchido então calcula a diferença
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col))) > 0 Then

                'Preenche o valor da diferenca com diferenca-quant*CustoUnitario
                GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = Format((GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)) * dFator * (GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col)), "Standard")

            Else
                GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = ""
            End If
        Else
            GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = ""
            GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col) = ""
        End If

    Else

        'Limpa a QuantidadeEstoque ,Quantidadediferenca e valorDiferenca da Tela
        GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = ""
        GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = ""

    End If

    QuantEstoque_Calcula = SUCESSO

    Exit Function

Erro_QuantEstoque_Calcula:

    QuantEstoque_Calcula = gErr

    Select Case gErr

        Case 41094, 41095, 41097, 41099, 41100

        Case 41096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 41098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 55432

        Case 55469
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_TEM_PRODUTO", gErr, objAlmoxarifado.sNomeReduzido, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161922)

    End Select

    Exit Function

End Function

Public Function QuantLote_Calcula(sProduto As String, ByVal iLinha As Integer, sLote As String, iFilialOP As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastroLoteSaldo As New ClassRastreamentoLoteSaldo
Dim objProduto As New ClassProduto
Dim dCustoMedioAtual As Double
Dim sTipo As String

On Error GoTo Erro_QuantLote_Calcula
    
    objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
    
    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41094
    
    objProduto.sCodigo = sProdutoFormatado
    
    sTipo = GridItens.TextMatrix(iLinha, iGrid_Tipo_Col)
    
    If iLinha >= GridItens.FixedRows And Len(Trim(objAlmoxarifado.sNomeReduzido)) > 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then
    
        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 41095
    
        If lErro = 28030 Then gError 41096
    
        If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Or (objProduto.iRastro = PRODUTO_RASTRO_OP And iFilialOP <> 0) Then 'Alterado por Wagner
                    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 41097
        
            If lErro = 25056 Then gError 41098
        
            objRastroLoteSaldo.iAlmoxarifado = objAlmoxarifado.iCodigo
            objRastroLoteSaldo.sProduto = sProdutoFormatado
            objRastroLoteSaldo.sLote = sLote
            objRastroLoteSaldo.iFilialOP = iFilialOP
            
            'Lê o Saldo do Produto no Lote Filial Passado
            lErro = CF("RastreamentoLoteSaldo_Le", objRastroLoteSaldo)
            If lErro <> SUCESSO And lErro <> 78633 Then gError 41099
        
            '??? If lErro = 78633 Then gError 55469 '?????
        
            sUnidadeMed = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
        
            lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
            If lErro <> SUCESSO Then gError 41100
        
            Select Case sTipo
        
                Case STRING_QUANT_DISPONIVEL_NOSSA, STRING_QUANT_DISPONIVEL_NOSSA_CI
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque((objRastroLoteSaldo.dQuantDispNossa + objRastroLoteSaldo.dQuantReservada) / dFator)
                Case STRING_QUANT_RECEB_INDISP, STRING_QUANT_RECEB_INDISP_CI
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastroLoteSaldo.dQuantRecIndl / dFator)
                Case STRING_QUANT_OUTRAS_INDISP, STRING_QUANT_OUTRAS_INDISP_CI
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastroLoteSaldo.dQuantIndOutras / dFator)
                Case STRING_QUANT_DEFEIT, STRING_QUANT_DEFEIT_CI
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastroLoteSaldo.dQuantDefeituosa / dFator)
                Case STRING_QUANT_3_CONSIG, STRING_QUANT_3_CONSIG_CI
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque((objRastroLoteSaldo.dQuantConsig3 + objRastroLoteSaldo.dQuantReservadaConsig) / dFator)
                Case STRING_QUANT_3_DEMO
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastroLoteSaldo.dQuantDemo3 / dFator)
                Case STRING_QUANT_3_CONSERTO
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastroLoteSaldo.dQuantConserto3 / dFator)
                Case STRING_QUANT_3_OUTRAS
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastroLoteSaldo.dQuantOutras3 / dFator)
                Case STRING_QUANT_3_BENEF
                    GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastroLoteSaldo.dQuantBenef3 / dFator)
                Case Else
                    Error 55432
                    
            End Select
    
            'Calcula a Quantidade diferença
            If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))) > 0 Then
                GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col) = Formata_Estoque(CDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) - CDbl(GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col)))
    
                'Se o custo unitário estiver preenchido então calcula a diferença
                If Len(Trim(GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col))) > 0 Then
    
                    'Preenche o valor da diferenca com diferenca-quant*CustoUnitario
                    GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = Format((GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)) * dFator * (GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col)), "Standard")
    
                Else
                    GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = ""
                End If
            Else
                GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = ""
                GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col) = ""
            End If

        End If
        
    Else

        'Limpa a QuantidadeEstoque ,Quantidadediferenca e valorDiferenca da Tela
        GridItens.TextMatrix(iLinha, iGrid_ValorDiferenca_col) = ""
        GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col) = ""
        GridItens.TextMatrix(iLinha, iGrid_QuantEstoque_Col) = ""

    End If

    QuantLote_Calcula = SUCESSO

    Exit Function

Erro_QuantLote_Calcula:

    QuantLote_Calcula = gErr

    Select Case gErr

        Case 41094, 41095, 41097, 41099, 41100

        Case 41096
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 41098
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 55432

        Case 55469
            'lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_TEM_PRODUTO_ALMOXARIFADO", gErr, objRastreamentoSaldo.sLote, objRastreamentoSaldo.sProduto, objRastreamentoSaldo.iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161923)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim dCusto As Double
Dim sTipo As String
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199347

        Produto.Text = sProduto

        lErro = CF("Produto_Critica_Estoque", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 41101

        If lErro = 25077 Then gError 41102

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 41103

            sTipo = GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col)

            'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
            lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
            If lErro <> SUCESSO Then gError 49607

            GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)
            
            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
                'Calcula a Quantidade Disponível
                lErro = QuantEstoque_Calcula(Produto.Text, GridItens.Row)
                If lErro <> SUCESSO Then gError 41104
            Else
                'Calcula a Quantidade Disponível
                lErro = QuantLote_Calcula(Produto.Text, GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78900
            End If
            
        End If

        If objProduto.iRastro = PRODUTO_RASTRO_OP Then
            GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41105
    
    Call Preenche_ContaContabil
                
    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 41101, 41103, 41104, 41105, 49607, 78900
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41102
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161924)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.ClipText)) > 0 Then

        lErro = Valor_Double_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 41106

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Quantidade.Text

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
        'Calcula a Quantidade Disponível
        lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
        If lErro <> SUCESSO Then gError 49617
    Else
        'Calcula a Quantidade Disponível
        lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78901
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41107

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 41106, 41107, 49617, 78901
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161925)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    If Len(UnidadeMed.Text) > 0 Then
    
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 55426
        Else
            'Calcula a Quantidade Disponível
            lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78902
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41108

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 41108, 55426, 78902
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161926)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ContaContabil

    Set objGrid.objControle = ContaContabil

    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", ContaContabil.Text, ContaContabil.ClipText, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49573

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49574

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaMascarada
            ContaContabil.PromptInclude = True

        'se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49575

            'conta não cadastrada
            If lErro = 5700 Then gError 49576

        End If

    Else

        ContaContabil.PromptInclude = False
        ContaContabil.Text = ""
        ContaContabil.PromptInclude = True

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49644

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = gErr

    Select Case gErr

        Case 49573, 49575, 49644
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49574
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49576
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161927)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.ClipText)) > 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 41109

        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 41110

        If lErro = 25157 Then gError 41111

        If lErro = 25162 Then gError 41112

        GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 41113
        Else
            'Calcula a Quantidade Disponível
            lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78904
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41114
    
    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    Call Preenche_ContaContabil
    
    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 41109, 41110, 41113, 41114, 78904
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41111

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 41112

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161928)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_CustoUnitario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sTipo As String

On Error GoTo Erro_Saida_Celula_CustoUnitario

    Set objGridInt.objControle = CustoUnitario

    If Len(Trim(CustoUnitario.ClipText)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(CustoUnitario.Text)
        If lErro <> SUCESSO Then gError 41115

    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = CustoUnitario.FormattedText

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
        'Calcula a Quantidade Disponível
        lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
        If lErro <> SUCESSO Then gError 49607
    Else
        'Calcula a Quantidade Disponível
        lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78905
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41116

    Saida_Celula_CustoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitario:

    Saida_Celula_CustoUnitario = gErr

    Select Case gErr

        Case 41115, 41116, 49607, 78905
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161929)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Etiqueta(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Etiqueta

    Set objGridInt.objControle = Etiqueta

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41117

    Saida_Celula_Etiqueta = SUCESSO

    Exit Function

Erro_Saida_Celula_Etiqueta:

    Saida_Celula_Etiqueta = gErr

    Select Case gErr

        Case 41117

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161930)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_Tipo(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sTipo As String
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim dCusto As Double

On Error GoTo Erro_Saida_Celula_Tipo

    Set objGridInt.objControle = Tipo

    sTipo = Tipo.Text

    GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = Tipo.Text
    
    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 49623

    If iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then

        objProduto.sCodigo = sProdutoFormatado

        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 49608

        'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
        lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
        If lErro <> SUCESSO Then gError 49609

        GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)
        
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 49650
        Else
            'Calcula a Quantidade Disponível
            lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78906
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 41118

    Saida_Celula_Tipo = SUCESSO

    Exit Function

Erro_Saida_Celula_Tipo:

    Saida_Celula_Tipo = gErr

    Select Case gErr

        Case 41118, 49608, 49609, 49623, 49650, 78906
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161931)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 41119

    Call Limpa_Tela_Inventario

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 41119

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161932)

    End Select

    Exit Sub

End Sub

Public Sub Limpa_Tela_Inventario()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Inventario

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Set gobjAnotacao = Nothing

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_Inventario:

    Select Case gErr

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161933)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objInventario As New ClassInventario
Dim sTipo As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 41120

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 41121

    'Verifica se Lote foi preenchido
    If Len(Trim(Lote.ClipText)) = 0 Then gError 41122

    'Verifica se há Algum Item de Movimento de Estoque Informado no GridItens
    If objGrid.iLinhasExistentes = 0 Then gError 41123

    objInventario.sCodigo = Codigo.Text

    'Para cada linha do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 41126

        'Verifica se o Almoxarifado foi informado
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 41127

        sTipo = GridItens.TextMatrix(iIndice, iGrid_Tipo_Col)
        If Len(Trim(sTipo)) = 0 Then gError 41128

        If sTipo = STRING_QUANT_3_DEMO Or sTipo = STRING_QUANT_3_CONSERTO Or sTipo = STRING_QUANT_3_OUTRAS Or sTipo = STRING_QUANT_3_BENEF Or _
           sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI Or sTipo = STRING_QUANT_RECEB_INDISP_CI Or sTipo = STRING_QUANT_OUTRAS_INDISP_CI Or sTipo = STRING_QUANT_DEFEIT_CI Or sTipo = STRING_QUANT_3_CONSIG_CI Then
            
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))) = 0 Then gError 41129
        End If

        'verifica se existe no inventario algum lançamento do mesmo produto, almoxarifado e tipo
        lErro = Verifica_Item_Repetido(iIndice)
        If lErro <> SUCESSO Then gError 55449

    Next

    'Move os dados da tela para o objInventario
    lErro = Move_Tela_Memoria(objInventario)
    If lErro <> SUCESSO Then gError 41130

    'Grava no BD os dados
    lErro = CF("InventarioPendente_Grava", objInventario)
    If lErro <> SUCESSO Then gError 41131

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_INVENTARIOLOTE
            gobjAnotacao.sID = CStr(objInventario.iFilialEmpresa) & "," & objInventario.sCodigo
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 41131
            
        End If
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 41120
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 41121
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 41122
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iIndice)

        Case 41123
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INVENTARIO_NAO_INFORMADO", gErr)

        Case 41126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 41127
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 41128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", gErr, iIndice)

        Case 41129
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOUNITARIO_NAO_PREENCHIDO", gErr, iIndice)

        Case 41130, 41131, 55449

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161934)

    End Select

    Exit Function

End Function

Public Function Move_Tela_Memoria(objInventario As ClassInventario) As Long
'Preenche objInventario (inclusive colItens)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(Codigo.Text)) > 0 Then objInventario.sCodigo = Codigo.Text

    If Len(Trim(Data.ClipText)) > 0 Then
        objInventario.dtData = CDate(Data.Text)
    Else
        objInventario.dtData = DATA_NULA
    End If

'hora
    If Len(Trim(HORA.ClipText)) > 0 Then
        objInventario.dtHora = CDate(HORA.Text)
    Else
        objInventario.dtHora = Time
    End If

    If Len(Trim(Lote.ClipText)) > 0 Then objInventario.iLote = Lote.Text

    objInventario.iFilialEmpresa = giFilialEmpresa

    For iIndice = 1 To objGrid.iLinhasExistentes

        lErro = Move_Itens_Memoria(iIndice, objInventario)
        If lErro <> SUCESSO Then gError 41132

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 41132

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161935)

    End Select

    Exit Function

End Function

Public Function Move_Itens_Memoria(iIndice As Integer, objInventario As ClassInventario) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sTipo As String, iTipo As Integer
Dim dCusto As Double, sEtiqueta As String
Dim objItemInventario As New ClassItemInventario
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim sContaAjusteFormatada As String
Dim sLote As String
Dim iFilialOP As Integer
Dim objProduto As New ClassProduto
Dim iIndice1 As Integer

On Error GoTo Erro_Move_Itens_Memoria

    'Pega o Código do Almoxarifado
    objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO And lErro <> 25056 Then gError 41133

    If lErro = 25056 Then gError 41134

    'Formata o Produto para BD
    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 41135

    'Formata conta para o BD
    If Len(GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col)) > 0 Then

            'Guarda a conta corrente
            lErro = CF("Conta_Formata", GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col), sContaFormatada, iContaPreenchida)
            If lErro <> SUCESSO Then gError 52077

    Else
        sContaFormatada = ""
    End If

    'Formata conta Ajuste para o BD
    If Len(GridItens.TextMatrix(iIndice, iGrid_ContaAjuste_Col)) > 0 Then

        'Guarda a conta Ajuste de estoque
        lErro = CF("Conta_Formata", GridItens.TextMatrix(iIndice, iGrid_ContaAjuste_Col), sContaAjusteFormatada, iContaPreenchida)
        If lErro <> SUCESSO Then gError 52076
    Else
        sContaAjusteFormatada = ""
    End If

    sTipo = GridItens.TextMatrix(iIndice, iGrid_Tipo_Col)
    For iIndice1 = 0 To Tipo.ListCount - 1
        If Tipo.List(iIndice1) = sTipo Then
            iTipo = Tipo.ItemData(iIndice1)
            Exit For
        End If
    Next

    If sTipo = STRING_QUANT_3_DEMO Or sTipo = STRING_QUANT_3_CONSERTO Or sTipo = STRING_QUANT_3_OUTRAS Or sTipo = STRING_QUANT_3_BENEF Or _
       sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI Or sTipo = STRING_QUANT_RECEB_INDISP_CI Or sTipo = STRING_QUANT_OUTRAS_INDISP_CI Or sTipo = STRING_QUANT_DEFEIT_CI Or sTipo = STRING_QUANT_3_CONSIG_CI Then
    
        dCusto = CDbl(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))
    End If

    objProduto.sCodigo = sProdutoFormatado
    
    'Le o Produto para verificar que Tipo de Rastro ele possui
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 78471
    
    'Se não encontrou o produto --> Erro
    If lErro = 28030 Then gError 78472
    
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then
        
        sLote = GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col)
        
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
    
        sLote = GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col)
        iFilialOP = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_FilialOP_Col))
    
    End If

    Set objItemInventario = objInventario.colItens.Add(0, sProdutoFormatado, GridItens.TextMatrix(iIndice, iGrid_DescricaoItem_Col), GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)), dCusto, objAlmoxarifado.iCodigo, GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col), iTipo, GridItens.TextMatrix(iIndice, iGrid_Etiqueta_Col), sContaFormatada, sContaAjusteFormatada, sLote, iFilialOP)

    objItemInventario.dQuantEst = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col))

    If Len(Trim(sLote)) > 0 Then

        For iIndice1 = 0 To Atualiza.ListCount - 1
            If Atualiza.List(iIndice1) = GridItens.TextMatrix(iIndice, iGrid_Atualiza_Col) Then
                objItemInventario.iAtualizaSoLote = Atualiza.ItemData(iIndice1)
                Exit For
            End If
        Next
    
    End If
    
    Move_Itens_Memoria = SUCESSO

    Exit Function

Erro_Move_Itens_Memoria:

    Move_Itens_Memoria = gErr

    Select Case gErr

        Case 41133, 41135, 52076, 52077, 78471

        Case 41134
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 78472
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 78473, 78474
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTREAMENTO_NAO_PREENCHIDO", gErr, iIndice)
        
        Case 78475
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161936)

    End Select

    Exit Function

End Function

Public Function Excluir_Registro() As Long

Dim lErro As Long
Dim iAchou As Integer
Dim iIndice As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objInventario As New ClassInventario
Dim sTipo As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Excluir_Registro

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 41204

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 41205

    'Verifica se Lote foi preenchida
    If Len(Trim(Lote.ClipText)) = 0 Then gError 41206

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridItens
    If objGrid.iLinhasExistentes = 0 Then gError 41207

    objInventario.sCodigo = Codigo.Text
    objInventario.iLote = StrParaInt(Lote.Text)
    
    'Pede ao usuário que confirme a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_INVLOTE", objInventario.sCodigo)
    If vbMsgRes = vbNo Then gError 57765

    lErro = CF("InventarioPendente_Le", objInventario)
    If lErro <> SUCESSO And lErro <> 41157 Then gError 41208

    'Se não encontrar
    If lErro = 41157 Then gError 41209

    'Exclui o Registro
    lErro = CF("InventarioPendente_Exclui", objInventario)
    If lErro <> SUCESSO Then gError 41210

    Excluir_Registro = SUCESSO

    Exit Function

Erro_Excluir_Registro:

    Excluir_Registro = gErr

    Select Case gErr

        Case 41204
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 41205
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 41206
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_PREENCHIDO", gErr)

        Case 41207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INVENTARIO_NAO_INFORMADO", gErr)

        Case 41208, 41210, 57765

        Case 41209
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INV_PENDENTE_NAO_CADASTRADO", gErr, objInventario.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161937)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 41136

    Call Limpa_Tela_Inventario

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 41136

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161938)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim lErro As Long '

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
        
    lErro = Excluir_Registro()
    If lErro <> SUCESSO Then gError 41137

    Call Limpa_Tela_Inventario

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 41137

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161939)

    End Select

    Exit Sub

End Sub

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

    Set objGrid = Nothing

    Set objEventoCodigo = Nothing
    Set objEventoLoteInv = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoAlmoxPadrao = Nothing
    Set objEventoRastroLote = Nothing

    Set objEventoContaContabil = Nothing
    
    Set gobjAnotacao = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Almoxarifado_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub AlmoxPadrao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CustoUnitario_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Etiqueta_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Lote_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)

    End If

End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)

Dim iLinhasExistentesAnterior As Integer
Dim iLinhaAnterior As Integer

    'Verifica se a Tecla apertada foi Del
    If KeyCode = vbKeyDelete Then
        'Guarda iLinhasExistentes
        iLinhasExistentesAnterior = objGrid.iLinhasExistentes

        'Guarda o índice da Linha a ser Excluída
        iLinhaAnterior = GridItens.Row

    End If

    Call Grid_Trata_Tecla1(KeyCode, objGrid)

End Sub

Public Sub GridItens_EnterCell()

    Call Grid_Entrada_Celula(objGrid, iAlterado)

End Sub

Public Sub GridItens_GotFocus()

    Call Grid_Recebe_Foco(objGrid)

End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        If objGridInt.objGrid Is GridItens Then

            Select Case GridItens.Col

                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 41138

                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 41139

                Case iGrid_Almoxarifado_Col
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 41140

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 41141

                Case iGrid_Etiqueta_Col
                    lErro = Saida_Celula_Etiqueta(objGridInt)
                    If lErro <> SUCESSO Then gError 41142

                Case iGrid_CustoUnitario_Col
                    lErro = Saida_Celula_CustoUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 41143

                Case iGrid_Tipo_Col
                    lErro = Saida_Celula_Tipo(objGridInt)
                    If lErro <> SUCESSO Then gError 41144

                Case iGrid_ContaContabil_Col
                    lErro = Saida_Celula_ContaContabil(objGridInt)
                    If lErro <> SUCESSO Then gError 49577

                Case iGrid_ContaAjuste_Col
                    lErro = Saida_Celula_ContaAjuste(objGridInt)
                    If lErro <> SUCESSO Then gError 52077
                
                Case iGrid_LoteProduto_Col
                    lErro = Saida_Celula_LoteProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 78577

                Case iGrid_FilialOP_Col
                    lErro = Saida_Celula_FilialOP(objGridInt)
                    If lErro <> SUCESSO Then gError 78578
                
                Case iGrid_Atualiza_Col
                    lErro = Saida_Celula_Atualiza(objGridInt)
                    If lErro <> SUCESSO Then gError 90832
             
            End Select

        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 41146

        iAlterado = REGISTRO_ALTERADO

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 39640, 41138, 41139, 41140, 41141, 41142, 41143, 41144, 78577, 78578, 90832

        Case 41145, 49577, 52077

        Case 41146
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161940)

    End Select

    Exit Function

End Function

Public Sub GridItens_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid)
    
End Sub

Public Sub GridItens_Scroll()

    Call Grid_Scroll(objGrid)

End Sub

Public Sub Almoxarifado_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Almoxarifado_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Almoxarifado_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Almoxarifado
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Produto_GotFocus()

Dim lErro As Long

    Call Grid_Campo_Recebe_Foco(objGrid)

    If gobjEST.iInventarioCodBarrAuto = 1 Then

        If objGrid.lErroSaidaCelula = 0 Then

            lErro = CF2(Me, "INV_Trata_CodigoBarras1")

            Call Grid_Entrada_Celula(objGrid, iAlterado)

            If lErro <> SUCESSO Then
    
                objGrid.lErroSaidaCelula = 1
            End If

        Else
    
            objGrid.lErroSaidaCelula = 0
    
        End If
        
    End If

End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ContaContabil_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub ContaContabil_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub ContaContabil_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = ContaContabil
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Quantidade_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Quantidade_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Quantidade
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UnidadeMed_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub UnidadeMed_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub UnidadeMed_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = UnidadeMed
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub UnidadeMed_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub CustoUnitario_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub CustoUnitario_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub CustoUnitario_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoUnitario
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Etiqueta_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Etiqueta_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Etiqueta_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Etiqueta
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Tipo_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub Tipo_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub Tipo_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Tipo
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Atualiza_Custo_Inventario(ByVal iLinha As Integer, objItemInventario As ClassItemInventario) As Long
'Verifica o Tipo, e verifica se produto é produzido ou comprado
'De acordo com isso prenche o custo com o custo medio ou com o que esta no bd
'e calcula a diferenca

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dCusto As Double
Dim sTipo As String

On Error GoTo Erro_Atualiza_Custo_Inventario

    objProduto.sCodigo = objItemInventario.sProduto

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO Then gError 49640
    
    Select Case objItemInventario.iTipo
    
        Case TIPO_QUANT_DISPONIVEL_NOSSA
            sTipo = STRING_QUANT_DISPONIVEL_NOSSA
            
        Case TIPO_QUANT_RECEB_INDISP
            sTipo = STRING_QUANT_RECEB_INDISP
            
        Case TIPO_QUANT_OUTRAS_INDISP
            sTipo = STRING_QUANT_OUTRAS_INDISP
            
        Case TIPO_QUANT_DEFEIT
            sTipo = STRING_QUANT_DEFEIT
            
        Case TIPO_QUANT_3_CONSIG
            sTipo = STRING_QUANT_3_CONSIG
            
        Case TIPO_QUANT_3_DEMO
            sTipo = STRING_QUANT_3_DEMO
            
        Case TIPO_QUANT_3_CONSERTO
            sTipo = STRING_QUANT_3_CONSERTO
        
        Case TIPO_QUANT_3_OUTRAS
            sTipo = STRING_QUANT_3_OUTRAS

        Case TIPO_QUANT_3_BENEF
            sTipo = STRING_QUANT_3_BENEF
                        
        Case TIPO_QUANT_DISPONIVEL_NOSSA_CI
            sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI
            
        Case TIPO_QUANT_RECEB_INDISP_CI
            sTipo = STRING_QUANT_RECEB_INDISP_CI
            
        Case TIPO_QUANT_OUTRAS_INDISP_CI
            sTipo = STRING_QUANT_OUTRAS_INDISP_CI
            
        Case TIPO_QUANT_DEFEIT_CI
            sTipo = STRING_QUANT_DEFEIT_CI
            
        Case TIPO_QUANT_3_CONSIG_CI
            sTipo = STRING_QUANT_3_CONSIG_CI
                        
    End Select
    
    If objItemInventario.iTipo = TIPO_QUANT_DISPONIVEL_NOSSA_CI Or objItemInventario.iTipo = TIPO_QUANT_RECEB_INDISP_CI Or objItemInventario.iTipo = TIPO_QUANT_OUTRAS_INDISP_CI Or objItemInventario.iTipo = TIPO_QUANT_DEFEIT_CI Or objItemInventario.iTipo = TIPO_QUANT_3_CONSIG_CI Then
        dCusto = objItemInventario.dCusto
    End If
    
    'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
    lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
    If lErro <> SUCESSO Then gError 49641

    GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)

    Atualiza_Custo_Inventario = SUCESSO

    Exit Function

Erro_Atualiza_Custo_Inventario:

    Atualiza_Custo_Inventario = gErr

    Select Case gErr

        Case 49640, 49641

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161941)

    End Select

    Exit Function

End Function

'??? excluido 24/04/00 Jones
''''inicio contabilidade
'''
'''Private Sub CTBBotaoModeloPadrao_Click()
'''
'''    Call objContabil.Contabil_BotaoModeloPadrao_Click
'''
'''End Sub
'''
'''Private Sub CTBModelo_Click()
'''
'''    Call objContabil.Contabil_Modelo_Click
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_Click()
'''
'''    Call objContabil.Contabil_GridContabil_Click
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_EnterCell()
'''
'''    Call objContabil.Contabil_GridContabil_EnterCell
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_GotFocus()
'''
'''    Call objContabil.Contabil_GridContabil_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_GridContabil_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_KeyDown(KeyCode As Integer, Shift As Integer)
'''
'''    Call objContabil.Contabil_GridContabil_KeyDown(KeyCode)
'''
'''End Sub
'''
'''
'''Private Sub CTBGridContabil_LeaveCell()
'''
'''        Call objContabil.Contabil_GridContabil_LeaveCell
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_GridContabil_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_RowColChange()
'''
'''    Call objContabil.Contabil_GridContabil_RowColChange
'''
'''End Sub
'''
'''Private Sub CTBGridContabil_Scroll()
'''
'''    Call objContabil.Contabil_GridContabil_Scroll
'''
'''End Sub
'''
'''Private Sub CTBConta_Change()
'''
'''    Call objContabil.Contabil_Conta_Change
'''
'''End Sub
'''
'''Private Sub CTBConta_GotFocus()
'''
'''    Call objContabil.Contabil_Conta_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBConta_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_Conta_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBConta_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_Conta_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBCcl_Change()
'''
'''    Call objContabil.Contabil_Ccl_Change
'''
'''End Sub
'''
'''Private Sub CTBCcl_GotFocus()
'''
'''    Call objContabil.Contabil_Ccl_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBCcl_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_Ccl_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBCcl_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_Ccl_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBCredito_Change()
'''
'''    Call objContabil.Contabil_Credito_Change
'''
'''End Sub
'''
'''Private Sub CTBCredito_GotFocus()
'''
'''    Call objContabil.Contabil_Credito_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBCredito_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_Credito_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBCredito_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_Credito_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBDebito_Change()
'''
'''    Call objContabil.Contabil_Debito_Change
'''
'''End Sub
'''
'''Private Sub CTBDebito_GotFocus()
'''
'''    Call objContabil.Contabil_Debito_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBDebito_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_Debito_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBDebito_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_Debito_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBSeqContraPartida_Change()
'''
'''    Call objContabil.Contabil_SeqContraPartida_Change
'''
'''End Sub
'''
'''Private Sub CTBSeqContraPartida_GotFocus()
'''
'''    Call objContabil.Contabil_SeqContraPartida_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBSeqContraPartida_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_SeqContraPartida_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBSeqContraPartida_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_SeqContraPartida_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)
'''
'''    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)
'''
'''End Sub
'''
'''Private Sub CTBHistorico_Change()
'''
'''    Call objContabil.Contabil_Historico_Change
'''
'''End Sub
'''
'''Private Sub CTBHistorico_GotFocus()
'''
'''    Call objContabil.Contabil_Historico_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBHistorico_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_Historico_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBHistorico_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_Historico_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBAglutina_GotFocus()
'''
'''    Call objContabil.Contabil_Aglutina_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBAglutina_KeyPress(KeyAscii As Integer)
'''
'''    Call objContabil.Contabil_Aglutina_KeyPress(KeyAscii)
'''
'''End Sub
'''
'''Private Sub CTBAglutina_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_Aglutina_Validate(Cancel)
'''
'''End Sub
'''
'''Private Sub CTBTvwContas_NodeClick(ByVal Node As MSComctlLib.Node)
'''
'''    Call objContabil.Contabil_TvwContas_NodeClick(Node)
'''
'''End Sub
'''
'''Private Sub CTBTvwCcls_NodeClick(ByVal Node As MSComctlLib.Node)
'''
'''    Call objContabil.Contabil_TvwCcls_NodeClick(Node)
'''
'''End Sub
'''
'''Private Sub CTBListHistoricos_DblClick()
'''
'''    Call objContabil.Contabil_ListHistoricos_DblClick
'''
'''End Sub
'''
'''Private Sub CTBBotaoLimparGrid_Click()
'''
'''    Call objContabil.Contabil_Limpa_GridContabil
'''
'''End Sub
'''
'''Private Sub CTBLote_Change()
'''
'''    Call objContabil.Contabil_Lote_Change
'''
'''End Sub
'''
'''Private Sub CTBLote_GotFocus()
'''
'''    Call objContabil.Contabil_Lote_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBLote_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_Lote_Validate(Cancel, Parent)
'''
'''End Sub
'''
'''Private Sub CTBDataContabil_Change()
'''
'''    Call objContabil.Contabil_DataContabil_Change
'''
'''End Sub
'''
'''Private Sub CTBDataContabil_GotFocus()
'''
'''    Call objContabil.Contabil_DataContabil_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBDataContabil_Validate(Cancel As Boolean)
'''
'''    Call objContabil.Contabil_DataContabil_Validate(Cancel, Parent)
'''
'''End Sub
'''
'''Private Sub objEventoLote_evSelecao(obj1 As Object)
''''traz o lote selecionado para a tela
'''
'''    Call objContabil.Contabil_objEventoLote_evSelecao(obj1)
'''
'''End Sub
'''
'''Private Sub objEventoDoc_evSelecao(obj1 As Object)
'''
'''    Call objContabil.Contabil_objEventoDoc_evSelecao(obj1)
'''
'''End Sub
'''
'''Private Sub CTBDocumento_Change()
'''
'''    Call objContabil.Contabil_Documento_Change
'''
'''End Sub
'''
'''Private Sub CTBDocumento_GotFocus()
'''
'''    Call objContabil.Contabil_Documento_GotFocus
'''
'''End Sub
'''
'''Private Sub CTBBotaoImprimir_Click()
'''
'''    Call objContabil.Contabil_BotaoImprimir_Click
'''
'''End Sub
'''
'''Private Sub CTBUpDown_DownClick()
'''
'''    Call objContabil.Contabil_UpDown_DownClick
'''
'''End Sub
'''
'''Private Sub CTBUpDown_UpClick()
'''
'''    Call objContabil.Contabil_UpDown_UpClick
'''
'''End Sub
'''
'''Private Sub CTBLabelDoc_Click()
'''
'''    Call objContabil.Contabil_LabelDoc_Click
'''
'''End Sub
'''
'''Private Sub CTBLabelLote_Click()
'''
'''    Call objContabil.Contabil_LabelLote_Click
'''
'''End Sub
'''
'''Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long
'''
'''Dim lErro As Long
'''Dim iLinha As Integer
'''Dim dQuantidadeConvertida As Double
'''Dim dQuantidade As Double
'''Dim sProduto As String
'''Dim sUm As String
'''
'''On Error GoTo Erro_Calcula_Mnemonico
'''
'''    Select Case objMnemonicoValor.sMnemonico
'''
'''        Case QUANT_ESTOQUE
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)) > 0 Then
'''
'''                    lErro = CF("UMEstoque_Conversao",GridItens.TextMatrix(iLinha, iGrid_Produto_Col), GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)), dQuantidadeConvertida)
'''                    If lErro <> SUCESSO then gError 64207
'''
'''                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
'''                Else
'''                    objMnemonicoValor.colValor.Add 0
'''                End If
'''            Next
'''
'''        Case CODIGO1
'''            If Len(Codigo.Text) > 0 Then
'''                objMnemonicoValor.colValor.Add Codigo.Text
'''            Else
'''                objMnemonicoValor.colValor.Add ""
'''            End If
'''
'''        Case DATA1
'''            If Len(Data.ClipText) > 0 Then
'''                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
'''            Else
'''                objMnemonicoValor.colValor.Add DATA_NULA
'''            End If
'''
'''        Case PRODUTO1
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''        Case UNIDADE_MED
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''        Case QUANTIDADE1
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
'''
'                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
'                    sProduto = GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
'                    sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
'                    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
'
'                    'Converte a quantidade para UM padrão estoque
'                    lErro = CF("UMEstoque_Conversao",sProduto, sUM, dQuantidade, dQuantidadeConvertida)
'                    If lErro <> SUCESSO Then gError 79943
'
'                    objMnemonicoValor.colValor.Add dQuantidadeConvertida

'''                Else
'''                    objMnemonicoValor.colValor.Add 0
'''                End If
'''            Next
'''
'''        Case QUANT_DIFERENCA
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col))
'''                Else
'''                    objMnemonicoValor.colValor.Add 0
'''                End If
'''            Next
'''
'''        Case DESCRICAO_ITEM
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_DescricaoItem_Col)) > 0 Then
'''                objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_DescricaoItem_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''        Case ALMOXARIFADO1
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)) > 0 Then
'''                objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''        Case TIPO1
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_Tipo_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Tipo_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''        Case CUSTO_UNITARIO
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col))
'''                Else
'''                    objMnemonicoValor.colValor.Add 0
'''                End If
'''            Next
'''
'''        Case ETIQUETA1
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_Etiqueta_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Etiqueta_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''         Case CONTACONTABIL1
'''
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_ContaContabil_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_ContaContabil_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''        Case CONTA_AJUSTE_INV
'''
'''            For iLinha = 1 To objGrid.iLinhasExistentes
'''                If Len(GridItens.TextMatrix(iLinha, iGrid_ContaAjuste_Col)) > 0 Then
'''                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_ContaAjuste_Col)
'''                Else
'''                    objMnemonicoValor.colValor.Add ""
'''                End If
'''            Next
'''
'''        Case Else
'''            Error 39655
'''
'''    End Select
'''
'''    Calcula_Mnemonico = SUCESSO
'''
'''    Exit Function
'''
'''Erro_Calcula_Mnemonico:
'''
'''    Calcula_Mnemonico = gErr
'''
'''    Select Case gErr
'''
'''        Case 39655
'''            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
'''
'''        Case 64207
'''
'''        Case Else
'''            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161942)
'''
'''    End Select
'''
'''    Exit Function
'''
'''End Function

Public Function Verifica_Item_Repetido(ByVal iIndice) As Long
'verifica se existem itens dentro do grid com o mesmo produto, almoxarifado e tipo do item iIndice

Dim iIndice1 As Integer
Dim sProdutoFormatado As String
Dim sProdutoFormatado1 As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Verifica_Item_Repetido

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 55450

    For iIndice1 = 1 To objGrid.iLinhasExistentes
    
        If iIndice <> iIndice1 Then
    
            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice1, iGrid_Produto_Col), sProdutoFormatado1, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 55451
    
            If sProdutoFormatado = sProdutoFormatado1 And GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = GridItens.TextMatrix(iIndice1, iGrid_Almoxarifado_Col) And GridItens.TextMatrix(iIndice, iGrid_Tipo_Col) = GridItens.TextMatrix(iIndice1, iGrid_Tipo_Col) Then gError 55452
    
        End If
        
    Next
    
    Verifica_Item_Repetido = SUCESSO

    Exit Function

Erro_Verifica_Item_Repetido:

    Verifica_Item_Repetido = gErr

    Select Case gErr

        Case 55450, 55451

        Case 55452
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_INVENTARIO_REPETIDO", gErr, iIndice, iIndice1)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161943)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_INVENTARIO_LOTES_LANCAMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Inventário - Lote"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "InventarioLote"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub Unload(objme As Object)

   RaiseEvent Unload

End Sub

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled() = New_Enabled
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'**** fim do trecho a ser copiado *****


Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
Dim lErro As Long

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is AlmoxPadrao Then
            Call AlmoxPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is ContaContabil Or Me.ActiveControl Is ContaAjuste Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Lote Then
            Call LoteLabel_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        ElseIf Me.ActiveControl Is LoteProduto Then
            Call BotaoLote_Click
        End If
    
    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        lErro = CF2(Me, "INV_Trata_CodigoBarras1")
    
    End If

End Sub

'Já existe em Rastreamento Lote
Function RastreamentoLote_Le(objRastroLote As ClassRastreamentoLote) As Long
'Lê rastreamento do lote a partir do produto, filialOP e código do lote passados

Dim lErro As Long
Dim lComando As Long
Dim tRastroLote As typeRastreamentoLote

On Error GoTo Erro_RastreamentoLote_Le

    'Abertura dos comandos
    lComando = Comando_Abrir()
    If lErro <> SUCESSO Then gError 75707

    tRastroLote.sObservacao = String(STRING_NOME, 0)

    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
    lErro = Comando_Executar(lComando, "SELECT DataValidade, DataEntrada, DataFabricacao, Observacao FROM RastreamentoLote WHERE Produto = ? AND Lote = ? AND FilialOP = ?", tRastroLote.dtDataValidade, tRastroLote.dtDataEntrada, tRastroLote.dtDataFabricacao, tRastroLote.sObservacao, objRastroLote.sProduto, objRastroLote.sCodigo, objRastroLote.iFilialOP)
    If lErro <> AD_SQL_SUCESSO Then gError 75708

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 75709

    'Se não encontrou, erro
    If lErro = AD_SQL_SEM_DADOS Then gError 75710

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    RastreamentoLote_Le = SUCESSO

    Exit Function

Erro_RastreamentoLote_Le:

    RastreamentoLote_Le = gErr
    
    Select Case gErr

        Case 75707
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case 75708, 75709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTE", gErr)

        Case 75710 'RastreamentoLote não cadastrado

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161944)

    End Select

    'Fechamento dos comandos
    Call Comando_Fechar(lComando)

    Exit Function

End Function

''Copiada da class CTInventario de GlobaisTelasEst
'Function RastreamentoLoteSaldo_Le(objRastroLoteSaldo As ClassRastreamentoLoteSaldo) As Long
''Lê a tabela de Rastreamento Lote Saldo
'
'Dim lErro As Long
'Dim lComando As Long
'Dim tRastroLoteSaldo As typeRastreamentoLoteSaldo
'
'On Error GoTo Erro_RastreamentoLoteSaldo_Le
'
'    'Abertura dos comandos
'    lComando = Comando_Abrir()
'    If lErro <> SUCESSO Then gError 78630
'
'    tRastroLoteSaldo.sProduto = String(STRING_PRODUTO, 0)
'
'    'Lê dados de RastrementoLote a partir de Produto, FilialOP e Lote
'    lErro = Comando_Executar(lComando, "SELECT RastreamentoLoteSaldo.Produto, Almoxarifado, NumIntDocLote, QuantDispNossa, QuantReservada, QuantReservadaConsig, QuantEmpenhada, QuantPedida, QuantRecIndl, QuantIndOutras, QuantDefeituosa, QuantConsig3, QuantConsig, QuantDemo3, QuantDemo, QuantConserto3, QuantConserto, QuantOutras3, QuantOutras, QuantOP, QuantBenef, QuantBenef3 FROM RastreamentoLoteSaldo, RastreamentoLote WHERE RastreamentoLote.NumIntDoc = RastreamentoLoteSaldo.NumIntDocLote AND RastreamentoLoteSaldo.Produto = ? AND Almoxarifado = ? AND RastreamentoLote.Lote = ? AND RastreamentoLote.FilialOP = ?", _
'    tRastroLoteSaldo.sProduto, tRastroLoteSaldo.iAlmoxarifado, tRastroLoteSaldo.lNumIntDocLote, tRastroLoteSaldo.dQuantDispNossa, tRastroLoteSaldo.dQuantReservada, tRastroLoteSaldo.dQuantReservadaConsig, tRastroLoteSaldo.dQuantEmpenhada, tRastroLoteSaldo.dQuantPedida, tRastroLoteSaldo.dQuantRecIndl, tRastroLoteSaldo.dQuantIndOutras, tRastroLoteSaldo.dQuantDefeituosa, tRastroLoteSaldo.dQuantConsig3, tRastroLoteSaldo.dQuantConsig, tRastroLoteSaldo.dQuantDemo3, tRastroLoteSaldo.dQuantDemo, tRastroLoteSaldo.dQuantConserto3, tRastroLoteSaldo.dQuantConserto, tRastroLoteSaldo.dQuantOutras3, tRastroLoteSaldo.dQuantOutras, tRastroLoteSaldo.dQuantOP, tRastroLoteSaldo.dQuantBenef, tRastroLoteSaldo.dQuantBenef3, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote, objRastroLoteSaldo.iFilialOP)
'    If lErro <> AD_SQL_SUCESSO Then gError 78631
'
'    lErro = Comando_BuscarPrimeiro(lComando)
'    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 78632
'
'    'Se não encontrou, erro
'    If lErro = AD_SQL_SEM_DADOS Then gError 78633
'
'    objRastroLoteSaldo.sProduto = tRastroLoteSaldo.sProduto
'    objRastroLoteSaldo.iAlmoxarifado = tRastroLoteSaldo.iAlmoxarifado
'    objRastroLoteSaldo.lNumIntDocLote = tRastroLoteSaldo.lNumIntDocLote
'    objRastroLoteSaldo.dQuantDispNossa = tRastroLoteSaldo.dQuantDispNossa
'    objRastroLoteSaldo.dQuantReservada = tRastroLoteSaldo.dQuantReservada
'    objRastroLoteSaldo.dQuantReservadaConsig = tRastroLoteSaldo.dQuantReservadaConsig
'    objRastroLoteSaldo.dQuantEmpenhada = tRastroLoteSaldo.dQuantEmpenhada
'    objRastroLoteSaldo.dQuantPedida = tRastroLoteSaldo.dQuantPedida
'    objRastroLoteSaldo.dQuantRecIndl = tRastroLoteSaldo.dQuantRecIndl
'    objRastroLoteSaldo.dQuantIndOutras = tRastroLoteSaldo.dQuantIndOutras
'    objRastroLoteSaldo.dQuantDefeituosa = tRastroLoteSaldo.dQuantDefeituosa
'    objRastroLoteSaldo.dQuantConsig3 = tRastroLoteSaldo.dQuantConsig3
'    objRastroLoteSaldo.dQuantConsig = tRastroLoteSaldo.dQuantConsig
'    objRastroLoteSaldo.dQuantDemo3 = tRastroLoteSaldo.dQuantDemo3
'    objRastroLoteSaldo.dQuantDemo = tRastroLoteSaldo.dQuantDemo
'    objRastroLoteSaldo.dQuantConserto3 = tRastroLoteSaldo.dQuantConserto3
'    objRastroLoteSaldo.dQuantConserto = tRastroLoteSaldo.dQuantConserto
'    objRastroLoteSaldo.dQuantOutras3 = tRastroLoteSaldo.dQuantOutras3
'    objRastroLoteSaldo.dQuantOutras = tRastroLoteSaldo.dQuantOutras
'    objRastroLoteSaldo.dQuantOP = tRastroLoteSaldo.dQuantOP
'    objRastroLoteSaldo.dQuantBenef = tRastroLoteSaldo.dQuantBenef
'    objRastroLoteSaldo.dQuantBenef3 = tRastroLoteSaldo.dQuantBenef3
'
'    'Fechamento dos comandos
'    Call Comando_Fechar(lComando)
'
'    RastreamentoLoteSaldo_Le = SUCESSO
'
'    Exit Function
'
'Erro_RastreamentoLoteSaldo_Le:
'
'    RastreamentoLoteSaldo_Le = gErr
'
'    Select Case gErr
'
'        Case 78630
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
'
'        Case 78631, 78632
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RASTREAMENTOLOTESALDO", gErr, objRastroLoteSaldo.sProduto, objRastroLoteSaldo.iAlmoxarifado, objRastroLoteSaldo.sLote)
'
'        Case 78633 'RastreamentoLoteSaldo não cadastrado
'
'        Case Else
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 161945)
'
'    End Select
'
'    'Fechamento dos comandos
'    Call Comando_Fechar(lComando)
'
'    Exit Function
'
'End Function

'################################################################
Public Sub BotaoLote_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim sLote As String

On Error GoTo Erro_BotaoLote_Click

    If (GridItens.Row = 0) Then gError 140223

    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)
    sLote = GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 140224

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 140225
    
    If Len(Trim(sLote)) > 0 Then
        objRastroLoteSaldo.sLote = sLote
    End If

    colSelecao.Add sProdutoFormatado

    Call Chama_Tela("RastroLoteSaldoLista", colSelecao, objRastroLoteSaldo, objEventoRastroLote, "Produto = ?")

    Exit Sub

Erro_BotaoLote_Click:

    Select Case gErr

        Case 140224
        
        Case 140223
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
                    
        Case 140225
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165459)

    End Select

    Exit Sub

End Sub

Public Sub objEventoRastroLote_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo
Dim objProduto As New ClassProduto

On Error GoTo Erro_objEventoRastroLote_evSelecao

    Set objRastroLoteSaldo = obj1

    If (GridItens.Row > 0) Then
        LoteProduto.Text = objRastroLoteSaldo.sLote
        GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col) = objRastroLoteSaldo.sLote
    End If

    Me.Show

    Exit Sub

Erro_objEventoRastroLote_evSelecao:

    Select Case gErr
    
        Case 140226 To 140228
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 165460)

    End Select

    Exit Sub

End Sub
'######################################################################

'###################################################
'Inserido por Wagner 23/02/2007
Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("Inventario_Automatico", sCodigo, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 185936

    Codigo.Text = sCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 185936
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185937)
    
    End Select

    Exit Sub

End Sub
'###################################################

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_INVENTARIOLOTE
    If Len(Trim(Codigo.Text)) > 0 Then
        objAnotacao.sID = CStr(giFilialEmpresa) & "," & Codigo.Text
    Else
        objAnotacao.sID = ""
        If Not (gobjAnotacao Is Nothing) Then
            objAnotacao.sTextoCompleto = gobjAnotacao.sTextoCompleto
            objAnotacao.sTitulo = gobjAnotacao.sTitulo
        End If
    End If
    
    Exit Sub
     
Erro_Anotacao_Extrai:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158093)
     
    End Select
     
    Exit Sub

End Sub

Public Sub Anotacao_Preenche(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Preenche

    'guarda o texto digitado
    Set gobjAnotacao = objAnotacao
        
    Exit Sub
     
Erro_Anotacao_Preenche:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158094)
     
    End Select
     
    Exit Sub

End Sub



