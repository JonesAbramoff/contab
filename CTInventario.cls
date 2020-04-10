VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTInventario"
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

Public gobjAnotacao As ClassAnotacoes

'inicio contabilidade

Dim objGrid1 As AdmGrid
Dim objContabil As New ClassContabil

Private WithEvents objEventoLote As AdmEvento
Attribute objEventoLote.VB_VarHelpID = -1
Private WithEvents objEventoDoc As AdmEvento
Attribute objEventoDoc.VB_VarHelpID = -1

'Mnemonicos
Private Const CODIGO1 As String = "Codigo"
Private Const DATA1 As String = "Data"
Private Const PRODUTO1 As String = "Produto_Codigo"
Private Const UNIDADE_MED As String = "Unidade_Med"
Private Const QUANTIDADE1 As String = "Quantidade"
Private Const QUANT_DIFERENCA As String = "QuantDiferenca"
Private Const DESCRICAO_ITEM As String = "Descricao_Item"
Private Const ALMOXARIFADO1 As String = "Almoxarifado"
Private Const TIPO1 As String = "Tipo"
Private Const CUSTO_UNITARIO As String = "Custo_Unitario"
Private Const ETIQUETA1 As String = "Etiqueta"
Private Const CONTACONTABIL1 As String = "ContaContabilEst"
Const CTAAJUSTEINV As String = "CtaAjusteInv"
Private Const QUANT_ESTOQUE As String = "Quant_Estoque"
Private Const QUANT_DISPONIVEL As String = "Quant_Disponivel"
Private Const QUANT_CONSIGTERC As String = "Quant_ConsigTerc"
Private Const QUANT_DEMONSTERC As String = "Quant_DemonsTerc"
Private Const QUANT_CONSERTOTERC As String = "Quant_ConsertoTerc"
Private Const QUANT_OUTRASTERC As String = "Quant_OutrasTerc"
Private Const QUANT_BENEFTERC As String = "Quant_BenefTerc"

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

'Constantes públicas dos tabs
Private Const TAB_Lancamentos = 1
Private Const TAB_Contabilizacao = 2

'--- inicio dos properties get dos controles da tela

Public Property Get Frame1() As Object
     Set Frame1 = objUserControl.Controls("Frame1")
End Property

Public Property Get Codigo() As Object
     Set Codigo = objUserControl.Controls("Codigo")
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

'--- fim dos properties get dos controles da tela

Public Sub BotaoPlanoConta_Click()

Dim lErro As Long
Dim iContaPreenchida As Integer
Dim sConta As String
Dim objPlanoConta As New ClassPlanoConta
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoPlanoConta_Click

    If GridItens.Row = 0 Then gError 43710

    If GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = "" Then gError 43745

    sConta = String(STRING_CONTA, 0)

    lErro = CF("Conta_Formata", ContaContabil.Text, sConta, iContaPreenchida)
    If lErro <> SUCESSO Then gError 43749

    If iContaPreenchida = CONTA_PREENCHIDA Then objPlanoConta.sConta = sConta

    'Chama PlanoContaESTLista
    Call Chama_Tela("PlanoContaESTLista", colSelecao, objPlanoConta, objEventoContaContabil)

    Exit Sub

Erro_BotaoPlanoConta_Click:

    Select Case gErr

        Case 43710
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43745
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 43749

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155929)

    End Select

    Exit Sub

End Sub

Public Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objInventario As New ClassInventario
Dim vbMsg As VbMsgBoxResult
Dim iIndice As Integer
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dCusto As Double

On Error GoTo Erro_Codigo_Validate

    'se o codigo foi trocado
    If sCodigoAntigo <> Trim(Codigo.Text) Then

        If Len(Trim(Codigo.Text)) > 0 Then

            Call Limpa_colItensNumIntDoc(colItensNumIntDoc)

            objInventario.sCodigo = Codigo.Text

            lErro = CF("Inventario_Le", objInventario)
            If lErro <> SUCESSO And lErro <> 41011 Then gError 55436

            If lErro = SUCESSO Then

                vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PREENCHER_TELA")

                If vbMsg = vbNo Then gError 55437

                lErro = Preenche_Tela(objInventario)
                If lErro <> SUCESSO Then gError 55438

            Else

                For iIndice = 1 To objGrid.iLinhasExistentes

'                    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
'
'                        lErro = QuantEstoque_Calcula(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), iIndice)
'                        If lErro <> SUCESSO Then gError 55444
'
'                    Else
'
'                        lErro = QuantLote_Calcula(GridItens.TextMatrix(iIndice, iGrid_Produto_Col), iIndice, GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_FilialOP_Col)))
'                        If lErro <> SUCESSO Then gError 78662
'
'                    End If

                    'Verifica se o produto está preenchido
                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 55455

                    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

                        objProduto.sCodigo = sProdutoFormatado

                        'Lê o produto no BD
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 55457

                        If lErro = 28030 Then gError 55458

                    End If

                Next

            End If

        End If

        sCodigoAntigo = Trim(Codigo.Text)

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case gErr

        Case 55436, 55437, 55438, 55444, 55455, 55456, 55457, 78662

        Case 55458
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155930)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    sCodigoAntigo = ""

    Set objEventoCodigo = New AdmEvento
    Set objEventoAlmoxPadrao = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoEstoque = New AdmEvento
    Set objEventoContaContabil = New AdmEvento

    Set colItensNumIntDoc = New Collection

    'Inicializa o iFrameAtual
    iFrameAtual = 1

    'Preenche list de ComboBox Tipo com o ItemData Correspondente
    Call CargaCombo_Tipo(Tipo)

    'Preenche a Data
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True

    'Inicializa a Máscara do Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 34918

    'Mascara da conta contabil
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaContabil)
    If lErro <> SUCESSO Then gError 49582

    'Inicializa Mascara Conta Ajuste
    lErro = CF("Inicializa_Mascara_Conta_MaskEd", ContaAjuste)
    If lErro <> SUCESSO Then gError 52056

    'Formato para a Quantidade
    Quantidade.Format = FORMATO_ESTOQUE
    QuantEstoque.Format = FORMATO_ESTOQUE
    QuantDiferenca.Format = FORMATO_ESTOQUE

    'Formato para o Custo
    CustoUnitario.Format = FORMATO_CUSTO
    ValorDiferenca.Format = FORMATO_CUSTO

    'Carrega a combo de Filial O.P.
    lErro = Carrega_FilialOP()
    If lErro <> SUCESSO Then gError 78374
    
    'Inicializa o GridItens
    Set objGrid = New AdmGrid

    lErro = Inicializa_GridItens(objGrid)
    If lErro <> SUCESSO Then gError 34919

    'inicializacao da parte de contabilidade
    lErro = objContabil.Contabil_Inicializa_Contabilidade(Me, objGrid1, objEventoLote, objEventoDoc, MODULO_ESTOQUE)
    If lErro <> SUCESSO Then gError 39635

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 34918, 34919, 39635, 49582, 52056, 78374

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155931)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Function Carrega_FilialOP() As Long
'Carrega a combobox FilialOP

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialOP

    'Lê o Código e o Nome de toda FilialOP do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 78375

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

        Case 78375
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155932)

    End Select

    Exit Function

End Function

Private Sub CargaCombo_Tipo(objComboTipo As ComboBox)
'inicializa a combo com os tipos de estoque possiveis

Dim lErro As Long

On Error GoTo Erro_CargaCombo_Tipo

    With objComboTipo

'        If giTipoVersao = VERSAO_FULL Then
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
            .AddItem STRING_QUANT_DISPONIVEL_NOSSA_CI2P
            .ItemData(.NewIndex) = TIPO_QUANT_DISPONIVEL_NOSSA_CI2P
           
           'Wagner 20/10/2011
           'Para não ter que criar mais 20 tipos de movimento de estoque vou
           'deixar só as movimentações com custo informado disponíveis
'            .AddItem STRING_QUANT_NOSSO_CONSIG
'            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_CONSIG
'            .AddItem STRING_QUANT_NOSSO_DEMO
'            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_DEMO
'            .AddItem STRING_QUANT_NOSSO_CONSERTO
'            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_CONSERTO
'            .AddItem STRING_QUANT_NOSSO_OUTRAS
'            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_OUTRAS
'            .AddItem STRING_QUANT_NOSSO_BENEF
'            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_BENEF

            .AddItem STRING_QUANT_NOSSO_CONSIG_CI
            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_CONSIG_CI
            .AddItem STRING_QUANT_NOSSO_DEMO_CI
            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_DEMO_CI
            .AddItem STRING_QUANT_NOSSO_CONSERTO_CI
            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_CONSERTO_CI
            .AddItem STRING_QUANT_NOSSO_OUTRAS_CI
            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_OUTRAS_CI
            .AddItem STRING_QUANT_NOSSO_BENEF_CI
            .ItemData(.NewIndex) = TIPO_QUANT_NOSSO_BENEF_CI

'        ElseIf giTipoVersao = VERSAO_LIGHT Then
'
'            .AddItem STRING_QUANT_DISPONIVEL_NOSSA
'            .ItemData(.NewIndex) = TIPO_QUANT_DISPONIVEL_NOSSA
'            .AddItem STRING_QUANT_3_CONSIG
'            .ItemData(.NewIndex) = TIPO_QUANT_3_CONSIG
'            .AddItem STRING_QUANT_3_OUTRAS
'            .ItemData(.NewIndex) = TIPO_QUANT_3_OUTRAS
'
'        End If
    End With

    Exit Sub

Erro_CargaCombo_Tipo:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155933)

    End Select

    Exit Sub

End Sub

Public Sub ContaContabil_Change()

    iAlterado = REGISTRO_ALTERADO

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

Private Function Saida_Celula_ContaContabil(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sContaFormatada As String
Dim sContaMascarada As String
Dim objPlanoConta As New ClassPlanoConta
Dim vbMsgRes As VbMsgBoxResult
Dim iContaPreenchida As Integer

On Error GoTo Erro_Saida_Celula_ContaContabil

    Set objGrid.objControle = ContaContabil

    If Len(Trim(ContaContabil.ClipText)) > 0 Then

        'Verifica se é uma conta simples e se está em condições de receber lançamentos. Devolve os dados da ContaSimples em objPlanoConta
        lErro = CF("ContaSimples_Critica_Modulo", sContaFormatada, ContaContabil.Text, objPlanoConta, MODULO_ESTOQUE)
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 49584

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 49585

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaMascarada
            ContaContabil.PromptInclude = True

        'Se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaContabil.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 49586

            'Conta não cadastrada
            If lErro = 5700 Then gError 49587

        End If

    Else

        ContaContabil.PromptInclude = False
        ContaContabil.Text = ""
        ContaContabil.PromptInclude = True

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49645

    Saida_Celula_ContaContabil = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaContabil:

    Saida_Celula_ContaContabil = gErr

    Select Case gErr

        Case 49584, 49586, 49645
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49585
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 49587
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaContabil.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155934)

    End Select

    Exit Function

End Function

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

Private Function Saida_Celula_ContaAjuste(objGridInt As AdmGrid) As Long

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
        If lErro <> SUCESSO And lErro <> 44096 And lErro <> 44098 Then gError 52057

        If lErro = SUCESSO Then

            sContaFormatada = objPlanoConta.sConta

            'mascara a conta
            sContaMascarada = String(STRING_CONTA, 0)

            lErro = Mascara_RetornaContaEnxuta(objPlanoConta.sConta, sContaMascarada)
            If lErro <> SUCESSO Then gError 52058

            ContaAjuste.PromptInclude = False
            ContaAjuste.Text = sContaMascarada
            ContaAjuste.PromptInclude = True

        'Se não encontrou a conta simples
        ElseIf lErro = 44096 Or lErro = 44098 Then

            'Critica o formato da conta, sua presença no BD e capacidade de receber lançamentos
            lErro = CF("Conta_Critica", ContaAjuste.Text, sContaFormatada, objPlanoConta, MODULO_ESTOQUE)
            If lErro <> SUCESSO And lErro <> 5700 Then gError 52059

            'Conta não cadastrada
            If lErro = 5700 Then gError 52060

        End If

    Else

        ContaAjuste.PromptInclude = False
        ContaAjuste.Text = ""
        ContaAjuste.PromptInclude = True

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 52061

    Saida_Celula_ContaAjuste = SUCESSO

    Exit Function

Erro_Saida_Celula_ContaAjuste:

    Saida_Celula_ContaAjuste = gErr

    Select Case gErr

        Case 52057, 52059, 52061
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52058
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 52060
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONTA_INEXISTENTE", ContaAjuste.Text)

            If vbMsgRes = vbYes Then
                objPlanoConta.sConta = sContaFormatada

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("PlanoConta", objPlanoConta)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155935)

    End Select

    Exit Function

End Function

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Tipo Estoque")
    objGridInt.colColuna.Add ("Lote \ O.P.")
    objGridInt.colColuna.Add ("Filial O.P.")
    objGridInt.colColuna.Add ("Atualiza")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quant. Encontrada")
    objGridInt.colColuna.Add ("Quant. Estoque")
    objGridInt.colColuna.Add ("Diferença-Quant")
    objGridInt.colColuna.Add ("Custo Unitário")
    objGridInt.colColuna.Add ("Diferença-Valor")
    objGridInt.colColuna.Add ("Conta de Estoque")
    objGridInt.colColuna.Add ("Conta de Ajuste")
    objGridInt.colColuna.Add ("Etiqueta")
    'A etiqueta está em baixo

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoItem.Name)
    objGridInt.colCampo.Add (Almoxarifado.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (LoteProduto.Name)
    objGridInt.colCampo.Add (FilialOP.Name)
    objGridInt.colCampo.Add (Atualiza.Name)
    objGridInt.colCampo.Add (UnidadeMed.Name)
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
    iGrid_LoteProduto_Col = 5
    iGrid_FilialOP_Col = 6
    iGrid_Atualiza_Col = 7
    iGrid_UnidadeMed_Col = 8
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

Private Function Preenche_ContaAjuste() As Long

Dim lErro As Long
Dim objMnemonicoCTBValor As New ClassMnemonicoCTBValor
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaAjuste

    objMnemonicoCTBValor.sMnemonico = CTAAJUSTEINV

    lErro = CF("MnemonicoCTBValor_Le", objMnemonicoCTBValor)
    If lErro <> SUCESSO And lErro <> 39690 Then gError 52066

    GridItens.TextMatrix(GridItens.Row, iGrid_ContaAjuste_Col) = objMnemonicoCTBValor.sValor

    Preenche_ContaAjuste = SUCESSO

    Exit Function

Erro_Preenche_ContaAjuste:

    Preenche_ContaAjuste = gErr

    Select Case gErr

        Case 52066

        Case 52067
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objMnemonicoCTBValor.sValor)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155936)

    End Select

    Exit Function

End Function

Public Function Preenche_ContaContabil() As Long
'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto
'Caso nao encontre -----> não tratar erro

Dim lErro As Long
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sContaEnxuta As String

On Error GoTo Erro_Preenche_ContaContabil

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col))) = 0 Or Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then Exit Function

    'Preenche o objEstoqueProduto
    objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col)

    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
    If lErro <> SUCESSO Then gError 49593

    'Formata o Produto para BD
    lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 49594

    objEstoqueProduto.sProduto = sProdutoFormatado
    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo

    'Lê o Estoque do Produto
    lErro = CF("EstoqueProdutoCC_Le", objEstoqueProduto)
    If lErro <> SUCESSO And lErro <> 49991 Then gError 49595

    If lErro = SUCESSO Then

        lErro = Mascara_RetornaContaEnxuta(objEstoqueProduto.sContaContabil, sContaEnxuta)
        If lErro <> SUCESSO Then gError 49601

        ContaContabil.PromptInclude = False
        ContaContabil.Text = sContaEnxuta
        ContaContabil.PromptInclude = True

        GridItens.TextMatrix(GridItens.Row, iGrid_ContaContabil_Col) = ContaContabil.Text

    End If

    Preenche_ContaContabil = SUCESSO

    Exit Function

Erro_Preenche_ContaContabil:

    Preenche_ContaContabil = gErr

    Select Case gErr

        Case 49593, 49594, 49595

        Case 49601
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objEstoqueProduto.sContaContabil)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155937)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no Banco de Dados

Dim lErro As Long
Dim iIndice As Integer
Dim vCodigo As Variant
Dim objInventario As New ClassInventario

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Invent"

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

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do Banco de Dados), tamanho do campo
    'no Banco de Dados no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objInventario.sCodigo, STRING_INVENTARIO_CODIGO, "Codigo"
    colCampoValor.Add "Data", objInventario.dtData, 0, "Data"
'hora
    colCampoValor.Add "Hora", CDbl(objInventario.dtHora), 0, "Hora"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155938)

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
    objInventario.iFilialEmpresa = giFilialEmpresa

    lErro = Preenche_Tela(objInventario)
    If lErro <> SUCESSO Then gError 34920

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 34920

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155939)

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

    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection

    'Lê os ítens do Inventario
    lErro = CF("InventarioItens_Le", objInventario)
    If lErro <> SUCESSO And lErro <> 41006 Then gError 34921

    If lErro = 41006 Then gError 55439

    'Passa as Informações de NumIntDoc de colItens para colItensNumIntDoc
    For iIndice = 1 To objInventario.colItens.Count

        colItensNumIntDoc.Add objInventario.colItens.Item(iIndice).lNumIntDoc

    Next

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

    lErro = Preenche_GridItens(objInventario.colItens)
    If lErro <> SUCESSO Then gError 34922

    'Traz os dados contábeis para a tela (contabilidade)
    lErro = objContabil.Contabil_Traz_Doc_Tela(objInventario.colItens(1).lNumIntDoc)
    If lErro <> SUCESSO And lErro <> 36326 Then gError 39636

    iAlterado = 0
    sCodigoAntigo = objInventario.sCodigo

    Preenche_Tela = SUCESSO

    Exit Function

Erro_Preenche_Tela:

    Preenche_Tela = gErr

    Select Case gErr

        Case 34921, 34922, 39636

        Case 55439
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INVENTARIO_NAO_CADASTRADO", gErr, objInventario.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155940)

    End Select

    Exit Function

End Function

Private Function Preenche_GridItens(colItens As ColItemInventario) As Long

Dim lErro As Long
Dim iIndice As Integer, iPosicao As Integer
Dim sProdutoMascarado As String, sCclMascarado As String, sContaMascara As String
Dim objItemInventario As ClassItemInventario
Dim sContaEnxuta As String
Dim sContaAjusteEnxuta As String
Dim objProduto As New ClassProduto
Dim dFator As Double
Dim objFilialOP As New AdmFiliais
Dim colRatreamentoMovto As New Collection
Dim objRatreamentoMovto As New ClassRastreamentoMovto
Dim objMovEstoque As New ClassMovEstoque
Dim objItemMovEst As ClassItemMovEstoque

On Error GoTo Erro_Preenche_GridItens

    iIndice = 0

    'Preenche GridItens
    For Each objItemInventario In colItens

        iIndice = iIndice + 1

        sProdutoMascarado = String(STRING_PRODUTO, 0)

        'Formata o Produto
        lErro = Mascara_RetornaProdutoEnxuto(objItemInventario.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 34923

''        lErro = Mascara_MascararProduto(objItemInventario.sProduto, sProdutoMascarado)
''        If lErro <> SUCESSO Then gError 34923

        sContaMascara = String(STRING_CONTA, 0)

        If objItemInventario.sContaContabilEst <> "" Then

            lErro = Mascara_RetornaContaEnxuta(objItemInventario.sContaContabilEst, sContaEnxuta)
            If lErro <> SUCESSO Then gError 49588

            ContaContabil.PromptInclude = False
            ContaContabil.Text = sContaEnxuta
            ContaContabil.PromptInclude = True

            GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col) = ContaContabil.Text

        End If

        sContaAjusteEnxuta = String(STRING_CONTA, 0)

        If objItemInventario.sContaContabilInv <> "" Then

            lErro = Mascara_RetornaContaEnxuta(objItemInventario.sContaContabilInv, sContaAjusteEnxuta)
            If lErro <> SUCESSO Then gError 52065

            ContaAjuste.PromptInclude = False
            ContaAjuste.Text = sContaAjusteEnxuta
            ContaAjuste.PromptInclude = True

            GridItens.TextMatrix(iIndice, iGrid_ContaAjuste_Col) = ContaAjuste.Text

        End If

        Produto.PromptInclude = False
        Produto.Text = sProdutoMascarado
        Produto.PromptInclude = True

        GridItens.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text 'sProdutoMascarado
        GridItens.TextMatrix(iIndice, iGrid_DescricaoItem_Col) = objItemInventario.sProdutoDesc
        GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col) = objItemInventario.sSiglaUM
        GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col) = Formata_Estoque(objItemInventario.dQuantidade)
        GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objItemInventario.dQuantEst)
        GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = Formata_Estoque(objItemInventario.dQuantidade - objItemInventario.dQuantEst)
        GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = objItemInventario.sAlmoxarifadoNomeRed

        GridItens.TextMatrix(iIndice, iGrid_Etiqueta_Col) = objItemInventario.sEtiqueta
        
        Set colRatreamentoMovto = New Collection
        
        objMovEstoque.iFilialEmpresa = giFilialEmpresa
        
        'Lê o movimento de estoque associado ao NumIntDocOrigem, TipoNumIntDocOrigem em questão para descobrir o codigo do movimento de estoque associado ao inventário
        lErro = CF("MovEstoque_Le_NumIntDocOrigem", objMovEstoque, MOVEST_TIPONUMINTDOCORIGEM_INVENTARIO, objItemInventario.lNumIntDoc)
        If lErro <> SUCESSO And lErro <> 89850 Then gError 90830
        
        Set objItemMovEst = objMovEstoque.colItens(1)
        
        'Le o Rastreamento e preenche o grid com o Número do Lote e o Numero da Filial OP
        lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEst.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRatreamentoMovto)
        If lErro <> SUCESSO And lErro <> 78414 Then gError 78916
        
        If colRatreamentoMovto.Count > 0 Then
        
            'Seta o primeiro Lote
            Set objRatreamentoMovto = colRatreamentoMovto(1)
        
            If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col) = objRatreamentoMovto.sLote
            
            If objRatreamentoMovto.iFilialOP > 0 Then
            
                objFilialOP.iCodFilial = objRatreamentoMovto.iFilialOP
    
                'Le a Filial Empresa da OP para pegar a descrição
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO Then gError 78402
    
                GridItens.TextMatrix(iIndice, iGrid_FilialOP_Col) = objFilialOP.iCodFilial & SEPARADOR & objFilialOP.sNome
            
            End If
            
            If Len(Trim(objRatreamentoMovto.sLote)) > 0 Then
            
                For iPosicao = 0 To Atualiza.ListCount - 1
                    If Atualiza.ItemData(iPosicao) = objItemInventario.iAtualizaSoLote Then
                        GridItens.TextMatrix(iIndice, iGrid_Atualiza_Col) = Atualiza.List(iPosicao)
                        Exit For
                    End If
                Next
                                  
            End If
                                  
        End If
            
        If objItemInventario.dQuantidade - objItemInventario.dQuantEst < 0 Then
            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = Format(-objItemInventario.dCusto, "Standard")
        Else
            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = Format(objItemInventario.dCusto, "Standard")
        End If

        objProduto.sCodigo = objItemInventario.sProduto

        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 55441

        If lErro = 28030 Then gError 55442

        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, objItemInventario.sSiglaUM, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 55443

        If objItemInventario.dQuantidade - objItemInventario.dQuantEst = 0 Then
            GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col) = Formata_Custo(0)
        Else
            GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col) = Formata_Custo(objItemInventario.dCusto / (Abs(objItemInventario.dQuantidade - objItemInventario.dQuantEst) * dFator))
        End If

        For iPosicao = 0 To Tipo.ListCount - 1
            If Tipo.ItemData(iPosicao) = objItemInventario.iTipo Then
                GridItens.TextMatrix(iIndice, iGrid_Tipo_Col) = Tipo.List(iPosicao)
                Exit For
            End If
        Next
       
    Next

    objGrid.iLinhasExistentes = colItens.Count

    lErro = Grid_Refresh_Checkbox(objGrid)
    If lErro <> SUCESSO Then gError 34924

    Preenche_GridItens = SUCESSO

    Exit Function

Erro_Preenche_GridItens:

    Preenche_GridItens = gErr

    Select Case gErr

        Case 34924, 49713, 49588, 52065, 55441, 55443, 78402, 90830

        Case 34923
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objItemInventario.sProduto)

        Case 55442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155941)

    End Select

    Exit Function

End Function

Public Sub CodigoLabel_Click()

Dim lErro As Long
Dim objInventario As New ClassInventario
Dim colSelecao As New Collection

    If Len(Trim(Codigo.Text)) > 0 Then objInventario.sCodigo = Codigo.Text

    Call Chama_Tela("InventarioLista", colSelecao, objInventario, objEventoCodigo)

End Sub

Public Sub Data_GotFocus()

    Call MaskEdBox_TrataGotFocus(Data, iAlterado)

End Sub

Public Sub GridItens_RowColChange()

    Call Grid_RowColChange(objGrid)


End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objInventario As ClassInventario

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objInventario = obj1

    lErro = Preenche_Tela(objInventario)
    If lErro <> SUCESSO Then gError 34925

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 34925

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155942)

    End Select

    Exit Sub

End Sub

Public Sub AlmoxPadraoLabel_Click()

Dim objAlmoxarifado As New ClassAlmoxarifado
Dim colSelecao As New Collection

    Call Chama_Tela("AlmoxarifadoLista_Consulta", colSelecao, objAlmoxarifado, objEventoAlmoxPadrao)

End Sub

Private Sub objEventoAlmoxPadrao_evSelecao(obj1 As Object)

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

    If GridItens.Row = 0 Then gError 43708

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) > 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 34926

        If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

        objProduto.sCodigo = sProduto

    End If

    sSelecao = "ControleEstoque<>?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE

    Call Chama_Tela("ProdutoEstoqueLista", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub

Erro_BotaoProdutos_Click:

    Select Case gErr

        Case 34926

        Case 43708
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155943)

    End Select

    Exit Sub

End Sub

Private Sub objEventoContaContabil_evSelecao(obj1 As Object)

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
        If lErro <> SUCESSO Then gError 55434

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

        Case 55434
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNACONTAENXUTA", gErr, objPlanoConta.sConta)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155944)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim sProdutoMascarado As String
Dim sTipo As String
Dim dCusto As Double

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col))) = 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34927

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then

            sProdutoMascarado = String(STRING_PRODUTO, 0)

            lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 34298

            'Lê os demais atributos do Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 34929

            If lErro = 28030 Then gError 34930

            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True

            

            If Not (Me.ActiveControl Is Produto) Then

                GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col) = Produto.Text

                'Preenche a Linha do Grid
                lErro = ProdutoLinha_Preenche(objProduto)
                If lErro <> SUCESSO Then gError 34931

                sTipo = GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col)

                'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
                lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
                If lErro <> SUCESSO Then gError 49615

                GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)


                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then

                    lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
                    If lErro <> SUCESSO Then gError 34932

                Else

                    lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                    If lErro <> SUCESSO Then gError 78663

                End If

                lErro = Preenche_ContaContabil()
                If lErro <> SUCESSO Then gError 52240

            End If

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 34927, 34929, 34931, 34932, 49615, 52240, 78663

        Case 34928
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", gErr, objProduto.sCodigo)

        Case 34930
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155945)

    End Select

    Exit Sub

End Sub

Public Function ProdutoLinha_Preenche(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iCclPreenchida As Integer
Dim sCclFormata As String
Dim sAlmoxarifadoPadrao As String

On Error GoTo Erro_ProdutoLinha_Preenche

    lErro = Preenche_ContaAjuste()
    If lErro <> SUCESSO Then gError 52068

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
        If lErro <> SUCESSO Then gError 52239

        GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = sAlmoxarifadoPadrao

    End If

    '??? disp nossa vai ser sempre a 1a entao o indice vai ser sempre zero
     GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = Tipo.List(0)

    'ALTERAÇÃO DE LINHAS EXISTENTES
    If (GridItens.Row - GridItens.FixedRows) = objGrid.iLinhasExistentes Then
        If objGrid.iLinhasExistentes + 1 = GridItens.Rows Then
            GridItens.Rows = GridItens.Rows + NUM_GRUPO_ITENS_INVENTARIO
        End If
        objGrid.iLinhasExistentes = objGrid.iLinhasExistentes + 1
        colItensNumIntDoc.Add 0
    End If

    ProdutoLinha_Preenche = SUCESSO

    Exit Function

Erro_ProdutoLinha_Preenche:

    ProdutoLinha_Preenche = gErr

    Select Case gErr

        Case 52068, 52239

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155946)

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

    If GridItens.Row = 0 Then gError 43709

    sCodProduto = GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 34933

    'Se na Linha corrente Produto estiver preenchido
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

        colSelecao.Add sProdutoFormatado

        Call Chama_Tela("EstoqueProdutoFilialLista", colSelecao, objEstoqueProduto, objEventoEstoque)
    Else
        Error 43734
    End If

    Exit Sub

Erro_BotaoEstoque_Click:

    Select Case gErr

        Case 34933

        Case 43709
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case 43734
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155947)

    End Select

    Exit Sub

End Sub

Private Sub objEventoEstoque_evselecao(obj1 As Object)

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
        If lErro <> SUCESSO Then gError 34934

        'Verifica se o produto está preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Preenche o Nome do Almoxarifado
            GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            Almoxarifado.Text = objEstoqueProduto.sAlmoxarifadoNomeReduzido

            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then

                'Calcula a Quantidade Disponível nesse Almoxarifado
                lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
                If lErro <> SUCESSO Then gError 34935

            Else

                'Calcula a Quantidade Disponível nesse Almoxarifado
                lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78664

            End If

            'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
            lErro = Preenche_ContaContabil()
            If lErro <> SUCESSO Then gError 49597

        End If

    End If

    Me.Show

    Exit Sub

Erro_objEventoEstoque_evselecao:

    Select Case gErr

        Case 34934, 34935, 49597, 78664

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155948)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objInventario As ClassInventario) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um Movestoque passado como parâmetro
    If Not objInventario Is Nothing Then

        'Lê o Inventario no Banco de Dados
        lErro = CF("Inventario_Le", objInventario)
        If lErro <> SUCESSO And lErro <> 41011 Then gError 34936

        If lErro = SUCESSO Then 'Se ele existe

            lErro = Preenche_Tela(objInventario)
            If lErro <> SUCESSO Then gError 34937

        ElseIf lErro = 41011 Then

            'Se ele não existe exibe apenas o código
            Codigo.Text = objInventario.sCodigo
            sCodigoAntigo = objInventario.sCodigo

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 34936, 34937, 34938

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155949)

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
        If lErro <> SUCESSO And lErro <> 25136 And lErro <> 25143 Then gError 34939

        If lErro = 25136 Then gError 34940

        If lErro = 25143 Then gError 34941

    End If

    Exit Sub

Erro_AlmoxPadrao_Validate:

    Cancel = True


    Select Case gErr

        Case 34939

        Case 34940, 34941
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, AlmoxPadrao.Text)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155950)

    End Select

    Exit Sub

End Sub

Public Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 34942

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 34942

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155951)

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
    If lErro <> SUCESSO Then gError 89803

    Exit Sub

Erro_Hora_Validate:

    Cancel = True

    Select Case gErr

        Case 89803

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155952)

    End Select

    Exit Sub

End Sub

Public Sub Opcao_Click()

    'Se frame selecionado não for o atual
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub

        'Esconde o frame atual, mostra o novo
        Frame1(Opcao.SelectedItem.Index).Visible = True
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index

        'se estiver selecionando o tabstrip de contabilidade e o usuário não alterou a contabilidade ==> carrega o modelo padrao
        If Opcao.SelectedItem.Caption = TITULO_TAB_CONTABILIDADE Then Call objContabil.Contabil_Carga_Modelo_Padrao

        Select Case iFrameAtual

            Case TAB_Lancamentos
                Parent.HelpContextID = IDH_INVENTARIO_LANCAMENTOS

            Case TAB_Contabilizacao
                Parent.HelpContextID = IDH_INVENTARIO_CONTABILIZACAO

        End Select

    End If

End Sub

Public Sub Tipo_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Atualiza_Click()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub UpDownData_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
        If lErro <> SUCESSO Then gError 34943

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 34943

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155953)

    End Select

    Exit Sub

End Sub

Public Sub UpDownData_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    If Len(Trim(Data.ClipText)) > 0 Then

        lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
        If lErro <> SUCESSO Then gError 34944

        iAlterado = REGISTRO_ALTERADO

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 34944

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155954)

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
Dim lNumIntDoc As Long

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridItens.TextMatrix(iLinha, iGrid_Produto_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 34945

    If colItensNumIntDoc.Count >= GridItens.Row Then
        lNumIntDoc = colItensNumIntDoc.Item(GridItens.Row)
    Else
        lNumIntDoc = 0
    End If

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
            If lErro <> SUCESSO And lErro <> 28030 Then gError 34946

            If lErro = 28030 Then gError 34947

            objClasseUM.iClasse = objProduto.iClasseUM

            'Preenche a List da Combo UnidadeMed com as UM's do Produto
            lErro = CF("UnidadesDeMedidas_Le_ClasseUM", objClasseUM, colSiglas)
            If lErro <> SUCESSO Then gError 34948

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

            If lNumIntDoc = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        End If

    ElseIf objControl.Name = "CustoUnitario" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 And _
        (GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_DEMO Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_CONSERTO Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_OUTRAS Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_BENEF Or _
        GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_NOSSO_CONSERTO_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_NOSSO_DEMO_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_NOSSO_CONSIG_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_NOSSO_OUTRAS_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_NOSSO_BENEF_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_DISPONIVEL_NOSSA_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_RECEB_INDISP_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_OUTRAS_INDISP_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_DEFEIT_CI Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_3_CONSIG_CI) Or GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col) = STRING_QUANT_DISPONIVEL_NOSSA_CI2P Then
            
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = "Quantidade" Or objControl.Name = "Almoxarifado" Or objControl.Name = "Tipo" Or objControl.Name = "Etiqueta" Or objControl.Name = "ContaContabil" Or objControl.Name = "ContaAjuste" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then
            objControl.Enabled = True
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = "LoteProduto" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78385

            If lErro = 28030 Then gError 78386

            If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
        Else
            objControl.Enabled = False
        End If

    ElseIf objControl.Name = "FilialOP" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO And lNumIntDoc = 0 Then

            objProduto.sCodigo = sProdutoFormatado

            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78387

            If lErro = 28030 Then gError 78388

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

        Case 34945, 34946, 34948, 78385, 78387, 92478

        Case 34947, 78386, 78388, 92479
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155955)

    End Select

    Exit Sub

End Sub

Public Function QuantLote_Calcula(sProduto As String, ByVal iIndice As Integer, sLote As String, iFilialOP As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim sUnidadeMed As String
Dim dFator As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objRastreamentoSaldo As New ClassRastreamentoLoteSaldo
Dim objProduto As New ClassProduto
Dim dCustoMedioAtual As Double
Dim sTipo As String

On Error GoTo Erro_QuantLote_Calcula

    'Maristela???
    'If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then

        objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)
    
        'Verifica se o produto está preenchido
        lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78635
    
        objProduto.sCodigo = sProdutoFormatado
    
        sTipo = GridItens.TextMatrix(iIndice, iGrid_Tipo_Col)
    
        If iIndice >= GridItens.FixedRows And Len(Trim(objAlmoxarifado.sNomeReduzido)) > 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then
    
            'Lê o produto no BD para obter UM de estoque
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78636
    
            If lErro = 28030 Then gError 78637
    
            lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 78638
    
            If lErro = 25056 Then gError 78639
    
            If (Len(Trim(sLote)) > 0 And objProduto.iRastro <> PRODUTO_RASTRO_OP) Or (Len(Trim(sLote)) > 0 And objProduto.iRastro = PRODUTO_RASTRO_OP And iFilialOP) > 0 Then
    
                objRastreamentoSaldo.iAlmoxarifado = objAlmoxarifado.iCodigo
                objRastreamentoSaldo.sProduto = sProdutoFormatado
                objRastreamentoSaldo.iFilialOP = iFilialOP
                objRastreamentoSaldo.sLote = sLote
    
                sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)
    
                lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
                If lErro <> SUCESSO Then gError 78642
    
                'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
                lErro = CF("RastreamentoLoteSaldo_Le", objRastreamentoSaldo)
                If lErro <> SUCESSO And lErro <> 78633 Then gError 78640
    
                'Se não encontrou ---> Erro
                If lErro = 78633 Then
                
                    GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(0)
    
                    'Calcula a Quantidade diferenca
                    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
                        GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = Formata_Estoque(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))
        
                        'se o custo unitario estiver preenchido entao calcula a diferenca
                        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))) > 0 Then
        
                            'Preenche o valor da diferença com diferenca-quant*CustoUnitario
                            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = Format(CDbl(GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col)) * dFator * CDbl(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col)), "Standard")
        
                        Else
                            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
                        End If
                    Else
                        GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
                        GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = ""
                    End If
    
                Else
            
                    Select Case sTipo
        
                        Case STRING_QUANT_DISPONIVEL_NOSSA, STRING_QUANT_DISPONIVEL_NOSSA_CI, STRING_QUANT_DISPONIVEL_NOSSA_CI2P
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque((objRastreamentoSaldo.dQuantDispNossa + objRastreamentoSaldo.dQuantReservada) / dFator)
                        Case STRING_QUANT_RECEB_INDISP, STRING_QUANT_RECEB_INDISP_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantRecIndl / dFator)
                        Case STRING_QUANT_OUTRAS_INDISP, STRING_QUANT_OUTRAS_INDISP_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantIndOutras / dFator)
                        Case STRING_QUANT_DEFEIT, STRING_QUANT_DEFEIT_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantDefeituosa / dFator)
                        Case STRING_QUANT_3_CONSIG, STRING_QUANT_3_CONSIG_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque((objRastreamentoSaldo.dQuantConsig3 + objRastreamentoSaldo.dQuantReservadaConsig) / dFator)
                        Case STRING_QUANT_3_DEMO
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantDemo3 / dFator)
                        Case STRING_QUANT_3_CONSERTO
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantConserto3 / dFator)
                        Case STRING_QUANT_3_OUTRAS
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantOutras3 / dFator)
                        Case STRING_QUANT_3_BENEF
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantBenef3 / dFator)
                        Case STRING_QUANT_NOSSO_BENEF, STRING_QUANT_NOSSO_BENEF_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantBenef / dFator)
                        Case STRING_QUANT_NOSSO_OUTRAS, STRING_QUANT_NOSSO_OUTRAS_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantOutras / dFator)
                        Case STRING_QUANT_NOSSO_CONSIG, STRING_QUANT_NOSSO_CONSIG_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantConsig / dFator)
                        Case STRING_QUANT_NOSSO_DEMO, STRING_QUANT_NOSSO_DEMO_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantDemo / dFator)
                        Case STRING_QUANT_NOSSO_CONSERTO, STRING_QUANT_NOSSO_CONSERTO_CI
                            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objRastreamentoSaldo.dQuantConserto / dFator)
                        Case Else
                            Error 78643
                     End Select
        
                    'Calcula a Quantidade diferenca
                    If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
                        GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = Formata_Estoque(CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) - CDbl(GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col)))
        
        
                        'se o custo unitario estiver preenchido entao calcula a diferenca
                        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))) > 0 Then
        
                            'Preenche o valor da diferença com diferenca-quant*CustoUnitario
                            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = Format(CDbl(GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col)) * dFator * CDbl(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col)), "Standard")
        
                        Else
                            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
                        End If
                    Else
                        GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
                        GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = ""
                    End If
    
                End If
    
            Else
    
                'Limpa a QuantidadeEstoque ,Quantidadediferenca e valorDiferenca da Tela
                GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
                GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = ""
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = ""
    
            End If
            
        Else
    
            'Limpa a QuantidadeEstoque ,Quantidadediferenca e valorDiferenca da Tela
            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
            GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = ""
            GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = ""
    
        End If

'    Else
'
'        'Limpa a QuantidadeEstoque ,Quantidadediferenca e valorDiferenca da Tela
'        GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
'        GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = ""
'        GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = ""
'
'    End If

    QuantLote_Calcula = SUCESSO

    Exit Function

Erro_QuantLote_Calcula:

    QuantLote_Calcula = gErr

    Select Case gErr

        Case 78635, 78636, 78638, 78640, 78642

        Case 78637
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 78639
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 78641
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_NAO_TEM_PRODUTO_ALMOXARIFADO", gErr, objRastreamentoSaldo.sLote, objRastreamentoSaldo.sProduto, objRastreamentoSaldo.iAlmoxarifado)

        Case 78643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_ESTOQUE_INVALIDO", gErr, sTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155956)

    End Select

    Exit Function

End Function

Public Function QuantEstoque_Calcula(sProduto As String, ByVal iIndice As Integer) As Long

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

    objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)

    'Verifica se o produto está preenchido
    lErro = CF("Produto_Formata", sProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 49608

    objProduto.sCodigo = sProdutoFormatado

    sTipo = GridItens.TextMatrix(iIndice, iGrid_Tipo_Col)

    If iIndice >= GridItens.FixedRows And Len(Trim(objAlmoxarifado.sNomeReduzido)) > 0 And iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then

        'Lê o produto no BD para obter UM de estoque
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 49609

        If lErro = 28030 Then gError 49610

        lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 49611

        If lErro = 25056 Then gError 49612

        objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
        objEstoqueProduto.sProduto = sProdutoFormatado

        'Lê o Estoque Produto correspondente ao Produto e ao Almoxarifado
        lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
        If lErro <> SUCESSO And lErro <> 21306 Then gError 49613

        If lErro = 21306 Then gError 55468

        sUnidadeMed = GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col)

        lErro = CF("UM_Conversao_Trans", objProduto.iClasseUM, sUnidadeMed, objProduto.sSiglaUMEstoque, dFator)
        If lErro <> SUCESSO Then gError 49614

        Select Case sTipo

            Case STRING_QUANT_DISPONIVEL_NOSSA, STRING_QUANT_DISPONIVEL_NOSSA_CI, STRING_QUANT_DISPONIVEL_NOSSA_CI2P
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque((objEstoqueProduto.dQuantDispNossa + objEstoqueProduto.dQuantReservada) / dFator)
            Case STRING_QUANT_RECEB_INDISP, STRING_QUANT_RECEB_INDISP_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantRecIndl / dFator)
            Case STRING_QUANT_OUTRAS_INDISP, STRING_QUANT_OUTRAS_INDISP_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantInd / dFator)
            Case STRING_QUANT_DEFEIT, STRING_QUANT_DEFEIT_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantDefeituosa / dFator)
            Case STRING_QUANT_3_CONSIG, STRING_QUANT_3_CONSIG_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque((objEstoqueProduto.dQuantConsig3 + objEstoqueProduto.dQuantReservadaConsig) / dFator)
            Case STRING_QUANT_3_DEMO
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantDemo3 / dFator)
            Case STRING_QUANT_3_CONSERTO
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantConserto3 / dFator)
            Case STRING_QUANT_3_OUTRAS
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantOutras3 / dFator)
            Case STRING_QUANT_3_BENEF
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantBenef3 / dFator)
            Case STRING_QUANT_NOSSO_BENEF, STRING_QUANT_NOSSO_BENEF_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantBenef / dFator)
            Case STRING_QUANT_NOSSO_CONSIG, STRING_QUANT_NOSSO_CONSIG_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantConsig / dFator)
            Case STRING_QUANT_NOSSO_CONSERTO, STRING_QUANT_NOSSO_CONSERTO_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantConserto / dFator)
            Case STRING_QUANT_NOSSO_DEMO, STRING_QUANT_NOSSO_DEMO_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantDemo / dFator)
            Case STRING_QUANT_NOSSO_OUTRAS, STRING_QUANT_NOSSO_OUTRAS_CI
                GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = Formata_Estoque(objEstoqueProduto.dQuantOutras / dFator)
            Case Else
                Error 55433
         End Select

        'Calcula a Quantidade diferenca
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) > 0 Then
            GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = Formata_Estoque(CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)) - CDbl(GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col)))


            'se o custo unitario estiver preenchido entao calcula a diferenca
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))) > 0 Then

                'Preenche o valor da diferença com diferenca-quant*CustoUnitario
                GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = Format(CDbl(GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col)) * dFator * CDbl(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col)), "Standard")

            Else
                GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
            End If
        Else
            GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
            GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = ""
        End If

    Else

        'Limpa a QuantidadeEstoque ,Quantidadediferenca e valorDiferenca da Tela
        GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col) = ""
        GridItens.TextMatrix(iIndice, iGrid_QuantDiferenca_Col) = ""
        GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col) = ""

    End If

    QuantEstoque_Calcula = SUCESSO

    Exit Function

Erro_QuantEstoque_Calcula:

    QuantEstoque_Calcula = gErr

    Select Case gErr

        Case 49608, 49609, 49611, 49613, 49614

        Case 49610
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 49612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 55433
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_ESTOQUE_INVALIDO", gErr, sTipo)

        Case 55468
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_TEM_PRODUTO", gErr, objAlmoxarifado.sNomeReduzido, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155957)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim vbMsg As VbMsgBoxResult
Dim sTipo As String
Dim dCusto As Double
Dim sProduto As String

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) <> 0 Then

        sProduto = Produto.Text

        lErro = CF("Trata_Segmento_Produto", sProduto)
        If lErro <> SUCESSO Then gError 199347

        Produto.Text = sProduto

        lErro = CF("Produto_Critica_Estoque", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25077 Then gError 34957

        If lErro = 25077 Then gError 34958

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            lErro = ProdutoLinha_Preenche(objProduto)
            If lErro <> SUCESSO Then gError 34959

            sTipo = GridItens.TextMatrix(GridItens.Row, iGrid_Tipo_Col)

            'calcula o custo atual do produto passado como parametro e devolve-o em dCusto
            lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
            If lErro <> SUCESSO Then gError 49616

            GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)

            If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then

                'Calcula a Quantidade Disponível
                lErro = QuantEstoque_Calcula(Produto.Text, GridItens.Row)
                If lErro <> SUCESSO Then gError 41043

            Else

                'Calcula a Quantidade Disponível
                lErro = QuantLote_Calcula(Produto.Text, GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
                If lErro <> SUCESSO Then gError 78655

            End If

        End If

        If objProduto.iRastro = PRODUTO_RASTRO_OP Then
            GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col) = giFilialEmpresa & SEPARADOR & gsNomeFilialEmpresa
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34960

    lErro = Preenche_ContaContabil()
    If lErro <> SUCESSO Then gError 52241

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 34957, 34959, 34960, 41043, 49616, 52241, 78655, 199347
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 34958
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)

            If vbMsg = vbYes Then
                objProduto.sCodigo = Produto.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Produto", objProduto)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155958)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Quantidade(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim dQuantTotal As Double

On Error GoTo Erro_Saida_Celula_Quantidade

    Set objGridInt.objControle = Quantidade

    If Len(Trim(Quantidade.ClipText)) > 0 Then

        lErro = Valor_NaoNegativo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 34961

        Quantidade.Text = Formata_Estoque(Quantidade.Text)

    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_Quantidade_Col) = Quantidade.Text

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then

        'Calcula a Quantidade Disponível
        lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
        If lErro <> SUCESSO Then gError 49618

    Else

        'Calcula a Quantidade Disponível
        lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78656

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34962

    Saida_Celula_Quantidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Quantidade:

    Saida_Celula_Quantidade = gErr

    Select Case gErr

        Case 34961, 34962, 49618, 78656
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155959)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_UnidadeMed(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_UnidadeMed

    Set objGridInt.objControle = UnidadeMed

    GridItens.TextMatrix(GridItens.Row, iGrid_UnidadeMed_Col) = UnidadeMed.Text

    If Len(UnidadeMed.Text) > 0 Then

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then

            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 55425

        Else

            lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78657

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34999

    Saida_Celula_UnidadeMed = SUCESSO

    Exit Function

Erro_Saida_Celula_UnidadeMed:

    Saida_Celula_UnidadeMed = gErr

    Select Case gErr

        Case 34999, 55425, 78657
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155960)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Almoxarifado(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Almoxarifado

    Set objGridInt.objControle = Almoxarifado

    If Len(Trim(Almoxarifado.ClipText)) > 0 Then

        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 34963

        lErro = TP_Almoxarifado_Filial_Produto_Grid(sProdutoFormatado, Almoxarifado, objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25157 And lErro <> 25162 Then gError 34964

        If lErro = 25157 Then gError 34965

        If lErro = 25162 Then gError 34966

        GridItens.TextMatrix(GridItens.Row, iGrid_Almoxarifado_Col) = objAlmoxarifado.sNomeReduzido

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 41044
        Else
            lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78658
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34967

    'Conta contabil ---> vem como PADRAO da  tabela EstoqueProduto se o Produto e o Almoxarifado estiverem Preenchidos
    lErro = Preenche_ContaContabil()
    If lErro <> SUCESSO Then gError 49704

    Saida_Celula_Almoxarifado = SUCESSO

    Exit Function

Erro_Saida_Celula_Almoxarifado:

    Saida_Celula_Almoxarifado = gErr

    Select Case gErr

        Case 34965

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE", Almoxarifado.Text)

            If vbMsg = vbYes Then

                objAlmoxarifado.sNomeReduzido = Almoxarifado.Text

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)

            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 34966

            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_ALMOXARIFADO_INEXISTENTE1", CInt(Almoxarifado.Text))

            If vbMsg = vbYes Then

                objAlmoxarifado.iCodigo = CInt(Almoxarifado.Text)

                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("Almoxarifado", objAlmoxarifado)


            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)

            End If

        Case 34963, 34964, 34967, 41044, 49704, 78658
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155961)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_CustoUnitario(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim sTipo As String

On Error GoTo Erro_Saida_Celula_CustoUnitario

    Set objGridInt.objControle = CustoUnitario

    If Len(Trim(CustoUnitario.ClipText)) > 0 Then

        'lErro = Valor_NaoNegativo_Critica(CustoUnitario.Text)
        lErro = Valor_Double_Critica(CustoUnitario.Text)
        If lErro <> SUCESSO Then gError 34968

    End If

    GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = CustoUnitario.FormattedText

    If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then

        'Calcula a Quantidade Disponível
        lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
        If lErro <> SUCESSO Then gError 49712

    Else

        'Calcula a Quantidade Disponível
        lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
        If lErro <> SUCESSO Then gError 78659

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34969

    Saida_Celula_CustoUnitario = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoUnitario:

    Saida_Celula_CustoUnitario = gErr

    Select Case gErr

        Case 34968, 34969, 49712, 78659
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 41002
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOUNITARIO_PREENCHIDO", gErr, GridItens.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155962)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Etiqueta(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Etiqueta

    Set objGridInt.objControle = Etiqueta

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 34997

    Saida_Celula_Etiqueta = SUCESSO

    Exit Function

Erro_Saida_Celula_Etiqueta:

    Saida_Celula_Etiqueta = gErr

    Select Case gErr

        Case 34968, 34969

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155963)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Tipo(objGridInt As AdmGrid) As Long

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
    If lErro <> SUCESSO Then gError 49619

    If iProdutoPreenchido = PRODUTO_PREENCHIDO And Len(Trim(sTipo)) > 0 Then

            objProduto.sCodigo = sProdutoFormatado

            'Lê o produto no BD
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 55459

            If lErro = 28030 Then gError 55460

            'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
            lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, sTipo)
            If lErro <> SUCESSO Then gError 55430

            GridItens.TextMatrix(GridItens.Row, iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then

            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 55427

        Else

            lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78660

        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 49622

    Saida_Celula_Tipo = SUCESSO

    Exit Function

Erro_Saida_Celula_Tipo:

    Saida_Celula_Tipo = gErr

    Select Case gErr

        Case 49619, 49620, 49622, 55427, 55430, 55459, 78660
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 55460
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155964)

    End Select

    Exit Function

End Function

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 34970

    Call Limpa_Tela_Inventario

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 34970

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155965)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExcluir_Click()

Dim objInventario As New ClassInventario
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 89942

    objInventario.sCodigo = Codigo.Text
    objInventario.iFilialEmpresa = giFilialEmpresa
    
    'trata a exclusão do movimento de estoque
    lErro = CF("Inventario_Exclui", objInventario, objContabil)
    If lErro <> SUCESSO Then gError 89943

    Call Limpa_Tela_Inventario

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 89942
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 89943
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155966)
    
    End Select
    
    Exit Sub

End Sub


Private Sub Limpa_Tela_Inventario()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Inventario

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Função genérica para Limpar a Tela
    Call Limpa_Tela(Me)

    'Limpa o Grid
    Call Grid_Limpa(objGrid)

    'Remove os ítens de colItensNumIntDoc
    Set colItensNumIntDoc = New Collection

    'Coloca a Data Atual na Tela
    Data.PromptInclude = False
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Data.PromptInclude = True
    
    Set gobjAnotacao = Nothing

    'limpeza da área relativa à contabilidade
    Call objContabil.Contabil_Limpa_Contabilidade

    iAlterado = 0
    sCodigoAntigo = ""

    Exit Sub

Erro_Limpa_Tela_Inventario:

    Select Case gErr

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155967)

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
Dim objItemMovEstoque As New ClassItemMovEstoque
Dim sTipo As String

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se o código foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then gError 34972

    'Verifica se a Data foi preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 34975

    'Verifica se há Algum Ítem de Movimento de Estoque Informado no GridItens
    If objGrid.iLinhasExistentes = 0 Then gError 34976

    objInventario.sCodigo = Codigo.Text

'    lErro = CF("Inventario_Le", objInventario)
'    If lErro <> SUCESSO And lErro <> 41011 Then gError 34973
'
'    'Se encontrar
'    If lErro = SUCESSO Then gError 34974

    'Para cada linha do Grid
    For iIndice = 1 To objGrid.iLinhasExistentes

        'Verifica se a Quantidade foi informada
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col))) = 0 Then gError 34977

        'Verifica se o Almoxarifado foi informado
        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col))) = 0 Then gError 34978

        If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_Tipo_Col))) = 0 Then gError 34979

        sTipo = GridItens.TextMatrix(iIndice, iGrid_Tipo_Col)
        If sTipo = STRING_QUANT_3_DEMO Or sTipo = STRING_QUANT_3_CONSERTO Or sTipo = STRING_QUANT_3_OUTRAS Or sTipo = STRING_QUANT_3_BENEF Or sTipo = STRING_QUANT_NOSSO_BENEF_CI Or sTipo = STRING_QUANT_NOSSO_CONSERTO_CI Or sTipo = STRING_QUANT_NOSSO_CONSIG_CI Or sTipo = STRING_QUANT_NOSSO_OUTRAS_CI Or sTipo = STRING_QUANT_NOSSO_DEMO_CI Or _
           sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI Or sTipo = STRING_QUANT_RECEB_INDISP_CI Or sTipo = STRING_QUANT_OUTRAS_INDISP_CI Or sTipo = STRING_QUANT_DEFEIT_CI Or sTipo = STRING_QUANT_3_CONSIG_CI Or sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI2P Then
            If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))) = 0 Then gError 34980
        End If

        'verifica se existe no inventario algum lançamento do mesmo produto, almoxarifado e tipo
        lErro = Verifica_Item_Repetido(iIndice)
        If lErro <> SUCESSO Then gError 55445

    Next

    lErro = Move_Tela_Memoria(objInventario)
    If lErro <> SUCESSO Then gError 34981

    'verifica se a data contábil é igual a data da tela ==> se não for, dá um aviso
    lErro = objContabil.Contabil_Testa_Data(CDate(Data.Text))
    If lErro <> SUCESSO Then gError 92048
    
    'Grava no BD os dados (inclusive os dados contábeis)
    lErro = CF("Inventario_Grava", objInventario, objContabil)
    If lErro <> SUCESSO Then gError 34982

    'gravar anotacao, se houver
    If Not (gobjAnotacao Is Nothing) Then
    
        If Len(Trim(gobjAnotacao.sTextoCompleto)) <> 0 Or Len(Trim(gobjAnotacao.sTitulo)) <> 0 Then
        
            gobjAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_INVENTARIO
            gobjAnotacao.sID = CStr(objInventario.iFilialEmpresa) & "," & objInventario.sCodigo
            gobjAnotacao.dtDataAlteracao = gdtDataHoje
            
            lErro = CF("Anotacoes_Grava", gobjAnotacao)
            If lErro <> SUCESSO Then gError 34982
            
        End If
        
    End If
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 34972
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 34973, 34981, 34982, 55445, 92048

'        Case 34974
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_INVENTARIO_CADASTRADO", gErr, objInventario.sCodigo)

        Case 34975
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_SEM_PREENCHIMENTO", gErr)

        Case 34976
            lErro = Rotina_Erro(vbOKOnly, "ERRO_INVENTARIO_NAO_INFORMADO", gErr)

        Case 34977
            lErro = Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDA", gErr, iIndice)

        Case 34978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO", gErr, iIndice)

        Case 34979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPO_NAO_PREENCHIDO", gErr, iIndice)

        Case 34980
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOUNITARIO_NAO_PREENCHIDO", gErr, iIndice)

        Case 41045
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ETIQUETA_NAO_PREENCHIDO", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155968)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objInventario As ClassInventario) As Long
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

    objInventario.iFilialEmpresa = giFilialEmpresa

    For iIndice = 1 To objGrid.iLinhasExistentes

        lErro = Move_Itens_Memoria(iIndice, objInventario)
        If lErro <> SUCESSO Then gError 34983

    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 34983

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155969)

    End Select

    Exit Function

End Function

Private Function Move_Itens_Memoria(iIndice As Integer, objInventario As ClassInventario) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim sTipo As String, iTipo As Integer
Dim dCusto As Double, sEtiqueta As String
Dim objItemInventario As New ClassItemInventario
Dim iIndice1 As Integer
Dim sContaFormatada As String
Dim iContaPreenchida As Integer
Dim sContaAjusteFormatada As String
Dim sLote As String
Dim iFilialOP As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Move_Itens_Memoria

     'Pega o Código do Almoxarifado
     objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col)

     lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
     If lErro <> SUCESSO And lErro <> 25056 Then gError 34984

     If lErro = 25056 Then gError 34985

     'Formata conta para o BD
     If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col))) > 0 Then

         'Guarda a conta contabil de estoque
         lErro = CF("Conta_Formata", GridItens.TextMatrix(iIndice, iGrid_ContaContabil_Col), sContaFormatada, iContaPreenchida)
         If lErro <> SUCESSO Then gError 52063
     Else
         sContaFormatada = ""
     End If

     'Formata conta Ajuste para o BD
     If Len(Trim(GridItens.TextMatrix(iIndice, iGrid_ContaAjuste_Col))) > 0 Then

         'Guarda a conta Ajuste de estoque
         lErro = CF("Conta_Formata", GridItens.TextMatrix(iIndice, iGrid_ContaAjuste_Col), sContaAjusteFormatada, iContaPreenchida)
         If lErro <> SUCESSO Then gError 52064
     Else
         sContaAjusteFormatada = ""
     End If

     'Formata o Produto para BD
     lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
     If lErro <> SUCESSO Then gError 34986

     sTipo = GridItens.TextMatrix(iIndice, iGrid_Tipo_Col)

     For iIndice1 = 0 To Tipo.ListCount - 1
         If Tipo.List(iIndice1) = sTipo Then
             iTipo = Tipo.ItemData(iIndice1)
             Exit For
         End If
     Next

     If sTipo = STRING_QUANT_3_DEMO Or sTipo = STRING_QUANT_3_CONSERTO Or sTipo = STRING_QUANT_3_OUTRAS Or sTipo = STRING_QUANT_3_BENEF Or sTipo = STRING_QUANT_NOSSO_BENEF_CI Or sTipo = STRING_QUANT_NOSSO_CONSERTO_CI Or sTipo = STRING_QUANT_NOSSO_CONSIG_CI Or sTipo = STRING_QUANT_NOSSO_OUTRAS_CI Or sTipo = STRING_QUANT_NOSSO_DEMO_CI Or _
        sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI Or sTipo = STRING_QUANT_RECEB_INDISP_CI Or sTipo = STRING_QUANT_OUTRAS_INDISP_CI Or sTipo = STRING_QUANT_DEFEIT_CI Or sTipo = STRING_QUANT_3_CONSIG_CI Or sTipo = STRING_QUANT_DISPONIVEL_NOSSA_CI2P Then
'        dCusto = CDbl(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))
        dCusto = Abs(CDbl(GridItens.TextMatrix(iIndice, iGrid_ValorDiferenca_col)))
        If StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col)) < 0 Then
            dCusto = -dCusto 'Se está com custo negativo o movimento deve acompanhar o sinal
        End If
     End If
     
     

     sEtiqueta = GridItens.TextMatrix(iIndice, iGrid_Etiqueta_Col)

     objProduto.sCodigo = sProdutoFormatado

     'Le o Produto para verificar que Tipo de Rastro ele possui
     lErro = CF("Produto_Le", objProduto)
     If lErro <> SUCESSO And lErro <> 28030 Then gError 78376

     'Se não encontrou o produto --> Erro
     If lErro = 28030 Then gError 78377

    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

         sLote = GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col)

     ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then

         sLote = GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col)
         iFilialOP = Codigo_Extrai(GridItens.TextMatrix(iIndice, iGrid_FilialOP_Col))

     End If
                     
     Set objItemInventario = objInventario.colItens.Add(0, sProdutoFormatado, GridItens.TextMatrix(iIndice, iGrid_DescricaoItem_Col), GridItens.TextMatrix(iIndice, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iIndice, iGrid_Quantidade_Col)), dCusto, objAlmoxarifado.iCodigo, GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col), iTipo, sEtiqueta, sContaFormatada, sContaAjusteFormatada, sLote, iFilialOP)

     objItemInventario.dQuantEst = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_QuantEstoque_Col))
     objItemInventario.dCustoUnitario = StrParaDbl(GridItens.TextMatrix(iIndice, iGrid_CustoUnitario_Col))

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

        Case 34984, 34986, 52063, 52064, 78376

        Case 34985
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_INEXISTENTE", gErr, objAlmoxarifado.sNomeReduzido)

        Case 78377
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 78378, 78379
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LOTE_GRID_NAO_PREENCHIDO", gErr, iIndice)

        Case 78380
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_OP_NAO_PREENCHIDA", gErr, iIndice)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155970)

    End Select

    Exit Function

End Function

Public Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 34986

    Call Limpa_Tela_Inventario

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 34986, 34987

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155971)

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

    Set colItensNumIntDoc = Nothing
    Set objEventoCodigo = Nothing
    Set objEventoAlmoxPadrao = Nothing
    Set objEventoProduto = Nothing
    Set objEventoEstoque = Nothing
    Set objEventoContaContabil = Nothing

    'eventos associados a contabilidade
    Set objEventoLote = Nothing
    Set objEventoDoc = Nothing

    Set objGrid = Nothing
    Set objGrid1 = Nothing
    Set objContabil = Nothing

    Set colItensNumIntDoc = Nothing
    
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

Public Sub DescricaoItem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Etiqueta_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Quantidade_Change()

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
Dim lNumIntDoc As Long

    If colItensNumIntDoc.Count >= GridItens.Row Then
        lNumIntDoc = colItensNumIntDoc.Item(GridItens.Row)
    Else
        lNumIntDoc = 0
    End If

    If lNumIntDoc = 0 Then

    'Verifica se a Tecla apertada foi Del
        If KeyCode = vbKeyDelete Then
            'Guarda iLinhasExistentes
            iLinhasExistentesAnterior = objGrid.iLinhasExistentes

            'Guarda o índice da Linha a ser Excluída
            iLinhaAnterior = GridItens.Row

        End If

        Call Grid_Trata_Tecla1(KeyCode, objGrid)

        'Verifica se a Linha foi realmente excluída
        If objGrid.iLinhasExistentes < iLinhasExistentesAnterior Then

            'Exclui de colItensNumIntDoc o Item correspondente, se houver
            colItensNumIntDoc.Remove iLinhaAnterior

        End If

    End If

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

        'tratamento de saida de celula da contabilidade
        lErro = objContabil.Contabil_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 39637

        If objGridInt.objGrid Is GridItens Then

            Select Case GridItens.Col

                Case iGrid_Produto_Col
                    lErro = Saida_Celula_Produto(objGridInt)
                    If lErro <> SUCESSO Then gError 34988

                Case iGrid_Quantidade_Col
                    lErro = Saida_Celula_Quantidade(objGridInt)
                    If lErro <> SUCESSO Then gError 34989

                Case iGrid_Almoxarifado_Col
                    lErro = Saida_Celula_Almoxarifado(objGridInt)
                    If lErro <> SUCESSO Then gError 34990

                Case iGrid_UnidadeMed_Col
                    lErro = Saida_Celula_UnidadeMed(objGridInt)
                    If lErro <> SUCESSO Then gError 34991

                Case iGrid_Etiqueta_Col
                    lErro = Saida_Celula_Etiqueta(objGridInt)
                    If lErro <> SUCESSO Then gError 34992

                Case iGrid_CustoUnitario_Col
                    lErro = Saida_Celula_CustoUnitario(objGridInt)
                    If lErro <> SUCESSO Then gError 34993

                Case iGrid_Tipo_Col
                    lErro = Saida_Celula_Tipo(objGridInt)
                    If lErro <> SUCESSO Then gError 34994

                Case iGrid_ContaContabil_Col
                    lErro = Saida_Celula_ContaContabil(objGridInt)
                    If lErro <> SUCESSO Then gError 49583

                Case iGrid_ContaAjuste_Col
                    lErro = Saida_Celula_ContaAjuste(objGridInt)
                    If lErro <> SUCESSO Then gError 52062

                Case iGrid_LoteProduto_Col
                    lErro = Saida_Celula_LoteProduto(objGridInt)
                    If lErro <> SUCESSO Then gError 78366

                Case iGrid_FilialOP_Col
                    lErro = Saida_Celula_FilialOP(objGridInt)
                    If lErro <> SUCESSO Then gError 78367
                    
                Case iGrid_Atualiza_Col
                    lErro = Saida_Celula_Atualiza(objGridInt)
                    If lErro <> SUCESSO Then gError 90831

            End Select

            lErro = Grid_Finaliza_Saida_Celula(objGridInt)
            If lErro Then gError 34996

        End If

        iAlterado = REGISTRO_ALTERADO

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 34988, 34989, 34990, 34991, 34992, 34993, 34994, 34995, 78366, 78367, 90831

        Case 34996
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 39637, 49583, 52062

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155972)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOP(objGridInt As AdmGrid) As Long
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
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 78587

            'Se não encontrou o ítem com o código informado
            If lErro = 6730 Then

                objFilialOP.iCodFilial = iCodigo

                'Pesquisa se existe FilialOP com o codigo extraido
                lErro = CF("FilialEmpresa_Le", objFilialOP)
                If lErro <> SUCESSO And lErro <> 27378 Then gError 78588

                'Se não encontrou a FilialOP
                If lErro = 27378 Then gError 78589

                'coloca na tela
                FilialOP.Text = iCodigo & SEPARADOR & objFilialOP.sNome

            End If

            'Não encontrou valor informado que era STRING
            If lErro = 6731 Then gError 78590

        End If

        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) > 0 Then

            lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 78594

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

                objRastroLote.sCodigo = GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col)
                objRastroLote.sProduto = sProdutoFormatado
                objRastroLote.iFilialOP = Codigo_Extrai(FilialOP.Text)

                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78595

                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78596

                'Preenche a Quantidade do Lote
                lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col), Codigo_Extrai(FilialOP.Text))
                If lErro <> SUCESSO Then gError 78670
            
            End If

        End If
    
    Else
    
        'se o lote também não está preenchido
        If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_LoteProduto_Col))) = 0 Then
    
            'Preenche a Quantidade do produto em estoque
            lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
            If lErro <> SUCESSO Then gError 78671
            
        Else
        
            GridItens.TextMatrix(GridItens.Row, iGrid_QuantEstoque_Col) = ""
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78597

    Saida_Celula_FilialOP = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOP:

    Saida_Celula_FilialOP = gErr

    Select Case gErr

        Case 78587, 78588, 78591, 78594, 78595, 78597, 78671
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78589
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78590
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, FilialOP.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78592
            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78593
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78596
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155973)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Atualiza(objGridInt As AdmGrid) As Long
'Faz a saida de celula de Atualiza

Dim lErro As Long
Dim iAtualiza As Integer

On Error GoTo Erro_Saida_Celula_Atualiza

    Set objGridInt.objControle = Atualiza
    
    GridItens.TextMatrix(GridItens.Row, iGrid_Atualiza_Col) = Atualiza.Text

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 90829

    Saida_Celula_Atualiza = SUCESSO

    Exit Function

Erro_Saida_Celula_Atualiza:

    Saida_Celula_Atualiza = gErr

    Select Case gErr

        Case 90829
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155974)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_LoteProduto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_LoteProduto

    Set objGridInt.objControle = LoteProduto

    If Len(Trim(LoteProduto.Text)) > 0 Then

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 78598

        objProduto.sCodigo = sProdutoFormatado

        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 78599

        If lErro = 28030 Then gError 78600

        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then

            'Se for rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Or objProduto.iRastro = PRODUTO_RASTRO_NUM_SERIE Then

                objRastroLote.sCodigo = LoteProduto.Text
                objRastroLote.sProduto = sProdutoFormatado

                'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 78601

                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 78602

            'Se for rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then

                If Len(Trim(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))) > 0 Then

                    objRastroLote.sCodigo = LoteProduto.Text
                    objRastroLote.sProduto = sProdutoFormatado
                    objRastroLote.iFilialOP = Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col))

                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 78606

                    'Se não encontrou --> Erro
                    If lErro = 75710 Then gError 78607

                End If


            End If
            
            'Preenche a Quantidade do Lote
            lErro = QuantLote_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row, LoteProduto.Text, Codigo_Extrai(GridItens.TextMatrix(GridItens.Row, iGrid_FilialOP_Col)))
            If lErro <> SUCESSO Then gError 78661
            
            For iIndice = 0 To Atualiza.ListCount - 1
                If Atualiza.ItemData(iIndice) = INVENTARIO_ATUALIZA_LOTE_ESTOQUE Then
                    Atualiza.ListIndex = iIndice
                    Exit For
                End If
            Next
            
            GridItens.TextMatrix(GridItens.Row, iGrid_Atualiza_Col) = Atualiza.Text
            
        End If

    Else
        'Preenche a Quantidade do Lote
        lErro = QuantEstoque_Calcula(GridItens.TextMatrix(GridItens.Row, iGrid_Produto_Col), GridItens.Row)
        If lErro <> SUCESSO Then gError 78651

        Atualiza.ListIndex = -1
        GridItens.TextMatrix(GridItens.Row, iGrid_Atualiza_Col) = Atualiza.Text

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 78608

    Saida_Celula_LoteProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_LoteProduto:

    Saida_Celula_LoteProduto = gErr

    Select Case gErr

        Case 78598, 78599, 78601, 78556, 78558, 78606, 78608, 78650, 78651, 78661 '78603
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78600
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 78602, 78607
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

'        Case 78604
'            lErro = Rotina_Erro(vbYesNo, "ERRO_OPCODIGO_NAO_CADASTRADO", gErr, objOrdemProducao.sCodigo)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
'
'        Case 78605
'            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, objOrdemProducao.sCodigo)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 155975)

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

            objGrid.iExecutaRotinaEnable = GRID_NAO_EXECUTAR_ROTINA_ENABLE

            Call Grid_Entrada_Celula(objGrid, iAlterado)

            objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

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

Public Sub DescricaoItem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Public Sub DescricaoItem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Public Sub DescricaoItem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = DescricaoItem
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

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

Public Sub CTBTvwContas_Expand(ByVal Node As MSComctlLib.Node)

    Call objContabil.Contabil_TvwContas_Expand(Node, CTBTvwContas.Nodes)

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
Dim dQuantidadeConvertida As Double
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim dQuantidade As Double
Dim sTipoEstoque As String
Dim sProduto As String
Dim sUM As String

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case QUANT_ESTOQUE
            For iLinha = 1 To objGrid.iLinhasExistentes

                If Len(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)) > 0 Then

                    lErro = CF("UMEstoque_Conversao", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), CDbl(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)), dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 64205

                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case QUANT_DISPONIVEL, QUANT_CONSIGTERC, QUANT_DEMONSTERC, QUANT_CONSERTOTERC, QUANT_OUTRASTERC, QUANT_BENEFTERC

            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)) > 0 Then

                    objAlmoxarifado.sNomeReduzido = GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                    lErro = CF("Almoxarifado_Le_NomeReduzido", objAlmoxarifado)
                    If lErro <> SUCESSO And lErro <> 25060 Then gError 75441

                    'Se não encontrou o almoxarifado, erro
                    If lErro = 25060 Then gError 75442

                    lErro = CF("Produto_Formata", GridItens.TextMatrix(iLinha, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
                    If lErro <> SUCESSO Then gError 75443

                    'Lê as quantidades do Produto No almoxarifado de origem
                    objEstoqueProduto.iAlmoxarifado = objAlmoxarifado.iCodigo
                    objEstoqueProduto.sProduto = sProdutoFormatado
                    lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
                    If lErro <> SUCESSO And lErro <> 21306 Then gError 75444

                    'Se não encontrou Estoque Produto, erro
                    If lErro = 21306 Then gError 75445

                    If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then

                        If objMnemonicoValor.sMnemonico = QUANT_DISPONIVEL Then
                            dQuantidade = objEstoqueProduto.dQuantDisponivel
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_CONSIGTERC Then
                            dQuantidade = objEstoqueProduto.dQuantConsig3
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_DEMONSTERC Then
                            dQuantidade = objEstoqueProduto.dQuantDemo3
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_CONSERTOTERC Then
                            dQuantidade = objEstoqueProduto.dQuantConserto3
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_OUTRASTERC Then
                            dQuantidade = objEstoqueProduto.dQuantOutras3
                        ElseIf objMnemonicoValor.sMnemonico = QUANT_BENEFTERC Then
                            dQuantidade = objEstoqueProduto.dQuantBenef3
                        End If

                        lErro = CF("UMEstoque_Conversao", sProdutoFormatado, GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col), dQuantidade, dQuantidadeConvertida)
                        If lErro <> SUCESSO Then gError 75446

                        objMnemonicoValor.colValor.Add dQuantidadeConvertida
                    Else
                        objMnemonicoValor.colValor.Add 0
                    End If
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case CODIGO1
            If Len(Codigo.Text) > 0 Then
                objMnemonicoValor.colValor.Add Codigo.Text
            Else
                objMnemonicoValor.colValor.Add ""
            End If

        Case DATA1
            If Len(Data.ClipText) > 0 Then
                objMnemonicoValor.colValor.Add CDate(Data.FormattedText)
            Else
                objMnemonicoValor.colValor.Add DATA_NULA
            End If

        Case PRODUTO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case UNIDADE_MED
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case QUANTIDADE1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col)) > 0 Then
                    
                    'Guarda os valores que serão passados como parâmetros em UMEstoque_Conversao
                    sProduto = GridItens.TextMatrix(iLinha, iGrid_Produto_Col)
                    sUM = GridItens.TextMatrix(iLinha, iGrid_UnidadeMed_Col)
                    dQuantidade = StrParaDbl(GridItens.TextMatrix(iLinha, iGrid_Quantidade_Col))
                    
                    'Converte a quantidade para UM padrão estoque
                    lErro = CF("UMEstoque_Conversao", sProduto, sUM, dQuantidade, dQuantidadeConvertida)
                    If lErro <> SUCESSO Then gError 79943
                    
                    objMnemonicoValor.colValor.Add dQuantidadeConvertida
                                                            
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case QUANT_DIFERENCA
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_QuantDiferenca_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case DESCRICAO_ITEM
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_DescricaoItem_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_DescricaoItem_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case ALMOXARIFADO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Almoxarifado_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case TIPO1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Tipo_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Tipo_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case CUSTO_UNITARIO
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add CDbl(GridItens.TextMatrix(iLinha, iGrid_CustoUnitario_Col))
                Else
                    objMnemonicoValor.colValor.Add 0
                End If
            Next

        Case ETIQUETA1
            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Etiqueta_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_Etiqueta_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case CONTACONTABIL1

            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_ContaContabil_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_ContaContabil_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case CTAAJUSTEINV

            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_ContaAjuste_Col)) > 0 Then
                    objMnemonicoValor.colValor.Add GridItens.TextMatrix(iLinha, iGrid_ContaAjuste_Col)
                Else
                    objMnemonicoValor.colValor.Add ""
                End If
            Next

        Case ESCANINHO_CUSTO

            For iLinha = 1 To objGrid.iLinhasExistentes
                If Len(GridItens.TextMatrix(iLinha, iGrid_Produto_Col)) > 0 Then

                    Select Case GridItens.TextMatrix(iLinha, iGrid_Tipo_Col)

                        Case STRING_QUANT_3_CONSIG, STRING_QUANT_3_CONSIG_CI
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSIGNACAO

                        Case STRING_QUANT_3_DEMO
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_DEMO

                        Case STRING_QUANT_3_CONSERTO
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_CONSERTO

                        Case STRING_QUANT_3_OUTRAS
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_OUTROS

                        Case STRING_QUANT_3_BENEF
                            objMnemonicoValor.colValor.Add ESCANINHO_3_EM_BENEF

                        Case STRING_QUANT_NOSSO_CONSIG, STRING_QUANT_NOSSO_CONSIG_CI
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_CONSIGNACAO

                        Case STRING_QUANT_NOSSO_DEMO, STRING_QUANT_NOSSO_DEMO_CI
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_DEMO

                        Case STRING_QUANT_NOSSO_CONSERTO, STRING_QUANT_NOSSO_CONSERTO_CI
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_CONSERTO

                        Case STRING_QUANT_NOSSO_OUTRAS, STRING_QUANT_NOSSO_OUTRAS_CI
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_OUTROS

                        Case STRING_QUANT_NOSSO_BENEF, STRING_QUANT_NOSSO_BENEF_CI
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO_EM_BENEF
                            
                        Case Else
                            objMnemonicoValor.colValor.Add ESCANINHO_NOSSO

                    End Select

                Else
                    objMnemonicoValor.colValor.Add 0

                End If
            Next

        Case Else
            gError 39654

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr

    Select Case gErr

        Case 39654
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case 64205, 75441, 75443, 75444, 75446

        Case 75442
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO1", gErr, objAlmoxarifado.sNomeReduzido)

        Case 75445
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEPRODUTO_NAO_CADASTRADO", gErr, objEstoqueProduto.sProduto, objEstoqueProduto.iAlmoxarifado)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155976)

    End Select

    Exit Function

End Function

Private Sub Limpa_colItensNumIntDoc(colItensNumIntDoc As Collection)

Dim lErro As Long
Dim iCount As Integer
Dim iIndice As Integer

On Error GoTo Erro_Limpa_colItensNumIntDoc

    iCount = colItensNumIntDoc.Count
    Set colItensNumIntDoc = New Collection

    For iIndice = 1 To iCount

        colItensNumIntDoc.Add 0

    Next

    Exit Sub

Erro_Limpa_colItensNumIntDoc:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 155977)

    End Select

    Exit Sub

End Sub

Private Function Verifica_Item_Repetido(iIndice) As Long
'verifica se existem itens dentro do grid com o mesmo produto, almoxarifado e tipo do item iIndice

Dim iIndice1 As Integer
Dim sProdutoFormatado As String
Dim sProdutoFormatado1 As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long

On Error GoTo Erro_Verifica_Item_Repetido

    lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice, iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 55446

    For iIndice1 = 1 To objGrid.iLinhasExistentes

        If iIndice <> iIndice1 Then

            lErro = CF("Produto_Formata", GridItens.TextMatrix(iIndice1, iGrid_Produto_Col), sProdutoFormatado1, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 55447

            If sProdutoFormatado = sProdutoFormatado1 And GridItens.TextMatrix(iIndice, iGrid_Almoxarifado_Col) = GridItens.TextMatrix(iIndice1, iGrid_Almoxarifado_Col) And GridItens.TextMatrix(iIndice, iGrid_Tipo_Col) = GridItens.TextMatrix(iIndice1, iGrid_Tipo_Col) And GridItens.TextMatrix(iIndice, iGrid_LoteProduto_Col) = GridItens.TextMatrix(iIndice1, iGrid_LoteProduto_Col) Then gError 55448

        End If

    Next

    Verifica_Item_Repetido = SUCESSO

    Exit Function

Erro_Verifica_Item_Repetido:

    Verifica_Item_Repetido = gErr

    Select Case gErr

        Case 55446, 55447

        Case 55448
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEM_INVENTARIO_REPETIDO", gErr, iIndice, iIndice1)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155978)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_INVENTARIO_LANCAMENTOS
    Set Form_Load_Ocx = Me
    Caption = "Inventário"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "Inventario"

End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

Public Sub Unload(objme As Object)

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

Dim lErro As Long

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is AlmoxPadrao Then
            Call AlmoxPadraoLabel_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call BotaoProdutos_Click
        ElseIf Me.ActiveControl Is ContaContabil Or Me.ActiveControl = ContaAjuste Then
            Call BotaoPlanoConta_Click
        ElseIf Me.ActiveControl Is Almoxarifado Then
            Call BotaoEstoque_Click
        End If
    
    ElseIf KeyCode = KEYCODE_CODBARRAS Then
        lErro = CF2(Me, "INV_Trata_CodigoBarras1")
        
        
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

'###################################################
'Inserido por Wagner
Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim sCodigo As String

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo lote disponível
    lErro = CF("Inventario_Automatico", sCodigo, giFilialEmpresa)
    If lErro <> SUCESSO Then gError 185934

    Codigo.Text = sCodigo

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 185934
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 185935)
    
    End Select

    Exit Sub

End Sub
'###################################################


Public Function INV_Trata_CodigoBarras1(ByVal objCT As Object) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoEnxuto As String
Dim sCodBarras As String
Dim sCodBarrasOriginal As String
Dim dCusto As Double

On Error GoTo Erro_INV_Trata_CodigoBarras1

    If objCT.objGrid.iLinhasExistentes + 1 = objCT.GridItens.Row Then
    
        'Verifica se o Produto está preenchido
        If Len(Trim(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col))) = 0 Then
            
            If objCT.ActiveControl Is objCT.Produto Then
                    
                    Set objCT.objGrid.objControle = objCT.Produto
            
                    lErro = Grid_Abandona_Celula(objCT.objGrid)
                    If lErro <> SUCESSO Then gError 199540
                    
            End If
            
            objProduto.lErro = 1
    
            Call Chama_Tela_Modal("CodigoBarras", objProduto)
    
            
            If objProduto.sCodigoBarras <> "Cancel" Then
                If objProduto.lErro = SUCESSO Then
    
                    lErro = CF("INV_Trata_CodigoBarras", objProduto)
                    If lErro <> SUCESSO Then gError 199366
    
                End If
    
                'Lê os demais atributos do Produto
                lErro = CF("Produto_Le", objProduto)
                If lErro <> SUCESSO And lErro <> 28030 Then gError 199305
    
                'Se não encontrou o Produto --> Erro
                If lErro = 28030 Then gError 199306
    
                lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
                If lErro <> SUCESSO Then gError 199307
        
                objCT.Show
        
                objCT.Produto.PromptInclude = False
                objCT.Produto.Text = sProdutoEnxuto
                objCT.Produto.PromptInclude = True
                
                objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col) = objCT.Produto.Text
                
                gError 210864
                
                
'                If Not objCT.ActiveControl Is objCT.Produto Then
'                    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col) = objCT.Produto.Text
'
'                    'Preenche a Linha do Grid
'                    lErro = objCT.ProdutoLinha_Preenche(objProduto)
'                    If lErro <> SUCESSO Then gError 199305
'
'                    'calcula o custo atual do produto passado como parametro e devolve-o de dCusto
'                    lErro = CF("Calcula_Custo_Atual", objProduto, dCusto, objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Tipo_Col))
'                    If lErro <> SUCESSO Then gError 199305
'
'                    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_CustoUnitario_Col) = Formata_Custo(dCusto)
'
'                    If Len(Trim(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_LoteProduto_Col))) = 0 Then
'
'                        lErro = objCT.QuantEstoque_Calcula(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col), objCT.GridItens.Row)
'                        If lErro <> SUCESSO Then gError 199305
'
'                    Else
'
'                        lErro = objCT.QuantLote_Calcula(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col), objCT.GridItens.Row, objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_LoteProduto_Col), Codigo_Extrai(objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_FilialOP_Col)))
'                        If lErro <> SUCESSO Then gError 199305
'
'                    End If
'
'                    lErro = objCT.Preenche_ContaContabil()
'                    If lErro <> SUCESSO Then gError 52240
                
'
'                objCT.GridItens.SetFocus
'                objCT.GridItens.FocusRect = flexFocusHeavy
    
    
            Else
            
                gError 199834
    
    
            End If
    
        End If
    
    End If

    INV_Trata_CodigoBarras1 = SUCESSO

    Exit Function

Erro_INV_Trata_CodigoBarras1:

    INV_Trata_CodigoBarras1 = gErr

'    objCT.GridItens.TextMatrix(objCT.GridItens.Row, objCT.iGrid_Produto_Col) = ""

    Select Case gErr

        Case 199305, 199308, 199363, 199366, 199540, 199834, 210864

        Case 199306, 199364
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case 199307
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 199309)

    End Select

    Exit Function

End Function

Public Sub Anotacao_Extrai(ByVal objAnotacao As ClassAnotacoes)

Dim lErro As Long

On Error GoTo Erro_Anotacao_Extrai

    objAnotacao.iTipoDocOrigem = ANOTACAO_ORIGEM_INVENTARIO
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


